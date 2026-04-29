import { useState, useEffect, useRef, useCallback, Fragment } from "react";
import * as XLSX from "xlsx";
import { PDFDocument } from "pdf-lib";
import { supabase, isSupabaseConfigured } from "./supabase";

const SK_API_KEY = "servial-anthropic-key";
const SK_GEMINI_KEY = "servial-gemini-key";
const SK_AI_PROVIDER = "servial-ai-provider";
const getApiKey = () => localStorage.getItem(SK_API_KEY) || "";
const setApiKey = (k) => localStorage.setItem(SK_API_KEY, k);
const getGeminiKey = () => localStorage.getItem(SK_GEMINI_KEY) || "";
const setGeminiKey = (k) => localStorage.setItem(SK_GEMINI_KEY, k);
const getAiProvider = () => localStorage.getItem(SK_AI_PROVIDER) || "gemini";
const setAiProviderLS = (p) => localStorage.setItem(SK_AI_PROVIDER, p);
const apiHeaders = () => ({
  "Content-Type": "application/json",
  "x-api-key": getApiKey(),
  "anthropic-version": "2023-06-01",
  "anthropic-dangerous-direct-browser-access": "true",
});

/* ── Función genérica callAI ── */
const callAI = async (system, userContent, maxTokens = 4096, provider, _retry = 0) => {
  const MAX_RETRIES = 3;
  try {
    if (provider === "claude") {
      const key = getApiKey();
      if (!key) throw new Error("Cal configurar la API Key d'Anthropic (Claude).");
      const messages = [{role: "user", content: userContent}];
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: apiHeaders(),
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", system, messages, max_tokens: maxTokens }),
      });
      if ((res.status === 429 || res.status === 529) && _retry < MAX_RETRIES) {
        const wait = res.status === 429 ? 65000 : 15000;
        console.log(`Claude ${res.status}, reintentant en ${wait/1000}s (intent ${_retry+1}/${MAX_RETRIES})…`);
        await new Promise(r => setTimeout(r, wait));
        return callAI(system, userContent, maxTokens, provider, _retry + 1);
      }
      if (!res.ok) { const e = await res.text(); throw new Error(`Claude API HTTP ${res.status}: ${e}`); }
      const data = await res.json();
      return data.content.filter(b => b.type === "text").map(b => b.text).join("");
    } else {
      /* Gemini */
      const key = getGeminiKey();
      if (!key) throw new Error("Cal configurar la API Key de Google Gemini.");
      const parts = Array.isArray(userContent)
        ? userContent.map(b => {
            if (b.type === "text") return { text: b.text };
            if (b.type === "document" && b.source?.type === "base64") return { inline_data: { mime_type: b.source.media_type || "application/pdf", data: b.source.data } };
            return { text: JSON.stringify(b) };
          })
        : [{ text: typeof userContent === "string" ? userContent : JSON.stringify(userContent) }];
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${key}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            system_instruction: { parts: [{ text: system }] },
            contents: [{ parts }],
            generationConfig: { maxOutputTokens: maxTokens, temperature: 0.2 },
          }),
        }
      );
      if (res.status === 429 && _retry < MAX_RETRIES) {
        const wait = 65000;
        console.log(`Gemini 429, reintentant en ${wait/1000}s (intent ${_retry+1}/${MAX_RETRIES})…`);
        await new Promise(r => setTimeout(r, wait));
        return callAI(system, userContent, maxTokens, provider, _retry + 1);
      }
      if (!res.ok) { const e = await res.text(); throw new Error(`Gemini API HTTP ${res.status}: ${e}`); }
      const data = await res.json();
      const cand = data.candidates?.[0];
      if (!cand) throw new Error("Gemini no ha retornat cap resposta.");
      return cand.content?.parts?.map(p => p.text || "").join("") || "";
    }
  } catch(err) {
    if ((err.message.includes("429") || err.message.includes("529")) && _retry < MAX_RETRIES) {
      console.log(`Error ${err.message}, reintentant en 65s (intent ${_retry+1}/${MAX_RETRIES})…`);
      await new Promise(r => setTimeout(r, 65000));
      return callAI(system, userContent, maxTokens, provider, _retry + 1);
    }
    throw err;
  }
};

/* ── System prompt per Consultes Legals ── */
const SYSTEM_LEGAL = `Ets un assessor jurídic expert en contractació pública espanyola i catalana, especialitzat en obres d'edificació.
Respon sempre en català. Cita els articles concrets de les lleis quan sigui rellevant. Si la pregunta fa referència a un plec concret, contextualitza la resposta amb les dades del plec.

═══════════════════════════════════════════════════════════════
MARC LEGAL COMPLET — NORMATIVA DE REFERÈNCIA
═══════════════════════════════════════════════════════════════

1. LLEI PRINCIPAL DE CONTRACTACIÓ PÚBLICA
   • Llei 9/2017, de 8 de novembre, de Contractes del Sector Públic (LCSP) — Transposició de les Directives UE 2014/23/UE i 2014/24/UE

2. REGLAMENTS DE CONTRACTACIÓ
   • Reial Decret 1098/2001, de 12 d'octubre — Reglament general de la Llei de contractes (RGLCAP), arts. 144, 155, 163, 164
   • Reial Decret 817/2009, de 8 de maig — Desenvolupa parcialment la Llei 30/2007
   • Reial Decret 773/2015, de 28 d'agost — Modifica preceptes del RGLCAP

3. CLASSIFICACIÓ I SOLVÈNCIA (LCSP arts. 65-85 + RD 1098/2001)
   • Classificació obligatòria en obres ≥ 500.000€ (art. 77 LCSP)
   • La classificació SUBSTITUEIX la solvència (econòmica + tècnica) en els subgrups que comprèn
   • Categories per anualitat mitjana: Cat.1 ≤150k | Cat.2 150-360k | Cat.3 360-840k | Cat.4 840k-2,4M | Cat.5 2,4-5M | Cat.6 >5M
   • Grups d'obres: A(Mov. terres) B(Ponts) C(Edificació) D(Ferrocarrils) E(Hidràuliques) F(Marítimes) G(Vials) H(Transports/instal.) I(Elèctriques) J(Mecàniques) K(Especials)

4. NORMATIVA D'ÀMBIT LOCAL I METROPOLITÀ
   • Llei 31/2010, de 3 d'agost, de l'Àrea Metropolitana de Barcelona (LAMB) — art. 14.F infraestructures d'interès metropolità
   • Text refós de la Llei d'Hisendes Locals — art. 174 despeses plurianuals
   • Llei 47/2003, de 26 de novembre, General Pressupostària

5. TRANSPARÈNCIA I BON GOVERN
   • Llei 19/2014, de 29 de desembre, de transparència, accés a la informació pública i bon govern de Catalunya

6. PROTECCIÓ DE DADES
   • Llei Orgànica 3/2018, de 5 de desembre (LOPD-GDD)
   • Reglament (UE) 2016/679 (RGPD)

7. PREVENCIÓ DE RISCOS LABORALS I SEGURETAT
   • Llei 31/1995, de 8 de novembre, de Prevenció de Riscos Laborals
   • Reial Decret 1627/1997, de 24 d'octubre — Disposicions mínimes de seguretat i salut en obres de construcció
   • Reial Decret 171/2004, de 30 de gener — Coordinació d'activitats empresarials

8. GESTIÓ DE RESIDUS
   • Reial Decret 105/2008, d'1 de febrer — Producció i gestió de residus de construcció i demolició

9. SUBCONTRACTACIÓ
   • Llei 32/2006, de 18 d'octubre — Reguladora de la subcontractació en el sector de la construcció

10. NORMATIVA FISCAL I TRIBUTÀRIA
    • Llei 58/2003, de 17 de desembre, General Tributària — art. 43.1.f responsabilitat subsidiària
    • Reial Decret 1619/2012, de 30 de novembre — Reglament d'obligacions de facturació
    • Llei 3/2004, de 29 de desembre — Mesures de lluita contra la morositat en operacions comercials

11. REVISIÓ DE PREUS
    • Reial Decret 55/2017, de 3 de febrer — Desenvolupa la Llei 2/2015 de desindexació de l'economia espanyola
    • LCSP arts. 103-104 — Revisió de preus en contractes públics

12. SUPORT ALS EMPRENEDORS
    • Llei 14/2013, de 27 de setembre — Suport als emprenedors i la seva internacionalització

13. SIGNATURA ELECTRÒNICA I CONTRACTACIÓ ELECTRÒNICA
    • Reglament (UE) 910/2014 (eIDAS) — Identificació electrònica i serveis de confiança
    • Decret Llei 3/2016 de la Generalitat de Catalunya — Disposició addicional primera
    • Decret 107/2005, de 31 de maig — Regulador del RELIC

14. PROCEDIMENT ADMINISTRATIU
    • Llei 39/2015, d'1 d'octubre — Procediment administratiu comú de les administracions públiques
    • Llei 29/1998, de 13 de juliol — Reguladora de la jurisdicció contenciosa administrativa

15. DIRECTIVES EUROPEES
    • Directiva 2014/23/UE — Adjudicació de contractes de concessió
    • Directiva 2014/24/UE — Contractació pública (transposada per la LCSP)

═══════════════════════════════════════════════════════════════

INSTRUCCIONS:
- Respon amb rigor jurídic però de forma comprensible per a un professional de la construcció
- Cita sempre l'article concret de la llei aplicable
- Si hi ha jurisprudència rellevant (TACRC, tribunals administratius), menciona-la
- Si la pregunta és ambigua, demana aclariment
- Si no tens certesa, indica-ho clarament
- Estructura la resposta amb capçaleres i punts per facilitar la lectura`;

const PROVINCIES = [
  { nom:"Barcelona", comarques:["Alt Penedès","Anoia","Bages","Baix Llobregat","Barcelonès","Berguedà","Garraf","Maresme","Moianès","Osona","Vallès Occidental","Vallès Oriental"]},
  { nom:"Girona", comarques:["Alt Empordà","Baix Empordà","Cerdanya","Garrotxa","Gironès","Pla de l'Estany","Ripollès","Selva"]},
  { nom:"Lleida", comarques:["Alt Urgell","Alta Ribagorça","Garrigues","Noguera","Pallars Jussà","Pallars Sobirà","Pla d'Urgell","Segarra","Segrià","Solsonès","Urgell","Val d'Aran"]},
  { nom:"Tarragona", comarques:["Alt Camp","Baix Camp","Baix Ebre","Baix Penedès","Conca de Barberà","Montsià","Priorat","Ribera d'Ebre","Tarragonès","Terra Alta"]},
];
const COMARQUES = PROVINCIES.flatMap(p => p.comarques);
// Mapa comarca → municipis principals (per poder filtrar lloc_execucio que conté municipis, no comarques)
const COMARCA_MUNICIPIS={
  "Alt Penedès":["Vilafranca del Penedès","Sant Sadurní d'Anoia","Gelida","Sant Pere de Riudebitlles","Olèrdola","Santa Margarida i els Monjos","Subirats","Torrelavit","Font-rubí","Avinyonet del Penedès","Sant Martí Sarroca","Sant Quintí de Mediona","Mediona","La Granada","Sant Cugat Sesgarrigues","Les Cabanyes","Sant Llorenç d'Hortons","Pacs del Penedès","Castellví de la Marca","Castellet i la Gornal"],
  "Anoia":["Igualada","Vilanova del Camí","Santa Margarida de Montbui","Piera","Masquefa","La Pobla de Claramunt","Capellades","Calaf","Òdena","Els Hostalets de Pierola","Jorba","El Bruc","Carme","Castellolí","Sant Martí de Tous","Copons","Montmaneu","La Torre de Claramunt","Rubió","Veciana"],
  "Bages":["Manresa","Sant Joan de Vilatorrada","Sallent","Santpedor","Navàs","Sant Fruitós de Bages","Súria","Cardona","Artés","Moià","Callús","Fonollosa","Avinyó","Castellbell i el Vilar","Castellgalí","Rajadell","Sant Salvador de Guardiola","El Pont de Vilomara i Rocafort","Monistrol de Montserrat","Balsareny"],
  "Baix Llobregat":["L'Hospitalet de Llobregat","Cornellà de Llobregat","El Prat de Llobregat","Sant Boi de Llobregat","Viladecans","Gavà","Esplugues de Llobregat","Sant Joan Despí","Sant Feliu de Llobregat","Molins de Rei","Sant Vicenç dels Horts","Castelldefels","Martorell","Sant Andreu de la Barca","Pallejà","Sant Just Desvern","Abrera","Olesa de Montserrat","Cervelló","Torrelles de Llobregat","Begues","Vallirana","Corbera de Llobregat","Santa Coloma de Cervelló","Collbató","Castellví de Rosanes","Sant Climent de Llobregat","La Palma de Cervelló","Sant Esteve Sesrovires"],
  "Barcelonès":["Barcelona","L'Hospitalet de Llobregat","Badalona","Santa Coloma de Gramenet","Sant Adrià de Besòs"],
  "Berguedà":["Berga","Gironella","Puig-reig","Avià","Casserres","Cercs","La Pobla de Lillet","Bagà","Olvan","Guardiola de Berguedà"],
  "Garraf":["Vilanova i la Geltrú","Sitges","Sant Pere de Ribes","Cubelles","Canyelles","Olivella"],
  "Maresme":["Mataró","Premià de Mar","El Masnou","Vilassar de Mar","Pineda de Mar","Arenys de Mar","Calella","Malgrat de Mar","Canet de Mar","Argentona","Sant Vicenç de Montalt","Premià de Dalt","Montgat","Alella","Vilassar de Dalt","Cabrera de Mar","Tiana","Sant Andreu de Llavaneres","Teià","Tordera","Santa Susanna","Palafolls","Sant Pol de Mar","Sant Cebrià de Vallalta","Caldes d'Estrac","Dosrius","Arenys de Munt","Sant Iscle de Vallalta","Òrrius","Cabrera de Mar"],
  "Moianès":["Moià","Calders","Santa Maria d'Oló","Castellterçol","L'Estany","Collsuspina","Castellcir","Granera","Monistrol de Calders","Sant Quirze Safaja"],
  "Osona":["Vic","Manlleu","Torelló","Centelles","Tona","Sant Pere de Torelló","Roda de Ter","Gurb","Sant Hipòlit de Voltregà","Tavertet","Les Masies de Voltregà","Santa Eugènia de Berga","Seva","Balenyà","Sant Julià de Vilatorta","Folgueroles","Lluçà","Prats de Lluçanès","Olost","Sant Bartomeu del Grau"],
  "Vallès Occidental":["Terrassa","Sabadell","Rubí","Cerdanyola del Vallès","Montcada i Reixac","Sant Cugat del Vallès","Ripollet","Barberà del Vallès","Badia del Vallès","Castellar del Vallès","Sant Quirze del Vallès","Polinyà","Sentmenat","Santa Perpètua de Mogoda","Palau-solità i Plegamans","Viladecavalls","Matadepera","Ullastrell","Vacarisses","Rellinars","Gallifa","Sant Llorenç Savall"],
  "Vallès Oriental":["Granollers","Mollet del Vallès","Sant Celoni","Parets del Vallès","Les Franqueses del Vallès","La Garriga","Cardedeu","La Llagosta","Caldes de Montbui","La Roca del Vallès","Llinars del Vallès","Sant Fost de Campsentelles","Montmeló","Montornès del Vallès","Canovelles","L'Ametlla del Vallès","Bigues i Riells","Santa Eulàlia de Ronçana","Vallromanes","Vilanova del Vallès","Sant Antoni de Vilamajor","Sant Pere de Vilamajor","Martorelles","Sant Feliu de Codines","Aiguafreda","Tagamanent","Figaró-Montmany"],
  "Alt Empordà":["Figueres","Roses","Castelló d'Empúries","L'Escala","Llançà","Vilafant","Peralada","Sant Pere Pescador","La Jonquera","Cadaqués","El Port de la Selva","Vilajuïga","Navata"],
  "Baix Empordà":["Palafrugell","La Bisbal d'Empordà","Sant Feliu de Guíxols","Torroella de Montgrí","Calonge i Sant Antoni","Palamós","Platja d'Aro","Begur","Pals","Forallac","Regencós","Corçà","Cruïlles","La Tallada d'Empordà"],
  "Cerdanya":["Puigcerdà","Bellver de Cerdanya","Alp","Das","Fontanals de Cerdanya","Llívia","Bolvir","Ger","Guils de Cerdanya","Isòvol","Meranges","Prats i Sansor","Prullans","Riu de Cerdanya","Urús"],
  "Garrotxa":["Olot","Sant Joan les Fonts","Besalú","Santa Pau","Les Preses","La Vall d'en Bas","Argelaguer","Tortellà","Castellfollit de la Roca","Sant Feliu de Pallerols","Riudaura","Mieres","Sant Jaume de Llierca","Montagut i Oix","Les Planes d'Hostoles"],
  "Gironès":["Girona","Salt","Sarrià de Ter","Cassà de la Selva","Celrà","Quart","Fornells de la Selva","Llagostera","Bescanó","Sant Julià de Ramis","Vilablareix","Sant Gregori","Aiguaviva","Bordils","Cervià de Ter","Flaçà","Juià","Llambilles","Madremanya","Sant Joan de Mollet","Sant Jordi Desvalls","Sant Martí de Llémena","Sant Martí Vell"],
  "Pla de l'Estany":["Banyoles","Porqueres","Cornellà del Terri","Crespià","Esponellà","Fontcoberta","Serinyà","Camós","Palol de Revardit"],
  "Ripollès":["Ripoll","Sant Joan de les Abadesses","Campdevànol","Ribes de Freser","Camprodon","Llanars","Vallfogona de Ripollès","Molló","Setcases","Gombrèn","Les Llosses","Ogassa","Pardines","Planoles","Queralbs","Toses","Vilallonga de Ter"],
  "Selva":["Blanes","Lloret de Mar","Santa Coloma de Farners","Tossa de Mar","Hostalric","Vidreres","Sils","Maçanet de la Selva","Arbúcies","Breda","Riells i Viabrea","Caldes de Malavella","Massanes","Anglès","Amer","Sant Hilari Sacalm"],
  "Alt Urgell":["La Seu d'Urgell","Oliana","Organyà","Montferrer i Castellbò","Ribera d'Urgellet","Coll de Nargó","Alàs i Cerc","Bassella","Peramola","Pont de Bar"],
  "Alta Ribagorça":["El Pont de Suert","Vilaller","La Vall de Boí"],
  "Garrigues":["Les Borges Blanques","Juneda","L'Albi","Arbeca","Belianes","Cervià de les Garrigues","El Cogul","La Floresta","La Granadella","Granyena de les Garrigues"],
  "Noguera":["Balaguer","Artesa de Segre","Camarasa","Cubells","Gerb","Os de Balaguer","Ponts","Térmens","Vallfogona de Balaguer","Vilanova de Meià"],
  "Pallars Jussà":["Tremp","La Pobla de Segur","Isona i Conca Dellà","Salàs de Pallars","Sant Esteve de la Sarga","Talarn","La Torre de Cabdella","Sarroca de Bellera"],
  "Pallars Sobirà":["Sort","Rialp","Llavorsí","Esterri d'Àneu","La Guingueta d'Àneu","Espot","Alt Àneu","Soriguera","Tírvia","Baix Pallars"],
  "Pla d'Urgell":["Mollerussa","Bell-lloc d'Urgell","Fondarella","Golmés","Linyola","Miralcamp","El Palau d'Anglesola","Vila-sana","Torregrossa","Sidamon","Castellnou de Seana"],
  "Segarra":["Cervera","Guissona","Sant Guim de Freixenet","Torà","Biosca","Ivorra","Les Oluges","Plans de Sió","Ribera d'Ondara","Sanaüja"],
  "Segrià":["Lleida","Alcarràs","Almacelles","Aitona","Alcoletge","Alfarràs","Alguaire","Almenar","Benavent de Segrià","Corbins","Gimenells i el Pla de la Font","Rosselló","Seròs","Soses","Sudanell","Torrefarrera","Torres de Segre","La Granja d'Escarp"],
  "Solsonès":["Solsona","Olius","Sant Llorenç de Morunys","Navès","Pinell de Solsonès","Clariana de Cardener","Riner","Castellar de la Ribera","La Coma i la Pedra","Lladurs","Pinós","La Molsosa"],
  "Urgell":["Tàrrega","Agramunt","Bellpuig","Anglesola","Verdú","Vilagrassa","Os de Balaguer","Preixens"],
  "Val d'Aran":["Vielha e Mijaran","Naut Aran","Les","Bossòst","Vilamòs","Es Bòrdes","Arres"],
  "Alt Camp":["Valls","El Pla de Santa Maria","Alcover","Aiguamúrcia","Bràfim","Cabra del Camp","El Pont d'Armentera","Figuerola del Camp","La Masó","Nulles","Puigpelat","El Rourell","Vila-rodona","Vilabella"],
  "Baix Camp":["Reus","Cambrils","Salou","Mont-roig del Camp","Riudoms","Vandellòs i l'Hospitalet de l'Infant","La Selva del Camp","Botarell","Castellvell del Camp","Les Borges del Camp","Montbrió del Camp","Pratdip","Vilanova d'Escornalbou","Vinyols i els Arcs"],
  "Baix Ebre":["Tortosa","Deltebre","Roquetes","L'Ampolla","El Perelló","L'Aldea","Camarles","Paüls","Benifallet","Xerta","Tivenys","Aldover"],
  "Baix Penedès":["El Vendrell","Calafell","Cunit","L'Arboç","Sant Jaume dels Domenys","Banyeres del Penedès","La Bisbal del Penedès","Bellvei","Llorenç del Penedès","Santa Oliva","Albinyana"],
  "Conca de Barberà":["Montblanc","L'Espluga de Francolí","Santa Coloma de Queralt","Barberà de la Conca","Blancafort","Conesa","Forès","Llorac","Passanant i Belltall","Pira","Rocafort de Queralt","Sarral","Savallà del Comtat","Senan","Solivella","Vallclara","Vilanova de Prades","Vimbodí i Poblet"],
  "Montsià":["Amposta","Sant Carles de la Ràpita","Alcanar","La Sénia","Ulldecona","Santa Bàrbara","Freginals","Godall","Mas de Barberans","Masdenverge"],
  "Priorat":["Falset","Cornudella de Montsant","Marçà","El Masroig","Gratallops","La Morera de Montsant","Porrera","Bellmunt del Priorat","La Bisbal de Falset","Cabacés","Capçanes","El Lloar","El Molar","La Torre de Fontaubella","Torroja del Priorat","Ulldemolins","La Vilella Alta","La Vilella Baixa"],
  "Ribera d'Ebre":["Móra d'Ebre","Móra la Nova","Flix","Ascó","Garcia","Ginestar","Miravet","Rasquera","Tivissa","La Torre de l'Espanyol","Vinebre"],
  "Tarragonès":["Tarragona","Vila-seca","Torredembarra","Constantí","El Morell","La Pobla de Mafumet","Altafulla","Creixell","El Catllar","La Canonja","La Nou de Gaià","La Riera de Gaià","La Secuita","Perafort","Renau","Roda de Berà","El Vendrell","Vespella de Gaià","Vilallonga del Camp"],
  "Terra Alta":["Gandesa","Batea","Bot","Caseres","Corbera d'Ebre","La Fatarella","Horta de Sant Joan","El Pinell de Brai","La Pobla de Massaluca","Prat de Comte","Vilalba dels Arcs"],
};
// Funció helper: donat un array de comarques seleccionades, retorna tots els termes a buscar (comarques + municipis + províncies)
function getGeoTerms(selectedComarques){
  const terms=new Set();
  for(const c of selectedComarques){
    terms.add(c.toLowerCase());
    const munis=COMARCA_MUNICIPIS[c];
    if(munis)for(const m of munis)terms.add(m.toLowerCase());
  }
  for(const p of PROVINCIES){
    if(p.comarques.every(c=>selectedComarques.includes(c)))terms.add(p.nom.toLowerCase());
  }
  return terms;
}
// Matching amb word boundaries per evitar falsos positius (ex: "vic" dins "servicio")
function matchGeoTerms(txt,geoTerms){
  const lower=txt.toLowerCase();
  for(const t of geoTerms){
    // Per termes curts (<=5 chars), requerim word boundary
    if(t.length<=5){
      const rx=new RegExp(`(?:^|[\\s,;.()/'"-])${t.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}(?:$|[\\s,;.()/'"-])`);
      if(rx.test(lower))return true;
    } else {
      if(lower.includes(t))return true;
    }
  }
  return false;
}
const CPV_GROUPS = [
  { label:"🏗️ Construcció general", items:[{code:"45000000",label:"45000000 – Construcció general"},{code:"45100000",label:"45100000 – Preparació del solar"},{code:"45110000",label:"45110000 – Demolicions i desmuntatges"},{code:"45111000",label:"45111000 – Treballs preparació de terrenys"},{code:"45112000",label:"45112000 – Excavació i terraplenament"}]},
  { label:"🏢 Edificació", items:[{code:"45200000",label:"45200000 – Edificis i eng. civil"},{code:"45210000",label:"45210000 – Construcció d'edificis"},{code:"45211000",label:"45211000 – Construcció d'habitatges"},{code:"45212000",label:"45212000 – Equipaments esportius i lleure"},{code:"45213000",label:"45213000 – Edificis comercials i industrials"},{code:"45214000",label:"45214000 – Edificis educatius"},{code:"45215000",label:"45215000 – Edificis sanitaris i socials"},{code:"45216000",label:"45216000 – Edificis serveis públics"},{code:"45260000",label:"45260000 – Cobertes"},{code:"45300000",label:"45300000 – Instal·lacions d'edificis"},{code:"45400000",label:"45400000 – Acabats d'edificis"},{code:"45410000",label:"45410000 – Arrebossats i enguixats"},{code:"45420000",label:"45420000 – Fusteria i ebenisteria"},{code:"45430000",label:"45430000 – Revestiments de sòls i parets"},{code:"45440000",label:"45440000 – Pintura i vidrieria"},{code:"45450000",label:"45450000 – Altres acabats"}]},
  { label:"🛣️ Urbanització i vials", items:[{code:"45230000",label:"45230000 – Autopistes, carreteres, vies"},{code:"45233000",label:"45233000 – Construcció i pavimentació"},{code:"45233200",label:"45233200 – Pavimentació"},{code:"45236000",label:"45236000 – Treballs de planament"},{code:"45112700",label:"45112700 – Paisatgisme"},{code:"45112500",label:"45112500 – Condicionament de terrenys"}]},
  { label:"💧 Xarxes i hidràulics", items:[{code:"45232000",label:"45232000 – Canonades i infraestructura"},{code:"45231300",label:"45231300 – Xarxes d'aigua i sanejament"},{code:"45232100",label:"45232100 – Abastament d'aigua"},{code:"45232400",label:"45232400 – Xarxes de clavegueram"},{code:"45246000",label:"45246000 – Regulació de cursos d'aigua"},{code:"45247000",label:"45247000 – Preses, canals i reg"}]},
  { label:"🌿 Jardineria", items:[{code:"77310000",label:"77310000 – Jardineria i zones verdes"},{code:"77300000",label:"77300000 – Serveis d'horticultura"},{code:"45112710",label:"45112710 – Jardineria en parcs"}]},
];
const CPV_OPTS = CPV_GROUPS.flatMap(g => g.items);
const CAT_LIMITS = {1:150000,2:360000,3:840000,4:2400000,5:5000000,6:null};
const CAT_LABELS  = {1:"150.000 €",2:"360.000 €",3:"840.000 €",4:"2.400.000 €",5:"5.000.000 €",6:"Sense límit"};
const SERVIAL_CLASS = {"A1":3,"C1":4,"C2":4,"C3":4,"C4":4,"C5":4,"C6":4,"C7":4,"C8":4,"C9":4,"E1":3,"E4":3,"E5":3,"E7":3,"G3":2,"G5":2,"G6":4,"K6":1};
const SERVIAL_CLASS_LABELS = {"A1":"Desmuntatges i buidatges","C1":"Demolicions","C2":"Estructura fàbrica/formigó","C3":"Estructures metàl·liques","C4":"Paleta, estucats i revestiments","C5":"Pedrera i marbre","C6":"Paviments i enrajolats","C7":"Aïllaments i impermeabilitzacions","C8":"Fusteria","C9":"Tancaments metàl·lics","E1":"Abastaments i sanejaments","E4":"Sèquies i desguassos","E5":"Defenses de marges i canalitzacions","E7":"Obres hidràuliques s/q.e.","G3":"Ferms de formigó hidràulic","G5":"Senyalitzacions i abalisaments","G6":"Obres vials s/q.e.","K6":"Jardineria i plantacions"};
const EMPTY_FILTERS = {importMin:"",importMax:"",comarques:[],cpvCodes:[],paraulesClau:"",nomesPotPresentar:false,nomesSuperiors:false};

const SK = "servial-licitacions-v4";
const SK_TIPUS = "servial-tipus-v1";
const ESTATS = ["PROPOSTA","NO PROPOSTA","EN ESTUDI","PRESENTADA","ADJUDICADA","NO PRESENTADA","NO ADJUDICADA","DESCARTADA"];
const EC = {"PROPOSTA":"bg-yellow-100 text-yellow-800","NO PROPOSTA":"bg-gray-200 text-gray-500","EN ESTUDI":"bg-purple-100 text-purple-800","PRESENTADA":"bg-blue-100 text-blue-800","ADJUDICADA":"bg-green-100 text-green-800","NO PRESENTADA":"bg-gray-200 text-gray-600","NO ADJUDICADA":"bg-orange-100 text-orange-800","DESCARTADA":"bg-red-100 text-red-800"};
const EBG = {"EN ESTUDI":"#E8F5E9","PROPOSTA":"#FFFDE7","PRESENTADA":"#E3F2FD","ADJUDICADA":"#C8E6C9","NO PRESENTADA":"#F5F5F5","NO ADJUDICADA":"#FFF3E0","DESCARTADA":"#FFFFFF","NO PROPOSTA":"#FFFFFF"};
const XBG = {"EN ESTUDI":"E8F5E9","PROPOSTA":"FFFDE7","PRESENTADA":"E3F2FD","ADJUDICADA":"C8E6C9","NO PRESENTADA":"F5F5F5","NO ADJUDICADA":"FFF3E0","DESCARTADA":"FFFFFF","NO PROPOSTA":"FFFFFF"};
const SEED = [
  {id:1,codi_obra:"26.03.001-ED",licitacio:"Habitatge tipus de 150 m2",client:"ALIASOL",public_privat:"PRIVAT",poblacio:"Indeterminada",estat:"EN ESTUDI",data_presentacio:"el client reclama l oferta",termini:"",import_pec_sense_iva:400000,classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"Raul Corpas 691 01 79 47",link_obra:"",analisi_completa:""},
  {id:2,codi_obra:"26.03.002-ED",licitacio:"Habitatge unifamiliar a Sant Quirze del Valles",client:"Quim Buxo",public_privat:"PRIVAT",poblacio:"Sant Quirze del Valles",estat:"EN ESTUDI",data_presentacio:"",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"rebudes ofertes pendent revisio",link_obra:"",analisi_completa:""},
  {id:3,codi_obra:"26.03.003-ED",licitacio:"3 habitatges unifamiliars a Sitges",client:"ALIASOL",public_privat:"PRIVAT",poblacio:"Sitges",estat:"EN ESTUDI",data_presentacio:"sin fecha",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"Pendiente editables del proyecto de Sitges.",link_obra:"",analisi_completa:""},
  {id:4,codi_obra:"26.03.004-ED",licitacio:"Casa unifamiliar Avgda. Mas Fuster 88 de Valldoreix",client:"Eldad Israel Erel",public_privat:"PRIVAT",poblacio:"Sant Cugat del Valles",estat:"EN ESTUDI",data_presentacio:"pendent parlar amb client",termini:"",import_pec_sense_iva:377500,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"PB: 75,53 m2 Total: 151,06 m2",link_obra:"",analisi_completa:""},
  {id:5,codi_obra:"26.03.005-ED",licitacio:"Casa a Pacs del Penedes",client:"Alfredo Lira",public_privat:"PRIVAT",poblacio:"Pacs del Penedes",estat:"EN ESTUDI",data_presentacio:"",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:6,codi_obra:"26.03.006-ED",licitacio:"Habitatge unifamiliar C/ Mirador n5 Sant Fost de Campsentelles",client:"Pau Pujolas Parset",public_privat:"PRIVAT",poblacio:"Sant Fost de Campsentelles",estat:"EN ESTUDI",data_presentacio:"projecte rebut 10/marc/26",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"628 598 618",link_obra:"",analisi_completa:""},
  {id:7,codi_obra:"26.03.007-ED",licitacio:"Reforma nau produccio industrial ELANFOODS",client:"Elanfoods",public_privat:"PRIVAT",poblacio:"Montcada i Reixach",estat:"EN ESTUDI",data_presentacio:"pendent rebre projecte modificat",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"No",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:8,codi_obra:"26.03.008-ED",licitacio:"Construccio edifici Prefectura Policia Local Caldes de Montbui Fase II",client:"Ajuntament de Caldes de Montbui",public_privat:"PUBLICA",poblacio:"Caldes de Montbui",estat:"EN ESTUDI",data_presentacio:"24/03/2026 23:53",termini:"8 mesos",import_pec_sense_iva:563925.81,classificacio:"C-2-3",criteris_puntuacio:"Preu 40pt + Termini garantia 15pt + Pla obres 45pt",aval:"Definitiva 5%",apertura:"",comentaris:"c/IVA: 682.350,23",link_obra:"https://contractaciopublica.cat/ca/detall-publicacio/300705890",analisi_completa:""},
  {id:9,codi_obra:"26.03.009-ED",licitacio:"Reforma banys i aules 3a planta Edifici Blanc Campus UAB",client:"Fundacio Universitat Autonoma de Barcelona",public_privat:"PUBLICA",poblacio:"Bellaterra",estat:"EN ESTUDI",data_presentacio:"15/04/2026 14:00",termini:"~8 setmanes",import_pec_sense_iva:231450.38,classificacio:"C-4-2 opcional",criteris_puntuacio:"Preu 60pt + Millores 10pt + Prog Treball 24pt + Pla Residus 6pt",aval:"Definitiva 5%",apertura:"",comentaris:"C-4-2 opcional substitueix solvencia",link_obra:"https://contractaciopublica.cat/ca/detall-publicacio/300704878",analisi_completa:""},
  {id:10,codi_obra:"",licitacio:"Renovacio tancaments Escola Berti",client:"Ajuntament de l Ametlla del Valles",public_privat:"PUBLICA",poblacio:"L Ametlla del Valles",estat:"PROPOSTA",data_presentacio:"07/04/2026 12:00",termini:"3 mesos",import_pec_sense_iva:245000,classificacio:"C-8-2 alternativa solvencia",criteris_puntuacio:"Preu 70pt + Termini execucio 30pt",aval:"Definitiva 5%",apertura:"",comentaris:"SERVIAL COMPLEIX C-8-4",link_obra:"https://contractaciopublica.cat/ca/detall-publicacio/300717406",analisi_completa:""},
  {id:11,codi_obra:"",licitacio:"Rehabilitacio Masia ca nAltimira",client:"Ajuntament Cerdanyola",public_privat:"PUBLICA",poblacio:"Cerdanyola del Valles",estat:"DESCARTADA",data_presentacio:"03/02/2026 14:00",termini:"",import_pec_sense_iva:2137343.70,classificacio:"K74",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"UTE Seranco Mucha alza no vamos",link_obra:"",analisi_completa:""},
  {id:12,codi_obra:"",licitacio:"Reforma seu Regio Sanitaria Nord CatSalut Sant Cugat",client:"CatSalut Servei Catala de la Salut",public_privat:"PUBLICA",poblacio:"Sant Cugat del Valles",estat:"DESCARTADA",data_presentacio:"07/04/2026 10:00",termini:"20 setmanes",import_pec_sense_iva:984715.83,classificacio:"C-4-3 J-2-3",criteris_puntuacio:"Preu 60pt + Tecnics 40pt",aval:"",apertura:"",comentaris:"FALTA CLASSIFICACIO J-2-3",link_obra:"https://contractaciopublica.cat/ca/detall-publicacio/300713083",analisi_completa:""},
  {id:13,codi_obra:"",licitacio:"Reforma vestidors Camp futbol Can Tito Fase 2 i 3",client:"Ajuntament Vilanova del Cami",public_privat:"PUBLICA",poblacio:"Vilanova del Cami",estat:"DESCARTADA",data_presentacio:"16/02/2026 23:59",termini:"",import_pec_sense_iva:435762.11,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:14,codi_obra:"",licitacio:"Rehabilitacio escola Emili Carles Tolra Castellar del Valles",client:"Ajuntament Castellar del Valles",public_privat:"PUBLICA",poblacio:"Castellar del Valles",estat:"DESCARTADA",data_presentacio:"17/02/2026 14:00",termini:"",import_pec_sense_iva:1494815.38,classificacio:"C64",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:15,codi_obra:"",licitacio:"Construccio Biblioteca Districte Passeig Valldaura Nou Barris",client:"BIMSA",public_privat:"PUBLICA",poblacio:"Barcelona",estat:"DESCARTADA",data_presentacio:"21/03/2026 13:00",termini:"",import_pec_sense_iva:18024840.94,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:16,codi_obra:"",licitacio:"2a Fase Casa de la Vila de Capellades",client:"Ajuntament Capellades",public_privat:"PUBLICA",poblacio:"Capellades",estat:"DESCARTADA",data_presentacio:"24/02/2026 23:59",termini:"",import_pec_sense_iva:749134.53,classificacio:"C43 I93",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"Fons Next Generation EU",link_obra:"",analisi_completa:""},
  {id:17,codi_obra:"",licitacio:"Nou menjador catering i banys Escola El Turo de Masquefa",client:"Departament Educacio Generalitat de Catalunya",public_privat:"PUBLICA",poblacio:"Masquefa",estat:"DESCARTADA",data_presentacio:"25/03/2026 20:00",termini:"240 dies naturals",import_pec_sense_iva:384095.38,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"Produccio baixa 45000 mes",link_obra:"",analisi_completa:""},
  {id:18,codi_obra:"",licitacio:"Remodelacio mitgera C Lluca 4 i Jardins Maria Manent",client:"Barcelona Infraestructures Municipals SA BIMSA",public_privat:"PUBLICA",poblacio:"Barcelona",estat:"DESCARTADA",data_presentacio:"27/03/2026 13:00",termini:"6 mesos",import_pec_sense_iva:986556.47,classificacio:"C-3-4",criteris_puntuacio:"Preu 30pt + Automatics 30pt + Judici valor 40pt",aval:"5% adjudicacio s/IVA",apertura:"",comentaris:"Solvencia volum negoci 828707 any",link_obra:"",analisi_completa:""},
  {id:19,codi_obra:"",licitacio:"Nau Industrial Plataforma PRIMA Gurb IREC",client:"Fundacio Institut de Recerca en Energia de Catalunya",public_privat:"PUBLICA",poblacio:"Gurb",estat:"DESCARTADA",data_presentacio:"09/03/2026 14:00",termini:"",import_pec_sense_iva:6935084.26,classificacio:"C25 4",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:20,codi_obra:"",licitacio:"Millora energetica escoles Cami del Mig Torre Llauder Cami del Cros Les Aigues",client:"Ajuntament de Mataro",public_privat:"PUBLICA",poblacio:"Mataro",estat:"DESCARTADA",data_presentacio:"09/03/2026 23:59",termini:"",import_pec_sense_iva:948316.95,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",analisi_completa:""},
  {id:21,codi_obra:"",licitacio:"Acord marc obres edificacio entitats locals de Catalunya",client:"Consorci Catala pel Desenvolupament Local",public_privat:"PUBLICA",poblacio:"Barcelona provincia",estat:"PROPOSTA",data_presentacio:"20/04/2026 15:00",termini:"",import_pec_sense_iva:1824809849.45,classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"133 lots pendent valorar",link_obra:"",analisi_completa:""},
];
const emptyForm = {codi_obra:"",licitacio:"",client:"",public_privat:"PUBLICA",poblacio:"",estat:"PROPOSTA",data_presentacio:"",termini:"",import_pec_sense_iva:"",classificacio:"",criteris_puntuacio:"",aval:"",apertura:"",comentaris:"",link_obra:"",link_publicacio:"",analisi_completa:""};

const parseCat = v => { if(!v&&v!==0)return 0; const s=String(v).toLowerCase().trim(); if(s==="a"||s==="1")return 1; if(s==="b"||s==="2")return 2; if(s==="c"||s==="3")return 3; if(s==="d"||s==="4")return 4; if(s==="e"||s==="5")return 5; if(s==="f"||s==="6")return 6; return parseInt(s)||0; };
const checkCompatibility = (requerida=[]) => { if(!requerida||requerida.length===0)return{pot:null,items:[]}; const items=requerida.map(r=>{const codi=(r.grup||"")+(r.subgrup||""),catReq=parseCat(r.categoria),catServial=SERVIAL_CLASS[codi]??null,denominacio=SERVIAL_CLASS_LABELS[codi]||`Grup ${r.grup} Subgrup ${r.subgrup}`,status=catServial===null?"absent":catServial>=catReq?"ok":"inferior";return{codi,denominacio,catReq,catServial,status};}); return{pot:items.every(i=>i.status==="ok"),items}; };

// Categoria per import (LCSP): 1=<=150k, 2=<=360k, 3=<=840k, 4=<=2.4M, 5=<=5M, 6=>5M
const catPerImport = imp => { if(imp<=150000)return 1; if(imp<=360000)return 2; if(imp<=840000)return 3; if(imp<=2400000)return 4; if(imp<=5000000)return 5; return 6; };

// Mapa CPV → grups de classificació probables
const CPV_TO_GRUPS = {
  "45100":["A1"],"45110":["A1","C1"],"45111":["A1"],"45112":["A1","K6"],
  "45200":["C2","C4"],"45210":["C2","C4"],"45211":["C2","C4"],"45212":["C2","C4"],
  "45213":["C2","C3"],"45214":["C2","C4"],"45215":["C2","C4"],"45216":["C2","C4"],
  "45220":["C2","C3"],"45230":["G6","G3"],"45231":["E1"],"45232":["E1","E4"],
  "45233":["G6","G3","G5"],"45236":["G6"],"45240":["E7"],"45246":["E7"],
  "45247":["E7","E5"],"45260":["C7"],"45300":["C4","C9"],"45310":["C9"],
  "45320":["C7"],"45330":["E1"],"45340":["C9"],"45400":["C4","C6"],
  "45410":["C4"],"45420":["C8"],"45430":["C6"],"45440":["C4"],"45450":["C4","C5"],
  "77300":["K6"],"77310":["K6"],"45000":["C2","C4"],
};
const inferClassificacio = (cpv, imp) => {
  if(imp<500000)return []; // No cal classificació per sota 500k
  const prefix = (cpv||"45000").replace(/-.*$/,"").slice(0,5);
  const grups = CPV_TO_GRUPS[prefix] || CPV_TO_GRUPS[prefix.slice(0,5)] || ["C2","C4"];
  const cat = catPerImport(imp);
  // Retorna el primer grup com a requeriment estimat
  return grups.slice(0,1).map(g => ({ grup: g.charAt(0), subgrup: g.slice(1), categoria: cat }));
};
const parseJSON = raw => { if(!raw)return null; if(raw.trim().startsWith("["))try{return JSON.parse(raw.trim());}catch(e){} const m1=raw.match(/\[[\s\S]*\]/); if(m1)try{return JSON.parse(m1[0]);}catch(e){try{return JSON.parse(m1[0].replace(/,(\s*[}\]])/g,"$1"));}catch(e2){}} const m2=raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/); if(m2)try{return JSON.parse(m2[1]);}catch(e){} return null; };
// Robust JSON parser for AI plec output: tries plain parse, code-fence strip,
// trailing-comma stripping, malformed-key recovery, and array/object extraction.
// Returns null on failure.
const parseJSONLenient = raw => {
  if(!raw)return null;
  const stripCommas = s => s.replace(/,(\s*[}\]])/g,"$1");
  // AI sometimes drops the `:"` between key and a value starting with `<` or `>`
  // (e.g. produces `"rang">20%"` instead of `"rang":">20%"`). Re-insert it.
  const fixOperatorKeys = s => s.replace(/"(\w+)"([<>])/g,'"$1":"$2');
  const tryAll = s => {
    if(!s)return null;
    try{return JSON.parse(s);}catch(e){}
    try{return JSON.parse(stripCommas(s));}catch(e){}
    try{return JSON.parse(fixOperatorKeys(s));}catch(e){}
    try{return JSON.parse(stripCommas(fixOperatorKeys(s)));}catch(e){}
    return null;
  };
  // 1) Direct parse
  const direct = tryAll(raw.trim());
  if(direct!==null)return direct;
  // 2) Strip code fence
  const fence = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
  if(fence){const r=tryAll(fence[1]);if(r!==null)return r;}
  // 3) Largest array
  const arr = raw.match(/\[[\s\S]*\]/);
  if(arr){const r=tryAll(arr[0]);if(r!==null)return r;}
  // 4) Largest object
  const obj = raw.match(/\{[\s\S]*\}/);
  if(obj){const r=tryAll(obj[0]);if(r!==null)return r;}
  return null;
};
const fmt = n => { const num=Number(n); if(!n&&n!==0)return"---"; if(isNaN(num))return"---"; return num.toLocaleString("ca-ES",{minimumFractionDigits:2,maximumFractionDigits:2})+" EUR"; };
const norm = s => String(s||"").toLowerCase().trim().replace(/\s+/g," ");
const sim = (a,b) => { a=norm(a);b=norm(b); if(!a||!b)return 0; if(a===b)return 1; const lo=a.length>b.length?a:b,sh=a.length>b.length?b:a; if(lo.includes(sh)&&sh.length>15)return 0.9; const wa=a.split(" ").filter(w=>w.length>3),wb=new Set(b.split(" ").filter(w=>w.length>3)); if(!wa.length)return 0; return wa.filter(w=>wb.has(w)).length/Math.max(wa.length,wb.size); };
const parseD = s => { if(!s)return null; const p=String(s).split("/"); if(p.length<3)return null; const d=parseInt(p[0],10),mo=parseInt(p[1],10),y=parseInt(p[2],10); if(isNaN(d)||isNaN(mo)||isNaN(y))return null; return new Date(y,mo-1,d); };
const parseDT = s => { if(!s)return null; const parts=String(s).split("/"); if(parts.length<3)return null; const rest=parts[2].trim().split(" "); if(rest.length<2)return null; const hm=rest[1].replace("h","").split(":"); if(hm.length<2)return null; return{day:parts[0].trim().padStart(2,"0"),mon:parts[1].trim().padStart(2,"0"),yr:rest[0],hr:hm[0].padStart(2,"0"),min:hm[1].padStart(2,"0")}; };
const codiNum = c => { if(!c)return 9999; const p=c.split("."); if(p.length<3)return 9999; const n=parseInt(p[2],10); return isNaN(n)?9999:n; };
const eOrder = e => ({"EN ESTUDI":0,"PROPOSTA":1,"PRESENTADA":2,"ADJUDICADA":3,"NO PRESENTADA":4,"NO ADJUDICADA":5,"NO PROPOSTA":6,"DESCARTADA":7}[e]||8);

const makeTableHTML = (data,noDesc) => {
  const date=new Date().toLocaleDateString("ca-ES"),suffix=noDesc?" (sense descartades)":"";
  const ths=["Codi","Licitacio","Client","P/P","Estat","Data Present.","Termini","Import s/IVA","Classif.","Aval","Comentaris"];
  const thead=ths.map(h=>`<th style='background:#1F3864;color:white;padding:4px;text-align:left;font-size:8px;white-space:nowrap'>${h}</th>`).join("");
  const rows=data.map(l=>{const bg=EBG[l.estat]||"#FFFFFF";const imp=l.import_pec_sense_iva?fmt(l.import_pec_sense_iva):"---";const vals=[l.codi_obra,l.licitacio,l.client,l.public_privat==="PUBLICA"?"PUB":"PRI",l.estat,l.data_presentacio,l.termini,imp,l.classificacio,l.aval,l.comentaris];const tds=vals.map((v,i)=>`<td style='padding:3px 4px;border-bottom:1px solid #D9E1F2;vertical-align:top;font-size:8px;text-align:${i===7?"right":"left"}'>${String(v||"---")}</td>`).join("");return`<tr style='background:${bg}'>${tds}</tr>`;}).join("");
  return`<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{font-family:Arial,sans-serif;margin:10mm;font-size:8px}table{width:100%;border-collapse:collapse}h2{color:#1F3864;font-size:14px;margin:0 0 4px 0}p{font-size:8px;color:#666;margin:0 0 8px 0}@media print{body{margin:5mm}}</style></head><body><h2>Licitacions Servial 2026</h2><p>${date} - ${data.length} licitacions${suffix}</p><table><thead><tr>${thead}</tr></thead><tbody>${rows}</tbody></table></body></html>`;
};

const addToCalendar = (l,setCal) => {
  const dt=parseDT(l.data_presentacio); if(!dt){alert("No s'ha pogut llegir la data: "+l.data_presentacio);return;}
  const s=`${dt.yr}-${dt.mon}-${dt.day}T${dt.hr}:${dt.min}:00`,eh=String(parseInt(dt.hr,10)+1).padStart(2,"0"),e=`${dt.yr}-${dt.mon}-${dt.day}T${eh}:${dt.min}:00`;
  const content=`Crea un event al Google Calendar:\nTitol: PRESENTACIO - ${l.licitacio||""}\nData inici: ${s}\nData fi: ${e}\nDescripcio: Client: ${l.client||"---"} | Import: ${l.import_pec_sense_iva?fmt(l.import_pec_sense_iva):"---"}${l.link_obra?" | "+l.link_obra:""}\nRecordatori: 48 hores abans.`;
  setCal("Creant event...");
  fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:apiHeaders(),body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:500,mcp_servers:[{type:"url",url:"https://gcal.mcp.claude.com/mcp",name:"gcal"}],messages:[{role:"user",content}]})})
    .then(r=>r.json()).then(d=>{const t=(d.content&&d.content.find(b=>b.type==="text")||{}).text||"Event creat.";setCal("");alert(t);}).catch(er=>{setCal("");alert("Error: "+er.message);});
};

// ── Parser CIDO: extreu licitacions del correu HTML de sgsi.svcSam@diba.cat ──
function parseCIDOHtml(rawBody){
  if(!rawBody)return[];
  let body=rawBody;
  try{body=rawBody.replace(/\\r\\n/g,"\n").replace(/\\r/g,"").replace(/\\n/g,"\n").replace(/\\t/g,"\t").replace(/\\"/g,'"').replace(/\\\//g,"/").replace(/\\\\/g,"\\");}catch(e){}
  const cleanHtml=t=>(t||"").replace(/<[^>]+>/g,"").replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&nbsp;/g," ").replace(/&#39;/g,"'").replace(/&quot;/g,'"').replace(/&[a-z]+;/g," ").replace(/\s+/g," ").trim();
  const licitacions=[];
  // Strategy 1: Find all item__descr blocks with CIDO links
  const descrRx=/<p[^>]*class="[^"]*item__descr[^"]*"[^>]*>\s*<a\s+href="(https?:\/\/cido\.diba\.cat\/contractacio\/[^"?]+)[^"]*"[^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while((m=descrRx.exec(body))!==null){
    const url=m[1];
    const title=cleanHtml(m[2]);
    if(!title||title.length<3)continue;
    // Look backwards for org name in preceding <td> that contains plain text (no item__descr)
    const before=body.substring(Math.max(0,m.index-2000),m.index);
    const orgMatch=before.match(/<td[^>]*>\s*([^<]{3,}?)\s*<\/td>\s*(?:<\/tr>[\s\S]*?<tr[^>]*>[\s\S]*?)?$/i);
    const organisme=orgMatch?cleanHtml(orgMatch[1]):"";
    licitacions.push({organisme,titol:title,url});
  }
  // Strategy 2 fallback: find all href to cido.diba.cat/contractacio
  if(licitacions.length===0){
    const linkRx=/<a\s+href="(https?:\/\/cido\.diba\.cat\/contractacio\/[^"?]+)[^"]*"[^>]*>([\s\S]*?)<\/a>/gi;
    while((m=linkRx.exec(body))!==null){
      const url=m[1];
      const title=cleanHtml(m[2]);
      if(title&&title.length>3)licitacions.push({organisme:"",titol:title,url});
    }
  }
  return licitacions;
}

/* ── Helpers Gmail API ── */
function extractGmailHtmlBody(payload){
  // Gmail messages can have nested parts
  if(payload.mimeType==="text/html"&&payload.body?.data){
    return atob(payload.body.data.replace(/-/g,"+").replace(/_/g,"/"));
  }
  if(payload.parts){
    for(const part of payload.parts){
      const result=extractGmailHtmlBody(part);
      if(result)return result;
    }
  }
  // Fallback: try plain text
  if(payload.mimeType==="text/plain"&&payload.body?.data){
    return atob(payload.body.data.replace(/-/g,"+").replace(/_/g,"/"));
  }
  return null;
}

function parseGencatToResults(html){
  // Parse plataforma.contractacio@gencat.cat emails
  // These emails contain licitació notifications with links to contractaciopublica.cat
  const results=[];
  if(!html)return results;
  // Look for links to contractaciopublica.cat
  const linkRx=/<a[^>]*href=["']([^"']*contractaciopublica\.cat[^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while((m=linkRx.exec(html))!==null){
    const url=m[1].replace(/&amp;/g,"&");
    const text=m[2].replace(/<[^>]+>/g,"").replace(/&[a-z]+;/g," ").replace(/\s+/g," ").trim();
    if(text.length>10&&!text.includes("Dessubscriure")&&!text.includes("baixa")){
      results.push({expedient:"",objecte:text,organisme:"",import_eur:0,data_publicacio:new Date().toLocaleDateString("ca-ES"),termini:"",cpv:"45000000",comarca_municipi:"",tipologia:"",font:"Gencat-Email",classificacio_requerida:[],puntuacio:6,justificacio:"📧 Correu Gencat",url});
    }
  }
  // If no contractaciopublica links, try generic parsing like CIDO
  if(results.length===0){
    const blockRx=/item__title[^>]*>([^<]+)<[\s\S]*?item__descr[^>]*>([\s\S]*?)<\/td>/g;
    while((m=blockRx.exec(html))!==null){
      const org=m[1].trim();
      const content=m[2].replace(/<[^>]+>/g," ").replace(/&[a-z]+;/g," ").replace(/\s+/g," ").trim();
      if(org&&content.length>5){
        const urlM=m[2].match(/href=["']([^"']+)["']/);
        results.push({expedient:"",objecte:content.slice(0,300),organisme:org,import_eur:0,data_publicacio:new Date().toLocaleDateString("ca-ES"),termini:"",cpv:"45000000",comarca_municipi:"",tipologia:"",font:"Gencat-Email",classificacio_requerida:[],puntuacio:6,justificacio:"📧 Correu Gencat",url:urlM?urlM[1]:""});
      }
    }
  }
  // Filter only construction-related
  const paraules=["obres","obra","construcció","construccio","reforma","rehabilitació","rehabilitacio","urbanització","urbanitzacio","paviment","reurbanitz","condicionament","enderroc","sanejament","col·lector","canonada","pont","vestidor","edifici","local","equipament","instal·lació","instal·lacions","adequació","adequacio","millora","ampliació","ampliacio","substitució","substitucio","impermeabilitz","drenatge","clavegueram","xarxa","vial","asfalt","demolició","demolicio","mur","coberta","façana","facana","forjat","fonament","estructura","terraplenat","esplanada","ferm"];
  const norm=s=>s.normalize("NFC").replace(/\s+/g," ").trim().toLowerCase();
  const filtered=results.filter(r=>{
    const t=norm(r.objecte||"");
    if(t.includes("subministrament")||t.includes("suministro"))return false;
    if(/servei/.test(t))return false;
    if(/servicio/.test(t))return false;
    if(t.includes("consultoria")||t.includes("assessorament")||t.includes("formaci"))return false;
    return true;
  });
  return filtered;
}

function FR({label,span2,children}){return(<div className={span2?"col-span-2":""}><label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">{label}</label><div className="mt-0.5">{children}</div></div>);}
function Badge({score}){const cls=score>=8?"bg-green-100 text-green-800":score>=5?"bg-slate-100 text-amber-800":"bg-red-100 text-red-700";return<span className={`text-xs font-bold px-2 py-0.5 rounded-full ${cls}`}>{score}/10</span>;}
function CompatBadge({pot}){if(pot===null)return<span className="text-xs bg-green-50 text-green-700 font-semibold px-2 py-0.5 rounded-full">✅ Sense classif. requerida</span>;if(pot)return<span className="text-xs bg-blue-100 text-blue-700 font-semibold px-2 py-0.5 rounded-full">🔍 Probable compliment</span>;return<span className="text-xs bg-slate-100 text-slate-600 font-semibold px-2 py-0.5 rounded-full">⚠️ Probable classif. insuficient</span>;}
function CompatDetail({items}){if(!items.length)return null;return<div className="mt-2 space-y-1">{items.map((it,i)=>(<div key={i} className={`flex items-center gap-2 text-xs px-2 py-1 rounded-lg ${it.status==="ok"?"bg-green-50 text-green-800":it.status==="inferior"?"bg-amber-50 text-amber-800":"bg-red-50 text-red-700"}`}><span className="font-bold w-8 shrink-0">{it.codi}</span><span className="flex-1">{it.denominacio}</span><span className="shrink-0">{it.status==="ok"&&`✅ Cat.${it.catServial} (req.${it.catReq})`}{it.status==="inferior"&&`⚠️ Cat.${it.catServial}→Cat.${it.catReq}`}{it.status==="absent"&&`❌ No acreditat`}</span></div>))}</div>;}

function GestorTab({refreshKey=0}){
  const [lic,setLic]=useState([]);
  const [loaded,setLoaded]=useState(false);
  const [tipus,setTipus]=useState(["PUBLICA","PRIVAT"]);
  const [showTipusSettings,setShowTipusSettings]=useState(false);
  const [newTipus,setNewTipus]=useState("");
  const [view,setView]=useState("table");
  const [form,setForm]=useState(emptyForm);
  const [editId,setEditId]=useState(null);
  const [search,setSearch]=useState("");
  const [fEstat,setFEstat]=useState("ACTIVES");
  const [fTipus,setFTipus]=useState("");
  const [colFilters,setColFilters]=useState({});
  const [activeColFilter,setActiveColFilter]=useState(null);
  const toggleColFilter=(col,val)=>setColFilters(prev=>{const cur=prev[col]||[];return{...prev,[col]:cur.includes(val)?cur.filter(v=>v!==val):[...cur,val]};});
  const clearColFilter=col=>setColFilters(prev=>{const n={...prev};delete n[col];return n;});
  const hasActiveColFilters=Object.values(colFilters).some(v=>v&&v.length>0);
  const ACTIVES=["EN ESTUDI","PROPOSTA"];const SENSE_DESC=ESTATS.filter(e=>e!=="DESCARTADA");
  const [showExport,setShowExport]=useState(false);
  const [analisi,setAnalisi]=useState("");
  const [showAnalisi,setShowAnalisi]=useState(false);
  const [confModal,setConfModal]=useState(null);
  const [dupModal,setDupModal]=useState(null);
  const [showDup,setShowDup]=useState(false);
  const [delId,setDelId]=useState(null);
  const [calSt,setCalSt]=useState("");
  const [backupMsg,setBackupMsg]=useState("");
  const importRef=useRef(null);

  const exportBackup=()=>{
    const data={version:1,date:new Date().toISOString(),licitacions:lic,tipus};
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);
    a.download=`backup_servial_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    setBackupMsg("✅ Backup descarregat");setTimeout(()=>setBackupMsg(""),3000);
  };
  const generarCorreuSetmanal=async()=>{
    const win=window.open("","_blank");
    if(!win){alert("Activa les finestres emergents per generar el correu.");return;}
    win.document.write("<html><body><p>Generant informe...</p></body></html>");
    try{
    const fmtE=v=>v?Number(v).toLocaleString("ca-ES",{minimumFractionDigits:2})+" €":"—";
    const subtotal=arr=>arr.reduce((s,l)=>s+(parseFloat(l.import_pec_sense_iva)||0),0);
    // MOVIMENTS dels últims 7 dies (del log + de la taula licitacions per created_at)
    const fa7=new Date();fa7.setDate(fa7.getDate()-7);
    let moviments=[];
    let licsRecents=[];
    try{const{data:logs}=await supabase.from("licitacions_log").select("*").gte("created_at",fa7.toISOString()).order("created_at",{ascending:false});moviments=logs||[];}catch(e){console.error("Log error:",e);}
    try{const{data:lr}=await supabase.from("licitacions").select("id,created_at,estat").gte("created_at",fa7.toISOString());licsRecents=lr||[];}catch(e){console.error("LicsRecents error:",e);}
    const getLic=id=>{const nid=Number(id);return lic.find(l=>l.id===nid||l.id===id);};
    const isNew=m=>!m.estat_anterior||m.estat_anterior===""||m.estat_anterior==="(buit)";
    const dedupById=arr=>{const seen=new Set();return arr.filter(l=>{if(!l||seen.has(l.id))return false;seen.add(l.id);return true;});};
    // Noves licitacions = totes les entrades a la base aquesta setmana (combina log + created_at de Supabase)
    const idsNoves=new Set();
    moviments.filter(m=>isNew(m)).forEach(m=>idsNoves.add(Number(m.licitacio_id)));
    licsRecents.forEach(r=>idsNoves.add(Number(r.id)));
    const novesAll=dedupById([...idsNoves].map(id=>getLic(id)));
    // Separar segons l'estat actual
    const movNovesPropostes=novesAll.filter(l=>l.estat==="PROPOSTA");
    const movNovesEnEstudi=novesAll.filter(l=>l.estat==="EN ESTUDI");
    const movNovesAltres=novesAll.filter(l=>l.estat!=="PROPOSTA"&&l.estat!=="EN ESTUDI");
    // Passen a En Estudi = només canvis d'estat (no noves directes a En Estudi)
    const movPassenEstudi=dedupById(moviments.filter(m=>!isNew(m)&&m.estat_nou==="EN ESTUDI").map(m=>getLic(m.licitacio_id)));
    const movPresentades=dedupById(moviments.filter(m=>m.estat_nou==="PRESENTADA").map(m=>getLic(m.licitacio_id)));
    const movNoPresentades=dedupById(moviments.filter(m=>m.estat_nou==="NO PRESENTADA").map(m=>getLic(m.licitacio_id)));
    const movAdjudicades=dedupById(moviments.filter(m=>m.estat_nou==="ADJUDICADA").map(m=>getLic(m.licitacio_id)));
    const movNoAdjudicades=dedupById(moviments.filter(m=>m.estat_nou==="NO ADJUDICADA").map(m=>getLic(m.licitacio_id)));
    const movDescartades=dedupById(moviments.filter(m=>m.estat_nou==="DESCARTADA").map(m=>getLic(m.licitacio_id)));
    const hiHaMoviments=movNovesPropostes.length+movNovesEnEstudi.length+movNovesAltres.length+movPassenEstudi.length+movPresentades.length+movNoPresentades.length+movAdjudicades.length+movNoAdjudicades.length+movDescartades.length>0;
    // Agrupar licitacions
    const presentades=lic.filter(l=>l.estat==="PRESENTADA");
    const noPresentades=lic.filter(l=>l.estat==="NO PRESENTADA");
    const propostesPub=lic.filter(l=>l.estat==="PROPOSTA"&&l.public_privat==="PUBLICA");
    const propostesPriv=lic.filter(l=>l.estat==="PROPOSTA"&&l.public_privat!=="PUBLICA");
    const estudiPub=lic.filter(l=>l.estat==="EN ESTUDI"&&l.public_privat==="PUBLICA");
    const estudiPriv=lic.filter(l=>l.estat==="EN ESTUDI"&&l.public_privat!=="PUBLICA");
    const adjudicades=lic.filter(l=>l.estat==="ADJUDICADA");
    const noAdjudicades=lic.filter(l=>l.estat==="NO ADJUDICADA");
    // Estils comuns
    const thS='style="background-color:#1a3d6e;color:#ffffff;font-weight:bold;border:1px solid #b5b5b5;padding:6px 8px;text-align:left;"';
    const thR='style="background-color:#1a3d6e;color:#ffffff;font-weight:bold;border:1px solid #b5b5b5;padding:6px 8px;text-align:right;"';
    const tdS='style="border:1px solid #b5b5b5;padding:6px 8px;vertical-align:top;"';
    const tdR='style="border:1px solid #b5b5b5;padding:6px 8px;vertical-align:top;text-align:right;white-space:nowrap;"';
    const secTitle=(num,title,count,total)=>`<p style="font-family:Arial,Helvetica,sans-serif;font-size:12pt;font-weight:bold;color:#1a3d6e;margin:24px 0 8px 0;border-bottom:2px solid #1a3d6e;padding-bottom:4px;">${num}) ${title} — ${count} licitaci${count!==1?"ons":"ó"} · ${fmtE(total)}</p>`;
    const emptyMsg=`<p style="font-style:italic;color:#777;background-color:#fafafa;padding:8px;border:1px dashed #c0c0c0;margin:6px 0 14px 0;">— Cap licitació en aquest apartat. Subtotal: 0,00 € —</p>`;
    const subRow=(cols,total)=>`<tr style="background-color:#eef2f7;"><td colspan="${cols-1}" ${tdS} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;font-weight:bold;">Subtotal</td><td ${tdR} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;white-space:nowrap;font-weight:bold;">${fmtE(total)}</td><td ${tdS}></td></tr>`;
    const rowBg=(i)=>i%2===1?` style="background-color:#fafbfc;"`:"";
    // Taula presentades (Codi, Licitació, Client, Import, Presentada, Execució)
    const tblPresentades=()=>{
      if(!presentades.length)return emptyMsg;
      let h=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;margin:6px 0 14px 0;border:1px solid #b5b5b5;"><tr><th ${thS}>Codi</th><th ${thS}>Licitació</th><th ${thS}>Client</th><th ${thR}>Import</th><th ${thS}>Presentada</th><th ${thS}>Execució</th></tr>`;
      presentades.forEach((l,i)=>{h+=`<tr${rowBg(i)}><td ${tdS}>${l.codi_obra||"—"}</td><td ${tdS}>${l.licitacio||"—"}</td><td ${tdS}>${l.client||"—"}</td><td ${tdR}>${fmtE(l.import_pec_sense_iva)}</td><td ${tdS}>${l.data_presentacio||"—"}</td><td ${tdS}>${l.termini||"—"}</td></tr>`;});
      h+=`<tr style="background-color:#eef2f7;"><td colspan="3" ${tdS} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;font-weight:bold;">Subtotal</td><td ${tdR} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;white-space:nowrap;font-weight:bold;">${fmtE(subtotal(presentades))}</td><td colspan="2" ${tdS}></td></tr></table>`;
      return h;
    };
    // Taula no presentades (Codi, Licitació, Client, Import, Termini)
    const tblNoPresentades=()=>{
      if(!noPresentades.length)return emptyMsg;
      let h=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;margin:6px 0 14px 0;border:1px solid #b5b5b5;"><tr><th ${thS}>Codi</th><th ${thS}>Licitació</th><th ${thS}>Client</th><th ${thR}>Import</th><th ${thS}>Termini</th></tr>`;
      noPresentades.forEach((l,i)=>{h+=`<tr${rowBg(i)}><td ${tdS}>${l.codi_obra||"—"}</td><td ${tdS}>${l.licitacio||"—"}</td><td ${tdS}>${l.client||"—"}</td><td ${tdR}>${fmtE(l.import_pec_sense_iva)}</td><td ${tdS}>${l.data_presentacio||"—"}</td></tr>`;});
      h+=`<tr style="background-color:#eef2f7;"><td colspan="3" ${tdS} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;font-weight:bold;">Subtotal</td><td ${tdR} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;white-space:nowrap;font-weight:bold;">${fmtE(subtotal(noPresentades))}</td><td ${tdS}></td></tr></table>`;
      return h;
    };
    // Taula propostes (Licitació, Client, Import, Termini, Execució)
    const tblPropostes=(items)=>{
      if(!items.length)return emptyMsg;
      let h=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;margin:6px 0 14px 0;border:1px solid #b5b5b5;"><tr><th ${thS}>Licitació</th><th ${thS}>Client</th><th ${thR}>Import</th><th ${thS}>Termini</th><th ${thS}>Execució</th></tr>`;
      items.forEach((l,i)=>{h+=`<tr${rowBg(i)}><td ${tdS}>${l.licitacio||"—"}</td><td ${tdS}>${l.client||"—"}</td><td ${tdR}>${fmtE(l.import_pec_sense_iva)}</td><td ${tdS}>${l.data_presentacio||"—"}</td><td ${tdS}>${l.termini||"—"}</td></tr>`;});
      h+=`<tr style="background-color:#eef2f7;"><td colspan="2" ${tdS} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;font-weight:bold;">Subtotal</td><td ${tdR} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;white-space:nowrap;font-weight:bold;">${fmtE(subtotal(items))}</td><td colspan="2" ${tdS}></td></tr></table>`;
      return h;
    };
    // Taula en estudi (Codi, Licitació, Client, Import, Situació)
    const tblEstudi=(items)=>{
      if(!items.length)return emptyMsg;
      let h=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;margin:6px 0 14px 0;border:1px solid #b5b5b5;"><tr><th ${thS}>Codi</th><th ${thS}>Licitació</th><th ${thS}>Client</th><th ${thR}>Import</th><th ${thS}>Situació</th></tr>`;
      items.forEach((l,i)=>{h+=`<tr${rowBg(i)}><td ${tdS}>${l.codi_obra||"—"}</td><td ${tdS}>${l.licitacio||"—"}</td><td ${tdS}>${l.client||"—"}</td><td ${tdR}>${fmtE(l.import_pec_sense_iva)}</td><td ${tdS}>${l.comentaris?.slice(0,80)||l.data_presentacio||"Sense novetats"}</td></tr>`;});
      h+=`<tr style="background-color:#eef2f7;"><td colspan="3" ${tdS} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;font-weight:bold;">Subtotal</td><td ${tdR} style="border:1px solid #b5b5b5;padding:6px 8px;text-align:right;white-space:nowrap;font-weight:bold;">${fmtE(subtotal(items))}</td><td ${tdS}></td></tr></table>`;
      if(items.some(l=>!l.import_pec_sense_iva))h+=`<p style="font-size:9.5pt;color:#666;font-style:italic;margin:0 0 14px 0;">Nota: els registres sense import encara no tenen PEC assignat al quadre.</p>`;
      return h;
    };
    // Quadre resum
    const seccions=[
      ["Presentades",presentades.length,subtotal(presentades)],
      ["No presentades",noPresentades.length,subtotal(noPresentades)],
      ["Propostes — client públic",propostesPub.length,subtotal(propostesPub)],
      ["Propostes — clients privats",propostesPriv.length,subtotal(propostesPriv)],
      ["En estudi — clients públics",estudiPub.length,subtotal(estudiPub)],
      ["En estudi — clients privats",estudiPriv.length,subtotal(estudiPriv)],
    ];
    if(adjudicades.length)seccions.push(["Adjudicades",adjudicades.length,subtotal(adjudicades)]);
    if(noAdjudicades.length)seccions.push(["No adjudicades",noAdjudicades.length,subtotal(noAdjudicades)]);
    const totalN=seccions.reduce((s,r)=>s+r[1],0);
    const totalI=seccions.reduce((s,r)=>s+r[2],0);
    let quadre=`<p style="font-family:Arial,Helvetica,sans-serif;font-size:12pt;font-weight:bold;color:#1a3d6e;margin:24px 0 8px 0;border-bottom:2px solid #1a3d6e;padding-bottom:4px;">Quadre resum</p>`;
    quadre+=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;margin:6px 0 14px 0;border:1px solid #b5b5b5;"><tr><th ${thS} style="background-color:#1a3d6e;color:#fff;font-weight:bold;border:1px solid #b5b5b5;padding:6px 8px;text-align:left;width:40px;">#</th><th ${thS}>Apartat</th><th ${thR} style="background-color:#1a3d6e;color:#fff;font-weight:bold;border:1px solid #b5b5b5;padding:6px 8px;text-align:right;width:70px;">Nº</th><th ${thR} style="background-color:#1a3d6e;color:#fff;font-weight:bold;border:1px solid #b5b5b5;padding:6px 8px;text-align:right;width:180px;">Import (sense IVA)</th></tr>`;
    seccions.forEach((s,i)=>{quadre+=`<tr${rowBg(i)}><td ${tdS}>${i+1}</td><td ${tdS}>${s[0]}</td><td ${tdR}>${s[1]}</td><td ${tdR}>${fmtE(s[2])}</td></tr>`;});
    quadre+=`<tr style="background-color:#1a3d6e;"><td colspan="2" style="border:1px solid #b5b5b5;padding:6px 8px;color:#fff;font-weight:bold;">TOTAL</td><td style="border:1px solid #b5b5b5;padding:6px 8px;color:#fff;font-weight:bold;text-align:right;">${totalN}</td><td style="border:1px solid #b5b5b5;padding:6px 8px;color:#fff;font-weight:bold;text-align:right;white-space:nowrap;">${fmtE(totalI)}</td></tr></table>`;
    // Construir HTML
    let body=`<!DOCTYPE html><html lang="ca"><head><meta charset="UTF-8"><title>Moviments licitacions 2026</title></head><body style="font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:#222;line-height:1.45;max-width:1100px;margin:24px auto;padding:0 16px;">`;
    body+=`<p style="margin:8px 0;">Bon dia,</p>`;
    body+=`<p style="margin:8px 0;">Us faig arribar els moviments de les licitacions d'edificació de l'última setmana. En primer terme, teniu un quadre resum de la setmana.</p>`;
    // SECCIÓ MOVIMENTS DELS ÚLTIMS 7 DIES
    if(hiHaMoviments){
      const tblMov=(items,cols=["Codi","Licitació","Client","Import"])=>{
        if(!items.length)return"";
        let h=`<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:10pt;margin:4px 0 12px 0;border:1px solid #b5b5b5;"><tr>${cols.map((c,i)=>`<th ${i===3?thR:thS}>${c}</th>`).join("")}</tr>`;
        items.forEach((l,i)=>{h+=`<tr${rowBg(i)}><td ${tdS}>${l.codi_obra||"—"}</td><td ${tdS}>${l.licitacio||"—"}</td><td ${tdS}>${l.client||"—"}</td><td ${tdR}>${fmtE(l.import_pec_sense_iva)}</td></tr>`;});
        h+=`</table>`;return h;
      };
      const fa7Str=`${String(fa7.getDate()).padStart(2,"0")}/${String(fa7.getMonth()+1).padStart(2,"0")}/${fa7.getFullYear()}`;
      body+=`<div style="background:#eef5ff;border:2px solid #1a3d6e;border-radius:8px;padding:14px 16px;margin:18px 0 22px 0;">`;
      body+=`<p style="font-size:13pt;font-weight:bold;color:#1a3d6e;margin:0 0 10px 0;">🔄 MOVIMENTS D'AQUESTA SETMANA (DES DEL ${fa7Str} FINS AVUI)</p>`;
      const sec=(emoji,title,items)=>{if(!items.length)return"";return`<p style="font-size:11pt;font-weight:bold;color:#1a3d6e;margin:12px 0 4px 0;">${emoji} ${title} (${items.length})</p>${tblMov(items)}`;};
      body+=sec("📥","Noves PROPOSTES (pendents de decisió de gerència)",movNovesPropostes);
      body+=sec("📘","Noves licitacions directament a EN ESTUDI",movNovesEnEstudi);
      if(movNovesAltres.length)body+=sec("📄","Noves licitacions en altres estats",movNovesAltres);
      body+=sec("📋","Passen de Proposta a En Estudi",movPassenEstudi);
      body+=sec("⏳","Pendents d'obertura (presentades sense resoldre)",presentades);
      body+=sec("✅","Passen a Presentada",movPresentades);
      body+=sec("❌","Passen a No Presentada",movNoPresentades);
      body+=sec("🏆","Adjudicades",movAdjudicades);
      body+=sec("⚠️","No Adjudicades",movNoAdjudicades);
      body+=sec("🚫","Descartades",movDescartades);
      body+=`</div>`;
    } else {
      body+=`<p style="background:#fff3cd;border:1px solid #ffeeba;padding:8px 12px;color:#856404;font-style:italic;border-radius:6px;">ℹ️ No s'han registrat moviments en els últims 7 dies.</p>`;
    }
    body+=`<p style="font-size:13pt;font-weight:bold;color:#1a3d6e;margin:24px 0 8px 0;">📊 ESTAT ACTUAL DE TOTES LES LICITACIONS</p>`;
    body+=secTitle(1,"Presentades",presentades.length,subtotal(presentades));
    body+=tblPresentades();
    body+=secTitle(2,"No presentades",noPresentades.length,subtotal(noPresentades));
    body+=tblNoPresentades();
    body+=secTitle(3,"Propostes — client públic",propostesPub.length,subtotal(propostesPub));
    body+=tblPropostes(propostesPub);
    body+=secTitle(4,"Propostes — clients privats",propostesPriv.length,subtotal(propostesPriv));
    body+=tblPropostes(propostesPriv);
    body+=secTitle(5,"En estudi — clients públics",estudiPub.length,subtotal(estudiPub));
    body+=tblEstudi(estudiPub);
    body+=secTitle(6,"En estudi — clients privats",estudiPriv.length,subtotal(estudiPriv));
    body+=tblEstudi(estudiPriv);
    body+=quadre;
    body+=`<p style="margin:8px 0;">Resto a disposició per revisar prioritats o preparar els propers terminis.</p>`;
    body+=`<p style="margin:8px 0;">Salutacions,</p>`;
    body+=`<div style="margin-top:20px;border-top:1px solid #ddd;padding-top:10px;"><button onclick="document.execCommand('selectAll');document.execCommand('copy');alert('Copiat!')" style="display:inline-block;margin:10px 5px;padding:8px 16px;background:#1a3d6e;color:white;border:none;border-radius:6px;cursor:pointer;font-size:13px;">📋 Copiar tot</button><button onclick="window.print()" style="display:inline-block;margin:10px 5px;padding:8px 16px;background:#1a3d6e;color:white;border:none;border-radius:6px;cursor:pointer;font-size:13px;">🖨 Imprimir / PDF</button></div>`;
    body+=`</body></html>`;
    win.document.open();win.document.write(body);win.document.close();win.focus();
    }catch(e){console.error("Error generant correu:",e);win.document.open();win.document.write("<html><body><p>Error: "+e.message+"</p></body></html>");win.document.close();}
  };
  const importBackup=(e)=>{
    const file=e.target.files?.[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      try{
        const data=JSON.parse(ev.target.result);
        if(!data.licitacions||!Array.isArray(data.licitacions))throw new Error("Format invàlid");
        const merged=[...data.licitacions];
        // Merge: add existing items not in backup by id
        lic.forEach(l=>{if(!merged.find(m=>m.id===l.id))merged.push(l);});
        save(merged);
        await saveAll(merged);
        if(data.tipus?.length>0)await saveTipus(data.tipus);
        setBackupMsg(`✅ Importat: ${data.licitacions.length} licitacions (${data.date?.slice(0,10)||"?"})`);
        setTimeout(()=>setBackupMsg(""),4000);
      }catch(err){setBackupMsg("❌ Error: "+err.message);setTimeout(()=>setBackupMsg(""),4000);}
    };
    reader.readAsText(file);
    e.target.value="";
  };

  useEffect(()=>{
    (async()=>{
      if(!isSupabaseConfigured()){
        // Fallback to localStorage when Supabase is not configured
        try{const t=localStorage.getItem(SK_TIPUS);if(t){const arr=JSON.parse(t);if(arr.length>0)setTipus(arr);}}catch(e){}
        try{const r=localStorage.getItem(SK);if(r){const data=JSON.parse(r);setLic(data);}else setLic(SEED);}catch(e){setLic(SEED);}
        setLoaded(true);return;
      }
      try{
        const{data:tipusData}=await supabase.from("tipus").select("nom");
        if(tipusData?.length>0)setTipus(tipusData.map(t=>t.nom));
      }catch(e){}
      try{
        const{data,error}=await supabase.from("licitacions").select("*");
        if(error)throw error;
        if(data&&data.length>0){setLic(data);}
        else{
          const cleanSeed=SEED.map(s=>({...s,import_pec_sense_iva:s.import_pec_sense_iva===""||s.import_pec_sense_iva==null?0:Number(s.import_pec_sense_iva)||0}));
          const{error:seedErr}=await supabase.from("licitacions").upsert(cleanSeed);
          if(seedErr)console.error("Seed error:",seedErr);
          setLic(SEED);
        }
      }catch(e){console.error("Load error:",e);setLic(SEED);}
      setLoaded(true);
    })();
    if(!isSupabaseConfigured())return;
    const channel=supabase.channel("licitacions-changes")
      .on("postgres_changes",{event:"*",schema:"public",table:"licitacions"},payload=>{
        if(payload.eventType==="INSERT")setLic(prev=>[payload.new,...prev.filter(l=>l.id!==payload.new.id)]);
        else if(payload.eventType==="UPDATE")setLic(prev=>prev.map(l=>l.id===payload.new.id?payload.new:l));
        else if(payload.eventType==="DELETE")setLic(prev=>prev.filter(l=>l.id!==payload.old.id));
      })
      .subscribe();
    return()=>{supabase.removeChannel(channel);};
  },[]);

  useEffect(()=>{
    if(!loaded)return;
    if(!isSupabaseConfigured())return;
    (async()=>{
      try{const{data}=await supabase.from("licitacions").select("*");if(data)setLic(data);}catch(e){}
    })();
  },[refreshKey,loaded]);

  const save=async list=>{setLic(list);if(!isSupabaseConfigured())try{localStorage.setItem(SK,JSON.stringify(list));}catch(e){}};
  const cleanForDB=item=>{const c={...item};if(c.import_pec_sense_iva===""||c.import_pec_sense_iva==null)c.import_pec_sense_iva=0;else c.import_pec_sense_iva=Number(c.import_pec_sense_iva)||0;return c;};
  const saveOne=async item=>{if(!isSupabaseConfigured())return;try{await supabase.from("licitacions").upsert(cleanForDB(item));}catch(e){console.error("Save error:",e);}};
  const logCanviEstat=async(licitacioId,estatAnterior,estatNou,codiObra)=>{try{await supabase.from("licitacions_log").insert({licitacio_id:licitacioId,estat_anterior:estatAnterior||"",estat_nou:estatNou,codi_obra:codiObra||""});}catch(e){console.error("Log error:",e);}};
  const saveAll=async list=>{if(!isSupabaseConfigured())return;try{await supabase.from("licitacions").upsert(list.map(cleanForDB));}catch(e){console.error("SaveAll error:",e);}};
  const saveTipus=async list=>{setTipus(list);if(!isSupabaseConfigured()){try{localStorage.setItem(SK_TIPUS,JSON.stringify(list));}catch(e){}return;}try{await supabase.from("tipus").delete().neq("id",0);await supabase.from("tipus").insert(list.map(t=>({nom:t})));}catch(e){console.error("SaveTipus error:",e);}};
  const ESTATS_ALLIBERA_CODI=["DESCARTADA"];
  const genCode=list=>{
    const now=new Date(),yy=String(now.getFullYear()).slice(-2),mm=String(now.getMonth()+1).padStart(2,"0");
    const prefixAny=`${yy}.`;// comptar per ANY complet (correlatiu anual)
    const usats=new Set();
    list.forEach(l=>{
      if(!l.codi_obra)return;
      // Codis NP s'alliberen — NO compten com a usats
      if(/^NP\d*-/.test(l.codi_obra))return;
      if(l.codi_obra.startsWith(prefixAny)){const n=codiNum(l.codi_obra);if(n<9999)usats.add(n);}
    });
    let num=1;
    while(usats.has(num))num++;
    return`${yy}.${mm}.${String(num).padStart(3,"0")}-ED`;
  };
  const onEstat=(id,estat)=>{
    const l=lic.find(x=>x.id===id);if(!l)return;
    // Passar a EN ESTUDI sense codi → assignar codi
    if(estat==="EN ESTUDI"&&!l.codi_obra){
      setConfModal({id,newEstat:estat,code:genCode(lic)});return;
    }
    // NO PRESENTADA → conserva el codi amb prefix NPx-, allibera el número per reutilitzar
    if(estat==="NO PRESENTADA"&&l.codi_obra&&!/^NP\d+-/.test(l.codi_obra)){
      // Trobar el pròxim número NP
      let maxNP=0;
      lic.forEach(x=>{const m=(x.codi_obra||"").match(/^NP(\d+)-/);if(m)maxNP=Math.max(maxNP,parseInt(m[1]));});
      const npCodi=`NP${maxNP+1}-${l.codi_obra}`;
      const updated=lic.map(x=>x.id===id?{...x,estat,codi_obra:npCodi}:x);
      save(updated);saveOne({...l,estat,codi_obra:npCodi});logCanviEstat(id,l.estat,estat,npCodi);return;
    }
    // DESCARTADA → alliberar el codi completament
    if(ESTATS_ALLIBERA_CODI.includes(estat)&&l.codi_obra){
      const updated=lic.map(x=>x.id===id?{...x,estat,codi_obra:""}:x);
      save(updated);saveOne({...l,estat,codi_obra:""});logCanviEstat(id,l.estat,estat,l.codi_obra);return;
    }
    const updated=lic.map(x=>x.id===id?{...x,estat}:x);
    save(updated);saveOne(updated.find(x=>x.id===id));logCanviEstat(id,l.estat,estat,l.codi_obra);
  };
  const confirmEstat=()=>{
    const codiExistent=lic.find(l=>l.codi_obra===confModal.code&&l.id!==confModal.id);
    if(codiExistent){alert(`⚠️ El codi ${confModal.code} ja està assignat a "${codiExistent.licitacio}". Tria un altre codi.`);return;}
    const lOld=lic.find(l=>l.id===confModal.id);
    const item={...lOld,estat:confModal.newEstat,codi_obra:confModal.code};
    save(lic.map(l=>l.id===confModal.id?item:l));saveOne(item);logCanviEstat(confModal.id,lOld?.estat||"",confModal.newEstat,confModal.code);setConfModal(null);
  };
  const findDups=(c,cid)=>lic.filter(l=>{if(l.id===cid)return false;if(c.codi_obra&&l.codi_obra&&norm(c.codi_obra)===norm(l.codi_obra))return true;const s=sim(c.licitacio,l.licitacio);if(s>=0.75&&norm(c.client)===norm(l.client))return true;if(s>=0.85)return true;return false;});
  const getPairs=()=>{const pairs=[],seen=new Set();for(let i=0;i<lic.length;i++)for(let j=i+1;j<lic.length;j++){const a=lic[i],b=lic[j],key=`${a.id}-${b.id}`;if(seen.has(key))continue;const s=sim(a.licitacio,b.licitacio),cm=a.codi_obra&&b.codi_obra&&norm(a.codi_obra)===norm(b.codi_obra);if(cm||(s>=0.6&&norm(a.client)===norm(b.client))||s>=0.75){pairs.push({a,b,s:Math.round(s*100)});seen.add(key);}}return pairs;};
  const doSave=(f,id)=>{const item=id!==null?{...f,id}:{...f,id:Date.now()};const list=id!==null?lic.map(l=>l.id===id?item:l):[item,...lic];save(list);saveOne(item);if(id===null)logCanviEstat(item.id,"",item.estat||"PROPOSTA",item.codi_obra);setView("table");setEditId(null);setForm(emptyForm);setAnalisi("");setDupModal(null);};
  const onSubmit=()=>{if(!form.licitacio&&!form.client)return;const dups=findDups(form,editId);if(dups.length>0){setDupModal({dups});return;}doSave(form,editId);};
  const onEdit=l=>{setForm({...l});setEditId(l.id);setAnalisi(l.analisi_completa||"");setShowAnalisi(false);setView("form");};
  const onDel=id=>{if(!confirm("Eliminar aquesta licitacio?"))return;save(lic.filter(l=>l.id!==id));supabase.from("licitacions").delete().eq("id",id).then(()=>{}).catch(e=>console.error("Del error:",e));};
  const imprimirPDF=(data,noDesc)=>{const html=makeTableHTML(data,noDesc);const win=window.open("","_blank");if(!win){alert("Activa les finestres emergents.");return;}win.document.write(html);win.document.close();win.focus();setTimeout(()=>win.print(),500);};
  const exportExcel=data=>{
    const cols=["CODI OBRA","LICITACIO","CLIENT","PRIVAT PUBLIC","POBLACIO","ESTAT","DATA PRESENTACIO","TERMINI","IMPORTE PEC Sin IVA","CLASSIF","CRITERIS PUNTUACIO","AVAL","APERTURA","COMENTARIOS","LINK OBRA"];
    const flds=["codi_obra","licitacio","client","public_privat","poblacio","estat","data_presentacio","termini","import_pec_sense_iva","classificacio","criteris_puntuacio","aval","apertura","comentaris","link_obra","link_publicacio"];
    const IC=8;
    const wsData=[cols,...data.map(l=>flds.map((f,fi)=>{if(fi===IC){const n=parseFloat(l[f]);return isNaN(n)?"":n;}return l[f]!=null?l[f]:""; }))];
    const ws=XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"]=[10,44,26,12,18,14,20,12,18,12,44,14,12,48,36].map(wch=>({wch}));
    const hs={font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{fgColor:{rgb:"1F3864"}},alignment:{horizontal:"center",vertical:"center",wrapText:true}};
    cols.forEach((_,ci)=>{const ref=XLSX.utils.encode_cell({r:0,c:ci});if(ws[ref])ws[ref].s=hs;});
    data.forEach((l,ri)=>{const fill=XBG[l.estat]||"FFFFFF";flds.forEach((f,ci)=>{const ref=XLSX.utils.encode_cell({r:ri+1,c:ci});if(!ws[ref])ws[ref]={t:"z",v:""};ws[ref].s={fill:{fgColor:{rgb:fill}},font:{sz:9},alignment:{horizontal:ci===IC?"right":"left",vertical:"top",wrapText:true}};});});
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Licitacions 2026");
    const b64=XLSX.write(wb,{bookType:"xlsx",type:"base64"});
    const a=document.createElement("a");a.href=`data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`;a.download="Licitacions_Servial_2026.xlsx";document.body.appendChild(a);a.click();document.body.removeChild(a);
    setShowExport(false);
  };
  const filtered=lic.filter(l=>{
    const q=search.toLowerCase();const ok=!q||[l.codi_obra,l.licitacio,l.client,l.poblacio].some(v=>v&&v.toLowerCase().includes(q));
    const estatOk=fEstat===""?true:fEstat==="ACTIVES"?ACTIVES.includes(l.estat):fEstat==="SENSE_DESC"?SENSE_DESC.includes(l.estat):l.estat===fEstat;
    if(!ok||!estatOk||fTipus&&l.public_privat!==fTipus)return false;
    // Filtres de columna (tipus Excel)
    const cf=colFilters;
    if(cf.estat?.length>0&&!cf.estat.includes(l.estat))return false;
    if(cf.public_privat?.length>0&&!cf.public_privat.includes(l.public_privat))return false;
    if(cf.poblacio?.length>0&&!cf.poblacio.includes(l.poblacio||"---"))return false;
    if(cf.client?.length>0&&!cf.client.includes(l.client||"---"))return false;
    if(cf.codi_obra?.length>0&&!cf.codi_obra.includes(l.codi_obra||"---"))return false;
    if(cf.data_presentacio?.length>0&&!cf.data_presentacio.includes(l.data_presentacio||"---"))return false;
    if(cf.termini?.length>0&&!cf.termini.includes(l.termini||"---"))return false;
    if(cf.import_pec?.length>0&&!cf.import_pec.includes(l.import_pec_sense_iva?fmt(l.import_pec_sense_iva):"---"))return false;
    if(cf.classificacio?.length>0&&!cf.classificacio.includes(l.classificacio||"---"))return false;
    if(cf.tecnica?.length>0){const val=l.tecnica!=null?l.tecnica:(l.criteris_puntuacio&&/100\s*%?\s*(auto|sobre)/i.test(l.criteris_puntuacio)?false:true);const tStr=val?"Sí":"No";if(!cf.tecnica.includes(tStr))return false;}
    return true;
  }).sort((a,b)=>{const da=parseD(a.data_presentacio),db=parseD(b.data_presentacio);if(da&&db)return da-db;if(da&&!db)return -1;if(!da&&db)return 1;const eo=eOrder(a.estat)-eOrder(b.estat);if(eo!==0)return eo;return codiNum(a.codi_obra)-codiNum(b.codi_obra);});
  const stats=(()=>{const pres=lic.filter(l=>l.estat==="PRESENTADA").length,adj=lic.filter(l=>l.estat==="ADJUDICADA").length,noAdj=lic.filter(l=>l.estat==="NO ADJUDICADA").length;return{total:lic.length,propostes:lic.filter(l=>l.estat==="PROPOSTA").length,estudi:lic.filter(l=>l.estat==="EN ESTUDI").length,presentades:pres+adj+noAdj,pendentObertura:pres,noPresentades:lic.filter(l=>l.estat==="NO PRESENTADA").length,adjudicades:adj,noAdjudicades:noAdj,descartades:lic.filter(l=>l.estat==="DESCARTADA").length};})();
  const dupPairs=showDup?getPairs():[];
  const inp="w-full border rounded-lg px-2 py-1.5 text-sm focus:outline-none focus:ring-1 focus:ring-blue-400";
  if(!loaded)return<div className="flex items-center justify-center h-64 text-gray-400 text-sm">Carregant...</div>;
  if(view==="form")return(
    <div className="max-w-2xl mx-auto p-4">
      <div className="flex items-center gap-2 mb-4"><button onClick={()=>{setView("table");setEditId(null);setForm(emptyForm);}} className="text-blue-600 hover:underline text-sm">← Tornar</button><h2 className="text-lg font-bold text-gray-800">{editId?"Editar licitacio":"Nova licitacio"}</h2></div>
      {analisi&&<div className="mb-4 border border-blue-200 rounded-xl overflow-hidden"><button onClick={()=>setShowAnalisi(!showAnalisi)} className="w-full flex items-center justify-between px-4 py-2.5 bg-blue-50 text-blue-800 font-semibold text-sm hover:bg-blue-100"><span>Informe d'analisi complet</span><span>{showAnalisi?"▲":"▼"}</span></button>{showAnalisi&&<div className="px-4 py-3 text-xs text-gray-700 whitespace-pre-wrap leading-relaxed max-h-80 overflow-y-auto bg-stone-50">{analisi}</div>}</div>}
      {dupModal&&<div className="mb-4 bg-orange-50 border-2 border-orange-300 rounded-xl p-4"><p className="font-bold text-orange-800 text-sm mb-2">Possible duplicat detectat:</p>{dupModal.dups.map(d=>(<div key={d.id} className="bg-stone-50 border border-orange-200 rounded-lg px-3 py-2 text-xs mb-1"><div className="font-semibold text-gray-800">{d.licitacio||"---"}</div><div className="text-gray-500">{d.client||"---"}{d.codi_obra?" - "+d.codi_obra:""} - {d.estat}</div></div>))}<div className="flex gap-2 mt-3"><button onClick={()=>setDupModal(null)} className="flex-1 px-3 py-2 border rounded-lg text-sm text-gray-600">Cancel</button><button onClick={()=>doSave(form,editId)} className="flex-1 px-3 py-2 bg-orange-500 hover:bg-orange-600 text-white font-bold rounded-lg text-sm">Guardar igualment</button></div></div>}
      <div className="bg-stone-50 border border-gray-200 rounded-xl p-4 shadow-sm">
        <div className="grid grid-cols-2 gap-3">
          <FR label="Codi Obra"><div className="flex gap-1"><input className={inp+(form.estat==="PROPOSTA"||form.estat==="NO PROPOSTA"?" bg-gray-100 text-gray-400 cursor-not-allowed":"")} value={form.estat==="PROPOSTA"||form.estat==="NO PROPOSTA"?"":form.codi_obra} placeholder={form.estat==="PROPOSTA"||form.estat==="NO PROPOSTA"?"S'assigna en passar a EN ESTUDI":"26.03.XXX-ED"} disabled={form.estat==="PROPOSTA"||form.estat==="NO PROPOSTA"} onChange={e=>setForm({...form,codi_obra:e.target.value})} />{form.estat!=="PROPOSTA"&&form.estat!=="NO PROPOSTA"&&!form.codi_obra&&<button type="button" onClick={()=>setForm({...form,codi_obra:genCode(lic)})} className="whitespace-nowrap px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg">+ Auto</button>}</div></FR>
          <FR label="Privat / Public"><select className={inp} value={form.public_privat} onChange={e=>setForm({...form,public_privat:e.target.value})}>{tipus.map(t=><option key={t}>{t}</option>)}</select></FR>
          <FR label="Licitacio" span2><textarea rows={2} className={inp+" resize-none"} value={form.licitacio} onChange={e=>setForm({...form,licitacio:e.target.value})} /></FR>
          <FR label="Client"><input className={inp} value={form.client} onChange={e=>setForm({...form,client:e.target.value})} /></FR>
          <FR label="Poblacio"><input className={inp} value={form.poblacio} onChange={e=>setForm({...form,poblacio:e.target.value})} /></FR>
          <FR label="Estat"><select className={inp} value={form.estat} onChange={e=>setForm({...form,estat:e.target.value})}>{ESTATS.map(s=><option key={s}>{s}</option>)}</select></FR>
          <FR label="Data Presentació"><div className="flex gap-2">
            <input type="datetime-local" className={inp+" flex-1"} value={(()=>{const v=form.data_presentacio||"";const m=v.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):?(\d{2})?)?/);if(m)return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}T${(m[4]||"00").padStart(2,"0")}:${(m[5]||"00").padStart(2,"0")}`;return"";})()}
              onChange={e=>{const d=e.target.value;if(!d)return;const dt=new Date(d);const formatted=`${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()} ${String(dt.getHours()).padStart(2,"0")}:${String(dt.getMinutes()).padStart(2,"0")}`;setForm({...form,data_presentacio:formatted});}} />
            <input className={inp+" flex-1"} value={form.data_presentacio} placeholder="Text lliure o DD/MM/YYYY HH:MM" onChange={e=>setForm({...form,data_presentacio:e.target.value})} />
          </div></FR>
          <FR label="Termini Execucio"><input className={inp} value={form.termini} placeholder="p.ex. 8 mesos" onChange={e=>setForm({...form,termini:e.target.value})} /></FR>
          <FR label="Import PEC s/IVA"><input type="number" className={inp} value={form.import_pec_sense_iva} onChange={e=>setForm({...form,import_pec_sense_iva:e.target.value})} /></FR>
          <FR label="Classificacio"><input className={inp} value={form.classificacio} placeholder="p.ex. C-2-3" onChange={e=>setForm({...form,classificacio:e.target.value})} /></FR>
          <FR label="Criteris Puntuacio" span2><input className={inp} value={form.criteris_puntuacio} onChange={e=>setForm({...form,criteris_puntuacio:e.target.value})} /></FR>
          <FR label="Aval"><input className={inp} value={form.aval} onChange={e=>setForm({...form,aval:e.target.value})} /></FR>
          <FR label="Apertura"><input className={inp} value={form.apertura} onChange={e=>setForm({...form,apertura:e.target.value})} /></FR>
          <FR label="Comentaris" span2><textarea rows={2} className={inp+" resize-none"} value={form.comentaris} onChange={e=>setForm({...form,comentaris:e.target.value})} /></FR>
          <FR label="Link Obra" span2><input className={inp} value={form.link_obra} placeholder="https://..." onChange={e=>setForm({...form,link_obra:e.target.value})} /></FR>
          <FR label="Link Publicació" span2><input className={inp} value={form.link_publicacio||""} placeholder="https://..." onChange={e=>setForm({...form,link_publicacio:e.target.value})} /></FR>
        </div>
        <div className="flex gap-2 mt-4"><button onClick={onSubmit} className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-xl text-sm">{editId?"Guardar canvis":"Afegir licitacio"}</button><button onClick={()=>{setView("table");setEditId(null);setForm(emptyForm);setDupModal(null);}} className="px-4 py-2 border rounded-xl text-sm text-gray-600 hover:bg-gray-50">Cancel</button></div>
      </div>
    </div>
  );
  return(
    <div className="p-3 max-w-full" onClick={()=>setShowExport(false)}>
      {showTipusSettings&&(<div className="fixed inset-0 bg-black bg-opacity-40 flex items-center justify-center z-50"><div className="bg-stone-50 rounded-2xl shadow-2xl p-6 max-w-sm w-full mx-4"><div className="flex items-center justify-between mb-4"><h3 className="text-base font-bold text-gray-900">Tipus de licitació</h3><button onClick={()=>{setShowTipusSettings(false);setNewTipus("");}} className="text-gray-400 hover:text-gray-600 text-2xl">×</button></div><div className="space-y-2 mb-4">{tipus.map((t,i)=>(<div key={t} className="flex items-center gap-2"><span className={"flex-1 px-3 py-1.5 rounded-lg text-sm font-medium border "+(i===0?"bg-blue-50 border-blue-200 text-blue-700":"bg-orange-50 border-orange-200 text-orange-700")}>{t}</span><button disabled={tipus.length<=1} onClick={()=>saveTipus(tipus.filter((_,j)=>j!==i))} className="text-red-400 hover:text-red-600 disabled:opacity-30 text-lg px-1">🗑️</button></div>))}</div><div className="flex gap-2"><input className="flex-1 border rounded-lg px-3 py-1.5 text-sm" placeholder="Nou tipus" value={newTipus} onChange={e=>setNewTipus(e.target.value.toUpperCase())} onKeyDown={e=>{if(e.key==="Enter"&&newTipus.trim()&&!tipus.includes(newTipus.trim())){saveTipus([...tipus,newTipus.trim()]);setNewTipus("");}}}/><button onClick={()=>{if(newTipus.trim()&&!tipus.includes(newTipus.trim())){saveTipus([...tipus,newTipus.trim()]);setNewTipus("");}}} disabled={!newTipus.trim()||tipus.includes(newTipus.trim())} className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 disabled:opacity-40 text-white font-bold rounded-lg text-sm">Afegir</button></div></div></div>)}
      {showDup&&(<div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4"><div className="bg-stone-50 rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col" style={{maxHeight:"90vh"}}><div className="flex items-center justify-between px-5 py-4 border-b"><h3 className="font-bold text-gray-900">Gestio de duplicats</h3><button onClick={()=>{setShowDup(false);setDelId(null);}} className="text-gray-400 hover:text-gray-600 text-2xl">×</button></div><div className="overflow-y-auto flex-1 p-4 space-y-4">{dupPairs.length===0?<p className="text-center text-gray-400 py-8">No s'han detectat duplicats.</p>:dupPairs.map((pair,pi)=>(<div key={pi} className="border border-orange-200 rounded-xl overflow-hidden"><div className="bg-orange-50 px-3 py-1.5 text-xs font-semibold text-orange-700">Similitud: {pair.s}%</div><div className="grid grid-cols-2 divide-x divide-gray-200">{[pair.a,pair.b].map(l=>(<div key={l.id} className="p-3"><div className="text-xs space-y-1 mb-3"><div className="font-semibold text-gray-800">{l.licitacio||"---"}</div><div className="text-gray-500">{l.client||"---"}</div>{l.codi_obra&&<div className="font-mono text-blue-600">{l.codi_obra}</div>}<span className={"inline-block px-1.5 py-0.5 rounded-full text-xs font-semibold "+(EC[l.estat]||"bg-gray-100 text-gray-600")}>{l.estat}</span></div>{delId===l.id?<div className="flex gap-1"><button className="flex-1 py-1.5 bg-gray-200 text-gray-700 text-xs font-bold rounded-lg" onClick={()=>setDelId(null)}>Cancel</button><button className="flex-1 py-1.5 bg-red-600 text-white text-xs font-bold rounded-lg" onClick={()=>{const id=l.id;setLic(cur=>cur.filter(x=>x.id!==id));supabase.from("licitacions").delete().eq("id",id).catch(()=>{});setDelId(null);setShowDup(false);}}>Confirmar</button></div>:<button className="w-full py-1.5 bg-red-500 hover:bg-red-600 text-white text-xs font-bold rounded-lg" onClick={()=>setDelId(l.id)}>Eliminar aquest</button>}</div>))}</div></div>))}</div><div className="px-5 py-3 border-t"><button onClick={()=>{setShowDup(false);setDelId(null);}} className="w-full py-2 border rounded-xl text-sm text-gray-600 hover:bg-gray-50">Tancar</button></div></div></div>)}
      {confModal&&(<div className="fixed inset-0 bg-black bg-opacity-40 flex items-center justify-center z-50"><div className="bg-stone-50 rounded-2xl shadow-2xl p-6 max-w-sm w-full mx-4"><div className="text-2xl mb-2 text-center">📋</div><h3 className="text-base font-bold text-gray-900 text-center mb-1">Assignar Codi d'Obra</h3><p className="text-sm text-gray-500 text-center mb-4">En passar a EN ESTUDI s'assignara el codi:</p><div className="bg-gray-50 border-2 border-blue-200 rounded-xl py-3 text-center mb-3"><span className="text-2xl font-mono font-bold text-blue-700">{confModal.code}</span></div><input className="w-full border rounded-lg px-3 py-1.5 text-sm font-mono text-center mb-4" value={confModal.code} onChange={e=>setConfModal({...confModal,code:e.target.value})} /><div className="flex gap-2"><button onClick={()=>setConfModal(null)} className="flex-1 px-4 py-2.5 border rounded-xl text-sm text-gray-600">Cancel</button><button onClick={confirmEstat} className="flex-1 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl text-sm">Confirmar</button></div></div></div>)}
      <div className="flex items-center justify-between mb-3">
        <div><h2 className="text-xl font-bold text-gray-900">SERVIAL. Licitacions d'edificació. Any 2026</h2></div>
        <div className="flex gap-2 flex-wrap items-center" onClick={e=>e.stopPropagation()}>
          <button onClick={()=>{setForm(emptyForm);setEditId(null);setAnalisi("");setView("form");}} className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-3 py-2 rounded-lg">+ Manual</button>
          <button onClick={()=>setShowTipusSettings(true)} className="bg-gray-500 hover:bg-gray-600 text-white text-sm font-semibold px-3 py-2 rounded-lg">⚙️</button>
          <button onClick={()=>setShowDup(true)} className="bg-orange-500 hover:bg-orange-600 text-white text-sm font-semibold px-3 py-2 rounded-lg">🧹</button>
          <button onClick={()=>imprimirPDF(filtered,fEstat==="SENSE_DESC")} className="bg-gray-600 hover:bg-gray-700 text-white text-sm font-semibold px-3 py-2 rounded-lg">🖨 PDF</button>
          <button onClick={()=>{navigator.clipboard.writeText("Servial7524A@").then(()=>alert("🔑 Clau copiada!")).catch(()=>{const t=document.createElement("textarea");t.value="Servial7524A@";document.body.appendChild(t);t.select();document.execCommand("copy");document.body.removeChild(t);alert("🔑 Clau copiada!");});}} className="bg-gray-800 hover:bg-gray-900 text-white text-sm font-semibold px-3 py-2 rounded-lg">🔑 Clau SERVIAL</button>
          <button onClick={generarCorreuSetmanal} className="bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold px-3 py-2 rounded-lg">📧 Correu</button>
          <button onClick={exportBackup} className="bg-teal-600 hover:bg-teal-700 text-white text-sm font-semibold px-3 py-2 rounded-lg">💾 Backup</button>
          <button onClick={()=>importRef.current?.click()} className="bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold px-3 py-2 rounded-lg">📂 Restaurar</button>
          <input ref={importRef} type="file" accept=".json" className="hidden" onChange={importBackup}/>
          <div className="relative" onClick={e=>e.stopPropagation()}>
            <button onClick={()=>setShowExport(!showExport)} className="bg-gray-700 hover:bg-gray-800 text-white text-sm font-semibold px-3 py-2 rounded-lg">Excel ▾</button>
            {showExport&&<div className="absolute right-0 mt-1 w-52 bg-stone-50 border border-gray-200 rounded-xl shadow-xl z-20"><button onClick={()=>exportExcel(filtered)} className="w-full text-left px-4 py-3 text-sm text-gray-700 hover:bg-gray-50 rounded-t-xl">Vista actual ({filtered.length} files)</button><div className="border-t border-gray-100"/><button onClick={()=>exportExcel(lic)} className="w-full text-left px-4 py-3 text-sm text-gray-700 hover:bg-gray-50 rounded-b-xl">Totes ({lic.length} files)</button></div>}
          </div>
        </div>
      </div>
      {backupMsg&&<div className="mb-3 px-3 py-2 bg-teal-50 border border-teal-200 rounded-lg text-sm text-teal-700 font-medium">{backupMsg}</div>}
      <div className="grid grid-cols-9 gap-2 mb-3">{[["Total",stats.total,"bg-gray-50 border-gray-200"],["Propostes",stats.propostes,"bg-yellow-50 border-yellow-200"],["En Estudi",stats.estudi,"bg-purple-50 border-purple-200"],["Presentades",stats.presentades,"bg-blue-50 border-blue-200"],["Pendent obertura",stats.pendentObertura,"bg-cyan-50 border-cyan-200"],["Adjudicades",stats.adjudicades,"bg-green-50 border-green-200"],["No Adjudicades",stats.noAdjudicades||0,"bg-amber-50 border-amber-200"],["No Presentades",stats.noPresentades,"bg-orange-50 border-orange-200"],["Descartades",stats.descartades,"bg-red-50 border-red-200"]].map(([label,val,cls])=>(<div key={label} className={`border rounded-lg p-2.5 text-center ${cls}`}><div className="text-xl font-bold">{val}</div><div className="text-xs text-gray-500">{label}</div></div>))}</div>
      <div className="flex gap-2 mb-3 flex-wrap">
        <input placeholder="Cerca codi, licitacio, client, poblacio..." value={search} onChange={e=>setSearch(e.target.value)} className="border rounded-lg px-3 py-1.5 text-sm flex-1 min-w-48"/>
        <select value={fEstat} onChange={e=>setFEstat(e.target.value)} className="border rounded-lg px-2 py-1.5 text-sm text-gray-600"><option value="ACTIVES">📌 Actives (En Estudi + Proposta)</option><option value="SENSE_DESC">Sense descartades</option><option value="">Tots els estats</option><option disabled>──────────</option>{ESTATS.map(s=><option key={s} value={s}>{s}</option>)}</select>
        <select value={fTipus} onChange={e=>setFTipus(e.target.value)} className="border rounded-lg px-2 py-1.5 text-sm text-gray-600"><option value="">Public / Privat</option>{tipus.map(t=><option key={t}>{t}</option>)}</select>
        {(fEstat!=="ACTIVES"||fTipus||search||hasActiveColFilters)&&<button onClick={()=>{setSearch("");setFEstat("ACTIVES");setFTipus("");setColFilters({});}} className="text-xs text-red-500 hover:underline px-1">Netejar filtres</button>}
        {hasActiveColFilters&&<span className="text-xs text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full">Filtres columna actius</span>}
      </div>
      {calSt&&<div className="mb-2 px-3 py-2 bg-blue-50 border border-blue-200 rounded-lg text-sm text-blue-700">{calSt}</div>}
      {filtered.length===0?<div className="text-center py-16 text-gray-400"><div className="text-4xl mb-2">📋</div><div className="font-semibold">{lic.length===0?"Cap licitacio":"Cap resultat"}</div></div>
        :<div className="gestor-table-wrap rounded-xl border border-gray-200 shadow-sm" style={{maxHeight:"75vh",overflowX:"scroll",overflowY:"auto"}} onClick={()=>setActiveColFilter(null)}><table className="text-xs" style={{minWidth:"1900px"}}><thead className="sticky top-0 z-10"><tr className="bg-gray-100 border-b border-gray-200 text-gray-500 uppercase">{(()=>{
          const FILTERABLE={codi_obra:{key:"codi_obra",vals:()=>[...new Set(lic.map(l=>l.codi_obra||"---"))].sort()},estat:{key:"estat",vals:()=>ESTATS},public_privat:{key:"public_privat",vals:()=>tipus},poblacio:{key:"poblacio",vals:()=>[...new Set(lic.map(l=>l.poblacio||"---"))].sort()},client:{key:"client",vals:()=>[...new Set(lic.map(l=>l.client||"---"))].sort()},data_presentacio:{key:"data_presentacio",vals:()=>[...new Set(lic.map(l=>l.data_presentacio||"---"))].sort()},termini:{key:"termini",vals:()=>[...new Set(lic.map(l=>l.termini||"---"))].sort()},import_pec:{key:"import_pec",vals:()=>[...new Set(lic.map(l=>l.import_pec_sense_iva?fmt(l.import_pec_sense_iva):"---"))].sort((a,b)=>{const na=parseFloat(a.replace(/\./g,"").replace(",",".")),nb=parseFloat(b.replace(/\./g,"").replace(",","."));return(isNaN(na)?0:na)-(isNaN(nb)?0:nb);})},classificacio:{key:"classificacio",vals:()=>[...new Set(lic.map(l=>l.classificacio||"---"))].sort()},tecnica:{key:"tecnica",vals:()=>["Sí","No"]}};
          const cols=[["Codi Obra","90px","codi_obra"],["Licitacio","280px"],["Client","180px","client"],["P/P","60px","public_privat"],["Poblacio","120px","poblacio"],["Estat","110px","estat"],["Data Present.","160px","data_presentacio"],["Termini","80px","termini"],["Import s/IVA","110px","import_pec"],["Classif.","80px","classificacio"],["Criteris","200px"],["Tècnica","65px","tecnica"],["Aval","70px"],["Sobres","120px"],["Comentaris","400px"],["🔗 Obra","80px"],["🌐 Publicació","80px"],["","70px"]];
          return cols.map(([h,w,filterKey])=>{
            const fDef=filterKey?FILTERABLE[filterKey]:null;
            const isActive=colFilters[filterKey]?.length>0;
            return<th key={h} style={{minWidth:w,width:w,position:"relative"}} className="px-2 py-2 text-left font-semibold whitespace-nowrap">
              {fDef?<div className="flex items-center gap-0.5">
                <span>{h}</span>
                <button onClick={e=>{e.stopPropagation();setActiveColFilter(activeColFilter===filterKey?null:filterKey);}} className={`ml-0.5 text-xs px-1 rounded ${isActive?"bg-blue-600 text-white":"text-gray-400 hover:text-gray-600 hover:bg-gray-200"}`}>{isActive?"▼✓":"▼"}</button>
                {activeColFilter===filterKey&&<div onClick={e=>e.stopPropagation()} className="absolute top-full left-0 mt-1 bg-white border border-gray-300 rounded-lg shadow-xl z-50 min-w-[160px] max-h-64 overflow-y-auto p-2" style={{textTransform:"none"}}>
                  <div className="flex gap-1 mb-1 border-b pb-1"><button onClick={()=>setColFilters(p=>({...p,[filterKey]:fDef.vals()}))} className="text-xs text-blue-600 hover:underline">Tot</button><button onClick={()=>clearColFilter(filterKey)} className="text-xs text-red-500 hover:underline">Cap</button></div>
                  {fDef.vals().map(v=><label key={v} className="flex items-center gap-1.5 py-0.5 px-1 hover:bg-gray-50 rounded cursor-pointer"><input type="checkbox" checked={(colFilters[filterKey]||[]).includes(v)} onChange={()=>toggleColFilter(filterKey,v)} className="rounded w-3 h-3"/><span className="text-xs text-gray-700 font-normal">{v}</span></label>)}
                </div>}
              </div>:h}
            </th>;
          });
        })()}</tr></thead><tbody className="divide-y divide-gray-100">{filtered.map(l=>(<tr key={l.id} className="hover:bg-gray-50 align-top"><td className="px-2 py-2 font-mono text-gray-600 whitespace-nowrap">{l.codi_obra||"---"}</td><td className="px-2 py-2 font-medium text-gray-800"><div title={l.licitacio}>{l.licitacio||"---"}</div></td><td className="px-2 py-2 text-gray-700">{l.client||"---"}</td><td className="px-2 py-2"><select value={l.public_privat||tipus[0]} className={"text-xs font-semibold px-1.5 py-0.5 rounded border-0 cursor-pointer "+(l.public_privat===tipus[0]?"bg-blue-100 text-blue-700":"bg-orange-100 text-orange-700")} onChange={e=>{const updated={...l,public_privat:e.target.value};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}>{tipus.map(t=><option key={t}>{t}</option>)}</select></td><td className="px-2 py-2 text-gray-600">{l.poblacio||"---"}</td><td className="px-2 py-2"><select value={l.estat} className={"text-xs font-semibold px-1.5 py-0.5 rounded-full border-0 cursor-pointer "+(EC[l.estat]||"bg-gray-100 text-gray-600")} onChange={e=>onEstat(l.id,e.target.value)}>{ESTATS.map(s=><option key={s}>{s}</option>)}</select></td><td className="px-2 py-2 text-gray-600"><div className="flex flex-col gap-1"><input type="datetime-local" className="text-xs border rounded px-1 py-0.5 w-full" value={(()=>{const v=l.data_presentacio||"";const m=v.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):?(\d{2})?)?/);if(m)return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}T${(m[4]||"00").padStart(2,"0")}:${(m[5]||"00").padStart(2,"0")}`;return"";})()} onChange={e=>{const d=e.target.value;if(!d)return;const dt=new Date(d);const formatted=`${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()} ${String(dt.getHours()).padStart(2,"0")}:${String(dt.getMinutes()).padStart(2,"0")}`;const updated={...l,data_presentacio:formatted};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}} /><input className="text-xs border rounded px-1 py-0.5 w-full" defaultValue={l.data_presentacio||""} placeholder="Text lliure" key={l.id+"-dp-"+l.data_presentacio} onBlur={e=>{const nv=e.target.value;if(nv!==(l.data_presentacio||"")){const updated={...l,data_presentacio:nv};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}}/></div></td><td className="px-2 py-2 text-gray-600">{l.termini||"---"}</td><td className="px-2 py-2 text-right font-medium text-gray-800 whitespace-nowrap">{(l.public_privat||tipus[0])!==tipus[0]?<input type="number" className="text-xs border rounded px-1 py-0.5 w-full text-right" defaultValue={l.import_pec_sense_iva||""} key={l.id+"-imp-"+l.import_pec_sense_iva} placeholder="Import" onBlur={e=>{const v=e.target.value?parseFloat(e.target.value):null;if(v!==(l.import_pec_sense_iva||null)){const updated={...l,import_pec_sense_iva:v};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}}/>:(l.import_pec_sense_iva?fmt(l.import_pec_sense_iva):"---")}</td><td className="px-2 py-2 font-mono">{l.classificacio||"---"}</td><td className="px-2 py-2 text-gray-600">{l.criteris_puntuacio||"---"}</td><td className="px-2 py-2">{(()=>{const priv=(l.public_privat||tipus[0])!==tipus[0];if(priv)return <span className="text-xs text-gray-300">—</span>;const val=l.tecnica!=null?l.tecnica:(l.criteris_puntuacio&&/100\s*%?\s*(auto|sobre)/i.test(l.criteris_puntuacio)?false:true);return <select value={val?"Sí":"No"} className={"text-xs font-semibold px-1.5 py-0.5 rounded border-0 cursor-pointer "+(val?"bg-slate-100 text-slate-600":"bg-green-100 text-green-700")} onChange={e=>{const updated={...l,tecnica:e.target.value==="Sí"};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}><option>Sí</option><option>No</option></select>})()}</td><td className="px-2 py-2 text-gray-600">{l.aval||"---"}</td><td className="px-2 py-2">{(()=>{
  const isPub=l.public_privat==="PUBLICA";
  const isPres=l.estat==="PRESENTADA"||l.estat==="ADJUDICADA"||l.estat==="NO ADJUDICADA";
  if(!isPub||!isPres)return<span className="text-xs text-gray-300">—</span>;
  const sobres=l.sobres||{};
  const opts=[["unic","Sobre únic"],["1A","Sobre 1/A"],["2B","Sobre 2/B"],["3C","Sobre 3/C"]];
  return<div className="space-y-0.5">{opts.map(([key,label])=>{const val=sobres[key];return<button key={key} onClick={()=>{const next=val===true?false:val===false?null:true;const newSobres={...sobres,[key]:next};const updated={...l,sobres:newSobres};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}} className={`block w-full text-left text-xs px-1.5 py-0.5 rounded ${val===true?"bg-green-100 text-green-700 font-semibold":val===false?"bg-red-100 text-red-600":"bg-gray-50 text-gray-400"}`}>{val===true?`✅ ${label}`:val===false?`❌ ${label}`:`⬜ ${label}`}</button>;})}</div>;
})()}</td><td className="px-2 py-2"><textarea rows={4} className="text-xs border rounded px-1 py-0.5 w-full resize-y" style={{minHeight:"60px",maxHeight:"200px"}} value={l.comentaris||""} placeholder="Comentaris" onBlur={e=>{if(e.target.value!==(l.comentaris||"")){const updated={...l,comentaris:e.target.value};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}} onChange={e=>{save(lic.map(x=>x.id===l.id?{...l,comentaris:e.target.value}:x));}}/></td><td className="px-2 py-2"><div className="flex items-center gap-1">{l.link_obra&&<a href={l.link_obra} target="_blank" rel="noreferrer" className="text-blue-500 hover:text-blue-700" title={l.link_obra}>🔗</a>}<input className="text-xs border rounded px-1 py-0.5 w-full" value={l.link_obra||""} placeholder="URL carpeta" onBlur={e=>{if(e.target.value!==(l.link_obra||"")){const updated={...l,link_obra:e.target.value};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}} onChange={e=>{save(lic.map(x=>x.id===l.id?{...l,link_obra:e.target.value}:x));}}/></div></td><td className="px-2 py-2"><div className="flex items-center gap-1">{l.link_publicacio&&<a href={l.link_publicacio} target="_blank" rel="noreferrer" className="text-green-500 hover:text-green-700" title={l.link_publicacio}>🌐</a>}<input className="text-xs border rounded px-1 py-0.5 w-full" value={l.link_publicacio||""} placeholder="URL publicació" onBlur={e=>{if(e.target.value!==(l.link_publicacio||"")){const updated={...l,link_publicacio:e.target.value};save(lic.map(x=>x.id===l.id?updated:x));saveOne(updated);}}} onChange={e=>{save(lic.map(x=>x.id===l.id?{...l,link_publicacio:e.target.value}:x));}}/></div></td><td className="px-2 py-2 whitespace-nowrap">{parseDT(l.data_presentacio)&&<button onClick={()=>addToCalendar(l,setCalSt)} title="Google Calendar" className="text-blue-400 hover:text-blue-600 mr-1">📅</button>}<button onClick={()=>onEdit(l)} className="text-blue-500 hover:text-blue-700 mr-1">✏️</button><button onClick={()=>onDel(l.id)} className="text-red-400 hover:text-red-600">🗑️</button></td></tr>))}</tbody></table></div>}
      <div className="mt-2 text-xs text-gray-400 text-right">{filtered.length} de {lic.length} licitacions — Sincronitzat amb Supabase</div>
    </div>
  );
}

export default function App(){
  const [activeTab,setActiveTab]=useState("filtres");
  const [filters,setFilters]=useState(EMPTY_FILTERS);
  const [results,setResults]=useState([]);
  const [loading,setLoading]=useState(false);
  const [statusMsg,setStatusMsg]=useState("");
  const [debugInfo,setDebugInfo]=useState("");
  const [expanded,setExpanded]=useState(null);
  const [showAdvanced,setShowAdvanced]=useState(false);
  const [showClassif,setShowClassif]=useState(false);
  const [recipients,setRecipients]=useState([{email:"vmata@servial.es",selected:false}]);
  const [newEmail,setNewEmail]=useState("");
  const [emailError,setEmailError]=useState("");
  const [emailStatus,setEmailStatus]=useState("");
  const [scheduleEnabled,setScheduleEnabled]=useState(false);
  const [scheduleTime,setScheduleTime]=useState("09:00");
  const [scheduleStatus,setScheduleStatus]=useState("");
  const [lastAutoRun,setLastAutoRun]=useState(null);
  const [countdown,setCountdown]=useState("");
  const [plecFiles,setPlecFiles]=useState([]);
  const [plecNom,setPlecNom]=useState("");
  const [plecOrgan,setPlecOrgan]=useState("");
  const [plecLoading,setPlecLoading]=useState(false);
  const [plecError,setPlecError]=useState("");
  const [plecStatus,setPlecStatus]=useState("");
  const [plecResults,setPlecResults]=useState([]);
  const [plecRawText,setPlecRawText]=useState("");
  const [plecView,setPlecView]=useState("taula");
  const [plecSavedMsg,setPlecSavedMsg]=useState("");
  const [plecHistorial,setPlecHistorial]=useState([]);
  const [plecHistorialOpen,setPlecHistorialOpen]=useState(false);
  const [legalQuery,setLegalQuery]=useState("");
  const [legalResponses,setLegalResponses]=useState([]);
  const [legalLoading,setLegalLoading]=useState(false);
  const consultarLegal=async()=>{
    if(!legalQuery.trim())return;
    setLegalLoading(true);
    try{
      const context=plecRawText?`CONTEXT DEL PLEC ANALITZAT (resum):\n${plecRawText.slice(0,6000)}\n\n`:"";
      const pregunta=legalQuery.trim();
      const userContent=[{type:"text",text:`${context}CONSULTA LEGAL:\n${pregunta}`}];
      const resposta=await callAI(SYSTEM_LEGAL,userContent,4096,aiProvider);
      setLegalResponses(prev=>[{pregunta,resposta,date:new Date().toLocaleTimeString("ca-ES")},...prev]);
      setLegalQuery("");
    }catch(e){setLegalResponses(prev=>[{pregunta:legalQuery,resposta:`❌ Error: ${e.message}`,date:new Date().toLocaleTimeString("ca-ES")},...prev]);}
    finally{setLegalLoading(false);}
  };
  useEffect(()=>{
    supabase.from("plec_historial").select("*").order("created_at",{ascending:false}).limit(20)
      .then(({data,error})=>{if(!error&&data)setPlecHistorial(data.map(d=>({id:d.id,nom:d.nom,date:d.date,results:d.results,rawText:d.raw_text})));});
    // Migrar localStorage a Supabase si hi ha dades locals
    try{const h=localStorage.getItem("servial-plec-historial");if(h){const local=JSON.parse(h);if(local.length>0){Promise.all(local.map(e=>supabase.from("plec_historial").upsert({id:e.id,nom:e.nom,date:e.date,results:e.results,raw_text:e.rawText},{onConflict:"id"}))).then(()=>{localStorage.removeItem("servial-plec-historial");supabase.from("plec_historial").select("*").order("created_at",{ascending:false}).limit(20).then(({data})=>{if(data)setPlecHistorial(data.map(d=>({id:d.id,nom:d.nom,date:d.date,results:d.results,rawText:d.raw_text})));});});}}}catch(e){}
  },[]);
  const guardarPlecHistorial=async(nom,results,rawText)=>{const r=results?.[0]||{};const label=[r.objecte,r.organisme,nom].filter(Boolean).join(" — ")||"Sense nom";const entry={id:Date.now(),nom:label,date:new Date().toLocaleDateString("ca-ES"),results,raw_text:rawText};await supabase.from("plec_historial").upsert(entry,{onConflict:"id"});const{data}=await supabase.from("plec_historial").select("*").order("created_at",{ascending:false}).limit(20);if(data)setPlecHistorial(data.map(d=>({id:d.id,nom:d.nom,date:d.date,results:d.results,rawText:d.raw_text})));};
  const eliminarPlecHistorial=async(id)=>{await supabase.from("plec_historial").delete().eq("id",id);setPlecHistorial(prev=>prev.filter(h=>h.id!==id));};
  const [showCidoImport,setShowCidoImport]=useState(false);
  const [cidoHtml,setCidoHtml]=useState("");
  const [showGencatImport,setShowGencatImport]=useState(false);
  const [gencatCount,setGencatCount]=useState(null);
  const gencatCacheRef=useRef(null);
  const getGencatCache=()=>gencatCacheRef.current;
  const setGencatCache=(results)=>{gencatCacheRef.current=results;try{localStorage.removeItem("servial-gencat-cache");}catch(e){}};
  const parseGencatEmail=(text)=>{
    // ESTRATÈGIA: El correu Gencat té dues seccions clares:
    // 1. "Plataforma de Serveis de Contractació Pública" → TOT és Catalunya → incloure tot
    // 2. "Publicacions de la plataforma de l'Estat" → Tot Espanya → filtrar per geografia catalana
    const isMeta=l=>/^(Data de publicació|Esmenat en data|Termini de presentació|Pressupost de licitació)/.test(l);
    const isOrgLine=l=>/^(Ajuntament|Universitat|Institut|Departament|Consell|Diputaci|Ferrocarril|Agència|ADIF|Aena|Consejería|Junta|Presidencia|Gobierno|Alcald|Pleno|Secretaria|Servicio|Subdirección|Dirección|Comité|Delegación|Empresa|Rectorado|TGSS|VISESA|AYUNTAMIENTO|Comisión|Consejo|Intendente|Axencia|Ayuntamiento|Alcaldía|Órgano|Infraestructur|Consorci|Patronat|Fundació|Generalitat|Societat|Autoritat|Mancomunitat|Entitat|Corporació|Port de |Ports de |Servei Català|INCASÒL|Agència Catalana|GENCAT|Govern|Parlament|Administra|el director|el secretari|la secretària|Ministerio|Gerencia|Subdelegación|Confederación|Demarcació|Gestió d|Organisme|Gestión|Sociedad|Autoridad|Comunidad|Ciudad Autónoma|Cabildo|Concejalía|Concejal|Rector|Parque|AUTOPISTAS|Consellería|Gerencia)/.test(l);

    // Separar seccions
    const secCatRx=/Plataforma de Serveis de Contractació Pública/i;
    const secEstatRx=/Publicacions de la plataforma de l'Estat/i;
    const lines=text.split("\n").map(l=>l.trim()).filter(Boolean);
    let secció="cat"; // per defecte tot és Catalunya
    const lics=[];let current=null;let lastOrg=""; // recorda últim organisme per agrupar licitacions seqüencials del mateix client

    for(let i=0;i<lines.length;i++){
      const line=lines[i];
      // Detectar canvi de secció
      if(secCatRx.test(line)){secció="cat";continue;}
      if(secEstatRx.test(line)){secció="estat";continue;}
      if(line==="Obres"&&lines[i-1]&&(secCatRx.test(lines[i-1])||secEstatRx.test(lines[i-1])))continue;
      if(line==="Obres")continue;

      // Detectar organisme
      if(isOrgLine(line)){
        if(current&&current.titol)lics.push({...current});
        lastOrg=line;
        current={organisme:line,titol:"",import_eur:0,data_presentacio:"",data_publicacio:"",secció};continue;
      }
      // Si no hi ha current però sí lastOrg, i la línia sembla títol (no meta, llarg) → és una nova licitació del mateix organisme
      if(!current&&lastOrg&&!isMeta(line)&&line.length>10&&!line.includes("€")&&!/^\d{2}\/\d{2}\/\d{4}/.test(line)){
        current={organisme:lastOrg,titol:line,import_eur:0,data_presentacio:"",data_publicacio:"",secció};continue;
      }
      // Fallback: organisme no reconegut (línia no-meta seguida de títol seguit de meta)
      if(!current&&!isMeta(line)&&line.length>5&&!line.includes("€")&&!/\d{2}\/\d{2}\/\d{4}/.test(line)){
        if(i+1<lines.length&&!isMeta(lines[i+1])&&lines[i+1].length>10&&i+2<lines.length&&isMeta(lines[i+2]?.trim())){
          lastOrg=line;
          current={organisme:line,titol:"",import_eur:0,data_presentacio:"",data_publicacio:"",secció};continue;
        }
      }
      if(current&&!current.titol&&line.length>10&&!isMeta(line)&&!line.includes("€")){current.titol=line;continue;}
      if(/Termini de presentació d'ofertes:\s*(.+)/.test(line)&&current){current.data_presentacio=line.match(/Termini de presentació d'ofertes:\s*(.+)/)[1].replace(/h\s*$/,"").trim();continue;}
      if(/Pressupost de licitació:\s*([\d.,]+)\s*€/.test(line)&&current){current.import_eur=parseFloat(line.match(/Pressupost de licitació:\s*([\d.,]+)\s*€/)[1].replace(/\./g,"").replace(",","."));lics.push({...current});current=null;continue;}
      if(/Data de publicació:\s*(.+)/.test(line)&&current){current.data_publicacio=line.match(/Data de publicació:\s*(.+)/)[1].trim();continue;}
      if(/Esmenat en data:\s*(.+)/.test(line)&&current){current.data_publicacio=line.match(/Esmenat en data:\s*(.+)/)[1].trim();continue;}
    }
    if(current&&current.titol)lics.push({...current});

    // Filtre tipus (exclou serveis, subministrament)
    const filtratTipus=lics.filter(l=>{
      const t=(l.titol||"").toLowerCase();
      if(/subministrament|suministro/.test(t))return false;
      if(/^servei|serveis\s|servicios?\s|servicio\s/.test(t))return false;
      if(/consultoria|assessorament|formaci/.test(t))return false;
      if(/arrendamiento de maquinaria/.test(t))return false;
      if(/alquiler de maquinaria/.test(t))return false;
      return true;
    });

    // Filtre geogràfic: secció "cat" → tot passa. Secció "estat" → només si menciona lloc català
    const geoTermsCat=getGeoTerms(COMARQUES);
    const catExtra=["catalunya","cataluña","generalitat","gencat","barcelona","girona","lleida","tarragona","el prat","reus","sabadell","terrassa","mataró","manresa","igualada","vic","granollers","vilanova","vilafranca","figueres","olot","tortosa","valls"];
    for(const ct of catExtra)geoTermsCat.add(ct);

    const filtratGeo=filtratTipus.filter(l=>{
      if(l.secció==="cat")return true; // Secció catalana → tot passa
      // Secció estatal → només si menciona localitat catalana al títol, organisme o data_presentacio
      const txt=(l.titol+" "+l.organisme).toLowerCase();
      return matchGeoTerms(txt,geoTermsCat);
    });

    return filtratGeo.map(l=>({expedient:"",objecte:l.titol,organisme:l.organisme,import_eur:l.import_eur,data_publicacio:l.data_publicacio||new Date().toLocaleDateString("ca-ES"),termini:"",data_presentacio_str:l.data_presentacio,cpv:"45000000",comarca_municipi:l.organisme||"",tipologia:"",font:"Gencat-Email",classificacio_requerida:[],puntuacio:l.secció==="cat"?6:5,justificacio:l.secció==="cat"?"📧 Correu Gencat (Catalunya)":"📧 Correu Gencat (Estatal)",url:""}));
  };
  const [gmailToken,setGmailToken]=useState(null);
  const [gmailConnected,setGmailConnected]=useState(false);
  const [gmailUser,setGmailUser]=useState("");
  const [cidoCount,setCidoCount]=useState(null);
  const [gestorRefreshKey,setGestorRefreshKey]=useState(0);
  const [apiKeyInput,setApiKeyInput]=useState(()=>getApiKey());
  const [geminiKeyInput,setGeminiKeyInput]=useState(()=>getGeminiKey());
  const [showApiKey,setShowApiKey]=useState(false);
  const [apiKeySaved,setApiKeySaved]=useState(!!getApiKey()||!!getGeminiKey());
  const [aiProvider,setAiProvider]=useState(()=>getAiProvider());
  const fileInputRef=useRef(null);
  const scheduleRef=useRef({enabled:false,time:"09:00",lastRun:null});

  useEffect(()=>{scheduleRef.current={enabled:scheduleEnabled,time:scheduleTime,lastRun:lastAutoRun};},[scheduleEnabled,scheduleTime,lastAutoRun]);
  const toggle=(key,val)=>setFilters(p=>({...p,[key]:p[key].includes(val)?p[key].filter(v=>v!==val):[...p[key],val]}));

  const parseCIDOToResults=useCallback((html)=>{
    const licitacions=parseCIDOHtml(html);
    if(!licitacions.length)return[];
    const paraules=["obres","obra","construcció","construccio","reforma","rehabilitació","rehabilitacio","urbanització","urbanitzacio","paviment","reurbanitz","condicionament","enderroc","sanejament","col·lector","canonada","pont","vestidor","edifici","local","equipament","instal·lació","instal·lacions","adequació","adequacio","millora","ampliació","ampliacio","substitució","substitucio","impermeabilitz","drenatge","clavegueram","xarxa","vial","asfalt","demolició","demolicio","mur","coberta","façana","facana","forjat","fonament","estructura","terraplenat","esplanada","ferm"];
    // Normalitza text: elimina accents codificats malament (Ã³→ó, etc.) i normalitza espais
    const norm=s=>s.normalize("NFC").replace(/\s+/g," ").trim().toLowerCase();
    const obres=licitacions.filter(l=>{
      const t=norm(l.titol);
      // Exclou si conté "subministrament" o "suministro" en qualsevol posició
      if(t.includes("subministrament")||t.includes("suministro")||t.includes("subministrament"))return false;
      // Exclou si conté "serveis" o "servei" seguit de qualsevol cosa (no obres)
      if(/servei/.test(t))return false;
      if(/servicio/.test(t))return false;
      // Exclou altres tipus no-obra
      if(t.includes("consultoria")||t.includes("assessorament")||t.includes("formaci"))return false;
      return true;
    });
    return obres.map(l=>({expedient:"",objecte:l.titol,organisme:l.organisme,import_eur:0,data_publicacio:new Date().toLocaleDateString("ca-ES"),termini:"",cpv:"45000000",comarca_municipi:l.organisme||"",tipologia:"",font:"CIDO-DIBA",classificacio_requerida:[],puntuacio:6,justificacio:"📧 CIDO correu",url:l.url}));
  },[]);

  const cidoCacheRef=useRef(null);
  const getCidoCache=()=>cidoCacheRef.current;
  const setCidoCache=(results)=>{cidoCacheRef.current=results;try{localStorage.removeItem("servial-cido-cache");}catch(e){}};

  const importarCIDO=useCallback(()=>{
    if(!cidoHtml.trim())return;
    const results=parseCIDOToResults(cidoHtml);
    setCidoCount(results.length);
    setCidoCache(results);
    setCidoHtml("");
    setShowCidoImport(false);
  },[cidoHtml,parseCIDOToResults]);

  /* ── Google OAuth + Gmail API directa (gratuït) ── */
  const connectGmail=useCallback(()=>{
    const clientId=import.meta.env.VITE_GOOGLE_CLIENT_ID;
    if(!clientId||!window.google?.accounts?.oauth2){alert("Google Identity Services no carregat. Recarrega la pàgina.");return;}
    const client=window.google.accounts.oauth2.initTokenClient({
      client_id:clientId,
      scope:"https://www.googleapis.com/auth/gmail.readonly",
      callback:(response)=>{
        if(response.error){
          console.error("Gmail OAuth error:",response);
          alert(`Error OAuth: ${response.error} - ${response.error_description||"Comprova la configuració a Google Cloud Console"}`);
          return;
        }
        if(response.access_token){
          setGmailToken(response.access_token);
          setGmailConnected(true);
          fetch("https://gmail.googleapis.com/gmail/v1/users/me/profile",{headers:{Authorization:`Bearer ${response.access_token}`}})
            .then(r=>r.json()).then(d=>setGmailUser(d.emailAddress||"")).catch(()=>{});
        }
      },
      error_callback:(err)=>{
        console.error("Gmail OAuth error_callback:",err);
        alert(`Error de connexió Gmail: ${err.type||err.message||JSON.stringify(err)}\n\nComprova:\n1. Popup no bloquejat\n2. Orígenes autoritzats a Google Cloud Console\n3. Usuari de prova afegit`);
      },
    });
    client.requestAccessToken({prompt:"consent"});
  },[]);

  const disconnectGmail=useCallback(()=>{
    if(gmailToken&&window.google?.accounts?.oauth2){
      window.google.accounts.oauth2.revoke(gmailToken);
    }
    setGmailToken(null);setGmailConnected(false);setGmailUser("");
  },[gmailToken]);

  /* ── Lectura automàtica de correus Gmail (CIDO + Gencat) ── */
  const fetchGmailLicitacions=useCallback(async()=>{
    if(!gmailToken)return[];
    try{
      // 1. Search for specific emails from CIDO and Gencat (last 3 days)
      const query=encodeURIComponent('(from:sgsi.svcSam@diba.cat subject:"Liictacions a Catalunya") OR (from:plataforma.contractacio@gencat.cat subject:"Correu diari de subscriptors generals") newer_than:3d');
      const listRes=await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${query}&maxResults=10`,{
        headers:{Authorization:`Bearer ${gmailToken}`}
      });
      if(!listRes.ok){if(listRes.status===401){setGmailConnected(false);setGmailToken(null);}return[];}
      const listData=await listRes.json();
      if(!listData.messages?.length)return[];

      // 2. Fetch each message body in parallel
      const msgPromises=listData.messages.map(m=>
        fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${m.id}?format=full`,{
          headers:{Authorization:`Bearer ${gmailToken}`}
        }).then(r=>r.json()).catch(()=>null)
      );
      const messages=await Promise.all(msgPromises);

      // 3. Extract HTML body from each message and parse
      let allResults=[];
      for(const msg of messages){
        if(!msg?.payload)continue;
        const htmlBody=extractGmailHtmlBody(msg.payload);
        if(!htmlBody)continue;
        // Determine source
        const fromHeader=(msg.payload.headers||[]).find(h=>h.name.toLowerCase()==="from")?.value||"";
        const isGencat=fromHeader.includes("gencat.cat");
        const isCido=fromHeader.includes("diba.cat");

        if(isCido){
          // Use existing CIDO parser
          const results=parseCIDOToResults(htmlBody);
          allResults.push(...results);
        }else if(isGencat){
          // Parse Gencat email - similar structure
          const results=parseGencatToResults(htmlBody);
          allResults.push(...results);
        }
      }
      return allResults;
    }catch(e){console.error("Gmail API error:",e);return[];}
  },[gmailToken,parseCIDOToResults]);

  const fetchTransparencia=useCallback(async()=>{
    const f=filters;
    const importMin=f.importMin?Number(f.importMin):0;
    const importMax=f.importMax?Number(f.importMax):0;
    const today=new Date().toISOString().slice(0,10);
    const since=new Date();since.setDate(since.getDate()-3);
    const sinceStr=since.toISOString().slice(0,10);
    let where=`tipus_contracte='Obres' AND fase_publicacio='Anunci de licitació' AND termini_presentacio_ofertes>='${today}' AND data_publicacio_anunci>='${sinceStr}'`;
    if(importMin>0)where+=` AND pressupost_licitacio_sense>=${importMin}`;
    if(importMax>0)where+=` AND pressupost_licitacio_sense<=${importMax}`;
    // Filtre per província a l'API directament (molt més fiable que filtrar per text)
    if(f.comarques.length>0&&f.comarques.length<COMARQUES.length){
      const provCodes=[];
      const bcnCom=PROVINCIES[0].comarques,girCom=PROVINCIES[1].comarques,lleCom=PROVINCIES[2].comarques,tarCom=PROVINCIES[3].comarques;
      if(f.comarques.some(c=>bcnCom.includes(c)))provCodes.push("'08'","'08 - Barcelona'","'8'");
      if(f.comarques.some(c=>girCom.includes(c)))provCodes.push("'17'","'17 - Girona'");
      if(f.comarques.some(c=>lleCom.includes(c)))provCodes.push("'25'","'25 - Lleida'");
      if(f.comarques.some(c=>tarCom.includes(c)))provCodes.push("'43'","'43 - Tarragona'");
      // Filtrar per codi provincial O per municipi de les comarques seleccionades
      const geoTerms=getGeoTerms(f.comarques);
      const municipis=[...geoTerms].filter(t=>t.length>5).slice(0,50);// màx 50 per no sobrecarregar la query
      const llocFilters=[...provCodes.map(p=>`lloc_execucio=${p}`),...municipis.map(m=>`lloc_execucio='${m.charAt(0).toUpperCase()+m.slice(1)}'`)];
      if(llocFilters.length>0)where+=` AND (${llocFilters.join(" OR ")})`;
    }
    const url=`https://analisi.transparenciacatalunya.cat/resource/ybgg-dgi6.json?$limit=500&$order=data_publicacio_anunci DESC&$where=${encodeURIComponent(where)}`;
    const res=await fetch(url);
    if(!res.ok)return[];
    const data=await res.json();
    const now=new Date();
    // Agrupar per expedient+objecte per detectar lots i prendre el valor estimat total si és més alt
    const grouped={};
    for(const r of data){
      const key=((r.codi_expedient||"")+"|"+(r.objecte_contracte||r.denominacio||"")).toLowerCase().replace(/[.\s]+$/,"");
      if(!grouped[key])grouped[key]={records:[],maxImport:0,maxValorEst:0};
      grouped[key].records.push(r);
      const pres=parseFloat(r.pressupost_licitacio_sense)||0;
      const val=parseFloat(r.valor_estimat_contracte)||parseFloat(r.valor_estimat_expedient)||0;
      grouped[key].maxImport=Math.max(grouped[key].maxImport,pres);
      grouped[key].maxValorEst=Math.max(grouped[key].maxValorEst,val);
    }
    // Calcular suma de tots els pressupostos per expedient (per casos de lots)
    for(const key of Object.keys(grouped)){
      const g=grouped[key];
      const exps=new Set(g.records.map(r=>(r.codi_expedient||"").replace(/\.+$/,"")));
      // Sumar pressupostos de registres únics (evitar duplicats mateixa fase)
      const seenByPres=new Set();
      let sum=0;
      g.records.forEach(r=>{const p=parseFloat(r.pressupost_licitacio_sense)||0;if(p>0&&!seenByPres.has(p)){seenByPres.add(p);sum+=p;}});
      g.sumImport=sum;
    }
    return data.map(r=>{
      const terminiRaw=r.termini_presentacio_ofertes||"";
      const terminiDate=terminiRaw?new Date(terminiRaw):null;
      const linkUrl=r.enllac_publicacio?.url||r.enllac_publicacio||"";
      const pres=parseFloat(r.pressupost_licitacio_sense)||0;
      const valEst=parseFloat(r.valor_estimat_contracte)||parseFloat(r.valor_estimat_expedient)||0;
      // Si hi ha valor estimat molt superior (>2x) al pressupost, usar el valor estimat
      const key=((r.codi_expedient||"")+"|"+(r.objecte_contracte||r.denominacio||"")).toLowerCase().replace(/[.\s]+$/,"");
      const g=grouped[key];
      let importFinal=pres;
      if(g&&g.sumImport>pres*1.5)importFinal=g.sumImport;// suma de lots
      else if(valEst>pres*2)importFinal=valEst;// valor estimat si >2x pressupost (projecte amb pròrrogues/lots)
      return{
        expedient:r.codi_expedient||"",
        objecte:r.objecte_contracte||r.denominacio||"",
        organisme:r.nom_organ||"",
        import_eur:importFinal,
        data_publicacio:r.data_publicacio_anunci?new Date(r.data_publicacio_anunci).toLocaleDateString("ca-ES"):"",
        termini:terminiDate?terminiDate.toLocaleDateString("ca-ES"):"",
        cpv:(r.codi_cpv||"45000000").split("-")[0],
        comarca_municipi:r.lloc_execucio||"",
        tipologia:r.procediment||"",
        font:"Transparència Catalunya",
        classificacio_requerida:[],
        puntuacio:6,
        justificacio:"API dades obertes Catalunya",
        url:linkUrl,
        _terminiDate:terminiDate
      };
    }).filter(r=>{
      // Only open tenders (termini in future or no termini)
      if(r._terminiDate&&r._terminiDate<now)return false;
      // Filter by keywords
      if(f.paraulesClau){const kw=f.paraulesClau.toLowerCase();if(!r.objecte.toLowerCase().includes(kw)&&!r.organisme.toLowerCase().includes(kw))return false;}
      // Filter by comarca/lloc
      if(f.comarques.length>0&&f.comarques.length<COMARQUES.length){
        const geoTerms=getGeoTerms(f.comarques);
        const loc=(r.comarca_municipi||"").toLowerCase();
        const org=(r.organisme||"").toLowerCase();
        const txt=loc+" "+org;
        if(!matchGeoTerms(txt,geoTerms))return false;
      }
      return true;
    }).map(({_terminiDate,...rest})=>rest).reduce((acc,r)=>{
      // Deduplicar per expedient (normalitzat sense punt final) + objecte
      const key=((r.expedient||"").replace(/\.+$/,"")+"|"+(r.objecte||"").slice(0,60)).toLowerCase();
      const existing=acc.find(x=>((x.expedient||"").replace(/\.+$/,"")+"|"+(x.objecte||"").slice(0,60)).toLowerCase()===key);
      if(existing){if(r.import_eur>existing.import_eur)Object.assign(existing,r);}
      else acc.push(r);
      return acc;
    },[]);
  },[filters]);

  const fetchPSCP=useCallback(async()=>{
    // Fetch tenders without import (pressupost IS NULL) — complementary to main query
    try{
      const today=new Date().toISOString().slice(0,10);
      const since=new Date();since.setDate(since.getDate()-3);const sinceStr=since.toISOString().slice(0,10);
      const where=`tipus_contracte='Obres' AND fase_publicacio='Anunci de licitació' AND termini_presentacio_ofertes>='${today}' AND data_publicacio_anunci>='${sinceStr}' AND pressupost_licitacio_sense IS NULL`;
      const url=`https://analisi.transparenciacatalunya.cat/resource/ybgg-dgi6.json?$limit=100&$order=termini_presentacio_ofertes ASC&$where=${encodeURIComponent(where)}`;
      const res=await fetch(url);
      if(!res.ok)return[];
      const data=await res.json();
      return data.map(r=>{
        const terminiRaw=r.termini_presentacio_ofertes||"";
        const terminiDate=terminiRaw?new Date(terminiRaw):null;
        const linkUrl=r.enllac_publicacio?.url||r.enllac_publicacio||"";
        return{
          expedient:r.codi_expedient||"",
          objecte:r.objecte_contracte||r.denominacio||"",
          organisme:r.nom_organ||"",
          import_eur:parseFloat(r.pressupost_licitacio_sense)||0,
          data_publicacio:r.data_publicacio_anunci?new Date(r.data_publicacio_anunci).toLocaleDateString("ca-ES"):"",
          termini:terminiDate?terminiDate.toLocaleDateString("ca-ES"):"",
          cpv:(r.codi_cpv||"45000000").split("-")[0],
          comarca_municipi:r.lloc_execucio||"",
          tipologia:r.procediment||"",
          font:"PSCP Catalunya",classificacio_requerida:[],puntuacio:5,
          justificacio:"Licitació oberta",url:linkUrl
        };
      });
    }catch(e){return[];}
  },[]);

  const buscarIEnviar=useCallback(async(auto=false)=>{
    if(auto)setScheduleStatus("⏳ Executant cerca automàtica…");
    setLoading(true);setResults([]);setStatusMsg("Cercant a contractaciopublica.cat i CIDO…");setDebugInfo("Iniciant cerca a APIs públiques (sense límit d'import)…");
    try{
      const f=filters;
      const cidoExtra=getCidoCache()||[];
      const gencatExtra=getGencatCache()||[];
      const [transpData,pscpData,gmailData]=await Promise.all([
        fetchTransparencia().catch(e=>{setDebugInfo(d=>d+`\n⚠️ Transparència: ${e.message}`);return[];}),
        fetchPSCP().catch(e=>{setDebugInfo(d=>d+`\n⚠️ PSCP: ${e.message}`);return[];}),
        fetchGmailLicitacions().catch(e=>{setDebugInfo(d=>d+`\n⚠️ Gmail: ${e.message}`);return[];})
      ]);
      if(gmailData.length>0){setCidoCache([...cidoExtra,...gmailData]);setCidoCount((cidoExtra.length||0)+gmailData.length);}
      const rawGmail=[...cidoExtra,...gmailData,...gencatExtra];
      // Aplicar filtres globals a resultats CIDO/Gmail (que no tenen filtre propi)
      // Per Gencat/CIDO: si no hi ha filtre geogràfic explícit, filtrar per TOTA Catalunya (no mostrar resta d'Espanya)
      const geoComarques=f.comarques.length>0?f.comarques:COMARQUES;// per defecte tot Catalunya
      const geoTerms=getGeoTerms(geoComarques);
      // Afegir termes genèrics de Catalunya per capturar organismes catalans no associats a municipi
      const catTerms=["catalunya","cataluña","generalitat","gencat","diputació de barcelona","diputació de girona","diputació de tarragona","diputació de lleida","diputacion de barcelona","diputacion de girona","diputacion de tarragona","diputacion de lleida"];
      for(const ct of catTerms)geoTerms.add(ct);
      const filteredGmail=rawGmail.filter(r=>{
        // Filtre per paraules clau
        if(f.paraulesClau){const kw=f.paraulesClau.toLowerCase();if(!r.objecte.toLowerCase().includes(kw)&&!r.organisme.toLowerCase().includes(kw))return false;}
        // Filtre geogràfic — SEMPRE actiu per fonts email (CIDO/Gencat)
        if(r.font==="Gencat-Email"||r.font==="CIDO-DIBA"){
          const txt=(r.objecte+" "+r.organisme+" "+(r.comarca_municipi||"")).toLowerCase();
          if(!matchGeoTerms(txt,geoTerms))return false;
        } else if(f.comarques.length>0&&f.comarques.length<COMARQUES.length){
          const txt=(r.objecte+" "+r.organisme+" "+(r.comarca_municipi||"")).toLowerCase();
          if(!matchGeoTerms(txt,geoTerms))return false;
        }
        // Filtre per import mínim/màxim
        if(f.importMin&&Number(f.importMin)>0&&r.import_eur&&r.import_eur<Number(f.importMin))return false;
        if(f.importMax&&Number(f.importMax)>0&&r.import_eur&&r.import_eur>Number(f.importMax))return false;
        return true;
      });
      setDebugInfo(`Transparència: ${transpData.length} | PSCP: ${pscpData.length} | CIDO/Gmail: ${rawGmail.length} (Gencat: ${gencatExtra.length}) | Filtrat: ${filteredGmail.length}`);
      const tots=[...transpData,...pscpData,...filteredGmail],vistos=new Set();
      const combinats=tots.filter(r=>{const k=(r.expedient||r.objecte||"").toLowerCase().trim().slice(0,50);if(!k||vistos.has(k))return false;vistos.add(k);return true;});
      if(combinats.length===0)throw new Error("No s'han trobat licitacions amb els filtres actuals.");
      const mapped=combinats.map(r=>{
        let classif=r.classificacio_requerida||[];
        if(classif.length===0&&r.import_eur>=500000)classif=inferClassificacio(r.cpv,r.import_eur);
        return{...r,classificacio_requerida:classif,_compat:checkCompatibility(classif)};
      });
      const sorted=[...mapped].sort((a,b)=>b.puntuacio-a.puntuacio);
      setResults(sorted);setActiveTab("resultats");
      setStatusMsg(`${sorted.length} licitació${sorted.length!==1?"ns":""} trobada${sorted.length!==1?"es":""} (Transp: ${transpData.length} | PSCP: ${pscpData.length} | CIDO/Gmail: ${filteredGmail.length})`);
      setDebugInfo(`✅ ${sorted.length} resultats totals`);
    }catch(e){setStatusMsg(`Error: ${e.message}`);setDebugInfo(`❌ ${e.message}`);}
    finally{setLoading(false);}
  },[fetchTransparencia,fetchPSCP,fetchGmailLicitacions,filters]);

  useEffect(()=>{
    const iv=setInterval(()=>{
      const{enabled,time,lastRun}=scheduleRef.current,now=new Date(),[h,m]=time.split(":").map(Number);
      if(enabled){const isTime=now.getHours()===h&&now.getMinutes()===m,alreadyRan=lastRun&&lastRun.toDateString()===now.toDateString()&&lastRun.getHours()===h&&lastRun.getMinutes()===m;if(isTime&&!alreadyRan)buscarIEnviar(true);}
      const target=new Date(now);target.setHours(h,m,0,0);if(target<=now)target.setDate(target.getDate()+1);
      const diff=target-now;setCountdown(`${Math.floor(diff/3600000)}h ${Math.floor((diff%3600000)/60000)}m`);
    },30000);
    return()=>clearInterval(iv);
  },[buscarIEnviar]);

  const addRecipient=()=>{const e=newEmail.trim().toLowerCase();if(!e)return;if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)){setEmailError("Adreça no vàlida");return;}if(recipients.find(r=>r.email===e)){setEmailError("Ja existeix");return;}setRecipients(p=>[...p,{email:e,selected:true}]);setNewEmail("");setEmailError("");};
  const toggleRecipient=email=>setRecipients(p=>p.map(r=>r.email===email?{...r,selected:!r.selected}:r));
  const removeRecipient=email=>setRecipients(p=>p.filter(r=>r.email!==email));
  const selectedEmails=recipients.filter(r=>r.selected).map(r=>r.email);

  const exportExcelRadar=(data,nom)=>{
    if(!data?.length)return;
    try{const ws=XLSX.utils.json_to_sheet(data);const cols=Object.keys(data[0]).map(k=>({wch:Math.min(60,Math.max(k.length,...data.map(r=>String(r[k]||"").length)))}));ws["!cols"]=cols;const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Licitacions");const b64=XLSX.write(wb,{bookType:"xlsx",type:"base64"});const a=document.createElement("a");a.href=`data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`;a.download=nom;document.body.appendChild(a);a.click();document.body.removeChild(a);}catch(e){alert(`Error: ${e.message}`);}
  };
  const exportarResultatsExcel=()=>exportExcelRadar(results.map(r=>({Expedient:r.expedient||"",Objecte:r.objecte||"",Organisme:r.organisme||"","Import (€)":r.import_eur||0,"Data publicació":r.data_publicacio||"","Termini presentació":r.termini||"",Font:r.font||"",CPV:r.cpv||"",Comarca:r.comarca_municipi||"","Classif.":(r.classificacio_requerida||[]).map(c=>`${c.grup}${c.subgrup} Cat.${c.categoria}`).join(" | "),"Servial pot":r._compat?.pot!==false?"SÍ":r._compat?.pot===false?"NO":"S/C",Puntuació:r.puntuacio||0,URL:r.url||""})),`licitacions_${new Date().toISOString().slice(0,10)}.xlsx`);

  const enviarEmail=async()=>{
    if(!results.length||!selectedEmails.length)return;setEmailStatus("Enviant…");
    const date=new Date().toLocaleDateString("ca-ES"),aptes=results.filter(r=>r._compat?.pot!==false);
    const body=`Bon dia,\n\nResum licitacions (${date}):\n\n✅ Aptes Servial (${aptes.length}):\n${aptes.map((r,i)=>`${i+1}. ${r.objecte}\n   ${r.organisme} — ${r.import_eur?.toLocaleString("ca-ES")||"N/D"} €\n   ${r.url||""}`).join("\n\n")}\n\nTotal: ${results.length}`;
    try{await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:apiHeaders(),body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:500,messages:[{role:"user",content:`Envia correu a: ${selectedEmails.join(", ")}\nAssumpte: Radar Licitacions PSCP — ${date}\n${body}`}],mcp_servers:[{type:"url",url:"https://gmail.mcp.claude.com/mcp",name:"gmail-mcp"}]})});setEmailStatus("Enviat ✓");setTimeout(()=>setEmailStatus(""),4000);}catch(e){setEmailStatus("Error");}
  };

  const guardarAlGestor=useCallback(async r=>{
    try{
      const nova={id:Date.now(),codi_obra:"",licitacio:r.objecte||"",client:r.organisme||"",public_privat:"PUBLICA",poblacio:"",estat:"PROPOSTA",data_presentacio:r.termini_presentacio||"",termini:r.termini_execucio||"",import_pec_sense_iva:r.import_sense_iva||"",classificacio:(r.classificacio_requerida||[]).map(c=>`${c.grup}${c.subgrup} Cat.${c.categoria}`).join(" | "),criteris_puntuacio:[...(r.criteris_automatics||[]).map(c=>`${c.nom} ${c.punts}pt`),...(r.criteris_judici_valor||[]).map(c=>`${c.nom} ${c.punts}pt`)].join(" + "),aval:r.garantia_definitiva||"",apertura:"",comentaris:[r.visita_obra&&r.visita_obra!=="No"?`📍 VISITA OBRA: ${r.visita_obra}`:"",r.diagnosic_servial?r.diagnosic_servial.slice(0,200):""].filter(Boolean).join(" | "),link_obra:"",link_publicacio:r.enllac_publicacio?.url||r.enllac_publicacio||"",analisi_completa:plecRawText||""};
      if(isSupabaseConfigured()){await supabase.from("licitacions").upsert(nova);await supabase.from("licitacions_log").insert({licitacio_id:nova.id,estat_anterior:"",estat_nou:nova.estat||"PROPOSTA",codi_obra:nova.codi_obra||""});}
      else{let llista=[];try{const r=localStorage.getItem(SK);if(r)llista=JSON.parse(r);}catch(e){}llista=[nova,...llista];localStorage.setItem(SK,JSON.stringify(llista));}
      setGestorRefreshKey(k=>k+1);setPlecSavedMsg("✅ Guardat al Gestor!");setTimeout(()=>setPlecSavedMsg(""),3000);
    }catch(e){setPlecSavedMsg("❌ Error: "+e.message);}
  },[plecRawText]);

  const analitzarPlec=async()=>{
    if(!plecFiles.length){setPlecError("Afegeix almenys un fitxer PDF.");return;}
    setPlecLoading(true);setPlecResults([]);setPlecRawText("");setPlecError("");setPlecStatus("Llegint fitxers PDF…");
    try{
      const MAX_PDF_PAGES=50;
      const rawDocs=await Promise.all(plecFiles.map(f=>new Promise((res,rej)=>{
        const reader=new FileReader();
        reader.onload=()=>res({name:f.name,arrayBuffer:reader.result});
        reader.onerror=()=>rej(new Error(`No s'ha pogut llegir ${f.name}`));
        reader.readAsArrayBuffer(f);
      })));
      setPlecStatus("Verificant pàgines dels PDFs…");
      const docs=[];
      for(const raw of rawDocs){
        const pdfDoc=await PDFDocument.load(raw.arrayBuffer);
        const totalPages=pdfDoc.getPageCount();
        if(totalPages<=MAX_PDF_PAGES){
          const b64=btoa(new Uint8Array(raw.arrayBuffer).reduce((s,b)=>s+String.fromCharCode(b),""));
          docs.push({name:raw.name,b64});
        }else{
          setPlecStatus(`PDF "${raw.name}" té ${totalPages} pàgines. Dividint en trossos de ${MAX_PDF_PAGES}…`);
          for(let start=0;start<totalPages;start+=MAX_PDF_PAGES){
            const end=Math.min(start+MAX_PDF_PAGES,totalPages);
            const newPdf=await PDFDocument.create();
            const pages=await newPdf.copyPages(pdfDoc,Array.from({length:end-start},(_,i)=>start+i));
            pages.forEach(p=>newPdf.addPage(p));
            const bytes=await newPdf.save();
            const b64=btoa(new Uint8Array(bytes).reduce((s,b)=>s+String.fromCharCode(b),""));
            docs.push({name:`${raw.name} (pàg. ${start+1}-${end})`,b64});
          }
        }
      }
      setPlecStatus(`Analitzant ${docs.length} document${docs.length!==1?"s":""} (${docs.length>rawDocs.length?"dividit automàticament":""}) amb IA…`);
      const PROMPT=`Analitza el PCAP/PPT d'aquesta licitació:
• Expedient: ${plecNom||"(no indicat)"}
• Òrgan de contractació: ${plecOrgan||"(no indicat)"}

═══════════════════════════════════════════════════════════════
MARC LEGAL — LCSP (Llei 9/2017) arts. 74-100 + RD 1098/2001
═══════════════════════════════════════════════════════════════

REGLA FONAMENTAL: Classificació i solvència NO són requisits independents acumulables.
La CLASSIFICACIÓ SUBSTITUEIX la solvència (econòmica + tècnica) en els grups/subgrups que comprèn.

QUAN S'EXIGEIX CLASSIFICACIÓ:
- Obres amb valor estimat ≥ 500.000 € → OBLIGATÒRIA (art. 77)
- Serveis → POTESTATIVA (l'òrgan pot exigir-la)
- Subministraments → NO aplica

QUAN NO HI HA CLASSIFICACIÓ (obres < 500.000 € o quan el plec no l'exigeix):
- S'exigeix solvència econòmica i financera (art. 87): volum negoci, assegurança, patrimoni net
- S'exigeix solvència tècnica (art. 88 obres): obres executades 10 anys, personal tècnic, maquinària, titulacions
- Les categories 1 i 2 es poden acreditar per solvència sense classificació

CATEGORIES DE CLASSIFICACIÓ (per anualitat mitjana):
Cat.1: fins 150.000€ | Cat.2: 150.001-360.000€ | Cat.3: 360.001-840.000€ | Cat.4: 840.001-2.400.000€ | Cat.5: 2.400.001-5.000.000€ | Cat.6: >5.000.000€

GRUPS I SUBGRUPS COMPLETS D'OBRES (RD 1098/2001):
A-Moviment de terres: A1=Desmunts/buidats, A2=Explanacions, A3=Pedreres, A4=Pous/galeries, A5=Túnels
B-Ponts/estructures: B1=Fàbrica/formigó massa, B2=Formigó armat, B3=Formigó pretensat, B4=Metàl·lics, B5=Mixtos
C-Edificació: C1=Demolicions, C2=Estructures fàbrica/formigó, C3=Estructures metàl·liques, C4=Paleta/arrebossats/revestiments, C5=Pedra/marbre, C6=Paviments/enrajolats, C7=Aïllaments/impermeabilitzacions, C8=Fusteria fusta, C9=Fusteria metàl·lica
D-Ferrocarrils: D1=Estesa vies, D2=Elevats sobre carril/cable, D3=Senyalització/enclavaments, D4=Electrificació, D5=Obres FC sense especialització
E-Hidràuliques: E1=Abastaments/sanejaments, E2=Preses, E3=Canals, E4=Sèquies/desguassos, E5=Defenses marges/canalitzacions, E6=Conduccions tuberia pressió gran Ø, E7=Obres hidràuliques sense especialització
F-Marítimes: F1=Dragats, F2=Esculleres, F3=Blocs formigó, F4=Caixons formigó armat, F5=Pilots/tablestacas, F6=Fars/senyalitzacions, F7=Obres marítimes sense especialització
G-Vials/pistes: G1=Autopistes/autovies, G2=Pistes aterratge, G3=Ferms formigó hidràulic, G4=Ferms mescles bituminoses, G5=Senyalització/balisament vial, G6=Obres vials sense especialització
H-Transports/instal·lacions: H1=Grans instal·lacions calefacció/refrigeració/ventilació, H2=Fontaneria, H3=Gas en edificis, H4=Salubritat/sanejament, H5=Contra incendis en edificis, H6=Climatització
I-Instal·lacions elèctriques: I1=Enllumenat/balisament lluminós, I2=Centrals producció energia, I3=Línies transport, I4=Subestacions, I5=Centres transformació/distribució AT, I6=Distribució BT, I7=Telecomunicacions/radioelèctriques, I8=Electròniques, I9=Sense especialització
J-Instal·lacions mecàniques: J1=Ascensors/escales mecàniques, J2=Transport pneumàtic, J3=Frigorífiques, J4=Bugaderies/neteja en sec, J5=Cuines industrials, J6=Esportives/recreatives, J7=Depuració/potabilització/tractament aigües, J8=Aprofitament solar, J9=Sense especialització
K-Especials: K1=Cimentacions especials, K2=Sondejos/injeccions/pilotatges, K3=Tablestacats, K4=Pintures/metal·litzacions, K5=Ornamentacions/decoracions, K6=Jardineria/plantacions, K7=Restauració béns immobles, K8=Estacions tractament aigües, K9=Sense especialització

GRUPS DE SERVEIS (classificació potestativa):
L=Serveis a persones | M=Telecomunicacions/TI | N=Consultoria/assessoria/gestió | O=Serveis tècnics enginyeria/arquitectura | P=I+D | Q=Culturals/recreatius/esportius | R=Manteniment/reparació

═══════════════════════════════════════════════════════════════

INSTRUCCIONS D'EXTRACCIÓ — segueix aquest esquema:

1. DADES BÀSIQUES
   - Objecte del contracte
   - Codi CPV
   - LOTS — SECCIÓ CRÍTICA:
     Si la licitació té LOTS, cada lot és una OBRA INDEPENDENT amb:
     → Objecte/descripció pròpia del lot
     → Import propi (pressupost base sense IVA)
     → Classificació pròpia (pot ser diferent per cada lot!)
     → Criteris d'adjudicació propis (si difereixen entre lots)
     → Termini d'execució propi (si difereix)
     ⚠️ TRACTA CADA LOT COM SI FOS UNA LICITACIÓ SEPARADA en l'informe.
     Exemple: Lot 1 = Obra civil (C4 Cat.4), Lot 2 = Instal·lacions (I1 Cat.3, J2 Cat.3)
     AL JSON FINAL: genera UN objecte JSON per CADA LOT dins l'array, NO un sol objecte per tot.
   - Pressupost base de licitació sense IVA i amb IVA (global i per lot si n'hi ha)
   - Valor estimat del contracte (pot diferir del pressupost si hi ha pròrrogues/lots)
   - Termini d'execució (global i per lot si difereix)
   - Termini de presentació d'ofertes (data i hora exacta, SEMPRE en format DD/MM/YYYY HH:MM, ex: 15/04/2026 14:00. Mai text descriptiu, només la data numèrica)

2. CLASSIFICACIÓ I SOLVÈNCIA — EXTRACCIÓ LITERAL DEL PCAP (no interpretació de la llei)

   ⚠️ REGLA PRINCIPAL: TRANSCRIU el que DIU literalment el PCAP. No apliquis teoria de la llei.
   El teu treball és llegir el PCAP i respondre a 3 preguntes concretes EXTRETES del document:

   PREGUNTA 1: ¿El PCAP EXIGEIX CLASSIFICACIÓ?
   → Busca al PCAP la secció de "Classificació empresarial" o "Solvència"
   → Si el PCAP exigeix classificació específica (ex: "C4 Cat.2") → "Sí, el PCAP exigeix classificació [grups/subgrups/cat.]"
   → Si el PCAP NO exigeix classificació → "No, el PCAP no exigeix classificació"
   → Omple al JSON: "exigeix_classificacio": true/false i "classificacio_requerida": [...]

   PREGUNTA 2: ¿El PCAP EXIGEIX SOLVÈNCIA ECONÒMICA I TÈCNICA?
   → Busca al PCAP les seccions "Solvència econòmica" i "Solvència tècnica"
   → Si el PCAP les exigeix com a requisits d'accés → "Sí, el PCAP exigeix solvència"
   → Si el PCAP no les menciona → "No, el PCAP no exigeix solvència"
   → TRANSCRIU literalment què demana: imports, mitjans d'acreditació, anys, perfils...
   → Omple al JSON: "solvencia_obligatoria": true/false, "solvencia_economica" i "solvencia_tecnica" amb transcripció

   PREGUNTA 3: ¿El PCAP permet SUBSTITUIR solvència per classificació (o viceversa)?
   → BUSCA literalment al PCAP frases com:
     - "podrà acreditar-se alternativament mitjançant classificació"
     - "s'admetrà la classificació en grup..."
     - "equivalent a la classificació en..."
     - "es podrà substituir per la classificació"
     - "alternativament, mitjançant la classificació"
     - "en lloc d'aquesta solvència, l'empresa podrà acreditar..."
     - "no serà necessari acreditar la solvència si l'empresa disposa de classificació..."
   → Si el PCAP ho permet → "Sí, el PCAP permet substitució. Text literal: [transcriu]"
   → Si el PCAP NO ho permet → "No, el PCAP no preveu substitució. La solvència és obligatòria."
   → Si no apareix res al PCAP sobre el tema → "El PCAP no menciona la possibilitat de substitució"
   → Omple al JSON: "permet_substitucio_per_classificacio": true/false/null

   CAMPS OBLIGATORIS AL JSON:
   → "exigeix_classificacio": true/false segons PCAP
   → "classificacio_requerida": [...] grup/subgrup/categoria segons PCAP
   → "solvencia_obligatoria": true/false segons PCAP
   → "solvencia_economica": transcripció literal de què demana el PCAP (imports, mitjans)
   → "solvencia_tecnica": transcripció literal de què demana el PCAP (obres similars, personal, etc.)
   → "permet_substitucio_per_classificacio": true/false/null segons PCAP
   → "text_substitucio_pcap": transcripció LITERAL del fragment del PCAP sobre substitució (si n'hi ha)
   → "subcas_explicacio": resum en 1-2 frases del que diu el PCAP

   IMPORTANT:
   • NO citis la llei (LCSP, arts. 77, 86, 87, 88...) com a resposta. Cita només el PCAP.
   • Si el PCAP és confús o contradictori, indica-ho: "El PCAP no és clar sobre aquest punt"
   • El teu paper és ser un lector fidel del PCAP, no un expert en lleis

   TAULA CPV ↔ GRUPS DE CLASSIFICACIÓ (correspondència orientativa):
   CPV 45110000 Demolicions/mov.terres → A (Mov.terres), C1 (Demolicions)
   CPV 45200000 Construcció edificis → C (Edificació: C2-C9)
   CPV 45210000 Edificis habitatges/equipaments → C (Edificació)
   CPV 45220000 Eng.civil/estructures → B (Ponts), C2-C3 (Estructures)
   CPV 45230000 Conduccions/canonades → E (Hidràuliques: E1,E4-E7)
   CPV 45233000 Pavimentació/vials → G (Vials: G3-G6)
   CPV 45240000 Projectes hidràulics → E (Hidràuliques)
   CPV 45260000 Cobertes → C7 (Aïllaments/impermeab.)
   CPV 45300000 Instal·lacions edificis → H,I,J (Instal·lacions)
   CPV 45310000 Instal·lació elèctrica → I (Elèctriques: I1-I9)
   CPV 45320000 Aïllament → C7 (Aïllaments)
   CPV 45330000 Fontaneria → H2 (Fontaneria), J (Mecàniques)
   CPV 45340000 Tanques/baranes → C9 (Tancaments metàl·lics)
   CPV 45350000 Instal·lacions mecàniques → J (Mecàniques: J1-J9)
   CPV 45400000 Acabats edificis → C4-C6,C8 (Paleta,pedra,paviments,fusteria)
   CPV 45410000 Arrebossats/guix → C4 (Paleta/revestiments)
   CPV 45420000 Fusteria → C8 (Fusteria fusta)
   CPV 45430000 Revestiments sòls/parets → C6 (Paviments/enrajolats)
   CPV 45440000 Pintura/vidrieria → K4 (Pintures), C4
   CPV 45450000 Altres acabats → C4,C5 (Pedra/marbre)
   CPV 77310000 Jardineria → K6 (Jardineria/plantacions)

   CLASSIFICACIÓ SERVIAL → CPV coberts:
   A1(Desmunts) → CPV 45110000-45112000
   C1-C9(Edificació) → CPV 45200000-45260000, 45400000-45450000
   E1,E4,E5,E7(Hidràuliques) → CPV 45230000-45240000
   G3,G5,G6(Vials) → CPV 45233000
   K6(Jardineria) → CPV 77310000
   ⚠️ Servial NO cobreix: CPV 45310000-45350000 (instal·lacions I,J,H) ni CPV 45234000 (ferrocarrils D)

   ATENCIÓ: La categoria depèn de l'anualitat mitjana (import ÷ anys), NO de l'import total.
   Ex: obra de 1.500.000€ en 2 anys → anualitat 750.000€ → Cat.3

3. CRITERIS D'ADJUDICACIÓ — EXTRACCIÓ EXHAUSTIVA

   3A. CRITERIS AUTOMÀTICS (Sobre econòmic / Sobre C):
   - Preu: fórmula EXACTA tal com apareix al plec (lineal, proporcional, amb/sense referència a mitjana)
   - Llindar de temeritat / oferta anormalment baixa: criteri i % o fórmula
   - Reducció de termini (si puntua): fórmula i límits
   - Altres automàtics: certificacions, garantia addicional, millores quantificables, etc.
   - Puntuació de cada criteri automàtic

   3B. CRITERIS DE JUDICI DE VALOR (Sobre tècnic / Sobre B) — SECCIÓ DECISIVA

   ⚠️ AQUESTA SECCIÓ ÉS LA MÉS IMPORTANT DE TOT L'INFORME.
   La qualitat d'aquesta extracció determina si l'empresa es presenta o no a la licitació.
   NO RESUMEIXIS. TRANSCRIU fidelment el que diu el plec per cada criteri.
   Si el plec és vague, indica-ho explícitament: "El plec no detalla aquest aspecte".

   ⚠️⚠️⚠️ PAS OBLIGATORI ABANS DE RESUMIR:
   Abans d'escriure res, LOCALITZA al PCAP la secció que llista els criteris subjectes a judici de valor i TRANSCRIU-LA LITERALMENT paraula per paraula amb tots els seus nivells. El PCAP habitualment llista cada criteri i els seus subcriteris amb una puntuació específica a cada línia.

   NO escrivis NOMÉS els criteris de primer nivell. Si al PCAP hi ha numeració del tipus "1.1, 1.1.1, 1.2, 2.1, 2.1.1..." HAS D'INCLOURE-HO TOT. Aquest és l'error més habitual: quedar-se al primer nivell. NO HO FACIS.

   EXEMPLE del nivell de detall que has de proporcionar (tret d'un PCAP real d'obra pública):
      Els criteris no avaluables segons xifres o fórmules i la seva puntuació son els següents:
      1. Procediment d'execució de l'obra (fins a 17 punts)
         1.1. Implantació i exposició de l'execució (fins a 13 punts)
            1.1.1. Implantació (fins a 2 punts)
            1.1.2. Execució (fins a 11 punts)
         1.2. Afectacions a l'entorn (fins a 4 punts)
      2. Planificació de l'obra i termini ofertat (fins a 8 punts)
         2.1. Planificació detallada (fins a 6,50 punts)
            2.1.1. Estratègica (fins a 3,25 punts)
            2.1.2. A curt termini (fins a 3,25 punts)
         2.2. Termini ofertat (fins a 1,50 punts)

   Si al PCAP apareix una estructura similar (amb els seus propis noms i valors), transcriu-la sencera. Si el PCAP té una estructura diferent (taules, lletres, bullets), adapta't però NO perdis cap nivell.

   Per CADA criteri de judici de valor, fes una FITXA COMPLETA:

   a) NOM del criteri i PUNTUACIÓ MÀXIMA
   b) SUBCRITERIS — EXTRACCIÓ ADAPTATIVA i EXHAUSTIVA

      Cada PCAP és diferent. Adapta't a l'estructura que trobis al document, sigui quina sigui (numerada, amb lletres, amb bullets, amb taules...). Però extreu TOTA la granularitat disponible al PCAP.

      REGLA FONAMENTAL: si el PCAP assigna punts a un subconcepte, ha d'aparèixer al teu informe. No agrupis, no resumeixis, no saltis nivells.

      PROCÉS D'EXTRACCIÓ:
      1. Localitza la secció de criteris de judici de valor al PCAP (amb els noms que usi el document).
      2. Identifica cada element que tingui una puntuació associada (sigui com sigui: "fins a X punts", "màxim X", "es valorarà amb X punts", taules amb columna de punts...).
      3. Detecta la jerarquia segons el format del propi document: pot ser numèrica (1, 1.1, 1.1.1), amb lletres (A, B, a, b), amb bullets indentats, amb taules niuades, o mixta.
      4. Normalitza internament a una jerarquia d'objectes amb "codi" (si existeix al PCAP, usa el codi literal; si no, genera un codi seqüencial), "nom" literal, "punts", "nivell" (enter) i "fills" (array).
      5. Si un element té fills amb puntuació → l'element és un contenidor; els fills són els criteris reals.
      6. Si un element NO té subdivisions amb punts → és terminal, "fills":[].

      VERIFICACIÓ DE COHERÈNCIA:
      • Suma dels punts dels fills ha de coincidir amb els punts del pare (amb tolerància de decimals).
      • La suma total ha de coincidir amb la puntuació màxima del sobre de judici de valor.
      • Si no quadren, revisa si t'has deixat algun subcriteri al PCAP.

      ❌ NO ACCEPTABLE: retornar només el nivell superior quan el PCAP desglossa més.
      ✅ ACCEPTABLE: respectar l'estructura exacta del document, amb tots els subniveles que contingui.

      FORMAT AL JSON (adaptatiu però consistent):
         "subcriteris":[
            {"codi":"<codi literal del PCAP o seqüencial>","nom":"<text literal>","punts":<número>,"nivell":<1/2/3/...>,"fills":[...]}
         ]
      → Decimals amb punt (6.5, NO 6,50)
      → Si el PCAP no desglossa un nivell, "fills":[] — NO inventis contingut
      → Si el PCAP té una taula amb files i subfiles, les subfiles van dins "fills" de la fila pare

      FORMAT A L'INFORME DE TEXT (a més del JSON):
      Transcriu amb indentació tot el que hi ha al PCAP, mantenint la numeració/codificació original:
      • Respecta el codi original si existeix (1.1, 1.1.1 o a), 1), i))
      • Indica els punts tal com ho fa el PCAP
      • Si el PCAP detalla criteris d'avaluació per subcriteri (ex: "es valorarà la concreció del plànning"), inclou aquesta informació

   c) QUÈ VALOREN EXACTAMENT — TRANSCRIU literalment o resumeix molt fidelment:
      - Text exacte del plec que descriu què es valora en aquest criteri
      - Si valoren exhaustivitat, concreció, innovació, adequació al projecte, experiència, metodologia...
      - Quins aspectes concreten donen més puntuació? Què diferencia una nota alta d'una baixa?

   d) REQUISITS D'EQUIP I PERSONAL — si el criteri valora personal:
      - Quins perfils professionals demanen? (cap d'obra, encarregat, tècnics especialistes...)
      - Exigeixen titulacions específiques? Quines?
      - Demanen anys d'experiència? Quants i en què?
      - Demanen CV o relació d'obres similars del personal?
      - Han d'estar adscrits a l'obra a temps complet o parcial?
      - Es valora la dedicació (hores/setmana)?
      - Es valora experiència en obres similars? Defineixen "similar"? (import, tipus, àmbit públic?)
      - Es demana relació nominal o només organigrama genèric?
      - Es valora formació addicional (PRL, qualitat, mediambiental)?

   e) REQUISITS DE PLANIFICACIÓ — si el criteri valora planificació:
      - Demanen planning d'obra? Quin format? (Gantt, camí crític, PERT, diagrama de barres?)
      - Han d'indicar fites intermèdies? Quines?
      - Valoren coherència amb el termini d'execució?
      - Demanen seqüència de tasques amb durades?
      - Cal justificar rendiments?
      - Cal indicar equips i maquinària per fase?

   f) REQUISITS DE METODOLOGIA — si el criteri valora procés constructiu:
      - Demanen descripció del procés constructiu? Amb quin nivell de detall?
      - Cal justificació tècnica dels procediments escollits?
      - Es valora l'adequació al context específic de l'obra (entorn urbà, edifici en ús, hospital, escola...)?
      - Demanen mesures de minimització d'impacte? (soroll, pols, vibracions, talls de servei)
      - Demanen pla de gestió de residus?
      - Demanen pla de qualitat? Amb quins continguts mínims?
      - Demanen pla de seguretat i salut adaptat? O només genèric?

   g) COM ES PUNTUA — barem de puntuació:
      - Si el plec defineix rangs (0-5: insuficient, 5-10: adequat, 10-15: excel·lent, etc.) TRANSCRIU-LOS
      - Si puntua per comparació entre ofertes (la millor rep el màxim, les altres proporcional)
      - Si hi ha criteri de consens de la mesa o puntuació individual dels membres
      - Si es ponderen criteris entre si
      - BUSCA inconsistències: si les bandes no cobreixen tot el rang de puntuació, ALERTA

   h) MILLORES VALORABLES — si el plec puntua millores:
      - Quines millores accepten? LLISTA COMPLETA
      - Catàleg tancat de millores o millores lliures?
      - Valoració: per import econòmic, per utilitat funcional, per innovació?
      - Límit de millores (nombre, import, % sobre pressupost)
      - Cost per a l'empresa: les millores les assumeix el contractista? O es paguen apart?
      - Es poden proposar millores no previstes al plec?

   LLINDAR MÍNIM TÈCNIC:
   - Hi ha puntuació mínima al sobre tècnic per obrir l'econòmic? Quina?
   - Si no s'arriba al llindar, l'oferta queda exclosa automàticament?
   - El llindar és sobre el total del sobre B o sobre algun criteri individual?

   3C. REQUISITS FORMALS DE LA MEMÒRIA TÈCNICA:
   - Limitació de pàgines (nombre màxim per memòria o per apartat)
   - Mida de la font / tipografia exigida
   - Format (A4, interlineat, marges, si permeten annexos fora del límit de pàgines)
   - Si es penalitza per excedir el límit de pàgines (no es llegeix? es penalitza? es retorna?)
   - Contingut obligatori: apartats que HA de tenir la memòria, índex suggerit o obligatori
   - Si demanen documentació gràfica (plànols, planning, organigrama, etc.)

   3D. RESUM ESTRATÈGIC:
   - Total punts automàtics vs judici de valor (punts i %)
   - Tipus de licitació: PREU DOMINANT (>60% auto) o QUALITAT DOMINANT (>50% judici valor)
   - On concentrar l'esforç: preu agressiu o memòria tècnica detallada?

4. CRITERIS D'EXCLUSIÓ I TEMERITAT — extracció ADAPTATIVA i EXHAUSTIVA

   Cada PCAP redacta la temeritat de manera diferent. Pot estar en una clàusula independent, dins de la secció de criteris d'adjudicació, en un annex, o en una taula. Adapta't al format que trobis i extreu el contingut LITERAL.

   ESTRATÈGIA DE RECERCA (lectura completa del PCAP):
   Cerca al document qualsevol referència a:
   • ofertes amb valors anormals/desproporcionats/baixos
   • presumpció de temeritat o baixes temeràries
   • supòsits o llindars que generin sospita d'oferta inviable
   • criteris numèrics, percentuals o per franges aplicats al preu
   • referències a l'art. 149 LCSP o desenvolupaments de la mateixa norma
   • taules o fórmules en clàusules de criteris d'adjudicació
   • redaccions menys habituals: "no es consideraran ofertes que superin...", "es rebutjaran ofertes inferiors a...", "es considerarà desproporcionada quan..."

   POSSIBLES FORMES DE FORMULAR-HO (no exhaustiu, només per orientar la recerca):
   • Percentatges fixos sobre el pressupost de licitació
   • Desviacions sobre la mitjana aritmètica de les ofertes
   • Comportament segons el nombre de licitadors
   • Combinacions amb baixes màximes admeses
   • Remissió íntegra a la normativa (art. 149 LCSP / art. 85 RGLCAP) sense afegir res
   • Taules amb varis trams

   EXTREU LITERALMENT (NO interpretis, NO simplifiquis):
   → Text exacte del PCAP que descriu el criteri de temeritat
   → Tots els valors numèrics o llindars
   → Procediment posterior (justificació, terminis, documents a aportar)
   → Conseqüències (exclusió automàtica, admissió condicional, etc.)
   → Qualsevol matís específic del document

   CRITERIS D'EXCLUSIÓ (extreu segons el que diu el PCAP):
   Cerca al PCAP causes d'exclusió de l'oferta. Cada PCAP en té de pròpies. Algunes habituals:
   • Documentació incompleta o errònia
   • Llindar mínim tècnic no assolit
   • Ofertes fora de termini
   • Incompliment de solvència/classificació
   • Format incorrecte de presentació
   • Manca de signatura electrònica vàlida
   • Incompliments específics del plec

   FORMAT AL JSON:
   "temeritat":{
     "formula":"Transcripció LITERAL del PCAP (no interpretis)",
     "percentatge":"Valors numèrics extrets, si n'hi ha",
     "procediment":"Termini per justificar, documents a aportar (literal del PCAP)",
     "consequencia":"Conseqüència segons el PCAP"
   }
   "criteris_exclusio":[llista literal de causes que apareixen al PCAP]

   ⚠️ SI després d'una lectura completa del PCAP NO trobes cap esment a la temeritat:
   • Indica clarament "El PCAP examinat no conté criteris específics de temeritat" — només si és cert
   • Tingues en compte: en obres és infreqüent que el PCAP no en parli, així que torna a buscar amb diferents paraules clau abans de concloure-ho

5. GARANTIES I CONDICIONS
   - Garantia provisional (si escau)
   - Garantia definitiva (% i base de càlcul)
   - Condicions especials d'execució
   - Visita d'obra obligatòria (sí/no, data si escau)
   - Penalitats rellevants

6. DIAGNÒSTIC PER A SERVIAL — SECCIÓ CRÍTICA, FES-LA AMB MÀXIM RIGOR

   Classificació acreditada SERVIAL (Exp. 202210137) — LLISTA EXHAUSTIVA, Servial NOMÉS té aquests subgrups:
   A1 Cat.3 | C1 Cat.4 | C2 Cat.4 | C3 Cat.4 | C4 Cat.4 | C5 Cat.4 | C6 Cat.4 | C7 Cat.4 | C8 Cat.4 | C9 Cat.4
   E1 Cat.3 | E4 Cat.3 | E5 Cat.3 | E7 Cat.3 | G3 Cat.2 | G5 Cat.2 | G6 Cat.4 | K6 Cat.1

   ⚠️ IMPORTANT: Servial NO TÉ classificació en els grups B, D, F, H, I, J ni en subgrups no llistats de A, E, G, K.
   Qualsevol subgrup NO present a la llista anterior = SERVIAL NO EL TÉ = INCOMPATIBLE per aquell subgrup.

   PROCEDIMENT OBLIGATORI (fes una taula):
   a) Llista TOTS els subgrups exigits pel plec en una taula amb columnes:
      | Subgrup exigit | Categoria exigida | Servial té? | Categoria Servial | Resultat |
      Per cada fila:
      - Cerca el subgrup EXACTE a la llista de Servial de dalt
      - Si NO el troba → Resultat = "❌ NO DISPOSA"
      - Si el troba però categoria insuficient → Resultat = "⚠️ CATEGORIA INSUFICIENT"
      - Si el troba i categoria suficient → Resultat = "✅ COMPLEIX"
   b) Calcula l'anualitat mitjana (import total ÷ anys execució) per verificar la categoria requerida.
   c) CONCLUSIÓ FINAL basada en la taula:
      - Si TOTS els subgrups són ✅ → COMPATIBLE
      - Si algun és ❌ o ⚠️ → INCOMPATIBLE o PARCIAL
      - Explica CLARAMENT quins subgrups falten i per què
   d) Si és PARCIAL o INCOMPATIBLE:
      - Suggereix si seria viable en UTE
      - Indica el perfil de soci necessari (quins grups/subgrups ha d'aportar)
      - Exemple: "Caldria un soci amb classificació I1 Cat.3 i J2 Cat.3 per cobrir instal·lacions"

7. ENLLAÇ DE PUBLICACIÓ
   - Si el document conté l'URL de publicació oficial (perfil del contractant, plataforma de contractació), inclou-lo al camp "enllac_publicacio" del JSON.

Al FINAL del text, afegeix el JSON entre aquests marcadors exactes (sense backticks ni text extra).

⚠️ REGLA IMPORTANT PER LOTS:
- Si la licitació té LOTS, genera un ARRAY amb un objecte JSON per CADA LOT: [{lot1},{lot2}]
- Si NO té lots, genera un ARRAY amb un sol objecte: [{objecte únic}]
- Cada objecte JSON del lot ha de tenir el camp "lot" amb el número/nom del lot

--JSON_INICI--
[{"lot":"","expedient":"","objecte":"","organisme":"","cpv":"","import_sense_iva":0,"import_amb_iva":0,"valor_estimat":0,"termini_execucio":"","termini_presentacio":"DD/MM/YYYY HH:MM","exigeix_classificacio":true,"classificacio_requerida":[{"grup":"C","subgrup":"2","categoria":3}],"classificacio_substitueix_solvencia":true,"solvencia_obligatoria":true,"permet_substitucio_per_classificacio":false,"text_substitucio_pcap":"","subcas_explicacio":"","solvencia_economica":"","solvencia_tecnica":"","criteris_automatics":[{"nom":"Preu","punts":60,"formula":""}],"criteris_judici_valor":[{"nom":"Planificació i seqüència d'execució","punts":15,"subcriteris":[{"codi":"1","nom":"Procediment d'execució","punts":17,"nivell":1,"fills":[{"codi":"1.1","nom":"Implantació i exposició","punts":13,"nivell":2,"fills":[{"codi":"1.1.1","nom":"Implantació","punts":2,"nivell":3,"fills":[]},{"codi":"1.1.2","nom":"Execució","punts":11,"nivell":3,"fills":[]}]},{"codi":"1.2","nom":"Afectacions a l'entorn","punts":4,"nivell":2,"fills":[]}]}],"que_valoren":"Coherència, adequació al projecte, compatibilitat","bandes_puntuacio":[{"rang":"0-2","descripcio":"proposta genèrica o poc adaptada"},{"rang":"3-4","descripcio":"correcta i coherent"},{"rang":"5-6","descripcio":"molt ben estructurada i específica"}],"inconsistencia":""}],"requisits_memoria":{"max_pagines":"","format":"DIN A4","font":"","interlineat":"","annexos":"","penalitzacio_excedir":""},"llindar_minim_tecnic":{"te_llindar":false,"puntuacio_minima":0,"consequencia":""},"total_punts_automatics":60,"total_punts_judici_valor":40,"total_punts":100,"temeritat":{"formula":"","percentatge":"","procediment":"","consequencia":""},"criteris_exclusio":[""],"garantia_definitiva":"5% preu adjudicació s/IVA","garantia_provisional":"","condicions_especials":"","visita_obra":"No","penalitats":"","diagnosic_servial":"","enllac_publicacio":""}]
--JSON_FI--`;
      const SYSTEM_PLEC=`Ets un expert en contractació pública espanyola (LCSP Llei 9/2017, RD 1098/2001). Regles clau:
1. CLASSIFICACIÓ SUBSTITUEIX SOLVÈNCIA (arts. 77-85): no són acumulables. En obres ≥500k€ la classificació és obligatòria.
2. CATEGORIA = anualitat mitjana (import÷anys execució), NO import total. Cat.1:≤150k | Cat.2:150-360k | Cat.3:360-840k | Cat.4:840k-2,4M | Cat.5:2,4-5M | Cat.6:>5M.
3. Coneixes TOTS els grups (A-K) i subgrups d'obres, i els grups de serveis (L-R).
4. Analitzes PCAP/PPT amb rigor jurídic per a Servial, constructora catalana amb classificació en grups A,C,E,G,K.`;
      let raw="";
      if(docs.length===1){
        const contentBlocks=[
          {type:"document",source:{type:"base64",media_type:"application/pdf",data:docs[0].b64}},
          {type:"text",text:PROMPT}
        ];
        setPlecStatus(`Analitzant amb ${aiProvider==="claude"?"Claude (Anthropic)":"Gemini (Google)"}…`);
        raw=await callAI(SYSTEM_PLEC,contentBlocks,16000,aiProvider);
      }else{
        // Múltiples trossos: analitzar cada part per separat i combinar
        const partials=[];
        for(let i=0;i<docs.length;i++){
          setPlecStatus(`Analitzant part ${i+1}/${docs.length} (${docs[i].name}) amb ${aiProvider==="claude"?"Claude":"Gemini"}…`);
          if(i<docs.length-1){
            // Parts intermèdies: extreure TOTES les dades rellevants, especialment TRANSCRIURE literalment els criteris i puntuacions
            const extractPrompt=`Aquesta és la part ${i+1}/${docs.length} del plec. EXTREU tota la informació rellevant que trobis.

⚠️ REGLA CRÍTICA: Si trobes la secció de CRITERIS D'ADJUDICACIÓ (criteris automàtics o criteris de judici de valor), TRANSCRIU LITERALMENT i COMPLETA l'arborescència amb tots els nivells i subnivells, codis i puntuacions exactes, sense resumir ni agrupar. Exemple del nivell de detall que vull:
  1. Procediment d'execució (fins a 17 punts)
     1.1. Implantació i exposició (fins a 13 punts)
        1.1.1. Implantació (fins a 2 punts)
        1.1.2. Execució (fins a 11 punts)
     1.2. Afectacions a l'entorn (fins a 4 punts)

Per la resta d'apartats (imports, terminis, classificació, solvència, garanties, temeritat, etc.), transcriu les dades literals tal com apareixen al document. NO resumeixis, NO simplifiquis, NO t'estalviïs subnivells.

NO generis JSON ni informe final en aquesta part; només la transcripció literal de les dades trobades.`;
            const contentBlocks=[
              {type:"document",source:{type:"base64",media_type:"application/pdf",data:docs[i].b64}},
              {type:"text",text:extractPrompt}
            ];
            const partRaw=await callAI(SYSTEM_PLEC,contentBlocks,12000,aiProvider);
            if(partRaw)partials.push(partRaw);
            setPlecStatus(`Part ${i+1} completada. Esperant 65s per rate limit…`);
            await new Promise(r=>setTimeout(r,65000));
          }else{
            // ÚLTIMA PART: genera l'informe ÚNIC i complet
            const finalPrompt=`Aquesta és la ÚLTIMA part (${i+1}/${docs.length}) del plec. A continuació tens les dades literals extretes de les parts anteriors (usa-les TOTES sense resumir, especialment els criteris i puntuacions):\n\n${partials.map((p,j)=>`=== DADES PART ${j+1} ===\n${p.slice(0,8000)}`).join("\n\n")}\n\n⚠️ IMPORTANT: Si les dades anteriors contenen la secció de criteris de judici de valor amb desglossament jeràrquic, TRANSCRIU-LA LITERALMENT a l'informe amb tots els subniveles, codis i puntuacions. NO resumeixis, NO agrupis, NO saltis nivells.\n\nAmb TOTES les dades anteriors + aquesta última part, genera UN SOL informe complet seguint l'estructura. NO repeteixis apartats. Cada apartat ha d'aparèixer UNA SOLA vegada.\n\n${PROMPT}`;
            const contentBlocks=[
              {type:"document",source:{type:"base64",media_type:"application/pdf",data:docs[i].b64}},
              {type:"text",text:finalPrompt}
            ];
            const partRaw=await callAI(SYSTEM_PLEC,contentBlocks,16000,aiProvider);
            if(partRaw)partials.push(partRaw);
          }
        }
        // Només l'últim partial conté l'informe complet
        raw=partials[partials.length-1]||"";
      }
      if(!raw)throw new Error("La IA no ha retornat cap resposta. Comprova que el PDF no estigui protegit.");
      setPlecRawText(raw);
      setPlecStatus("");
      const jsonMatch=raw.match(/--JSON_INICI--([\s\S]*?)--JSON_FI--/);
      let parsedResults=[];
      // Try parsing whatever's between markers; fall back to the full raw if markers missing.
      const candidate = jsonMatch ? jsonMatch[1] : raw;
      const parsed = parseJSONLenient(candidate);
      if(parsed){
        parsedResults = Array.isArray(parsed) ? parsed : [parsed];
        setPlecResults(parsedResults);
        setPlecView("taula");
      } else if(jsonMatch){
        // We had markers but couldn't recover JSON — log the snippet for debugging.
        try{JSON.parse(jsonMatch[1].trim());}catch(e){console.error("Plec JSON parse error:",e.message,"\nSnippet head:",jsonMatch[1].slice(0,200),"\nSnippet tail:",jsonMatch[1].slice(-200));}
        setPlecError("Informe generat però no s'ha pogut parsejar el JSON. Consulta la pestanya 📝 Informe.");
        setPlecView("text");
      } else {
        setPlecView("text");
      }
      guardarPlecHistorial(plecNom||plecOrgan||"Anàlisi",parsedResults,raw);
    }catch(e){
      setPlecError(`Error: ${e.message}`);
      setPlecStatus("");
    }finally{
      setPlecLoading(false);
    }
  };

  const exportarPlecExcel=()=>exportExcelRadar(plecResults.map(r=>({"Expedient":r.expedient||"","Objecte":r.objecte||"","Organisme":r.organisme||"","CPV":r.cpv||"","Import s/IVA":r.import_sense_iva||0,"Import c/IVA":r.import_amb_iva||0,"Termini execució":r.termini_execucio||"","Termini presentació":r.termini_presentacio||"","Classificació":(r.classificacio_requerida||[]).map(c=>`${c.grup}${c.subgrup} Cat.${c.categoria}`).join(" | "),"Solvència econòmica":r.solvencia_economica||"","Solvència tècnica":r.solvencia_tecnica||"","Criteris automàtics":(r.criteris_automatics||[]).map(c=>`${c.nom}(${c.punts}pts): ${c.formula}`).join(" | "),"Criteris judici valor":(r.criteris_judici_valor||[]).map(c=>`${c.nom}(${c.punts}pts)`).join(" | "),"Garantia definitiva":r.garantia_definitiva||"","Diagnòstic Servial":r.diagnosic_servial||""})),`analisi_plec_${plecNom||"licitacio"}_${new Date().toISOString().slice(0,10)}.xlsx`);

  const mdToHtml=md=>{if(!md)return"";
    // Convert markdown tables to HTML tables
    const lines=md.split('\n');const out=[];let inTable=false;
    for(let i=0;i<lines.length;i++){
      const line=lines[i].trim();
      if(line.startsWith('|')&&line.endsWith('|')){
        const cells=line.split('|').filter((_,j,a)=>j>0&&j<a.length-1).map(c=>c.trim());
        if(!inTable){inTable=true;out.push('<table><thead><tr>'+cells.map(c=>'<th>'+c+'</th>').join('')+'</tr></thead><tbody>');
          // Skip separator row (|---|---|)
          if(i+1<lines.length&&/^\|[\s\-:|]+\|$/.test(lines[i+1].trim()))i++;
        }else{out.push('<tr>'+cells.map(c=>'<td>'+c+'</td>').join('')+'</tr>');}
      }else{
        if(inTable){inTable=false;out.push('</tbody></table>');}
        out.push(line);
      }
    }
    if(inTable)out.push('</tbody></table>');
    return out.join('\n')
    .replace(/####\s+(.+)/g,'<h4>$1</h4>')
    .replace(/###\s+(.+)/g,'<h3>$1</h3>')
    .replace(/##\s+(.+)/g,'<h2>$1</h2>')
    .replace(/#\s+(.+)/g,'<h1>$1</h1>')
    .replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>')
    .replace(/\*(.+?)\*/g,'<em>$1</em>')
    .replace(/^[-•]\s+(.+)/gm,'<li>$1</li>')
    .replace(/(<li>.*<\/li>)/gs,m=>'<ul>'+m+'</ul>')
    .replace(/<\/ul>\s*<ul>/g,'')
    .replace(/═+/g,'<hr>')
    .replace(/─+/g,'<hr style="border-style:dashed">')
    .replace(/\n{2,}/g,'</p><p>')
    .replace(/\n/g,'<br>')
    .replace(/^/,'<p>').replace(/$/,'</p>');};
  const exportarPlecPDF=(res,rawText,nom)=>{
    const r=res[0]||{};const date=new Date().toLocaleDateString("ca-ES");
    const fmtI=v=>v?Number(v).toLocaleString("ca-ES",{minimumFractionDigits:2})+" €":"—";
    const classif=(r.classificacio_requerida||[]).map(c=>`${c.grup}${c.subgrup} Cat.${c.categoria}`).join(" | ")||"—";
    const critsAuto=(r.criteris_automatics||[]).map(c=>`<tr><td>${c.nom}</td><td class="num">${c.punts} pts</td><td class="small">${c.formula||""}</td></tr>`).join("");
    const critsJV=(r.criteris_judici_valor||[]).map(c=>`<tr><td>${c.nom}</td><td class="num">${c.punts} pts</td></tr>`).join("");
    const ptsAuto=(r.criteris_automatics||[]).reduce((s,c)=>s+(c.punts||0),0);
    const ptsJV=(r.criteris_judici_valor||[]).reduce((s,c)=>s+(c.punts||0),0);
    const informeText=(rawText||"").replace(/--JSON_INICI--[\s\S]*?--JSON_FI--/,"").trim();
    const informeHtml=mdToHtml(informeText);
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Anàlisi Plec - ${r.objecte||nom||""}</title>
<style>
@page{size:A4;margin:18mm 15mm}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:11px;color:#333;line-height:1.6;max-width:100%}
.header{background:linear-gradient(135deg,#1e3a5f,#2d5a8e);color:white;padding:20px 24px;border-radius:8px;margin-bottom:16px}
.header h1{font-size:18px;margin:0 0 4px 0;border:none;padding:0;color:white}
.header .meta{color:#a8c8e8;font-size:10px;margin:0}
.header .expedient{color:#d4e8f7;font-size:12px;margin-top:6px}
h2{font-size:13px;color:#1e3a5f;margin:20px 0 8px 0;padding:6px 10px;background:#f0f4f8;border-left:3px solid #1e3a5f;border-radius:0 4px 4px 0}
h3{font-size:12px;color:#2d5a8e;margin:14px 0 6px 0;padding-bottom:3px;border-bottom:1px dotted #ccc}
h4{font-size:11px;color:#444;margin:10px 0 4px 0;font-weight:600}
table{border-collapse:collapse;width:100%;margin:8px 0;border-radius:4px;overflow:hidden}
thead tr{background:#1e3a5f;color:white}
th{padding:6px 10px;text-align:left;font-size:10px;font-weight:600}
td{padding:5px 10px;border-bottom:1px solid #e8e8e8;font-size:10px}
tr:nth-child(even){background:#f8fafc}
tr:hover{background:#eef3f8}
td.num{text-align:right;font-weight:600;color:#1e3a5f}
td.small{font-size:9px;color:#666}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:10px;font-weight:bold;margin:2px}
.badge-ok{background:#d1fae5;color:#065f46}
.badge-ko{background:#fee2e2;color:#991b1b}
.badge-parcial{background:#fef3c7;color:#92400e}
.diag{background:linear-gradient(135deg,#fffbeb,#fef3c7);border:1px solid #f59e0b;border-radius:8px;padding:14px 16px;margin:16px 0}
.diag strong{color:#92400e;font-size:12px}
.diag p{margin:6px 0 0 0;color:#78350f}
.informe{margin-top:12px;page-break-before:always}
.informe h2{background:#f0f4f8}
.informe p{margin:4px 0;text-align:justify}
.informe ul{margin:4px 0 8px 16px;padding:0}
.informe li{margin:2px 0;list-style:disc}
.informe hr{border:none;border-top:1px solid #ddd;margin:12px 0}
.section-card{background:white;border:1px solid #e5e7eb;border-radius:8px;padding:14px 16px;margin:8px 0;box-shadow:0 1px 3px rgba(0,0,0,0.05)}
.footer{text-align:center;color:#999;font-size:9px;margin-top:24px;padding-top:8px;border-top:1px solid #eee}
strong{color:#1e3a5f}
</style></head><body>
<div class="header">
<h1>📋 Anàlisi de Plec</h1>
<div class="meta">Generat el ${date} · Servial Construccions</div>
${r.expedient?`<div class="expedient">Exp. ${r.expedient}</div>`:""}
</div>
<div class="section-card">
<h2>📌 Dades bàsiques</h2>
<table><tbody>
${r.objecte?`<tr><th width="30%">Objecte</th><td>${r.objecte}</td></tr>`:""}
${r.organisme?`<tr><th>Organisme</th><td>${r.organisme}</td></tr>`:""}
${r.cpv?`<tr><th>CPV</th><td>${r.cpv}</td></tr>`:""}
<tr><th>Import s/IVA</th><td class="num">${fmtI(r.import_sense_iva)}</td></tr>
<tr><th>Import c/IVA</th><td class="num">${fmtI(r.import_amb_iva)}</td></tr>
${r.valor_estimat?`<tr><th>Valor estimat</th><td class="num">${fmtI(r.valor_estimat)}</td></tr>`:""}
${r.termini_execucio?`<tr><th>Termini execució</th><td>${r.termini_execucio}</td></tr>`:""}
${r.termini_presentacio?`<tr><th>⏰ Termini presentació</th><td><strong>${r.termini_presentacio}</strong></td></tr>`:""}
</tbody></table>
</div>
<div class="section-card">
<h2>🏗️ Classificació i solvència</h2>
<p><strong>Classificació requerida:</strong> <span class="badge" style="background:#e0e7ff;color:#3730a3">${classif}</span></p>
${r.classificacio_substitueix_solvencia?`<p style="color:#065f46;font-size:10px">✅ La classificació substitueix la solvència</p>`:""}
${r.solvencia_economica?`<p><strong>Solvència econòmica:</strong> ${r.solvencia_economica}</p>`:""}
${r.solvencia_tecnica?`<p><strong>Solvència tècnica:</strong> ${r.solvencia_tecnica}</p>`:""}
</div>
${(critsAuto||critsJV)?`<div class="section-card">
<h2>📊 Criteris d'adjudicació</h2>
<p style="font-size:10px;color:#666">Total: <strong>${ptsAuto+ptsJV} punts</strong> · Automàtics: ${ptsAuto} pts (${Math.round(ptsAuto/(ptsAuto+ptsJV)*100)}%) · Judici de valor: ${ptsJV} pts (${Math.round(ptsJV/(ptsAuto+ptsJV)*100)}%)</p>
${critsAuto?`<table><thead><tr><th>Criteri automàtic</th><th width="60">Punts</th><th>Fórmula</th></tr></thead><tbody>${critsAuto}</tbody></table>`:""}
${critsJV?`<table><thead><tr><th>Criteri judici de valor</th><th width="60">Punts</th></tr></thead><tbody>${critsJV}</tbody></table>`:""}
</div>`:""}
<div class="section-card">
<h2>🔒 Garanties i condicions</h2>
${r.garantia_definitiva?`<p><strong>Garantia definitiva:</strong> ${r.garantia_definitiva}</p>`:""}
${r.garantia_provisional?`<p><strong>Garantia provisional:</strong> ${r.garantia_provisional}</p>`:""}
${r.visita_obra?`<p><strong>Visita d'obra:</strong> ${r.visita_obra}</p>`:""}
${r.condicions_especials?`<p><strong>Condicions especials:</strong> ${r.condicions_especials}</p>`:""}
${r.penalitats?`<p><strong>Penalitats:</strong> ${r.penalitats}</p>`:""}
</div>
${r.diagnosic_servial?`<div class="diag"><strong>⚠️ Diagnòstic per a Servial</strong><p>${r.diagnosic_servial}</p></div>`:""}
${informeText?`<div class="informe"><h2>📝 Informe complet de l'anàlisi</h2>${informeHtml}</div>`:""}
<div class="footer">SERVIAL Construccions · Anàlisi de plec generat automàticament · ${date}</div>
</body></html>`;
    const win=window.open("","_blank");if(!win){alert("Activa les finestres emergents.");return;}
    win.document.write(html);win.document.close();win.focus();setTimeout(()=>win.print(),600);
  };

  const totAptes=results.filter(r=>r._compat?.pot!==false).length;
  const totInsuf=results.filter(r=>r._compat?.pot===false).length;
  const visibleResults=filters.nomesPotPresentar?results.filter(r=>r._compat?.pot!==false):filters.nomesSuperiors?results.filter(r=>r._compat?.pot===false):results;

  return(
    <div className="min-h-screen bg-stone-300 text-sm">
      <div className="bg-blue-900 text-white px-6 py-4 flex items-center justify-between shadow-lg">
        <div className="flex items-center gap-2"><span className="text-2xl">🔍</span><div><h1 className="text-lg font-bold leading-none">Radar + Gestor de Licitacions</h1><p className="text-blue-300 text-xs mt-0.5">PSCP · CIDO-DIBA · PLACE · Gmail · Catalunya</p></div></div>
        <div className="flex items-center gap-3">
          <button onClick={()=>setShowApiKey(v=>!v)} className={`text-xs px-3 py-1.5 rounded-lg font-medium ${apiKeySaved?"bg-green-600 hover:bg-green-700 text-white":"bg-red-500 hover:bg-red-600 text-white animate-pulse"}`}>{apiKeySaved?`🤖 ${aiProvider==="claude"?"Claude":"Gemini"} ✓`:"🔑 Configura IA"}</button>
          <div className="text-right text-xs text-blue-200"><div className="font-semibold text-white text-sm">Servial</div><div>{new Date().toLocaleDateString("ca-ES",{day:"numeric",month:"long",year:"numeric"})}</div></div>
        </div>
      </div>
      {showApiKey&&<div className="bg-yellow-50 border-b border-yellow-200 px-6 py-4">
        <div className="max-w-6xl mx-auto space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="font-semibold text-gray-700 text-sm">🤖 Configuració IA — Proveïdor d'anàlisi</h3>
            <button onClick={()=>setShowApiKey(false)} className="text-gray-400 hover:text-gray-600 text-lg">✕</button>
          </div>
          <div className="flex gap-2">
            <button onClick={()=>{setAiProvider("gemini");setAiProviderLS("gemini");}} className={`flex-1 px-4 py-2.5 rounded-lg text-xs font-semibold border-2 transition-all ${aiProvider==="gemini"?"border-blue-600 bg-blue-50 text-blue-800 shadow-sm":"border-gray-200 bg-stone-50 text-gray-500 hover:border-gray-300"}`}>
              <div className="text-base mb-0.5">🟢 Google Gemini</div>
              <div className="font-normal opacity-75">Gratuït · Bona qualitat</div>
            </button>
            <button onClick={()=>{setAiProvider("claude");setAiProviderLS("claude");}} className={`flex-1 px-4 py-2.5 rounded-lg text-xs font-semibold border-2 transition-all ${aiProvider==="claude"?"border-purple-600 bg-purple-50 text-purple-800 shadow-sm":"border-gray-200 bg-stone-50 text-gray-500 hover:border-gray-300"}`}>
              <div className="text-base mb-0.5">🟣 Claude (Anthropic)</div>
              <div className="font-normal opacity-75">De pagament · Qualitat superior</div>
            </button>
          </div>
          {aiProvider==="gemini"&&<div className="flex items-center gap-3">
            <label className="text-xs font-semibold text-gray-600 shrink-0">Google Gemini API Key:</label>
            <input type="password" value={geminiKeyInput} onChange={e=>setGeminiKeyInput(e.target.value)} placeholder="AIza..." className="flex-1 border rounded-lg px-3 py-1.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-300"/>
            <button onClick={()=>{setGeminiKey(geminiKeyInput);setApiKeySaved(!!geminiKeyInput||!!getApiKey());setShowApiKey(false);}} className="bg-blue-700 hover:bg-blue-800 text-white text-xs font-semibold px-4 py-1.5 rounded-lg">Guardar</button>
          </div>}
          {aiProvider==="claude"&&<div className="flex items-center gap-3">
            <label className="text-xs font-semibold text-gray-600 shrink-0">Anthropic API Key:</label>
            <input type="password" value={apiKeyInput} onChange={e=>setApiKeyInput(e.target.value)} placeholder="sk-ant-..." className="flex-1 border rounded-lg px-3 py-1.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-300"/>
            <button onClick={()=>{setApiKey(apiKeyInput);setApiKeySaved(!!apiKeyInput||!!getGeminiKey());setShowApiKey(false);}} className="bg-purple-700 hover:bg-purple-800 text-white text-xs font-semibold px-4 py-1.5 rounded-lg">Guardar</button>
          </div>}
          <p className="text-xs text-gray-400">Les claus es guarden al localStorage del navegador. <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" className="text-blue-500 underline">Obtenir clau Gemini gratuïta</a> · <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer" className="text-purple-500 underline">Obtenir clau Claude</a></p>
        </div>
      </div>}
      <div className="bg-stone-50 border-b shadow-sm">
        <div className="max-w-6xl mx-auto px-4 flex">
          {[["filtres","⚙️ Criteris"],["resultats",`📋 Resultats${results.length?` (${results.length})`:""}`],["plecs","📄 Anàlisi Plecs"],["gestio","📁 Gestió"]].map(([id,label])=>(
            <button key={id} onClick={()=>setActiveTab(id)} className={`py-3 px-5 font-medium border-b-2 transition-colors ${activeTab===id?"border-blue-700 text-blue-700":"border-transparent text-gray-500 hover:text-gray-700"}`}>{label}</button>
          ))}
        </div>
      </div>
      {activeTab==="gestio"&&<GestorTab refreshKey={gestorRefreshKey}/>}
      {activeTab!=="gestio"&&(
      <div className="max-w-6xl mx-auto px-4 py-5 space-y-4">
        {activeTab==="filtres"&&(<>
          <div className="bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 text-xs text-blue-700 flex items-center gap-2"><span>💡</span><span>Cerca licitacions d'obres publicades les últimes 72h amb termini obert. Fonts: Transparència Catalunya + CIDO-DIBA (importació manual).</span></div>
          <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
            <div className="flex items-center justify-between mb-3"><h3 className="font-semibold text-gray-700">💰 Import del Contracte</h3><span className="text-xs text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full">Màxim Servial (Cat.4): 2.400.000 €</span></div>
            <div className="mb-3"><button onClick={()=>setFilters(p=>({...p,importMin:"",importMax:""}))} className={`text-xs px-3 py-1.5 rounded-full border font-semibold ${!filters.importMin&&!filters.importMax?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>📋 Totes (sense límit)</button></div>
            <div className="mb-3"><p className="text-xs text-gray-400 mb-1.5">Límits per categoria:</p><div className="flex flex-wrap gap-1.5">{Object.entries(CAT_LIMITS).map(([cat,limit])=>{const maxCat=Math.max(...Object.values(SERVIAL_CLASS)),isMax=parseInt(cat)===maxCat,isSel=filters.importMax===String(limit||""),disabled=parseInt(cat)>maxCat+1;return<button key={cat} disabled={disabled} onClick={()=>setFilters(p=>({...p,importMax:limit?String(limit):""}))} className={`text-xs px-2.5 py-1 rounded-full border font-medium ${disabled?"opacity-30 cursor-not-allowed bg-gray-100 text-gray-400 border-gray-200":isSel?"bg-blue-700 text-white border-blue-700":isMax?"bg-green-50 text-green-700 border-green-400":"bg-stone-50 text-gray-600 border-gray-300"}`}>{isMax&&"★ "}Cat.{cat} — {CAT_LABELS[cat]}</button>;})}
            </div></div>
            <div className="flex gap-3">{[["importMin","Import mínim (€)"],["importMax","Import màxim (€)"]].map(([k,label])=>(<div key={k} className="flex-1"><label className="text-xs text-gray-500 block mb-1">{label}</label><input type="number" value={filters[k]} onChange={e=>setFilters(p=>({...p,[k]:e.target.value}))} className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/></div>))}</div>
          </div>
          <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
            <h3 className="font-semibold text-gray-700 mb-3">🏆 Classificació Servial</h3>
            <div className="flex flex-col gap-2">
              <button onClick={()=>setFilters(p=>({...p,nomesPotPresentar:false,nomesSuperiors:false}))} className={`w-full flex items-center justify-center gap-2 text-sm font-semibold px-4 py-2.5 rounded-xl border ${!filters.nomesPotPresentar&&!filters.nomesSuperiors?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>📋 Totes les licitacions</button>
              <button disabled={results.length===0} onClick={()=>{setFilters(p=>({...p,nomesPotPresentar:true,nomesSuperiors:false}));setActiveTab("resultats");}} className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white font-semibold text-sm px-4 py-2.5 rounded-xl">🔍 Probable compliment Servial {results.length>0&&<span className="bg-stone-50 text-blue-700 font-bold text-xs px-2 py-0.5 rounded-full">{totAptes}</span>}</button>
              <button disabled={results.length===0} onClick={()=>{setFilters(p=>({...p,nomesSuperiors:true,nomesPotPresentar:false}));setActiveTab("resultats");}} className="w-full flex items-center justify-center gap-2 bg-amber-500 hover:bg-amber-600 disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white font-semibold text-sm px-4 py-2.5 rounded-xl">⚠️ Probable classif. insuficient {results.length>0&&<span className="bg-stone-50 text-slate-500 font-bold text-xs px-2 py-0.5 rounded-full">{totInsuf}</span>}</button>
              {results.length===0&&<p className="text-xs text-gray-400 text-center">Els comptadors s'activen després de cercar</p>}
            </div>
          </div>
          <div className="bg-stone-50 rounded-xl border shadow-sm p-4"><h3 className="font-semibold text-gray-700 mb-2">🔤 Paraules Clau</h3><input type="text" placeholder="ex. pavimentació, clavegueram, parc…" value={filters.paraulesClau} onChange={e=>setFilters(p=>({...p,paraulesClau:e.target.value}))} className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/></div>
          <div className="bg-purple-50 rounded-xl border border-purple-200 shadow-sm p-4">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2"><span className="text-lg">📧</span><h3 className="font-semibold text-purple-800">Correu CIDO-DIBA</h3></div>
              <div className="flex items-center gap-2">
                {cidoCount!==null&&<span className="text-xs font-semibold text-purple-600 bg-purple-100 px-2 py-0.5 rounded-full">{cidoCount} obres importades</span>}
                <button onClick={()=>setShowCidoImport(v=>!v)} className="bg-purple-600 hover:bg-purple-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg">{showCidoImport?"Tancar":"Importar correu"}</button>
              </div>
            </div>
            <div className={`rounded-lg px-3 py-2 mb-2 flex items-center justify-between ${gmailConnected?"bg-green-50 border border-green-200":"bg-amber-50 border border-slate-200"}`}>
              {gmailConnected?(<>
                <span className="text-xs text-green-700">✅ Gmail connectat: <strong>{gmailUser||"..."}</strong> — els correus de CIDO i Gencat es llegiran automàticament</span>
                <button onClick={disconnectGmail} className="text-xs text-red-500 hover:text-red-700 font-medium ml-2 shrink-0">Desconnectar</button>
              </>):(<>
                <span className="text-xs text-slate-600">Connecta Gmail per llegir automàticament els correus de licitacions</span>
                <button onClick={connectGmail} className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg ml-2 shrink-0">🔗 Connectar Gmail</button>
              </>)}
            </div>
            <p className="text-xs text-purple-600 mb-2">{gmailConnected?"O també pots importar manualment:":"Sense Gmail, pots importar manualment:"} Obre el correu de <strong>sgsi.svcSam@diba.cat</strong> al Gmail, selecciona tot el contingut (Ctrl+A), copia (Ctrl+C) i enganxa-ho aquí.</p>
            {showCidoImport&&<div className="space-y-2">
              <div className="w-full border-2 border-dashed border-purple-300 rounded-lg px-4 py-6 text-center cursor-text bg-stone-50 focus-within:ring-2 focus-within:ring-purple-400 focus-within:border-purple-400"
                contentEditable suppressContentEditableWarning
                onPaste={e=>{e.preventDefault();const html=e.clipboardData.getData("text/html")||e.clipboardData.getData("text/plain");if(html){const results=parseCIDOToResults(html);setCidoCount(results.length);setCidoCache(results);setShowCidoImport(false);if(results.length>0)setDebugInfo(`✅ CIDO: ${results.length} obres importades del correu`);else setDebugInfo("⚠️ CIDO: no s'han trobat obres al contingut enganxat");}}}
              ><span className="text-purple-400 text-sm pointer-events-none">{cidoCount!==null&&cidoCount>0?"✅ Enganxa un nou correu per actualitzar":"Fes Ctrl+V aquí per enganxar el correu CIDO"}</span></div>
              <p className="text-xs text-purple-500">Obre el correu CIDO al Gmail → Ctrl+A (selecciona tot) → Ctrl+C (copia) → clica aquí i Ctrl+V (enganxa). Es processarà automàticament.</p>
            </div>}
            {cidoCount!==null&&cidoCount>0&&!showCidoImport&&<p className="text-xs text-purple-500 mt-1">✅ Dades CIDO importades avui — s'inclouran a la propera cerca.</p>}
          </div>
          <div className="bg-green-50 rounded-xl border border-green-200 shadow-sm p-4">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2"><span className="text-lg">📧</span><h3 className="font-semibold text-green-800">Correu Gencat (plataforma.contractacio)</h3></div>
              <div className="flex items-center gap-2">
                {gencatCount!==null&&<span className="text-xs font-semibold text-green-600 bg-green-100 px-2 py-0.5 rounded-full">{gencatCount} obres importades</span>}
                <button onClick={()=>setShowGencatImport(v=>!v)} className="bg-green-600 hover:bg-green-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg">{showGencatImport?"Tancar":"📋 Importar correu"}</button>
              </div>
            </div>
            <p className="text-xs text-green-600 mb-2">Obre el correu de <strong>plataforma.contractacio@gencat.cat</strong> ("Correu diari de subscriptors generals"), selecciona tot el contingut (Ctrl+A), copia (Ctrl+C) i enganxa-ho aquí.</p>
            {showGencatImport&&<div className="space-y-2">
              <div className="w-full border-2 border-dashed border-green-300 rounded-lg px-4 py-6 text-center cursor-text bg-stone-50 focus-within:ring-2 focus-within:ring-green-400 focus-within:border-green-400"
                contentEditable suppressContentEditableWarning
                onPaste={e=>{e.preventDefault();const text=e.clipboardData.getData("text/plain");if(text){const results=parseGencatEmail(text);setGencatCount(results.length);setGencatCache(results);setShowGencatImport(false);if(results.length>0)setDebugInfo(`✅ Gencat: ${results.length} obres importades del correu`);else setDebugInfo("⚠️ Gencat: no s'han trobat obres al contingut enganxat");}}}
              ><span className="text-green-400 text-sm pointer-events-none">{gencatCount!==null&&gencatCount>0?"✅ Enganxa un nou correu per actualitzar":"Fes Ctrl+V aquí per enganxar el correu Gencat"}</span></div>
              <p className="text-xs text-green-500">Obre el correu Gencat → Ctrl+A (selecciona tot) → Ctrl+C (copia) → clica aquí i Ctrl+V (enganxa).</p>
            </div>}
            {gencatCount!==null&&gencatCount>0&&!showGencatImport&&<p className="text-xs text-green-500 mt-1">✅ Dades Gencat importades avui — s'inclouran a la propera cerca.</p>}
          </div>
          <div className={`rounded-xl border shadow-sm p-4 ${scheduleEnabled?"bg-blue-900 border-blue-700":"bg-stone-50"}`}>
            <div className="flex items-center justify-between mb-3"><div className="flex items-center gap-2"><span className={scheduleEnabled?"animate-pulse text-lg":"opacity-50 text-lg"}>🔔</span><h3 className={`font-semibold ${scheduleEnabled?"text-white":"text-gray-700"}`}>Cerca Automàtica Diària</h3></div><button onClick={()=>setScheduleEnabled(v=>!v)} className={`relative inline-flex h-6 w-11 items-center rounded-full ${scheduleEnabled?"bg-blue-400":"bg-gray-300"}`}><span className={`inline-block h-4 w-4 transform rounded-full bg-stone-50 shadow transition-transform ${scheduleEnabled?"translate-x-6":"translate-x-1"}`}/></button></div>
            <div className="flex items-center gap-4"><div><label className={`text-xs block mb-1 ${scheduleEnabled?"text-blue-200":"text-gray-500"}`}>Hora d'execució</label><input type="time" value={scheduleTime} onChange={e=>setScheduleTime(e.target.value)} className="border rounded-lg px-3 py-1.5 text-sm bg-stone-50"/></div>{scheduleEnabled&&<div className="flex-1"><p className="text-blue-200 text-xs mb-0.5">Propera execució en</p><p className="text-white font-bold text-lg">{countdown||"calculant…"}</p></div>}</div>
            {scheduleEnabled&&<div className="mt-3 space-y-1"><p className="text-xs text-blue-200">✅ Cada dia a les <strong className="text-white">{scheduleTime}</strong>.</p><p className="text-xs text-amber-300">⚠️ L'aplicació ha d'estar oberta al navegador.</p>{scheduleStatus&&<p className="text-xs text-green-300">{scheduleStatus}</p>}</div>}
          </div>
          <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
            <h3 className="font-semibold text-gray-700 mb-3">📧 Destinataris del Correu</h3>
            <div className="space-y-1.5 mb-3">{recipients.map(r=>(<div key={r.email} className={`flex items-center gap-2 px-3 py-2 rounded-lg border ${r.selected?"bg-blue-50 border-blue-200":"bg-gray-50 border-gray-200"}`}><input type="checkbox" checked={r.selected} onChange={()=>toggleRecipient(r.email)} className="rounded text-blue-600 shrink-0"/><span className={`flex-1 text-xs ${r.selected?"text-blue-800 font-medium":"text-gray-400"}`}>{r.email}</span><button onClick={()=>removeRecipient(r.email)} className="text-gray-300 hover:text-red-400">✕</button></div>))}</div>
            <div className="flex gap-2"><div className="flex-1"><input type="email" placeholder="nou@empresa.cat" value={newEmail} onChange={e=>{setNewEmail(e.target.value);setEmailError("");}} onKeyDown={e=>e.key==="Enter"&&addRecipient()} className={`w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 ${emailError?"border-red-400":"border-gray-300 focus:ring-blue-300"}`}/>{emailError&&<p className="text-xs text-red-500 mt-0.5">{emailError}</p>}</div><button onClick={addRecipient} className="shrink-0 bg-blue-700 hover:bg-blue-800 text-white text-xs font-semibold px-4 py-2 rounded-lg">+ Afegir</button></div>
          </div>
          <button onClick={()=>setShowClassif(v=>!v)} className="w-full flex items-center justify-center gap-2 bg-stone-50 border border-gray-300 hover:border-blue-400 text-gray-600 font-medium text-sm py-2.5 rounded-xl">{showClassif?"▲ Amagar":"▼ Classificació acreditada Servial"}</button>
          {showClassif&&<div className="space-y-3">{[{grup:"A",label:"Grup A — Moviment de terres",items:["A1"]},{grup:"C",label:"Grup C — Edificació",items:["C1","C2","C3","C4","C5","C6","C7","C8","C9"]},{grup:"E",label:"Grup E — Hidràulics",items:["E1","E4","E5","E7"]},{grup:"G",label:"Grup G — Vials i pistes",items:["G3","G5","G6"]},{grup:"K",label:"Grup K — Especials",items:["K6"]}].map(({grup,label,items})=>(<div key={grup} className="bg-stone-50 rounded-xl border shadow-sm p-4"><h3 className="font-semibold text-gray-700 mb-2">{label}</h3><div className="grid grid-cols-1 md:grid-cols-2 gap-2">{items.map(k=>(<div key={k} className="flex items-center gap-3 bg-slate-50 rounded-lg px-3 py-2"><span className="font-bold text-blue-700 w-8 shrink-0">{k}</span><span className="flex-1 text-gray-700 text-xs">{SERVIAL_CLASS_LABELS[k]}</span><span className="bg-blue-100 text-blue-800 font-bold text-xs px-2 py-0.5 rounded-full shrink-0">Cat.{SERVIAL_CLASS[k]}</span></div>))}</div></div>))}</div>}
          <button onClick={()=>setShowAdvanced(v=>!v)} className="w-full flex items-center justify-center gap-2 bg-stone-50 border border-gray-300 hover:border-blue-400 text-gray-600 font-medium text-sm py-2.5 rounded-xl">{showAdvanced?"▲ Amagar filtres avançats":"▼ Filtres avançats (CPV i àmbit geogràfic)"}</button>
          {showAdvanced&&(<>
            <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
              <div className="flex items-center justify-between mb-3"><h3 className="font-semibold text-gray-700">📍 Àmbit Geogràfic <span className="text-xs font-normal ml-1">{filters.comarques.length===0?<span className="text-blue-600">(tot Catalunya)</span>:`${filters.comarques.length} sel.`}</span></h3><button onClick={()=>setFilters(p=>({...p,comarques:p.comarques.length===COMARQUES.length?[]:COMARQUES}))} className={`text-xs px-2.5 py-1 rounded-full border font-medium ${filters.comarques.length===COMARQUES.length?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>{filters.comarques.length===COMARQUES.length?"✓ Tot Catalunya":"Tot Catalunya"}</button></div>
              <div className="flex gap-2 mb-3">{PROVINCIES.map(p=>{const allSel=p.comarques.every(c=>filters.comarques.includes(c)),someSel=!allSel&&p.comarques.some(c=>filters.comarques.includes(c));return<button key={p.nom} onClick={()=>setFilters(f=>({...f,comarques:allSel?f.comarques.filter(c=>!p.comarques.includes(c)):[...new Set([...f.comarques,...p.comarques])]}))} className={`flex-1 py-1.5 rounded-lg text-xs font-semibold border ${allSel?"bg-blue-700 text-white border-blue-700":someSel?"bg-blue-100 text-blue-700 border-blue-400":"bg-stone-50 text-gray-600 border-gray-300"}`}>{allSel?"✓ ":someSel?"– ":""}{p.nom}</button>;})}
              </div>
              <div className="grid grid-cols-4 gap-x-4 max-h-52 overflow-y-auto">{PROVINCIES.map(p=>(<div key={p.nom}><p className="text-xs font-semibold text-gray-400 uppercase tracking-wide mb-1">{p.nom}</p>{p.comarques.map(c=>(<label key={c} className="flex items-center gap-1 cursor-pointer py-0.5 hover:bg-gray-50 rounded px-0.5"><input type="checkbox" checked={filters.comarques.includes(c)} onChange={()=>toggle("comarques",c)} className="rounded text-blue-600 w-3 h-3 shrink-0"/><span className="text-xs text-gray-700 leading-tight">{c}</span></label>))}</div>))}</div>
            </div>
            <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
              <div className="flex items-center justify-between mb-2"><h3 className="font-semibold text-gray-700">🔢 Codis CPV</h3><button onClick={()=>setFilters(p=>({...p,cpvCodes:p.cpvCodes.length===CPV_OPTS.length?[]:CPV_OPTS.map(c=>c.code)}))} className={`text-xs px-2.5 py-1 rounded-full border font-medium ${filters.cpvCodes.length===CPV_OPTS.length?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>{filters.cpvCodes.length===CPV_OPTS.length?"✓ Tots":"Seleccionar tots"}</button></div>
              <div className="flex flex-wrap gap-1.5 mb-3">{CPV_GROUPS.map(g=>{const codes=g.items.map(i=>i.code),allSel=codes.every(c=>filters.cpvCodes.includes(c)),someSel=!allSel&&codes.some(c=>filters.cpvCodes.includes(c));return<button key={g.label} onClick={()=>setFilters(p=>({...p,cpvCodes:allSel?p.cpvCodes.filter(c=>!codes.includes(c)):[...new Set([...p.cpvCodes,...codes])]}))} className={`text-xs px-2.5 py-1 rounded-full border font-medium ${allSel?"bg-blue-700 text-white border-blue-700":someSel?"bg-blue-100 text-blue-700 border-blue-400":"bg-stone-50 text-gray-600 border-gray-300"}`}>{allSel?"✓ ":someSel?"– ":""}{g.label.replace(/^.{2}/,"")}</button>;})}
              </div>
              <div className="space-y-3 max-h-64 overflow-y-auto pr-1">{CPV_GROUPS.map(g=>(<div key={g.label}><p className="text-xs font-semibold text-gray-400 uppercase tracking-wide mb-1">{g.label}</p>{g.items.map(c=>(<label key={c.code} className="flex items-center gap-2 cursor-pointer py-0.5 pl-1"><input type="checkbox" checked={filters.cpvCodes.includes(c.code)} onChange={()=>toggle("cpvCodes",c.code)} className="rounded text-blue-600"/><span className="text-xs text-gray-700">{c.label}</span></label>))}</div>))}</div>
            </div>
          </>)}
          {debugInfo&&<div className="bg-gray-900 text-green-300 rounded-xl p-3 text-xs font-mono whitespace-pre-wrap break-all max-h-48 overflow-y-auto">🔧 {debugInfo}</div>}
          <button onClick={()=>buscarIEnviar(false)} disabled={loading} className="w-full bg-blue-800 hover:bg-blue-900 disabled:bg-blue-300 text-white font-semibold py-3.5 rounded-xl text-base flex items-center justify-center gap-2 shadow">
            {loading?<><svg className="animate-spin h-5 w-5" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>{statusMsg}</>:"🔍 Cercar Licitacions"}
          </button>
        </>)}
        {activeTab==="resultats"&&(<>
          {results.length>0&&<div className="grid grid-cols-3 gap-3"><div className="bg-stone-50 rounded-xl border shadow-sm p-3 text-center"><div className="text-2xl font-bold text-gray-800">{results.length}</div><div className="text-xs text-gray-500">Licitacions trobades</div></div><div className="bg-blue-50 rounded-xl border border-blue-200 shadow-sm p-3 text-center"><div className="text-2xl font-bold text-blue-700">{totAptes}</div><div className="text-xs text-blue-600">Probable compliment</div></div><div className="bg-amber-50 rounded-xl border border-slate-200 shadow-sm p-3 text-center"><div className="text-2xl font-bold text-slate-500">{totInsuf}</div><div className="text-xs text-slate-400">Probable classif. insuficient</div></div></div>}
          <div className="bg-stone-50 rounded-xl border shadow-sm px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
            <div className="flex items-center gap-3"><span className="text-gray-600 text-sm">{statusMsg||"Configura els filtres i cerca."}</span>{results.length>0&&<label className="flex items-center gap-1.5 cursor-pointer text-xs text-gray-600"><input type="checkbox" checked={filters.nomesPotPresentar} onChange={e=>setFilters(p=>({...p,nomesPotPresentar:e.target.checked,nomesSuperiors:false}))} className="rounded text-blue-600"/>Probable compliment</label>}</div>
            <div className="flex gap-2 shrink-0">
              <button onClick={()=>buscarIEnviar(false)} disabled={loading} className="bg-blue-700 hover:bg-blue-800 disabled:bg-gray-300 text-white text-xs font-medium px-3 py-2 rounded-lg">{loading?"…":"🔄 Cercar"}</button>
              <button onClick={exportarResultatsExcel} disabled={!results.length} className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-gray-300 text-white text-xs font-medium px-3 py-2 rounded-lg">📥 Excel</button>
              <button onClick={enviarEmail} disabled={!results.length||selectedEmails.length===0} className="bg-orange-500 hover:bg-orange-600 disabled:bg-gray-300 text-white text-xs font-medium px-3 py-2 rounded-lg">{emailStatus||"📧 Email"}</button>
            </div>
          </div>
          {results.length===0?<div className="bg-stone-50 rounded-xl border shadow-sm p-12 text-center text-gray-400"><div className="text-5xl mb-3">📋</div><p>Configura els criteris i clica <strong>Cercar</strong>.</p></div>
            :<div className="space-y-2">{visibleResults.map((r,i)=>(<div key={i} onClick={()=>setExpanded(expanded===i?null:i)} className={`bg-stone-50 rounded-xl border shadow-sm p-4 cursor-pointer hover:shadow-md transition-shadow ${r._compat?.pot!==false?"border-l-4 border-l-green-400":r._compat?.pot===false?"border-l-4 border-l-red-300":""}`}>
              <div className="flex items-start gap-3">
                <div className="flex-1 min-w-0">
                  <div className="flex flex-wrap items-center gap-1.5 mb-1"><Badge score={r.puntuacio}/><CompatBadge pot={r._compat?.pot}/><span className="text-xs text-gray-400">{r.expedient}</span>{r.justificacio?.includes("CIDO")&&<span className="text-xs bg-purple-100 text-purple-700 px-2 py-0.5 rounded-full font-semibold">📧 CIDO</span>}</div>
                  <p className="font-semibold text-gray-800 leading-snug">{r.objecte}</p>
                  <p className="text-xs text-gray-500 mt-0.5">{r.organisme}</p>
                  {r.classificacio_requerida?.length>0&&<p className="text-xs text-gray-400 mt-0.5">Classif.: {r.classificacio_requerida.map(c=>`${c.grup}${c.subgrup} Cat.${c.categoria}`).join(" | ")}</p>}
                  {expanded===i&&<div className="mt-2 space-y-2"><p className="text-xs text-gray-500 italic bg-gray-50 rounded p-2">{r.justificacio}</p>{r._compat?.items?.length>0&&<div><p className="text-xs font-semibold text-gray-600 mb-1">Anàlisi classificació Servial:</p><CompatDetail items={r._compat.items}/></div>}</div>}
                </div>
                <div className="text-right text-xs shrink-0 space-y-1 min-w-36">
                  <div className="font-bold text-gray-800 text-base">{r.import_eur?`${r.import_eur.toLocaleString("ca-ES")} €`:"N/D"}</div>
                  <div className="text-gray-400">📢 <span className="font-medium text-gray-600">{r.data_publicacio||"—"}</span></div>
                  <div className="text-red-600 font-semibold">⏰ {r.termini||"—"}</div>
                  <div className="text-gray-500">📍 {r.comarca_municipi||"—"}</div>
                  {r.font&&<div className="bg-blue-50 text-blue-700 font-semibold px-1.5 py-0.5 rounded text-xs text-center">{r.font}</div>}
                  {r.url&&<a href={r.url} target="_blank" rel="noopener noreferrer" onClick={e=>e.stopPropagation()} className="text-blue-600 hover:underline">Veure →</a>}
                </div>
              </div>
              {expanded!==i&&r._compat?.pot===false&&r._compat?.items?.length>0&&<p className="text-xs text-red-500 mt-1.5">Manca: {r._compat.items.filter(it=>it.status!=="ok").map(it=>it.status==="absent"?`${it.codi} (no acreditat)`:`${it.codi} Cat.${it.catServial}→${it.catReq}`).join(", ")}</p>}
            </div>))}</div>}
        </>)}
        {activeTab==="plecs"&&(<>
          <div className="bg-slate-50 rounded-xl border border-slate-300 shadow-sm p-4">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2"><span className="text-xl">⚖️</span><h3 className="font-semibold text-slate-800">Consultes Legals</h3>{plecRawText&&<span className="text-xs text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">Plec carregat</span>}{plecFiles.length>0&&!plecRawText&&<span className="text-xs text-blue-600 bg-blue-100 px-2 py-0.5 rounded-full">{plecFiles.length} PDF pendent d'anàlisi</span>}</div>
              <div className="flex items-center gap-2">
                <label className="bg-slate-200 hover:bg-slate-300 text-slate-800 text-xs font-semibold px-3 py-1.5 rounded-lg cursor-pointer">📎 Pujar PDF per consulta
                  <input type="file" accept=".pdf,application/pdf" multiple className="hidden" onChange={e=>{if(e.target.files?.length){setPlecFiles(Array.from(e.target.files));setPlecResults([]);setPlecRawText("");setPlecError("");setPlecStatus("");}}}/>
                </label>
              </div>
            </div>
            <p className="text-xs text-slate-600 mb-3">{plecRawText?"Fes preguntes sobre el plec analitzat o sobre contractació pública en general.":"Puja un PDF i/o fes preguntes sobre contractació pública. Si analitzes el plec primer, les respostes seran contextualitzades."}</p>
            <div className="flex gap-2">
              <textarea rows={2} className="flex-1 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-400 resize-y" placeholder="Escriu la teva consulta legal... Ex: Quina classificació necessito? Puc subcontractar? Quina garantia definitiva cal?" value={legalQuery} onChange={e=>setLegalQuery(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey&&!legalLoading){e.preventDefault();consultarLegal();}}}/>
              <button onClick={consultarLegal} disabled={legalLoading||!legalQuery.trim()} className="bg-slate-700 hover:bg-slate-800 disabled:bg-slate-300 text-white font-semibold px-4 py-2 rounded-lg text-sm shrink-0 self-end">{legalLoading?"Consultant...":"⚖️ Consultar"}</button>
            </div>
            {legalResponses.length>0&&<div className="mt-3 space-y-2 max-h-96 overflow-y-auto">{legalResponses.map((lr,i)=>(<div key={i} className="bg-white rounded-lg border border-slate-200 overflow-hidden">
              <div className="bg-slate-100 px-3 py-1.5 flex items-center gap-2"><span className="text-slate-600 font-bold text-xs">CONSULTA</span><span className="text-xs text-slate-400">{lr.date}</span></div>
              <div className="px-3 py-2 bg-slate-50 border-b border-slate-100"><p className="text-xs font-medium text-slate-800">{lr.pregunta}</p></div>
              <div className="px-3 py-2"><div className="text-xs text-gray-700 leading-relaxed whitespace-pre-wrap">{lr.resposta}</div></div>
            </div>))}</div>}
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 text-xs text-blue-700 flex items-center gap-2"><span>📄</span><span>Puja el PCAP/PPT en PDF. La IA extraurà totes les dades clau i farà un diagnòstic per a Servial. Podràs guardar-ho al Gestor com a Proposta.</span></div>
          {plecHistorial.length>0&&<div className="bg-stone-50 rounded-xl border shadow-sm p-3">
            <div className="flex items-center justify-between cursor-pointer" onClick={()=>setPlecHistorialOpen(!plecHistorialOpen)}>
              <h4 className="text-sm font-semibold text-gray-700">📋 Últims {plecHistorial.length} anàlisis</h4>
              <span className="text-gray-400 text-xs">{plecHistorialOpen?"▲":"▼"}</span>
            </div>
            {plecHistorialOpen&&<div className="mt-2 space-y-1">{plecHistorial.map((h,idx)=>{const r=h.results?.[0]||{};const imp=r.import_sense_iva?Number(r.import_sense_iva).toLocaleString("ca-ES",{minimumFractionDigits:0})+" €":"";return(<div key={h.id} className="flex items-center gap-2 bg-gray-50 rounded-lg px-3 py-2 text-xs hover:bg-gray-100 border border-gray-200">
              <span className="bg-blue-100 text-blue-700 font-bold px-1.5 py-0.5 rounded text-xs shrink-0">{idx+1}</span>
              <span className="text-gray-400 shrink-0">{h.date}</span>
              <div className="flex-1 min-w-0"><div className="font-medium text-gray-800 truncate" title={r.objecte||h.nom}>{r.objecte||h.nom}</div><div className="text-gray-400 text-xs">{r.organisme?r.organisme+" · ":""}{imp}{r.termini_presentacio?` · ⏰ ${r.termini_presentacio}`:""}</div></div>
              <button onClick={()=>{setPlecResults(h.results||[]);setPlecRawText(h.rawText||"");setPlecNom(h.nom);setPlecView(h.results?.length?"taula":"text");}} className="text-blue-500 hover:text-blue-700 shrink-0" title="Veure">👁️</button>
              <button onClick={()=>exportarPlecPDF(h.results||[],h.rawText||"",h.nom)} className="text-green-500 hover:text-green-700 shrink-0" title="Descarregar PDF">📥</button>
              <button onClick={()=>eliminarPlecHistorial(h.id)} className="text-red-400 hover:text-red-600 shrink-0" title="Eliminar">🗑️</button>
            </div>);})}</div>}
          </div>}
          <div className="bg-stone-50 rounded-xl border shadow-sm p-4 space-y-4">
            <h3 className="font-semibold text-gray-700">📂 Documents de la licitació</h3>
            <div className="grid grid-cols-2 gap-3">
              <div><label className="text-xs text-gray-500 block mb-1">Nom / codi expedient (opcional)</label><input type="text" placeholder="ex. SCS-2026-270" value={plecNom} onChange={e=>setPlecNom(e.target.value)} className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/></div>
              <div><label className="text-xs text-gray-500 block mb-1">Òrgan de contractació (opcional)</label><input type="text" placeholder="ex. Ajuntament de Barcelona" value={plecOrgan} onChange={e=>setPlecOrgan(e.target.value)} className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/></div>
            </div>
            <div onClick={()=>fileInputRef.current?.click()} className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-colors ${plecFiles.length>0?"border-green-400 bg-green-50":"border-blue-300 hover:border-blue-500 bg-blue-50 hover:bg-blue-100"}`}>
              <div className="text-3xl mb-2">{plecFiles.length>0?"✅":"📎"}</div>
              <p className="text-sm font-medium text-blue-700">{plecFiles.length>0?`${plecFiles.length} fitxer${plecFiles.length!==1?"s":""} seleccionat${plecFiles.length!==1?"s":""}. Clica per canviar.`:"Clica per seleccionar PDFs"}</p>
              <p className="text-xs text-blue-400 mt-1">PCAP, PPT, anunci… (múltiples fitxers permesos)</p>
              <input ref={fileInputRef} type="file" accept=".pdf,application/pdf" multiple className="hidden" onChange={e=>{if(e.target.files?.length){setPlecFiles(Array.from(e.target.files));setPlecResults([]);setPlecRawText("");setPlecError("");setPlecStatus("");}}}/>
            </div>
            {plecFiles.length>0&&<div className="space-y-1.5">{plecFiles.map((f,i)=>(<div key={i} className="flex items-center gap-2 bg-slate-50 rounded-lg px-3 py-2 text-xs"><span className="text-red-500">📄</span><span className="flex-1 font-medium text-gray-700 truncate">{f.name}</span><span className="text-gray-400 shrink-0">{(f.size/1024/1024).toFixed(2)} MB</span><button onClick={e=>{e.stopPropagation();setPlecFiles(prev=>prev.filter((_,j)=>j!==i));}} className="text-gray-300 hover:text-red-400 ml-1">✕</button></div>))}</div>}
            {plecError&&<div className="bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-xs text-red-700 flex items-start gap-2"><span className="shrink-0 mt-0.5">❌</span><span>{plecError}</span></div>}
            {plecStatus&&!plecError&&<div className="bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 text-xs text-blue-700 flex items-center gap-2"><svg className="animate-spin h-3 w-3 shrink-0" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg><span>{plecStatus}</span></div>}
            <button onClick={analitzarPlec} disabled={plecLoading||plecFiles.length===0} className="w-full bg-blue-800 hover:bg-blue-900 disabled:bg-blue-300 text-white font-semibold py-3 rounded-xl text-sm flex items-center justify-center gap-2">
              {plecLoading?<><svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>{plecStatus||"Analitzant plec…"}</>:plecFiles.length===0?"Afegeix un PDF primer":"🔎 Analitzar plec amb IA"}
            </button>
          </div>
          {(plecResults.length>0||plecRawText)&&(<>
            <div className="bg-stone-50 rounded-xl border shadow-sm px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
              <div className="flex gap-2"><button onClick={()=>setPlecView("taula")} className={`text-xs font-semibold px-3 py-1.5 rounded-full border ${plecView==="taula"?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>📊 Taula resum</button><button onClick={()=>setPlecView("text")} className={`text-xs font-semibold px-3 py-1.5 rounded-full border ${plecView==="text"?"bg-blue-700 text-white border-blue-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>📝 Informe complet</button><button onClick={()=>setPlecView("legal")} className={`text-xs font-semibold px-3 py-1.5 rounded-full border ${plecView==="legal"?"bg-slate-700 text-white border-amber-700":"bg-stone-50 text-gray-600 border-gray-300"}`}>⚖️ Consultes legals</button></div>
              <div className="flex items-center gap-2">
                {plecSavedMsg&&<span className="text-xs font-semibold text-green-700 bg-green-50 px-3 py-1.5 rounded-full">{plecSavedMsg}</span>}
                {plecResults.length>0&&<button onClick={()=>guardarAlGestor(plecResults[0])} className="bg-purple-600 hover:bg-purple-700 text-white text-xs font-semibold px-3 py-2 rounded-lg">📁 Guardar al Gestor</button>}
                {plecResults.length>0&&<button onClick={exportarPlecExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-medium px-3 py-2 rounded-lg">📥 Excel</button>}
                {(plecResults.length>0||plecRawText)&&<button onClick={()=>exportarPlecPDF(plecResults,plecRawText,plecNom)} className="bg-red-600 hover:bg-red-700 text-white text-xs font-medium px-3 py-2 rounded-lg">📄 PDF</button>}
              </div>
            </div>
            {plecView==="taula"&&plecResults.map((r,idx)=>(<div key={idx} className="space-y-3">
              {r.lot&&<div className="bg-indigo-600 text-white rounded-xl px-4 py-2.5 font-bold text-sm flex items-center gap-2"><span className="bg-white/20 rounded-full px-2.5 py-0.5 text-xs">LOT {idx+1}</span>{r.lot}{r.objecte&&r.lot!==r.objecte&&<span className="font-normal opacity-80">— {r.objecte}</span>}</div>}
              <div className="bg-stone-50 rounded-xl border shadow-sm p-4"><h3 className="font-semibold text-gray-700 mb-3">📋 Dades bàsiques</h3>{r.termini_presentacio&&<div className="bg-red-50 border border-red-200 rounded-lg px-3 py-2 mb-3 flex items-center gap-2"><span className="text-red-600 text-lg">⏰</span><span className="text-sm font-bold text-red-800">Data presentació oferta: {r.termini_presentacio}</span></div>}<div className="grid grid-cols-2 gap-x-6 gap-y-2 text-xs">{[["Expedient",r.expedient],["Objecte",r.objecte],["Organisme",r.organisme],["CPV",r.cpv],["Import s/IVA",r.import_sense_iva?`${Number(r.import_sense_iva).toLocaleString("ca-ES")} €`:""],["Import c/IVA",r.import_amb_iva?`${Number(r.import_amb_iva).toLocaleString("ca-ES")} €`:""],["Valor estimat",r.valor_estimat?`${Number(r.valor_estimat).toLocaleString("ca-ES")} €`:""],["Termini execució",r.termini_execucio]].filter(([,v])=>v).map(([k,v])=>(<div key={k} className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">{k}:</span><span className="font-medium text-gray-800">{v}</span></div>))}</div>
                <div className="mt-3 flex items-center gap-2 text-xs"><span className="text-gray-400 shrink-0 w-36">🌐 Enllaç publicació:</span>{r.enllac_publicacio?<a href={r.enllac_publicacio} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline truncate">{r.enllac_publicacio}</a>:<span className="text-gray-300">No detectat al PDF</span>}<input className="flex-1 border rounded px-2 py-1 text-xs" placeholder="Enganxa l'URL de la publicació" defaultValue={r.enllac_publicacio||""} onBlur={e=>{const updated=[...plecResults];updated[idx]={...r,enllac_publicacio:e.target.value};setPlecResults(updated);}}/></div>
              </div>
              <div className="bg-stone-50 rounded-xl border shadow-sm p-4"><h3 className="font-semibold text-gray-700 mb-3">🏆 Classificació i solvència (LCSP)</h3>
                {r.exigeix_classificacio!==false&&r.classificacio_requerida?.length>0?(<>
                  <div className="bg-blue-50 border border-blue-200 rounded-lg px-3 py-2 mb-3 text-xs text-blue-800">
                    <strong>📜 Classificació obligatòria</strong> — Substitueix la solvència econòmica i tècnica (art. 77-85 LCSP)
                  </div>
                  <div className="mb-3"><p className="text-xs text-gray-500 mb-1.5">Classificació requerida:</p><div className="flex flex-wrap gap-1.5">{r.classificacio_requerida.map((c,i)=>{const codi=`${c.grup}${c.subgrup}`,catS=SERVIAL_CLASS[codi]??null,catR=parseCat(c.categoria),ok=catS!==null&&catS>=catR;return<span key={i} className={`text-xs font-bold px-2.5 py-1 rounded-full ${catS===null?"bg-red-100 text-red-700":ok?"bg-green-100 text-green-800":"bg-slate-100 text-amber-800"}`}>{codi} Cat.{c.categoria} {catS===null?"❌ No disposem":ok?`✅ Servial Cat.${catS}`:`⚠️ Servial Cat.${catS} (insuf.)`}</span>;})}</div></div>
                  {(r.solvencia_economica||r.solvencia_tecnica)&&<div className="bg-amber-50 border border-slate-200 rounded-lg px-3 py-2 mb-2 text-xs text-amber-800"><strong>⚠️ Requisits addicionals</strong> (al marge de la classificació):</div>}
                </>):(<>
                  <div className="bg-gray-50 border border-gray-200 rounded-lg px-3 py-2 mb-2 text-xs text-gray-700">
                    <strong>📋 Sense classificació obligatòria per llei</strong> (valor &lt; 500.000€) — S'acredita mitjançant solvència (arts. 87-88 LCSP)
                  </div>
                </>)}
                {/* Bloc de substitució PCAP - mostrar SEMPRE quan hi ha solvència o quan hi ha classificació exigida */}
                {(r.solvencia_economica||r.solvencia_tecnica||r.exigeix_classificacio)&&<>
                  {r.permet_substitucio_per_classificacio===true&&<div className="bg-green-50 border border-green-300 rounded-lg px-3 py-2 mb-2 text-xs text-green-800">
                    <strong>✅ El PCAP permet SUBSTITUIR la solvència per classificació</strong>
                    {r.text_substitucio_pcap&&<div className="mt-1 pt-1 border-t border-green-200 italic text-green-700">Text del PCAP: "{r.text_substitucio_pcap}"</div>}
                  </div>}
                  {r.permet_substitucio_per_classificacio===false&&<div className="bg-red-50 border border-red-300 rounded-lg px-3 py-2 mb-2 text-xs text-red-800">
                    <strong>❌ El PCAP NO permet la substitució de la solvència per classificació</strong> — La solvència és requisit obligatori i ha de ser acreditada encara que es tingui classificació.
                  </div>}
                  {(r.permet_substitucio_per_classificacio===null||r.permet_substitucio_per_classificacio===undefined)&&<div className="bg-amber-50 border border-amber-300 rounded-lg px-3 py-2 mb-2 text-xs text-amber-800">
                    <strong>⚠️ El PCAP no és clar sobre la substitució</strong> — Revisar manualment el document.
                  </div>}
                  {r.subcas_explicacio&&<div className="text-xs text-gray-600 italic mb-2 px-1 bg-gray-50 rounded p-2 border border-gray-200">📄 Segons el PCAP: {r.subcas_explicacio}</div>}
                </>}
                <div className="space-y-1 text-xs">{r.solvencia_economica&&<div className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">Solvència econòmica:</span><span className="font-medium text-gray-800">{r.solvencia_economica}</span></div>}{r.solvencia_tecnica&&<div className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">Solvència tècnica:</span><span className="font-medium text-gray-800">{r.solvencia_tecnica}</span></div>}</div>
              </div>
              {((r.criteris_automatics?.length>0)||(r.criteris_judici_valor?.length>0))&&(()=>{
                const ptsAuto=(r.criteris_automatics||[]).reduce((s,c)=>s+(c.punts||0),0);
                const ptsJV=(r.criteris_judici_valor||[]).reduce((s,c)=>s+(c.punts||0),0);
                const total=r.total_punts||ptsAuto+ptsJV||100;
                const pctAuto=Math.round(ptsAuto/total*100);
                const pctJV=Math.round(ptsJV/total*100);
                const BAND_COLORS=["bg-purple-100 text-purple-800 border-purple-300","bg-slate-100 text-amber-800 border-slate-300","bg-green-100 text-green-800 border-green-300","bg-blue-100 text-blue-800 border-blue-300"];
                return <div className="space-y-3">
                {/* SOBRE C - AUTOMÀTICS */}
                <div className="bg-stone-50 rounded-xl border shadow-sm p-4">
                  <div className="flex justify-between items-start mb-3">
                    <div><h3 className="font-semibold text-gray-700">Sobre C — Criteris automàtics</h3>
                    {r.requisits_memoria?.max_pagines&&<p className="text-xs text-gray-400 mt-0.5">Sense memòria · Avaluació automàtica</p>}</div>
                    <div className="text-right"><span className="text-2xl font-bold text-blue-700">{ptsAuto} pts</span><div className="text-xs text-gray-400">{pctAuto}% del total</div></div>
                  </div>
                  <div className="space-y-2">{(r.criteris_automatics||[]).map((c,i)=>(<div key={i} className="bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5">
                    <div className="flex justify-between items-center"><span className="text-sm font-semibold text-blue-800">{c.nom}</span><span className="font-bold text-blue-700 text-sm">{c.punts} pts</span></div>
                    {c.formula&&<p className="text-blue-600 mt-1 text-xs font-mono bg-blue-100/50 rounded px-2 py-1">{c.formula}</p>}
                  </div>))}</div>
                </div>

                {/* SOBRE B - JUDICI DE VALOR */}
                {(r.criteris_judici_valor||[]).length>0&&<div className="bg-stone-50 rounded-xl border shadow-sm p-4">
                  <div className="flex justify-between items-start mb-1">
                    <div><h3 className="font-semibold text-gray-700">Sobre B — Criteris subjectes a judici de valor</h3>
                    {r.requisits_memoria?.max_pagines&&<p className="text-xs text-gray-400 mt-0.5">Memòria d'organització i execució · màx. {r.requisits_memoria.max_pagines} pàgines {r.requisits_memoria.format||"DIN A4"}</p>}</div>
                    <div className="text-right"><span className="text-2xl font-bold text-slate-600">{ptsJV} pts</span><div className="text-xs text-gray-400">{pctJV}% del total</div></div>
                  </div>
                  {r.llindar_minim_tecnic?.te_llindar&&<div className="bg-red-50 border border-red-200 rounded-lg px-3 py-2 mb-3 text-xs text-red-800"><strong>⚠️ Llindar mínim tècnic:</strong> {r.llindar_minim_tecnic.puntuacio_minima} pts — {r.llindar_minim_tecnic.consequencia||"Per sota del llindar, l'oferta queda exclosa"}</div>}
                  <div className="space-y-3 mt-3">{(r.criteris_judici_valor||[]).map((c,i)=>(<div key={i} className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="flex justify-between items-center px-4 py-3 bg-gradient-to-r from-gray-50 to-white border-b border-gray-100">
                      <div className="flex items-center gap-2"><span className="bg-gray-700 text-white text-xs font-bold w-6 h-6 rounded-full flex items-center justify-center">{i+1}</span><span className="font-semibold text-gray-800">{c.nom}</span></div>
                      <span className="font-bold text-slate-600 text-lg">{c.punts} pts</span>
                    </div>
                    <div className="px-4 py-3 space-y-2.5">
                      {c.subcriteris?.length>0&&(()=>{
                        const renderSub=(items,depth=0)=>items.map((sc,j)=>{
                          if(typeof sc!=="object")return<li key={`${depth}-${j}`} className="text-xs text-gray-600 flex gap-2" style={{marginLeft:`${depth*16}px`}}><span className="text-gray-400">•</span><span>{sc}</span></li>;
                          const codi=sc.codi||"";const nom=sc.nom||sc.name||"";const punts=sc.punts!=null?sc.punts:null;
                          return<Fragment key={`${depth}-${j}`}>
                            <li className="text-xs flex gap-2 py-0.5" style={{marginLeft:`${depth*16}px`}}>
                              <span className={depth===0?"text-indigo-600 font-bold":depth===1?"text-blue-600 font-semibold":"text-gray-500"}>{codi?codi+"."+" ":"• "}</span>
                              <span className={depth===0?"font-semibold text-gray-800":"text-gray-600"}>{nom}</span>
                              {punts!=null&&<span className={`ml-auto shrink-0 ${depth===0?"font-bold text-amber-700":"text-amber-600"}`}>({punts} pts)</span>}
                            </li>
                            {sc.fills?.length>0&&renderSub(sc.fills,depth+1)}
                          </Fragment>;
                        });
                        return<ul className="space-y-0.5 bg-gray-50 rounded p-2 border border-gray-100">{renderSub(c.subcriteris,0)}</ul>;
                      })()}
                      {c.que_valoren&&<div className="text-xs text-gray-500 italic bg-gray-50 rounded px-2 py-1.5">{c.que_valoren}</div>}
                      {c.bandes_puntuacio?.length>0&&<div className="flex flex-wrap gap-1.5 mt-1">{c.bandes_puntuacio.map((b,k)=>(<span key={k} className={`text-xs font-medium px-2.5 py-1 rounded-full border ${BAND_COLORS[k%BAND_COLORS.length]}`}>{b.rang} pts: {b.descripcio}</span>))}</div>}
                      {c.inconsistencia&&<div className="bg-yellow-50 border border-yellow-300 rounded-lg px-3 py-2 text-xs text-yellow-800 mt-1"><strong>⚠️ Inconsistència detectada al PCAP:</strong> {c.inconsistencia}</div>}
                    </div>
                  </div>))}</div>
                </div>}

                {/* REQUISITS MEMÒRIA */}
                {r.requisits_memoria&&(r.requisits_memoria.max_pagines||r.requisits_memoria.font||r.requisits_memoria.interlineat)&&<div className="bg-stone-50 rounded-xl border shadow-sm p-4">
                  <h3 className="font-semibold text-gray-700 mb-2">📝 Requisits formals de la memòria</h3>
                  <div className="grid grid-cols-2 gap-2 text-xs">{[["Màx. pàgines",r.requisits_memoria.max_pagines],["Format",r.requisits_memoria.format],["Font/mida",r.requisits_memoria.font],["Interlineat",r.requisits_memoria.interlineat],["Annexos",r.requisits_memoria.annexos],["Penalització excedir",r.requisits_memoria.penalitzacio_excedir]].filter(([,v])=>v).map(([k,v])=>(<div key={k} className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">{k}:</span><span className="font-medium text-gray-800">{v}</span></div>))}</div>
                </div>}

                {/* TEMERITAT I EXCLUSIÓ */}
                {(r.temeritat?.formula||r.criteris_exclusio?.filter(e=>e).length>0)&&<div className="bg-red-50 rounded-xl border border-red-200 shadow-sm p-4">
                  <h3 className="font-semibold text-red-800 mb-2">🚫 Criteris d'exclusió i temeritat</h3>
                  {r.temeritat?.formula&&<div className="text-xs space-y-1 mb-2"><div className="flex gap-2"><span className="text-red-400 shrink-0 w-36">Fórmula temeritat:</span><span className="font-medium text-red-800">{r.temeritat.formula}</span></div>{r.temeritat.percentatge&&<div className="flex gap-2"><span className="text-red-400 shrink-0 w-36">% baixa temerària:</span><span className="font-medium text-red-800">{r.temeritat.percentatge}</span></div>}{r.temeritat.procediment&&<div className="flex gap-2"><span className="text-red-400 shrink-0 w-36">Procediment:</span><span className="font-medium text-red-800">{r.temeritat.procediment}</span></div>}{r.temeritat.consequencia&&<div className="flex gap-2"><span className="text-red-400 shrink-0 w-36">Conseqüència:</span><span className="font-medium text-red-800">{r.temeritat.consequencia}</span></div>}</div>}
                  {r.criteris_exclusio?.filter(e=>e).length>0&&<div className="mt-2"><p className="text-xs font-semibold text-red-700 mb-1">Causes d'exclusió:</p><ul className="space-y-0.5">{r.criteris_exclusio.filter(e=>e).map((e,i)=>(<li key={i} className="text-xs text-red-700 flex gap-2"><span>❌</span><span>{e}</span></li>))}</ul></div>}
                </div>}

                {/* RESUM ESTRATÈGIC */}
                <div className="bg-gradient-to-r from-gray-700 to-gray-800 rounded-xl p-4 text-white">
                  <h3 className="font-semibold mb-2">📊 Resum estratègic</h3>
                  <div className="flex gap-4 mb-2">
                    <div className="flex-1 bg-white/10 rounded-lg p-2.5 text-center"><div className="text-xl font-bold">{ptsAuto}</div><div className="text-xs opacity-70">Automàtics ({pctAuto}%)</div></div>
                    <div className="flex-1 bg-white/10 rounded-lg p-2.5 text-center"><div className="text-xl font-bold">{ptsJV}</div><div className="text-xs opacity-70">Judici valor ({pctJV}%)</div></div>
                    <div className="flex-1 bg-white/10 rounded-lg p-2.5 text-center"><div className="text-xl font-bold">{total}</div><div className="text-xs opacity-70">Total punts</div></div>
                  </div>
                  <div className={`text-xs font-semibold px-3 py-1.5 rounded-full inline-block ${pctAuto>=60?"bg-blue-400/30":"bg-amber-400/30"}`}>{pctAuto>=60?"💰 PREU DOMINANT — Prioritzar oferta econòmica agressiva":"📝 QUALITAT DOMINANT — Prioritzar memòria tècnica detallada"}</div>
                </div>
              </div>;})()}
              <div className="bg-stone-50 rounded-xl border shadow-sm p-4"><h3 className="font-semibold text-gray-700 mb-3">🔒 Garanties i condicions</h3><div className="space-y-1 text-xs">{r.garantia_definitiva&&<div className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">Garantia definitiva:</span><span className="font-medium text-gray-800">{r.garantia_definitiva}</span></div>}{r.visita_obra&&<div className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">Visita d'obra:</span><span className="font-medium text-gray-800">{r.visita_obra}</span></div>}{r.condicions_especials&&<div className="flex gap-2"><span className="text-gray-400 shrink-0 w-36">Condicions especials:</span><span className="font-medium text-gray-800">{r.condicions_especials}</span></div>}</div></div>
              {r.diagnosic_servial&&<div className="bg-amber-50 border border-slate-200 rounded-xl p-4"><h3 className="font-semibold text-amber-800 mb-2">⚠️ Diagnòstic per a Servial</h3><p className="text-xs text-slate-800 leading-relaxed whitespace-pre-wrap">{r.diagnosic_servial}</p></div>}
            </div>))}
            {plecView==="text"&&plecRawText&&<div className="bg-stone-50 rounded-xl border shadow-sm p-5"><div className="text-gray-700 text-xs leading-relaxed whitespace-pre-wrap">{plecRawText.replace(/--JSON_INICI--[\s\S]*?--JSON_FI--/,"").trim()}</div></div>}
            {plecView==="legal"&&<div className="space-y-3">
              <div className="bg-slate-50 rounded-xl border border-slate-300 shadow-sm p-4">
                <div className="flex items-center gap-2 mb-3"><span className="text-xl">⚖️</span><h3 className="font-semibold text-slate-800">Consultes Legals</h3>{plecRawText&&<span className="text-xs text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">Plec carregat com a context</span>}</div>
                <p className="text-xs text-slate-600 mb-3">Fes preguntes sobre contractació pública, classificació, solvència, criteris d'adjudicació, garanties, subcontractació, etc. {plecRawText?"La resposta es contextualitzarà amb el plec analitzat.":"Analitza un plec primer per obtenir respostes contextualitzades."}</p>
                <div className="flex gap-2">
                  <textarea rows={3} className="flex-1 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-400 resize-y" placeholder="Escriu la teva consulta legal... Ex: Quina classificació necessito per presentar-me? Puc subcontractar? Quina garantia definitiva cal?" value={legalQuery} onChange={e=>setLegalQuery(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey&&!legalLoading){e.preventDefault();consultarLegal();}}}/>
                  <button onClick={consultarLegal} disabled={legalLoading||!legalQuery.trim()} className="bg-slate-700 hover:bg-slate-800 disabled:bg-slate-300 text-white font-semibold px-4 py-2 rounded-lg text-sm shrink-0 self-end">{legalLoading?"Consultant...":"⚖️ Consultar"}</button>
                </div>
              </div>
              {legalResponses.map((lr,i)=>(<div key={i} className="bg-stone-50 rounded-xl border shadow-sm overflow-hidden">
                <div className="bg-slate-100 border-b border-slate-200 px-4 py-2.5 flex items-center justify-between">
                  <div className="flex items-center gap-2"><span className="text-slate-600 font-bold text-xs">CONSULTA</span><span className="text-xs text-slate-500">{lr.date}</span></div>
                </div>
                <div className="px-4 py-3 bg-amber-50 border-b"><p className="text-sm font-medium text-slate-800">{lr.pregunta}</p></div>
                <div className="px-4 py-3"><div className="text-xs text-gray-700 leading-relaxed whitespace-pre-wrap">{lr.resposta}</div></div>
              </div>))}
              {legalResponses.length===0&&<div className="text-center py-8 text-gray-400"><div className="text-3xl mb-2">⚖️</div><p className="text-sm">Encara no has fet cap consulta legal</p><p className="text-xs mt-1">Escriu una pregunta sobre el plec o sobre contractació pública en general</p></div>}
            </div>}
          </>)}
        </>)}
      </div>
      )}
    </div>
  );
}
