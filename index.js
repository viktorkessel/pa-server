const express = require("express");
const pptxgen = require("pptxgenjs");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const os = require("os");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ─── CONFIG ────────────────────────────────────────────────────────────────

const SERVICE_ACCOUNT = {
  type: "service_account",
  project_id: "weighty-bounty-500220-k2",
  private_key_id: "f2bce7cf1113be36ed71da04832816e68cd0fe6b",
  private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDmVtkQkN8Mb+ar\nWuB0M6efeVfvFmxlQOsQTZ/htOTtp1liWPBme8vSWAIDys4OD/R2UXuj6ZSjp1KD\nmponFVItQ7pupxIAIjE9nmrg88WfBLAqzPOcNlJi99GBuOLV03O3FLIdHcOQS7aP\n1fcK7YyOQAm9a05aNcoeuXp1WSda9UfT1+VeQMzFpWYneDmk6Y+PHgb7UXNU3cF6\nHqbsReiDcjSGO6Qv742zo2G2Gqvjiqcb/byrDNd+iOdqqbz5prYLgaPlafakLv8C\nILR4UbDyJIHsdHiF7KlVGMkDm3qqNNDExdjQApQ9IItmDPLYbYUUHe13zm3qD6JU\nU5HRJVezAgMBAAECggEAJ3JVHyyoPjGFv+/Cw7O676kwwX3GF3sMJpEcoeNAV2DT\nAnA3z11BFSdERRD8w9TAqWif83zDm1lC3M3ju3C/z4erTFrpTsOMSkCwjhDBVc/Q\nOaw6xhRZfKBzLTOUCUl6pi5xqGHaBYKp460hbyoqaYwW/uz7EKDomeVqzuDzR/Sp\nXCeebV51kTB294uWgSYPMLv0GZMr7nqrM0y/KJk6MIEp0MdxZGHz06yNNcZEUDOp\n3HmJF8Ac+WTRZR4MRsHcI9uPQbbwUKS8aqjIl6KP6rfe406RGdk3/N0RXuBlEYpY\nXJNpuUJEIX8wrucl8jem2topIhCEaz4jp7DlCNud1QKBgQD7nXGPZeHjBHbm651y\njU87fdIfZqO/G7gW5eNXqMvXmZAYWr3y3dHASvMaDhQLXU9poFi9gk/fA7S0fxU5\nST0lPqXZuY9EHD4QEURSw9GODLnkFx7tUk4JZJZoXxxRoAo7IvpIFgbh2+2SNznE\nNfWLc/aPjyCWLAdJv8AKxCuKFwKBgQDqWnwJd9fQrw8VxlzboGZrCo2WdXsLX8LO\nieCh3AoQ38VLa66L1nKxmf5ewTU2h8KZtYxZiM5hYSlrUVJ+z/dFSP5X2zgRByuQ\nMYgywQfcnUaKjjFqe9CsDKHB+4IR2wWCn0192Q+idtZ6RRofWwFDNVCSd78zc9gV\nerWcYg0MxQKBgGyTLAWs+SR1MLnEwlREFgSd7qYOLGqzE2UOHmQdOpj2cu+PeJfs\nkC5iKTT9Djk5MqNpxeFA8iZR+xBIJmhp6jDAL+Eb6W6CoWVdjhYI2cxqQWzjbT0L\nXxE+jojLQfUXElMibM3ec2nS9Vk663p3rlVW3PZl1LH4ncx9Kel7bS5RAoGBAIqv\nF14HBLkjD5tEMulpp5JKwXU7eqBFgCP3pe5DJ9XKhpfWVom8U01Uoxpx3cqyVm/d\n5vuJIRCxjfBLRLV03pCq3JEjUq+I/besCDx4zO06SWdvNEEemmNmakdATj3Azib5\nnMMPs9fh2ZSr1gSeN0YNg+vZX1XUhRhY8KZlw/E9AoGBAO2uyvbLRdYjaYsdIyod\nE9hQ22xUCTU05aaSEzgzlosugh9loTQilnHS/RNLz79K4lEBEwNFu5kK+sdnpuJq\nxIHbjBjvYOk4wMSCHhEUHlDNBsJcfyCJ+/fiC/ijK5ktl2RawGX6KN7frevcH6cX\nfls2ZHgXi+ouCEW2EVpasrlP\n-----END PRIVATE KEY-----\n",
  client_email: "sb-conseil-pa@weighty-bounty-500220-k2.iam.gserviceaccount.com",
  client_id: "108283571791048614042",
  auth_uri: "https://accounts.google.com/o/oauth2/auth",
  token_uri: "https://oauth2.googleapis.com/token",
  auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  client_x509_cert_url: "https://www.googleapis.com/robot/v1/metadata/x509/sb-conseil-pa%40weighty-bounty-500220-k2.iam.gserviceaccount.com",
  universe_domain: "googleapis.com"
};

const DRIVE_FOLDERS = {
  "qynapse":     "1Bc2e_deAI3LMAeTBjVyVovb2Q55SPP8s",
  "biome":       "1eSMYXNilgzV1VModMQF9vMtMtZ8uhfZ_",
  "genepredict": "1fytSfQgiAKanvUJfCSEeAqzvrPa8vauW",
  "amorphous":   "1-73BW8nmEg3wUJdszjPA_HYDLLAlxklw",
  "boutin":      "18YL5XSKSo-5V5gefrPmAFz74F98He6-s",
  "medivie":     "1bsjICqrEK92sMSDE5O4rzSJgYWm1fpFA",
  "transversal": "1M-cPRxraEXZE3siLGOY7BAGws8AQMUmS"
};

// ─── COULEURS & CONSTANTES ─────────────────────────────────────────────────

const W = 8.27, H = 11.69, MARGIN = 0.3;
const COL_PRIO = 0.82, COL_ACTION = 4.7, COL_RESP = 1.35, COL_DATE = 0.68;
const TABLE_W = COL_PRIO + COL_ACTION + COL_RESP + COL_DATE;

const C = {
  teal:"028090", tealDark:"016070", tealLight:"E8F4F5", tealMid:"B8DDE0",
  orange:"E8A838", white:"FFFFFF", light:"F4F9FA", mid:"DDE8EA",
  text:"1A2E30", muted:"5A7A7D",
  urgent:"C0392B", prio:"C07000", normal:"1A6FA8", fait:"1E7E34", veille:"777777",
  iso:"4A5568", isoLight:"EEF0F4", hs:"5A6A7A", hsLight:"EFF1F4"
};

function badgeColor(p) {
  return {URGENT:C.urgent,PRIO:C.prio,NORMAL:C.normal,FAIT:C.fait,VEILLE:C.veille}[p]||C.muted;
}

// ─── GÉNÉRATION PPTX ───────────────────────────────────────────────────────

function header(slide, meta) {
  slide.addShape("rect",{x:0,y:0,w:W,h:0.48,fill:{color:C.teal},line:{color:C.teal}});
  slide.addText(meta.title,{x:MARGIN,y:0.04,w:5.5,h:0.24,fontSize:11,bold:true,color:C.white,fontFace:"Calibri",margin:0});
  slide.addText(meta.date,{x:5.5,y:0.04,w:W-5.5-MARGIN,h:0.24,fontSize:8,color:C.orange,fontFace:"Calibri",align:"right",margin:0});
  slide.addText(meta.sub,{x:MARGIN,y:0.29,w:W-MARGIN*2,h:0.16,fontSize:6.5,color:"C8E8EA",fontFace:"Calibri",margin:0});
}

function recentsBlock(slide, items, y) {
  const ROW_H = 0.135;
  slide.addShape("rect",{x:MARGIN,y,w:TABLE_W,h:0.17,fill:{color:C.teal},line:{color:C.teal}});
  slide.addText("ÉVOLUTIONS RÉCENTES",{x:MARGIN+0.05,y:y+0.02,w:4,h:0.13,fontSize:6.5,bold:true,color:C.white,fontFace:"Calibri",margin:0});
  y+=0.17;
  items.forEach((item,i)=>{
    slide.addShape("rect",{x:MARGIN,y,w:TABLE_W,h:ROW_H,fill:{color:i%2===0?C.white:C.tealLight},line:{color:C.mid,pt:0.4}});
    slide.addText("→  "+item,{x:MARGIN+0.06,y:y+0.01,w:TABLE_W-0.12,h:ROW_H-0.02,fontSize:6.5,color:C.text,fontFace:"Calibri",margin:0});
    y+=ROW_H;
  });
  return y+0.05;
}

function block(slide, bandColor, titleLeft, titleRight, rows, y, altLight) {
  const TITLE_H=0.19, HDR_H=0.14, ROW_H=0.165;
  const hWidths=[COL_PRIO,COL_ACTION,COL_RESP,COL_DATE];
  slide.addShape("rect",{x:MARGIN,y,w:TABLE_W,h:TITLE_H,fill:{color:bandColor},line:{color:bandColor}});
  slide.addText(titleLeft,{x:MARGIN+0.05,y:y+0.025,w:3.5,h:TITLE_H-0.05,fontSize:7.5,bold:true,color:C.white,fontFace:"Calibri",margin:0});
  if(titleRight) slide.addText(titleRight,{x:MARGIN+3.6,y:y+0.025,w:TABLE_W-3.65,h:TITLE_H-0.05,fontSize:6.5,color:C.orange,italic:true,fontFace:"Calibri",align:"right",margin:0});
  y+=TITLE_H;
  const hLabels=["PRIORITÉ","ACTION","RESPONSABLE","ÉCHÉANCE"];
  let x=MARGIN;
  hLabels.forEach((lbl,i)=>{
    slide.addShape("rect",{x,y,w:hWidths[i],h:HDR_H,fill:{color:altLight},line:{color:C.tealMid,pt:0.5}});
    slide.addText(lbl,{x:x+0.03,y:y+0.01,w:hWidths[i]-0.06,h:HDR_H-0.02,fontSize:6,bold:true,color:bandColor,fontFace:"Calibri",margin:0});
    x+=hWidths[i];
  });
  y+=HDR_H;
  rows.forEach((row,idx)=>{
    const [prio,action,resp,date]=row;
    const bg=idx%2===0?C.white:altLight;
    x=MARGIN;
    [prio,action,resp,date].forEach((val,i)=>{
      slide.addShape("rect",{x,y,w:hWidths[i],h:ROW_H,fill:{color:i===0?badgeColor(prio):bg},line:{color:C.mid,pt:0.4}});
      slide.addText(val,{x:x+0.03,y:y+0.015,w:hWidths[i]-0.06,h:ROW_H-0.03,fontSize:i===0?6.5:7,bold:i===0,color:i===0?C.white:C.text,fontFace:"Calibri",margin:0,wrap:true});
      x+=hWidths[i];
    });
    y+=ROW_H;
  });
  return y+0.05;
}

function footer(slide) {
  slide.addText("Document confidentiel — SB Conseil — stephaneboyer2302@gmail.com",{x:MARGIN,y:H-0.22,w:W-MARGIN*2,h:0.16,fontSize:6,color:C.muted,fontFace:"Calibri",align:"center",margin:0});
}

async function generatePPTX(data) {
  const pres = new pptxgen();
  pres.defineLayout({name:"A4P",width:W,height:H});
  pres.layout="A4P";
  const slide = pres.addSlide();
  slide.background={color:C.white};

  const meta = {
    title: `PLAN D'ACTIONS — ${(data.client||"CLIENT").toUpperCase()}`,
    sub: `SB Conseil  |  ${data.mission||"Mission"}  |  Depuis le ${data.depuis||"01/06/2026"}`,
    date: data.date || new Date().toLocaleDateString("fr-FR")
  };

  header(slide, meta);
  let y = 0.54;

  if (data.recents && data.recents.length > 0) {
    y = recentsBlock(slide, data.recents, y);
  }

  if (data.initiatives) {
    for (const init of data.initiatives) {
      y = block(slide, C.tealDark, `${init.num}. ${init.title}`, init.goal ? "Objectif : "+init.goal : null, init.rows, y, C.tealLight);
    }
  }

  if (data.actions_isolees && data.actions_isolees.length > 0) {
    y = block(slide, C.iso, "ACTIONS ISOLÉES", null, data.actions_isolees, y, C.isoLight);
  }

  if (data.hors_scope && data.hors_scope.length > 0) {
    y = block(slide, C.hs, "HORS SCOPE / IDÉES À DISCUTER", null, data.hors_scope, y, C.hsLight);
  }

  footer(slide);

  const tmpFile = path.join(os.tmpdir(), `PA_${Date.now()}.pptx`);
  await pres.writeFile({fileName: tmpFile});
  return tmpFile;
}

// ─── UPLOAD DRIVE ──────────────────────────────────────────────────────────

async function uploadToDrive(filePath, fileName, folderId) {
  const auth = new google.auth.GoogleAuth({
    credentials: SERVICE_ACCOUNT,
    scopes: ["https://www.googleapis.com/auth/drive"]
  });
  const drive = google.drive({version:"v3", auth});

  // Chercher si le fichier existe déjà
  const existing = await drive.files.list({
    q: `name='${fileName}' and '${folderId}' in parents and trashed=false`,
    fields: "files(id,name)"
  });

  const fileStream = fs.createReadStream(filePath);
  const mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  if (existing.data.files.length > 0) {
    // Mettre à jour le fichier existant
    const fileId = existing.data.files[0].id;
    await drive.files.update({
      fileId,
      media: {mimeType, body: fileStream}
    });
    return {action:"updated", fileId};
  } else {
    // Créer un nouveau fichier
    const res = await drive.files.create({
      requestBody: {name: fileName, parents: [folderId]},
      media: {mimeType, body: fileStream},
      fields: "id"
    });
    return {action:"created", fileId: res.data.id};
  }
}

// ─── ROUTES ────────────────────────────────────────────────────────────────

app.get("/health", (req, res) => res.json({status:"ok"}));

app.post("/generate-pa", async (req, res) => {
  try {
    const { client, data } = req.body;

    if (!client) return res.status(400).json({error:"client requis"});

    const clientKey = client.toLowerCase().replace(/[^a-z]/g,"");
    const folderId = DRIVE_FOLDERS[clientKey];
    if (!folderId) return res.status(400).json({error:`Client inconnu: ${client}. Clients disponibles: ${Object.keys(DRIVE_FOLDERS).join(", ")}`});

    // Générer le PPTX
    const tmpFile = await generatePPTX({client, ...data});

    // Nom du fichier
    const today = new Date().toLocaleDateString("fr-FR").replace(/\//g,"-");
    const fileName = `PA_${client}_${today}.pptx`;

    // Lire le fichier en base64
    const fileBuffer = fs.readFileSync(tmpFile);
    const base64Content = fileBuffer.toString("base64");

    // Nettoyage
    fs.unlinkSync(tmpFile);

    res.json({
      success: true,
      client,
      fileName,
      folderId,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      fileContent: base64Content
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({error: err.message});
  }
});

// ─── DONNÉES PA PAR CLIENT ─────────────────────────────────────────────────

const PA_DATA = {
  qynapse: {
    mission: "Partenariats industriels Europe & Série A",
    depuis: "01/06/2026",
    recents: [
      "17/06 — Call Philips (Nader, Charles) — intérêt partenariat confirmé",
      "18/06 — C/R interne validé — C/R externe envoyé à Nader et Charles",
      "18/06 — Call Brainomix (Jérôme Galbrun) — CEO veut rencontrer Qynapse",
      "22/06 — Call suivi Philips 8/07 à 9h confirmé",
      "22/06 — Call Olivier : Monaco acté, piste Gopi 2/HCL, bridge financement prioritaire"
    ],
    initiatives: [
      {
        num:"1", title:"PHILIPS", goal:"LOI/MOU avant JFR (oct. 2026)",
        rows:[
          ["URGENT","Préparer call 8/07 : Monaco pilote, périmètre JFR, piste Gopi 2/HCL Coton","Stéphane + Olivier","08/07"],
          ["URGENT","Confirmer Monaco — appel direct Charles","Stéphane + Charles","Avant 08/07"],
          ["PRIO","Activer piste Gopi 2 : intégrer Qynapse dans offre Philips-HCL Lyon","Stéphane + Olivier","08/07"],
          ["PRIO","LOI / MOU signé par Nader — levier Série A","Stéphane","30/09"],
          ["NORMAL","2-3 features exclusives Philips + biz model anti-margin squeeze","Stéphane + Olivier","30/09"]
        ]
      },
      {
        num:"2", title:"BRAINOMIX", goal:"Visio intro sem. 29/06 — term sheet H1 2027",
        rows:[
          ["URGENT","Créneaux visio à Jérôme (Olivier : lun, mar AM, jeu >15h)","Stéphane → Jérôme","Cette sem."],
          ["URGENT","Envoyer deck Qynapse EN à Jérôme","Olivier","Cette sem."],
          ["PRIO","Brief Olivier : complémentarité stroke/neuro-dégén., vigilance rachat","Stéphane","Avant visio"],
          ["NORMAL","Structure cible : co-distrib. 600 hôpitaux, rev-share, option intégration","Stéphane + Olivier","Avant visio"]
        ]
      },
      {
        num:"3", title:"CONSTRUCTEURS IRM", goal:"1er accord Europe H1 2027",
        rows:[
          ["PRIO","Canon Medical — contacter Franck Girard","Stéphane","Juillet"],
          ["NORMAL","United Imaging — contacter Alex Ripert","Stéphane + Olivier","Septembre"],
          ["VEILLE","Siemens (Pasquier, Lavirotte) — contact maintenu, pas de push","Stéphane","Q4 2026"]
        ]
      },
      {
        num:"4", title:"FINANCEMENT", goal:"Bridge 500-700K€ juillet — Série A Q4 2026",
        rows:[
          ["URGENT","Bridge 500-700K€ — actionnaires historiques, val. 5-10M€ pre-money","Olivier + Lionel","Juillet"],
          ["PRIO","Annexe A investisseurs (liste SB Conseil) → Olivier","Stéphane → Olivier","03/07"],
          ["PRIO","Session Lionel : narratif, deck, BA / family office / Capital Cell","Stéphane + Lionel","Juillet"]
        ]
      },
      {
        num:"5", title:"VOICE OF CUSTOMER", goal:"2-3 interviews avant fin septembre",
        rows:[
          ["PRIO","2-3 entretiens cliniciens sur QyScore v1 + attentes v2.0","Stéphane","30/09"],
          ["NORMAL","3 visites terrain avec Aline — synthèse VOC → R&D","Stéphane + Aline","30/09 / Oct."]
        ]
      },
      {
        num:"6", title:"GOUVERNANCE", goal:"Rythme installé avant fin juillet",
        rows:[
          ["FAIT","Visio Olivier / Stéphane 22/06 18h — PA validé","Stéphane + Olivier","22/06 ✓"],
          ["PRIO","Mail Qynapse + accès Slack — ce PA = document de suivi mission","Olivier","Juillet"],
          ["NORMAL","Rencontre physique équipe Qynapse (Aline)","Olivier","Q3 2026"]
        ]
      }
    ],
    actions_isolees: [
      ["PRIO","Préparer deck Qynapse EN pour Jérôme Galbrun (version investor-ready)","Stéphane + Olivier","Cette sem."],
      ["PRIO","Qualifier plateforme crowdfunding Capital Cell — présenter à Lionel","Stéphane → Lionel","Juillet"]
    ],
    hors_scope: [
      ["VEILLE","Marketplaces IA (Incepto, Aidoc) — à qualifier Q3 2026","Stéphane","Q3 2026"],
      ["VEILLE","Licensing pharma / CRO — process Olivier en cours, hors scope SB Conseil","Olivier","Oct. 2026"],
      ["VEILLE","GE Healthcare — déprioritisé post-acquisition Icometrix","—","—"]
    ]
  }
};

// Route simplifiée : juste {"client":"qynapse"} suffit
app.post("/generate-pa/simple", async (req, res) => {
  try {
    const { client } = req.body;
    if (!client) return res.status(400).json({error:"client requis"});

    const clientKey = client.toLowerCase().replace(/[^a-z]/g,"");
    const folderId = DRIVE_FOLDERS[clientKey];
    if (!folderId) return res.status(400).json({error:`Client inconnu: ${client}`});

    const paData = PA_DATA[clientKey];
    if (!paData) return res.status(400).json({error:`Pas de données PA pour: ${client}`});

    const today = new Date().toLocaleDateString("fr-FR").replace(/\//g,"-");
    const tmpFile = await generatePPTX({client, date: today, ...paData});
    const fileName = `PA_${client}_${today}.pptx`;

    const fileBuffer = fs.readFileSync(tmpFile);
    const base64Content = fileBuffer.toString("base64");
    fs.unlinkSync(tmpFile);

    res.json({
      success: true,
      client,
      fileName,
      folderId,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      fileContent: base64Content
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({error: err.message});
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PA Server running on port ${PORT}`));
