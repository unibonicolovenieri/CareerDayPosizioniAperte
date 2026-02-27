const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, ShadingType, Table, TableRow, TableCell, WidthType,
  LevelFormat
} = require('docx');
const fs = require('fs');

const aziende = [
  { n: "001", nome: "Accenture", id: "504292" },
  { n: "002", nome: "ADHR GROUP - Agenzia per il Lavoro S.p.a.", id: "804" },
  { n: "003", nome: "Aetna Group S.p.A.", id: "28" },
  { n: "004", nome: "Agrintesa Soc. Coop. Agricola", id: "921" },
  { n: "005", nome: "Akkodis", id: "744" },
  { n: "006", nome: "Akronos Technologies", id: "2621" },
  { n: "007", nome: "Alfasigma", id: "1722" },
  { n: "008", nome: "Allianz SpA", id: "745" },
  { n: "009", nome: "ALLSIDES", id: "2642" },
  { n: "010", nome: "Alpitronic", id: "1963" },
  { n: "011", nome: "ALSTOM", id: "503786" },
  { n: "012", nome: "Altea Federation", id: "383" },
  { n: "013", nome: "ALTEN ITALIA SPA", id: "503942" },
  { n: "014", nome: "Areajob Spa", id: "2601" },
  { n: "015", nome: "auxiell group", id: "1184" },
  { n: "016", nome: "AYES", id: "2141" },
  { n: "017", nome: "BCUBE", id: "101" },
  { n: "018", nome: "BDO Italia S.p.A.", id: "502788" },
  { n: "019", nome: "BECKHOFF AUTOMATION", id: "499917" },
  { n: "020", nome: "Benu Farmacia", id: "2721" },
  { n: "021", nome: "BERCO", id: "362" },
  { n: "022", nome: "BERLUTI", id: "406" },
  { n: "023", nome: "Bios Management Italia Srl", id: "2101" },
  { n: "024", nome: "BolognaFiere Group", id: "941" },
  { n: "025", nome: "Bonfiglioli S.p.A.", id: "501097" },
  { n: "026", nome: "BPER BANCA", id: "122" },
  { n: "027", nome: "Bureau Veritas Italia", id: "1781" },
  { n: "028", nome: "BV TECH", id: "2722" },
  { n: "029", nome: "CAMPLUS", id: "2245" },
  { n: "030", nome: "Cantiere del Pardo SPA", id: "1185" },
  { n: "031", nome: "Carraro", id: "2182" },
  { n: "032", nome: "Caterpillar", id: "2701" },
  { n: "033", nome: "Cattaneo Zanetto Pomposo & Co.", id: "1186" },
  { n: "034", nome: "Cefla", id: "24" },
  { n: "035", nome: "CEL Components", id: "2742" },
  { n: "036", nome: "Centro di Sperimentazione Laimburg", id: "2643" },
  { n: "037", nome: "CENTRO SOFTWARE SPA", id: "1187" },
  { n: "038", nome: "CFT", id: "1762" },
  { n: "039", nome: "CHIESI FARMACEUTICI", id: "498348" },
  { n: "040", nome: "CHIRON ENERGY", id: "2202" },
  { n: "041", nome: "CIMBRIA SRL", id: "1224" },
  { n: "042", nome: "CINECA", id: "503909" },
  { n: "043", nome: "Clauger-Technofrigo S.r.l", id: "1242" },
  { n: "044", nome: "CLEMENTONI", id: "1349" },
  { n: "045", nome: "COCCHI TECHNOLOGIES", id: "500320" },
  { n: "046", nome: "Coesia Spa", id: "503836" },
  { n: "047", nome: "COMECER SpA", id: "501174" },
  { n: "048", nome: "Conserve Italia", id: "1201" },
  { n: "049", nome: "Coop Alleanza 3.0", id: "281" },
  { n: "050", nome: "Coopservice S.Coop.p.A.", id: "499140" },
  { n: "051", nome: "COSWELL SPA", id: "499042" },
  { n: "052", nome: "Credem", id: "403" },
  { n: "053", nome: "Credit Agricole Italia S.p.A.", id: "2021" },
  { n: "054", nome: "Crif S.p.A", id: "503831" },
  { n: "055", nome: "DATALOGIC", id: "503083" },
  { n: "056", nome: "Davines Group", id: "2102" },
  { n: "057", nome: "DECATHLON", id: "501194" },
  { n: "058", nome: "DELOITTE", id: "501263" },
  { n: "059", nome: "d-fine s.r.l.", id: "1724" },
  { n: "060", nome: "Dometic", id: "2644" },
  { n: "061", nome: "Ducati", id: "498881" },
  { n: "062", nome: "ECO CERTIFICAZIONI SPA", id: "1189" },
  { n: "063", nome: "Edison", id: "1862" },
  { n: "064", nome: "Eli Lilly", id: "421" },
  { n: "065", nome: "ENGIE", id: "2646" },
  { n: "066", nome: "EPSOL SRL", id: "2648" },
  { n: "067", nome: "EssilorLuxottica", id: "502147" },
  { n: "068", nome: "Essity Italy", id: "2649" },
  { n: "069", nome: "Eurotec Srl", id: "2650" },
  { n: "070", nome: "EY", id: "498320" },
  { n: "071", nome: "F.B. S.p.A.", id: "1191" },
  { n: "072", nome: "Fileni", id: "2651" },
  { n: "073", nome: "Focaccia Group Automotive", id: "1980" },
  { n: "074", nome: "FOR S.p.A.", id: "2652" },
  { n: "075", nome: "FORTNA", id: "1284" },
  { n: "076", nome: "Forvis Mazars Spa", id: "2105" },
  { n: "077", nome: "Futur-A Group", id: "2126" },
  { n: "078", nome: "Ghibson Italia S.r.l.", id: "1404" },
  { n: "079", nome: "Gi Group", id: "2361" },
  { n: "080", nome: "Grant Thornton", id: "1767" },
  { n: "081", nome: "Grundfos Water Treatment Italy Srl", id: "2653" },
  { n: "082", nome: "Gruppo Amadori", id: "615313" },
  { n: "083", nome: "Gruppo BCC Iccrea", id: "1202" },
  { n: "084", nome: "Gruppo Colorobbia", id: "1249" },
  { n: "085", nome: "Gruppo Concorde Spa", id: "499823" },
  { n: "086", nome: "Gruppo Ferrovie dello Stato Italiane", id: "499034" },
  { n: "087", nome: "Gruppo Fini spa", id: "2654" },
  { n: "088", nome: "GRUPPO GRANTERRE", id: "2244" },
  { n: "089", nome: "Gruppo Hera", id: "502354" },
  { n: "090", nome: "Gruppo Tampieri", id: "2181" },
  { n: "091", nome: "Gruppo Teddy S.p.A.", id: "500569" },
  { n: "092", nome: "Gruppo Trevi", id: "30" },
  { n: "093", nome: "Gruppo Zucchetti", id: "41" },
  { n: "094", nome: "Hitachi Energy", id: "2106" },
  { n: "095", nome: "Horsa Spa", id: "630293" },
  { n: "096", nome: "HPE Group", id: "503839" },
  { n: "097", nome: "ICONSULTING SPA", id: "606814" },
  { n: "098", nome: "IEMA GROUP", id: "1743" },
  { n: "099", nome: "IKOS Consulting Italia", id: "1784" },
  { n: "100", nome: "Illumia", id: "366" },
  { n: "101", nome: "IMA SPA", id: "606374" },
  { n: "102", nome: "ING. FERRARI S.P.A.", id: "502191" },
  { n: "103", nome: "INRES SC", id: "2656" },
  { n: "104", nome: "ITALIAN EXHIBITION GROUP SPA", id: "1345" },
  { n: "105", nome: "KPMG", id: "2582" },
  { n: "106", nome: "Laboratori Guglielmo Marconi spa", id: "1441" },
  { n: "107", nome: "Lafarmacia.", id: "861" },
  { n: "108", nome: "Lavoropiu SpA", id: "606942" },
  { n: "109", nome: "Lidl Italia", id: "504777" },
  { n: "110", nome: "Lincotek", id: "1802" },
  { n: "111", nome: "LivaNova", id: "1764" },
  { n: "112", nome: "Ludovico Martelli s.p.a.", id: "2658" },
  { n: "113", nome: "MAIRE S.p.A.", id: "2524" },
  { n: "114", nome: "MAPEI SPA", id: "1765" },
  { n: "115", nome: "Marazzi Group S.r.l", id: "503960" },
  { n: "116", nome: "Marchesini Group SpA", id: "498344" },
  { n: "117", nome: "MARPOSS S.p.A.", id: "608471" },
  { n: "118", nome: "Marr", id: "105" },
  { n: "119", nome: "Mechinno srl", id: "505822" },
  { n: "120", nome: "MEC3 - OPTIMA SPA | GRUPPO CASA OPTIMA", id: "1702" },
  { n: "121", nome: "Moltiply Group Spa", id: "2659" },
  { n: "122", nome: "Motor Power Company", id: "504160" },
  { n: "123", nome: "NEXIA AUDIREVI", id: "1344" },
  { n: "124", nome: "Novomatic Italia S.p.a.", id: "1982" },
  { n: "125", nome: "NX Engineering Srl", id: "2161" },
  { n: "126", nome: "Officine Maccaferri Spa", id: "1225" },
  { n: "127", nome: "Openjobmetis S.p.A.", id: "2660" },
  { n: "128", nome: "OPTIT", id: "608377" },
  { n: "129", nome: "OT Consulting", id: "769" },
  { n: "130", nome: "Philip Morris International", id: "503671" },
  { n: "131", nome: "Pirola Pennuto Zei & Associati", id: "664" },
  { n: "132", nome: "Profilglass s.p.a.", id: "501261" },
  { n: "133", nome: "PwC", id: "500639" },
  { n: "134", nome: "Randstad Italia Spa", id: "608463" },
  { n: "135", nome: "Raytec Vision SpA", id: "2662" },
  { n: "136", nome: "REKEEP SPA", id: "505674" },
  { n: "137", nome: "Relatech", id: "2761" },
  { n: "138", nome: "REPLY", id: "499018" },
  { n: "139", nome: "Rete servizi agricoltura", id: "2723" },
  { n: "140", nome: "Revlon", id: "1341" },
  { n: "141", nome: "Riccoboni Holding Spa", id: "2602" },
  { n: "142", nome: "Risorse", id: "303" },
  { n: "143", nome: "Rothoblaas Srl", id: "1283" },
  { n: "144", nome: "SACMI Imola S.C.", id: "606824" },
  { n: "145", nome: "SAINT-GOBAIN ITALIA", id: "2107" },
  { n: "146", nome: "SCM Group", id: "61" },
  { n: "147", nome: "SCS Azioninnova", id: "242" },
  { n: "148", nome: "Segula Technologies", id: "1226" },
  { n: "149", nome: "Sherwin-Williams Italy srl", id: "1261" },
  { n: "150", nome: "SOLUX SRL", id: "2528" },
  { n: "151", nome: "Stanzani SpA", id: "1193" },
  { n: "152", nome: "STEF", id: "1121" },
  { n: "153", nome: "System Logistics", id: "498485" },
  { n: "154", nome: "TAS SPA", id: "803" },
  { n: "155", nome: "Teach For Italy ETS", id: "1704" },
  { n: "156", nome: "TECHINT COMPAGNIA TECNICA INTERNAZIONALE S.p.A.", id: "2663" },
  { n: "157", nome: "Tennant Company", id: "2743" },
  { n: "158", nome: "Tetra Pak", id: "498472" },
  { n: "159", nome: "Timac Agro Italia", id: "1194" },
  { n: "160", nome: "TMC Italia", id: "2741" },
  { n: "161", nome: "Toyota Material Handling", id: "509015" },
  { n: "162", nome: "Trenitalia Tper", id: "2664" },
  { n: "163", nome: "Truzzi S.p.A.", id: "2665" },
  { n: "164", nome: "Tyche Bank", id: "2530" },
  { n: "165", nome: "UMANA", id: "503878" },
  { n: "166", nome: "UNICOOP FIRENZE", id: "2242" },
  { n: "167", nome: "UniCredit", id: "502865" },
  { n: "168", nome: "Unigra", id: "1541" },
  { n: "169", nome: "Unimpiego Confindustria srl", id: "615650" },
  { n: "170", nome: "UNITEC S.p.A.", id: "502215" },
  { n: "171", nome: "Var Group", id: "2666" },
  { n: "172", nome: "VINAVIL S.p.A.", id: "2667" },
  { n: "173", nome: "Vislab Srl", id: "1770" },
  { n: "174", nome: "Wienerberger S.p.A.", id: "1403" },
  { n: "175", nome: "XTEL SRL", id: "2669" },
];

const baseUrl = "https://eventi.unibo.it/careerday/aziende-partecipanti-2026/view/";

const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function labelCell(text, width) {
  return new TableCell({
    borders: cellBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "EAF2FB", type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: "Arial" })] })]
  });
}

function valueCell(text, width) {
  return new TableCell({
    borders: cellBorders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, font: "Arial" })] })]
  });
}

function emptyValueCell(width) {
  return new TableCell({
    borders: cellBorders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 200, bottom: 200, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })]
  });
}

function buildAziendaTable(az) {
  const url = baseUrl + az.id;
  const labelW = 2400;
  const valueW = 6626;
  const total = labelW + valueW;

  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: [labelW, valueW],
    rows: [
      // Header row: numero + nome azienda
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: total, type: WidthType.DXA },
            columnSpan: 2,
            shading: { fill: "1F3864", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            children: [new Paragraph({
              children: [
                new TextRun({ text: `[${az.n}]  ${az.nome}`, bold: true, color: "FFFFFF", size: 22, font: "Arial" })
              ]
            })]
          })
        ]
      }),
      // URL
      new TableRow({
        children: [
          labelCell("URL Scheda", labelW),
          valueCell(url, valueW),
        ]
      }),
      // Stato
      new TableRow({
        children: [
          labelCell("Stato", labelW),
          new TableCell({
            borders: cellBorders,
            width: { size: valueW, type: WidthType.DXA },
            margins: cellMargins,
            children: [new Paragraph({
              children: [
                new TextRun({ text: "[ ] Da verificare    [X] Verificato    [NO] Non pertinente", size: 20, font: "Arial" })
              ]
            })]
          }),
        ]
      }),
      // Posizioni
      new TableRow({
        children: [
          labelCell("Posizioni aperte\n(tirocinio/tesi)", labelW),
          emptyValueCell(valueW),
        ]
      }),
      // Note
      new TableRow({
        children: [
          labelCell("Note", labelW),
          emptyValueCell(valueW),
        ]
      }),
    ]
  });
}

// Build document
const children = [
  // Title
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
    children: [new TextRun({ text: "CAREER DAY - UNIVERSITÀ DI BOLOGNA 2026", bold: true, size: 36, font: "Arial", color: "1F3864" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text: "Tirocini per Tesi — Ingegneria Informatica", size: 26, font: "Arial", color: "2E74B5" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [new TextRun({ text: `Totale aziende: 175  |  Ultimo aggiornamento: ___________`, size: 20, font: "Arial", color: "666666" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E74B5", space: 1 } },
    children: [new TextRun({ text: "Legenda stato:  [ ] Da verificare  |  [X] Verificato  |  [NO] Non pertinente", size: 18, font: "Arial", color: "888888" })]
  }),
];

// Add each company table with spacing
aziende.forEach((az, i) => {
  children.push(buildAziendaTable(az));
  children.push(new Paragraph({ spacing: { after: 240 }, children: [new TextRun("")] }));
});

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 900, bottom: 1080, left: 900 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/career_day_unibo_2026_tirocini.docx', buffer);
  console.log('Done!');
}).catch(err => {
  console.error(err);
  process.exit(1);
});
