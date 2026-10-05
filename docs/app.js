"use strict";

// Farben wie in der Vorlage
const RKI_BLAU = [4, 90, 166];
const BLAU = [68, 114, 196];
const ORANGE = [237, 125, 49];
const GITTER = [217, 217, 217];
const TEXT = [89, 89, 89];
const TITEL = [64, 64, 64];

// Reihenfolge und Titel wie in der Vorlage; [kursiver Teil, normaler Teil]
const DIAGRAMME = [
  ["3GCECOLI", "E. coli", " Cephalosporine 3. Gen."],
  ["MRSA", "", "MRSA"],
  ["CRKP", "K. pneumoniae", " Carbapeneme"],
  ["VRE", "", "VRE"],
];
const ZIELJAHR = 2030;

let daten = null;
let pdfUrl = null;
let pdfName = "bericht.pdf";

const $ = (id) => document.getElementById(id);

function zahlDe(v, nd) {
  return v.toFixed(nd).replace(".", ",");
}
function ohneNullen(s) {
  return s.includes(",") ? s.replace(/0+$/, "").replace(/,$/, "") : s;
}
function beschriftung(code, v) {
  // Vorlage: 2 Nachkommastellen, bei CRKP gekürzt (0,2 / 0,196)
  return code === "CRKP" ? ohneNullen(zahlDe(v, 3)) : zahlDe(v, 2);
}

function zuZahl(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() !== "") {
    const n = Number(v.trim().replace(",", "."));
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

// ---------- Excel lesen ----------

function leseTabelle(buffer) {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const zeilen = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: null });
  if (zeilen.length < 2) throw new Error("Die Tabelle enthält keine Datenzeilen.");

  const kopf = zeilen[0].map((h) => (h == null ? "" : String(h).trim()));
  const jahrSpalten = [];
  kopf.forEach((h, i) => {
    if (/^(19|20)\d\d$/.test(h)) jahrSpalten.push([Number(h), i]);
  });
  if (jahrSpalten.length < 2) throw new Error("Keine Jahresspalten (z. B. 2019, 2020, …) in der Kopfzeile gefunden.");

  // Zielwert: Spalte "target_abs" bzw. "Target 2030 …"; sonst letzte Spalte
  let zielSpalte = kopf.findIndex((h) => /target_abs|target\s*2030/i.test(h));
  if (zielSpalte < 0) zielSpalte = kopf.length - 1;
  if (jahrSpalten.some(([, i]) => i === zielSpalte)) throw new Error("Spalte mit dem Zielwert 2030 (target_abs) fehlt.");

  const ergebnis = {};
  for (const z of zeilen.slice(1)) {
    const code = z[0] == null ? "" : String(z[0]).trim();
    if (!code) continue;
    const werte = jahrSpalten.map(([j, i]) => [j, zuZahl(z[i])]);
    // Jahre ohne Wert am Ende (z. B. noch nicht vorhanden) werden ignoriert
    while (werte.length && werte[werte.length - 1][1] == null) werte.pop();
    if (werte.some(([, v]) => v == null)) throw new Error(`Fehlender oder ungültiger Jahreswert in Zeile „${code}“.`);
    const ziel = zuZahl(z[zielSpalte]);
    if (ziel == null) throw new Error(`Zielwert 2030 fehlt in Zeile „${code}“.`);
    ergebnis[code] = { jahre: werte.map((w) => w[0]), werte: werte.map((w) => w[1]), ziel };
  }
  const fehlend = DIAGRAMME.map((d) => d[0]).filter((c) => !ergebnis[c]);
  if (fehlend.length) throw new Error("Folgende Zeilen fehlen: " + fehlend.join(", "));
  return ergebnis;
}

function zeigeTabelle(d) {
  const alleJahre = [...new Set(DIAGRAMME.flatMap(([c]) => d[c].jahre))].sort();
  let html = "<thead><tr><th>drug_bug</th>" + alleJahre.map((j) => `<th>${j}</th>`).join("") + "<th>Ziel 2030</th></tr></thead><tbody>";
  for (const [c] of DIAGRAMME) {
    html += `<tr><td>${c}</td>`;
    for (const j of alleJahre) {
      const i = d[c].jahre.indexOf(j);
      html += `<td>${i < 0 ? "–" : String(d[c].werte[i]).replace(".", ",")}</td>`;
    }
    html += `<td>${String(d[c].ziel).replace(".", ",")}</td></tr>`;
  }
  $("tabelle").innerHTML = html + "</tbody>";
}

// ---------- PDF zeichnen (Maße in mm, A4 quer) ----------

function achse(vmax) {
  // Achsenmaximum und Schrittweite ähnlich der Excel-Automatik
  for (const schritt of [0.05, 0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50]) {
    const n = Math.round((vmax / schritt) * 1e6) / 1e6;
    const anzahl = Math.floor(n) + 1;
    if (anzahl <= 8) return { max: anzahl * schritt, schritt };
  }
  return { max: vmax * 1.1, schritt: vmax / 5 };
}

function achsenText(v, schritt) {
  if (schritt >= 1) return String(Math.round(v));
  const nd = String(schritt).split(".")[1].length;
  return ohneNullen(zahlDe(v, nd));
}

function zeichneDiagramm(doc, x, y, w, h, [code, kursiv, normal], d) {
  // Rahmen
  doc.setDrawColor(...GITTER);
  doc.setLineWidth(0.25);
  doc.setLineDashPattern([], 0);
  doc.rect(x, y, w, h);

  // Titel (kursiver + normaler Teil, gemeinsam zentriert)
  doc.setFontSize(9);
  doc.setTextColor(...TITEL);
  doc.setFont("helvetica", "bolditalic");
  const w1 = kursiv ? doc.getTextWidth(kursiv) : 0;
  doc.setFont("helvetica", "bold");
  const w2 = doc.getTextWidth(normal);
  let tx = x + w / 2 - (w1 + w2) / 2;
  if (kursiv) {
    doc.setFont("helvetica", "bolditalic");
    doc.text(kursiv, tx, y + 7);
    tx += w1;
    doc.setFont("helvetica", "bold");
  }
  doc.text(normal, tx, y + 7);

  // Plotbereich
  const px = x + 13, py = y + 13, pw = w - 19, ph = h - 32;
  const { jahre, werte, ziel } = d;
  const j0 = jahre[0];
  const sx = (j) => px + ((j - j0 + 0.5) / (ZIELJAHR - j0 + 1)) * pw;
  const { max, schritt } = achse(Math.max(...werte, ziel));
  const sy = (v) => py + ph - (v / max) * ph;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  doc.setTextColor(...TEXT);
  doc.text("Inzidenz", x + 3, y + 7);

  doc.setFontSize(6.5);
  for (let i = 0; i * schritt <= max + 1e-9; i++) {
    const v = i * schritt;
    doc.setDrawColor(...GITTER);
    doc.setLineWidth(0.2);
    doc.line(px, sy(v), px + pw, sy(v));
    doc.text(achsenText(v, schritt), px - 2, sy(v) + 0.8, { align: "right" });
  }
  for (let j = j0; j <= ZIELJAHR; j++) {
    doc.text(String(j), sx(j), py + ph + 4, { align: "center" });
  }

  // Start-Ziel-Linie (gepunktet)
  doc.setDrawColor(...ORANGE);
  doc.setLineWidth(0.3);
  doc.setLineDashPattern([0.4, 0.6], 0);
  doc.line(sx(j0), sy(werte[0]), sx(ZIELJAHR), sy(ziel));
  doc.setLineDashPattern([], 0);

  // Datenreihe
  doc.setDrawColor(...BLAU);
  doc.setFillColor(...BLAU);
  doc.setLineWidth(0.75);
  doc.setLineJoin("round");
  for (let i = 1; i < werte.length; i++) {
    doc.line(sx(jahre[i - 1]), sy(werte[i - 1]), sx(jahre[i]), sy(werte[i]));
  }
  for (let i = 0; i < werte.length; i++) doc.circle(sx(jahre[i]), sy(werte[i]), 0.8, "F");

  // Start- und Zielpunkt mit Beschriftung
  doc.setFillColor(...ORANGE);
  doc.circle(sx(j0), sy(werte[0]), 0.85, "F");
  doc.circle(sx(ZIELJAHR), sy(ziel), 0.85, "F");
  doc.setTextColor(...TEXT);
  doc.text(beschriftung(code, werte[0]), sx(j0), sy(werte[0]) - 2, { align: "center" });
  doc.text(beschriftung(code, ziel), sx(ZIELJAHR), sy(ziel) - 2, { align: "center" });

  // Legende
  const ly = y + h - 6;
  const eintraege = [
    [`${j0} - ${jahre[jahre.length - 1]}`, BLAU, "linie"],
    ["start - ziel", ORANGE, "linie"],
    ["Linear (start - ziel)", ORANGE, "punkte"],
  ];
  doc.setFontSize(6.5);
  const breiten = eintraege.map(([t]) => 10 + doc.getTextWidth(t));
  const gesamt = breiten.reduce((a, b) => a + b, 0) + 8 * (eintraege.length - 1);
  let lx = x + w / 2 - gesamt / 2;
  eintraege.forEach(([t, farbe, art], i) => {
    doc.setDrawColor(...farbe);
    doc.setFillColor(...farbe);
    if (art === "linie") {
      doc.setLineWidth(0.75);
      doc.line(lx, ly, lx + 8, ly);
      doc.circle(lx + 4, ly, 0.8, "F");
    } else {
      doc.setLineWidth(0.3);
      doc.setLineDashPattern([0.4, 0.6], 0);
      doc.line(lx, ly, lx + 8, ly);
      doc.setLineDashPattern([], 0);
    }
    doc.setTextColor(...TEXT);
    doc.text(t, lx + 10, ly + 0.8);
    lx += breiten[i] + 8;
  });
}

function erzeugePdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const titel = $("f-titel").value;
  const datum = $("f-datum").value;

  doc.setProperties({ title: `${titel} – ${$("f-fuss").value}` });

  // Folientitel
  doc.setFont("helvetica", "bold");
  doc.setFontSize(24);
  doc.setTextColor(...RKI_BLAU);
  doc.text(titel, 15, 45);

  // 2 × 2 Diagramme
  const bw = 106, bh = 63, x1 = 15, x2 = 126, y1 = 50, y2 = 116;
  const pos = [[x1, y1], [x2, y1], [x1, y2], [x2, y2]];
  DIAGRAMME.forEach((dg, i) => zeichneDiagramm(doc, pos[i][0], pos[i][1], bw, bh, dg, daten[dg[0]]));

  // Fußzeile
  doc.setDrawColor(...RKI_BLAU);
  doc.setLineWidth(0.25);
  for (const xl of [15, 84, 262, 281]) doc.line(xl, 184, xl, 197);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...RKI_BLAU);
  doc.text(datum, 18, 188);
  doc.text($("f-fuss").value, 87, 188);
  doc.text($("f-seite").value, 271.5, 188, { align: "center" });

  return doc;
}

// ---------- Oberfläche ----------

function meldung(text, art) {
  const m = $("meldung");
  m.textContent = text;
  m.className = "meldung " + (art || "");
}

async function dateiGewaehlt(datei) {
  if (!datei) return;
  try {
    daten = leseTabelle(await datei.arrayBuffer());
    zeigeTabelle(daten);
    $("schritt2").classList.remove("versteckt");
    $("schritt3").classList.remove("versteckt");
    $("vorschau").classList.add("versteckt");
    $("laden").disabled = true;
    pdfName = datei.name.replace(/\.[^.]+$/, "") + "_bericht.pdf";
    meldung(`„${datei.name}“ eingelesen.`, "ok");
  } catch (e) {
    daten = null;
    $("schritt2").classList.add("versteckt");
    $("schritt3").classList.add("versteckt");
    meldung("Fehler: " + e.message, "fehler");
  }
}

function vorlageHerunterladen(ev) {
  ev.preventDefault();
  const kopf = ["drug_bug", 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, "target_perc", "target_abs"];
  const zeilen = [kopf, ...DIAGRAMME.map(([c]) => [c])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(zeilen), "Tabelle1");
  XLSX.writeFile(wb, "vorlage_dart_ziele.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const heute = new Date();
  $("f-datum").value = [heute.getDate(), heute.getMonth() + 1].map((n) => String(n).padStart(2, "0")).join(".") + "." + heute.getFullYear();

  const drop = $("drop");
  drop.addEventListener("click", () => $("datei").click());
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("datei").click(); } });
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("aktiv"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("aktiv"));
  drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("aktiv"); dateiGewaehlt(e.dataTransfer.files[0]); });
  $("datei").addEventListener("change", (e) => dateiGewaehlt(e.target.files[0]));
  $("vorlage").addEventListener("click", vorlageHerunterladen);

  $("erzeugen").addEventListener("click", () => {
    if (!daten) return;
    const doc = erzeugePdf();
    if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    pdfUrl = URL.createObjectURL(doc.output("blob"));
    $("vorschau").src = pdfUrl;
    $("vorschau").classList.remove("versteckt");
    $("laden").disabled = false;
  });
  $("laden").addEventListener("click", () => {
    const a = document.createElement("a");
    a.href = pdfUrl;
    a.download = pdfName;
    a.click();
  });
});
