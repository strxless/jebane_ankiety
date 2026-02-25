// api/export.js  – faithful replica of the printed questionnaire
import { getClient, initSchema } from "./_db.js";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, ShadingType, BorderStyle, VerticalAlign,
  UnderlineType
} from "docx";
import JSZip from "jszip";

export const config = { runtime: "nodejs", maxDuration: 30 };

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const TICK  = "\u2612"; // ☒  checked box
const EMPTY = "\u2610"; // ☐  empty box
const FONT  = "Arial Narrow";
const SZ    = 16;  // 8pt in half-points — tight like the printed form
const SZ_SM = 14;  // 7pt for small items
const SZ_H  = 18;  // 9pt for section headers
const W     = 9800; // full content width in DXA (A4 narrow margins)

// ─── CHECKBOX HELPERS ─────────────────────────────────────────────────────────
function ck(answers, key, value) {
  const a = answers[key];
  if (!a) return EMPTY;
  const v = String(value).toLowerCase();
  const norm = x => String(x).toLowerCase();
  if (Array.isArray(a)) return a.some(x => norm(x) === v || norm(x).startsWith(v.split('.')[0] + '.')) ? TICK : EMPTY;
  return (norm(a) === v || norm(a).startsWith(v.split('.')[0] + '.')) ? TICK : EMPTY;
}
function yn(answers, key) {
  const v = answers[key];
  return { t: v === 'TAK' ? TICK : EMPTY, n: v === 'NIE' ? TICK : EMPTY };
}

// ─── TEXT / PARAGRAPH HELPERS ─────────────────────────────────────────────────
function r(text, opts = {}) {
  return new TextRun({
    text: String(text ?? ''),
    font: FONT,
    size: opts.sz ?? SZ,
    bold: opts.bold ?? false,
    color: opts.color,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function para(children, opts = {}) {
  const runs = (Array.isArray(children) ? children : [r(children)]);
  return new Paragraph({
    children: runs,
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: {
      before: opts.before ?? 30,
      after:  opts.after  ?? 30,
      line:   opts.line   ?? 220,
    },
    indent: opts.indent ? { left: opts.indent } : undefined,
  });
}

// ─── BORDER HELPERS ───────────────────────────────────────────────────────────
const B1  = { style: BorderStyle.SINGLE, size: 4,  color: "000000" };
const B_NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const BORDERS     = { top: B1, bottom: B1, left: B1, right: B1 };
const BORDERS_NONE= { top: B_NONE, bottom: B_NONE, left: B_NONE, right: B_NONE };
// inner borders (no outer) — for cells inside a merged section
const B_INNER_R   = { top: B_NONE, bottom: B_NONE, left: B_NONE, right: B1   };
const B_INNER_L   = { top: B_NONE, bottom: B_NONE, left: B1,     right: B_NONE };

// ─── CELL HELPERS ─────────────────────────────────────────────────────────────
function cell(children, { w = 4900, bg, borders = BORDERS, vAlign } = {}) {
  const paras = (Array.isArray(children) ? children : [children]).map(c =>
    c instanceof Paragraph ? c : para(c)
  );
  return new TableCell({
    children: paras,
    width: { size: w, type: WidthType.DXA },
    borders,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    verticalAlign: vAlign ?? VerticalAlign.TOP,
    ...(bg ? { shading: { fill: bg, type: ShadingType.CLEAR } } : {}),
  });
}

// Full-width header row (grey background)
function hdrRow(text, totalW = W, cols = 1) {
  return new TableRow({
    children: [new TableCell({
      children: [para([r(text, { bold: true, sz: SZ_H })], { before: 40, after: 40 })],
      columnSpan: cols,
      width:    { size: totalW, type: WidthType.DXA },
      borders:  BORDERS,
      margins:  { top: 40, bottom: 40, left: 80, right: 80 },
      shading:  { fill: "D9D9D9", type: ShadingType.CLEAR },
    })]
  });
}

// Two-column table row
function row2(leftChildren, rightChildren, lw = W/2, rw = W/2) {
  return new TableRow({
    children: [
      cell(Array.isArray(leftChildren)  ? leftChildren  : [leftChildren],  { w: lw }),
      cell(Array.isArray(rightChildren) ? rightChildren : [rightChildren], { w: rw }),
    ]
  });
}

// Three-column table row
function row3(a, b, c, w1 = W/3, w2 = W/3, w3 = W/3) {
  return new TableRow({ children: [
    cell(Array.isArray(a) ? a : [a], { w: w1 }),
    cell(Array.isArray(b) ? b : [b], { w: w2 }),
    cell(Array.isArray(c) ? c : [c], { w: w3 }),
  ]});
}

// Convenience: full table with single border
function tbl(rows, colWidths) {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: colWidths ?? [W],
    rows,
  });
}

// A checklist rendered as compact paragraphs
function checks(answers, key, opts, indent = 0) {
  return opts.map(opt =>
    para([r(`${ck(answers, key, opt)} ${opt}`, { sz: SZ })], {
      before: 20, after: 20, indent
    })
  );
}

// Two-column checklist table (splits opts in half)
function checks2col(answers, key, opts) {
  const mid  = Math.ceil(opts.length / 2);
  const left = opts.slice(0, mid);
  const rght = opts.slice(mid);
  const n    = Math.max(left.length, rght.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push(new TableRow({ children: [
      cell(left[i]  ? [para([r(`${ck(answers, key, left[i])} ${left[i]}`,  { sz: SZ })], { before: 20, after: 20 })] : [para('')], { w: W/2 }),
      cell(rght[i] ? [para([r(`${ck(answers, key, rght[i])} ${rght[i]}`, { sz: SZ })], { before: 20, after: 20 })] : [para('')], { w: W/2 }),
    ]}));
  }
  return tbl(rows, [W/2, W/2]);
}

// ─── MAIN DOCUMENT BUILDER ────────────────────────────────────────────────────
function buildDoc(record) {
  const a  = record.answers || {};
  const ch = []; // children array

  // ══ TITLE ══════════════════════════════════════════════════════════════════
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para(
        [r("Kwestionariusz osoby w kryzysie bezdomności w ramach Ogólnopolskiego badania liczby osób w kryzysie bezdomności – rok badania: 2026*", { bold: true, sz: SZ_H })],
        { align: AlignmentType.CENTER, before: 60, after: 60 }
      )],
      width: { size: W, type: WidthType.DXA },
      borders: BORDERS,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
    })] })
  ], [W]));

  // ══ WSTĘP ══════════════════════════════════════════════════════════════════
  ch.push(tbl([
    hdrRow("WSTĘP"),
    new TableRow({ children: [new TableCell({
      children: [para([
        r("W przypadku stwierdzenia przez ankietera zagrożenia życia lub zdrowia osoby bezdomnej należy niezwłocznie powiadomić odpowiednie służby, w tym policję – tel. 112 i 997", { bold: true, sz: SZ })
      ], { align: AlignmentType.CENTER, before: 50, after: 50 })],
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  // ══ WYWIAD / ZGODA / DUPLIKAT / SPOSÓB ════════════════════════════════════
  const { t: wT, n: wN } = yn(a, 'wywiad_dzisiaj');
  const { t: zT, n: zN } = yn(a, 'zgoda_udzial');
  const dupBox = a.duplikat === 'TAK' ? TICK : EMPTY;
  const spP    = ck(a, 'sposob_wypelnienia', 'Pełny');
  const spS    = ck(a, 'sposob_wypelnienia', 'Skrócony');

  ch.push(tbl([
    row2(
      [para([r("■ Czy był przeprowadzony z Panią/Panem taki wywiad dzisiaj?  ", { bold: true, sz: SZ }), r(`${wT} Tak,  ${wN} Nie`, { sz: SZ })], { before: 40, after: 20 }),
       para([r(`■ Jeśli tak, zakończyć i zaznaczyć duplikat: ${dupBox}`, { sz: SZ })], { before: 20, after: 40 })],
      [para([r("■ Czy zgadza się Pan/i na udział w badaniu?  ", { bold: true, sz: SZ }), r(`${zT} Tak,  ${zN} Nie`, { sz: SZ })], { before: 40, after: 20 }),
       para([r("■ Sposób wypełnienia:  ", { bold: true, sz: SZ }), r(`${spP} pełny / ${spS} skrócony`, { sz: SZ })], { before: 20, after: 40 })],
    )
  ], [W/2, W/2]));

  // UWAGA box
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([
        r("UWAGA!!! ", { bold: true, sz: SZ }),
        r("Pierwszym pytaniem, które należy zadać respondentowi jest pytanie czy w dniu dzisiejszym był badany tym wywiadem. Jeśli dana osoba już uczestniczyła w wywiadzie prosimy nie rozpoczynać wywiadu. Jeśli z osobą bezdomną z pewnych względów jest utrudniony kontakt bądź odmawia wzięcia udziału w badaniu, prosimy o wypełnienie kwestionariusza z zaznaczeniem miejsca przebywania, płci, szacowanego wieku. W przypadku dzieci (0-17 lat) wypełniamy tylko pytania 1-4 oraz miejsce przebywania.", { sz: SZ }),
      ], { before: 40, after: 40 })],
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  // ══ MIEJSCE ════════════════════════════════════════════════════════════════
  ch.push(tbl([
    hdrRow("MIEJSCE PRZEPROWADZENIA BADANIA / PRZEBYWANIA OSOBY W KRYZYSIE BEZDOMNOŚCI"),
    new TableRow({ children: [new TableCell({
      children: [para([r("Województwo – Pomorskie      Powiat – Gdynia      Gmina – Gdynia      Miejscowość – Gdynia", { bold: true, sz: SZ })], { before: 40, after: 40 })],
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  const { t: mT, n: mN } = yn(a, 'miasto_powyzej_100k');
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([r("Czy miasto powyżej 100 tysięcy mieszkańców?  ", { bold: true, sz: SZ }), r(`${mT} Tak,  ${mN} Nie`, { sz: SZ })])],
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  // Place checkboxes — 2 col exactly like the form
  const miejsca = [
    [1,"Noclegownia"],[2,"Ogrzewalnia"],[3,"Schronisko dla osób bezdomnych"],
    [4,"Schronisko dla osób bezdomnych z usługami opiekuńczymi"],
    [5,"Mieszkanie wspomagane"],[6,"Mieszkanie treningowe"],
    [7,"Dom dla matek z małoletnimi dziećmi i kobiet w ciąży"],
    [8,"Ośrodek interwencji kryzysowej"],
    [9,"Specjalistyczny ośrodek wsparcia dla osób doznających przemocy domowej"],
    [10,"Szpital, hospicjum, Zakład Opiekuńczo-Leczniczy (ZOL), inna placówka zdrowia"],
    [11,"Zakład karny, areszt śledczy"],[12,"Izba wytrzeźwień, pogotowie socjalne"],
    [13,"Instytucja zdrowia psychicznego/leczenia uzależnień"],
    [14,"Inna placówka/miejsce mieszkalne"],[15,"Pustostan"],
    [16,"Domek na działce, altana działkowa"],
    [17,"Miejsce niemieszkalne: ulica, klatka schodowa, dworzec PKP/PKS, altana śmietnikowa, piwnica, itp."],
  ];
  const stored = a.miejsce_pobytu || '';
  const leftM  = miejsca.slice(0, 8);
  const rightM = miejsca.slice(8);
  const n = Math.max(leftM.length, rightM.length);
  const placeRows = [];
  for (let i = 0; i < n; i++) {
    const lI = leftM[i];
    const rI = rightM[i];
    const lB = lI ? (stored.startsWith(`${lI[0]}.`) ? TICK : EMPTY) : '';
    const rB = rI ? (stored.startsWith(`${rI[0]}.`) ? TICK : EMPTY) : '';
    placeRows.push(row2(
      lI ? [para([r(`${lB} ${lI[0]}. ${lI[1]}`, { sz: SZ })], { before: 20, after: 20 })] : [para('')],
      rI ? [para([r(`${rB} ${rI[0]}. ${rI[1]}`, { sz: SZ })], { before: 20, after: 20 })] : [para('')],
    ));
  }
  ch.push(tbl(placeRows, [W/2, W/2]));

  const nazwa = a.miejsce_nazwa || '';
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([
        r("Po zaznaczeniu proszę wpisać nazwę miejsca/opisać je:  ", { sz: SZ }),
        r(nazwa || "___________________________________________________________________", { sz: SZ }),
      ], { before: 40, after: 40 })],
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  // ══ PYTANIA HEADER ════════════════════════════════════════════════════════
  ch.push(tbl([hdrRow("PYTANIA")], [W]));

  // ── P1 Płeć + P2 Wiek ────────────────────────────────────────────────────
  const kB = ck(a, 'p1_plec', '1.1. Kobieta');
  const mB = ck(a, 'p1_plec', '1.2. Mężczyzna');
  const age = a.p2_wiek_liczba || '......';
  const wD  = ck(a, 'p2_wiek_typ', 'Wiek deklarowany');
  const wO  = ck(a, 'p2_wiek_typ', 'Wiek oszacowany');
  const wKD = ck(a, 'p2_wiek_kategoria', 'Osoba dorosła (pow. 18 lat)');
  const wKDz= ck(a, 'p2_wiek_kategoria', 'Dziecko (0–17 lat)');

  ch.push(tbl([
    row2(
      [para([r("1. Płeć:  ", { bold: true, sz: SZ }), r(`1.1. kobieta ${kB}    1.2. mężczyzna ${mB}`, { sz: SZ })], { before: 40, after: 40 })],
      [para([r("2. Wiek:  ", { bold: true, sz: SZ }), r(`${age} lat   ${wD} wiek deklarowany   ${wO} wiek oszacowany`, { sz: SZ })], { before: 40, after: 20 }),
       para([r(`${wKD} osoba dorosła (pow. 18 lat)    ${wKDz} dziecko (0–17 lat)`, { sz: SZ })], { before: 20, after: 40 })],
    )
  ], [W/2, W/2]));

  // ── P3 Obywatelstwo ──────────────────────────────────────────────────────
  const obywOpts   = ["Polskie","Ukraińskie","Inne z Europy (wyłączając Ukrainę)","Inne z Azji","Inne z Afryki","Inne pozostałe lub brak"];
  const statusOpts = ["Uchodźcy","Ochrona tymczasowa","Stały pobyt","Nieuregulowany"];
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([r("3. Obywatelstwo i dane o statusie uchodźcy", { bold: true, sz: SZ })], { before: 40, after: 20 })],
      columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] }),
    row2(
      [para([r("3.1. Obywatelstwo", { bold: true, sz: SZ })], { before: 20, after: 20 }), ...checks(a, 'p3_obywatelstwo', obywOpts)],
      [para([r("3.2. Status cudzoziemca", { bold: true, sz: SZ })], { before: 20, after: 20 }), ...checks(a, 'p3_status_cudzoziemca', statusOpts)],
    ),
  ], [W/2, W/2]));

  // ── P4 Zameldowanie + P5 Czas bezdomności ────────────────────────────────
  ch.push(tbl([
    row2(
      [para([r("4. Czy posiada Pan(i) zameldowanie na pobyt stały?:", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p4_zameldowanie', [
         "4.1. tak, w gminie obecnego pobytu",
         "4.2. tak, poza gminą obecnego pobytu",
         "4.3. nie, ostatnie zameldowanie było w gminie obecnego pobytu",
         "4.4. nie, ostatnie zameldowanie było poza gminą obecnego pobytu",
       ])],
      [para([r("5. Jak długo doświadcza Pan/i bezdomności?", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p5_czas_bezdomnosci', [
         "5.1. do 3 miesięcy","5.2. od 3 do 6 miesięcy","5.3. od 6 do 12 miesięcy",
         "5.4. od 12 do 24 miesięcy","5.5. od 2 do 5 lat","5.6. od 5 do 10 lat",
         "5.7. od 10 lat do 20 lat","5.8. powyżej 20 lat",
       ])],
    )
  ], [W/2, W/2]));

  // ── P6 Stan cywilny + P7 Wykształcenie + P8 Gospodarstwo ────────────────
  const w3 = Math.floor(W/3);
  ch.push(tbl([
    row3(
      [para([r("6. Stan cywilny", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p6_stan_cywilny', ["6.1. kawaler/panna","6.2. żonaty/zamężna","6.3. rozwiedziony/rozwiedziona","6.4. wdowiec/wdowa","6.5. w wolnym związku","6.6. w separacji"])],
      [para([r("7. Wykształcenie", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p7_wyksztalcenie', ["7.1. niepełne podstawowe","7.2. podstawowe","7.3. gimnazjalne","7.4. zawodowe","7.5. średnie (techniczne też)","7.6. wyższe","7.7. nie wiem"])],
      [para([r("8. Z kim obecnie Pani/Pan gospodaruje:", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p8_gospodarstwo', ["8.1. samodzielnie/samotnie","8.2. partner/partnerka","8.3. kolega/koleżanka/znajomy/znajoma","8.4. małoletnie dzieci (0–17 lat)","8.5. dorosłe dzieci/członkowie dalszej rodziny","8.6. zbiorowo/w grupie"])],
      w3, W - w3*2, w3,
    )
  ], [w3, W - w3*2, w3]));

  // ── P9 Dochody ───────────────────────────────────────────────────────────
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([r("9. Jakie źródła dochodu Pan(i) posiada? (Można zaznaczyć dowolną liczbę odpowiedzi):", { bold: true, sz: SZ })], { before: 40, after: 20 })],
      columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] }),
    ...buildCheckRows(a, 'p9_dochody', [
      "9.1. zatrudnienie;","9.2. praca „na czarno\";",
      "9.3. praca chroniona/zatrudnienie wspierane;","9.4. zbieractwo;",
      "9.5. zasiłek z pomocy społecznej;","9.6. świadczenia ZUS;",
      "9.7. żebractwo;","9.8. alimenty;",
      "9.9. renta/emerytura;","9.10. nie posiadam dochodu",
      "9.11. odmowa odpowiedzi","",
    ]),
  ], [W/2, W/2]));

  // ── P10 Przyczyny ────────────────────────────────────────────────────────
  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: [para([r("10. Które wydarzenia były według Pana(i) przyczyną bezdomności? (proszę zaznaczyć maksymalnie 3):", { bold: true, sz: SZ })], { before: 40, after: 20 })],
      columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] }),
    ...buildCheckRows(a, 'p10_przyczyny', [
      "10.1. konflikt rodzinny","10.2. odejście/śmierć rodzica/opiekuna w dzieciństwie",
      "10.3. przemoc domowa","10.4. rozpad związku",
      "10.5. zadłużenie","10.6. bezrobocie, brak pracy, utrata pracy",
      "10.7. problemy wynikające z orientacji seksualnej","10.8. zły stan zdrowia, niepełnosprawność",
      "10.9. eksmisja, wymeldowanie z mieszkania","10.10. uzależnienie od alkoholu",
      "10.11. uzależnienie od narkotyków","10.12. uzależnienie od hazardu",
      "10.13. migracja/wyjazd na stałe do innego kraju","10.14. choroba/zaburzenia psychiczne inne niż uzależnienia",
      "10.15. opuszczenie placówki opiekuńczo-wychowawczej","10.16. opuszczenie zakładu karnego",
      "10.17. konflikt z prawem","10.18. inna przemoc niż domowa",
      "10.19. problemy wynikające ze zmiany wiary","10.20. odmowa odpowiedzi",
    ]),
  ], [W/2, W/2]));

  // ── P11 Pomoc + P12 Oczekiwane wsparcie ─────────────────────────────────
  const p12opts_l = ["12.1. żywnościowe","12.2. higieniczne (w tym dostęp do łaźni)","12.3. zdrowotne","12.4. schronienie","12.5. terapia uzależnień","12.6. wsparcie psychologiczne"];
  const p12opts_r = ["12.7. pomoc prawna","12.8. pomoc w znalezieniu pracy","12.9. finansowe","12.10. mieszkaniowe","12.11. wyjście z długów","12.12. nie oczekuję pomocy"];

  ch.push(tbl([
    row2(
      [para([r("11. Czy Pan(i) korzysta z pomocy i w jakiej postaci? (proszę zaznaczyć wszystkie formy, z których osoba korzysta):", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p11_pomoc', ["11.1. wsparcie finansowe","11.2. posiłek","11.3. odzież","11.4. schronienie","11.5. terapia uzależnień","11.6. opieka zdrowotna","11.7. nie korzystam"])],
      [para([r("12. W jakich obszarach oczekuje Pan(i) wsparcia/pomocy? (należy zaznaczyć maksymalnie 3 potrzeby)", { bold: true, sz: SZ })], { before: 40, after: 20 }),
       ...checks(a, 'p12_oczekiwane_wsparcie', [...p12opts_l, ...p12opts_r])],
    )
  ], [W/2, W/2]));

  // ── P13 Bezdomność ukryta ────────────────────────────────────────────────
  const { t: t1, n: n1 } = yn(a, 'p13_1_czy_pomieszkuje');
  const { t: t2, n: n2 } = yn(a, 'p13_2_czy_pomieszkiwal');
  const { t: t3, n: n3 } = yn(a, 'p13_3_czy_zna');
  const dOpts = ["do 3 miesięcy","od 3 do 12 miesięcy","od 1 roku do 2 lat","od 2 do 5 lat","powyżej 5 lat"];
  const ileOpts = ["1–2","3–5","6–10","więcej niż 10"];

  const p13rows = [
    para([r("13. Pytania o tzw. bezdomność ukrytą", { bold: true, sz: SZ })], { before: 40, after: 20 }),
    para([r("13.1. Czy obecnie Pan(i) pomieszkuje w domu/mieszkaniu u rodziny, znajomych, czy innych osób, tj. nie ma Pan(i) własnego miejsca zamieszkania i przebywa tymczasowo u innych osób?", { sz: SZ })], { before: 20, after: 10 }),
    para([r(`${t1} Tak    ${n1} Nie`, { sz: SZ })], { before: 10, after: 20 }),
  ];
  if (a.p13_1_czy_pomieszkuje === 'TAK') {
    p13rows.push(para([r("13.1.2. Jak długo obecnie Pan(i) pomieszkuje w domu/mieszkaniu u rodziny, znajomych, innych osób nie mając własnego miejsca zamieszkania?", { sz: SZ })], { before: 20, after: 10 }));
    p13rows.push(...dOpts.map((o, i) => para([r(`${ck(a,'p13_1_2_jak_dlugo', o)} 13.1.2.${i+1}. ${o}`, { sz: SZ })], { before: 10, after: 10 })));
  }
  p13rows.push(para([r("13.2. Czy w przeszłości Pan(i) tymczasowo pomieszkiwał(a) w domu/mieszkaniu u rodziny, znajomych, czy innych osób, nie mając własnego miejsca zamieszkania?", { sz: SZ })], { before: 20, after: 10 }));
  p13rows.push(para([r(`${t2} Tak    ${n2} Nie`, { sz: SZ })], { before: 10, after: 20 }));
  if (a.p13_2_czy_pomieszkiwal === 'TAK') {
    p13rows.push(para([r("13.2. Jak długo tymczasowo Pan(i) pomieszkiwał(a) w domu/mieszkaniu u rodziny, znajomych, innych osób, nie mając własnego miejsca zamieszkania?", { sz: SZ })], { before: 20, after: 10 }));
    p13rows.push(...dOpts.map((o, i) => para([r(`${ck(a,'p13_2_jak_dlugo', o)} 13.2.${i+1}. ${o}`, { sz: SZ })], { before: 10, after: 10 })));
  }
  p13rows.push(para([r("13.3. Czy zna Pan(i) inne osoby, które w ciągu ostatnich 12 miesięcy tymczasowo pomieszkiwały w domu/mieszkaniu u rodziny, znajomych, innych osób, nie mając własnego miejsca zamieszkania?", { sz: SZ })], { before: 20, after: 10 }));
  p13rows.push(para([r(`${t3} Tak    ${n3} Nie`, { sz: SZ })], { before: 10, after: 10 }));
  if (a.p13_3_czy_zna === 'TAK') {
    p13rows.push(para([
      r("jeśli Tak, ile Pan(i) zna takich osób:  ", { sz: SZ }),
      ...ileOpts.map(o => r(`${ck(a,'p13_3_ile_osob',o)} ${o}   `, { sz: SZ })),
    ], { before: 10, after: 20 }));
  }

  ch.push(tbl([
    new TableRow({ children: [new TableCell({
      children: p13rows,
      width: { size: W, type: WidthType.DXA }, borders: BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    })] })
  ], [W]));

  // ══ FUNKCJA ANKIETERA ═════════════════════════════════════════════════════
  ch.push(tbl([hdrRow("FUNKCJA ANKIETERA")], [W]));
  ch.push(tbl([
    row2(
      checks(a, 'funkcja_ankietera', ["Wolontariusz","Pracownik socjalny","Pracownik placówki dla bezdomnych","Inna"]),
      checks(a, 'funkcja_ankietera', ["Pracownik gminy","Strażnik miejski/policjant","Pracownik do spraw streetworkingu (Streetworker)"]),
    )
  ], [W/2, W/2]));

  // ══ FOOTER ════════════════════════════════════════════════════════════════
  ch.push(para([r("*Wzór kwestionariusza może ulec zmianie. W takim przypadku zostanie ona zakomunikowana w odpowiednim czasie przed badaniem.", { sz: SZ_SM })], { before: 80, after: 20 }));
  ch.push(para([r(`Wygenerowano automatycznie | ID: ${record.id} | ${(record.ts || '').slice(0, 10)}`, { sz: 12, color: "999999" })], { before: 0, after: 0 }));

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 500, bottom: 500, left: 600, right: 600 }
        }
      },
      children: ch,
    }]
  });
}

// ─── HELPER: split options into 2 col rows ────────────────────────────────────
function buildCheckRows(answers, key, opts) {
  const mid  = Math.ceil(opts.length / 2);
  const left = opts.slice(0, mid);
  const rght = opts.slice(mid);
  const n    = Math.max(left.length, rght.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const lT = left[i] || '';
    const rT = rght[i] || '';
    rows.push(new TableRow({ children: [
      cell(lT ? [para([r(`${ck(answers, key, lT)} ${lT}`, { sz: SZ })], { before: 20, after: 20 })] : [para('')], { w: W/2 }),
      cell(rT ? [para([r(`${ck(answers, key, rT)} ${rT}`, { sz: SZ })], { before: 20, after: 20 })] : [para('')], { w: W/2 }),
    ]}));
  }
  return rows;
}

// ─── HTTP HANDLER ─────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).end();

  try {
    await initSchema();
    const db = getClient();
    const { id, ids } = req.query;

    // ── single docx ──────────────────────────────────────────────────────────
    if (id) {
      const result = await db.execute({ sql: "SELECT * FROM responses WHERE id = ?", args: [Number(id)] });
      if (!result.rows.length) return res.status(404).json({ error: "Not found" });
      const row    = result.rows[0];
      const record = { id: row[0], ext_id: row[1], ts: row[2], created: row[3], answers: JSON.parse(row[4]) };
      const buf    = await Packer.toBuffer(buildDoc(record));
      const ts     = (record.ts || record.created || '').slice(0, 10);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="kwestionariusz_${String(record.id).padStart(4,'0')}_${ts}.docx"`);
      return res.status(200).send(Buffer.from(buf));
    }

    // ── ZIP (selected or all) ─────────────────────────────────────────────────
    let rows;
    if (ids) {
      const idList = ids.split(',').map(x => Number(x.trim())).filter(x => !isNaN(x));
      if (!idList.length) return res.status(400).json({ error: "Invalid ids" });
      const ph = idList.map(() => '?').join(',');
      rows = (await db.execute({ sql: `SELECT * FROM responses WHERE id IN (${ph}) ORDER BY id`, args: idList })).rows;
    } else {
      rows = (await db.execute("SELECT * FROM responses ORDER BY id")).rows;
    }
    if (!rows.length) return res.status(404).json({ error: "No responses" });

    const zip = new JSZip();
    for (const row of rows) {
      const record = { id: row[0], ext_id: row[1], ts: row[2], created: row[3], answers: JSON.parse(row[4]) };
      const buf    = await Packer.toBuffer(buildDoc(record));
      const ts     = (record.ts || record.created || '').slice(0, 10);
      zip.file(`kwestionariusz_${String(record.id).padStart(4,'0')}_${ts}.docx`, buf);
    }

    const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const date   = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="kwestionariusze_${date}.zip"`);
    return res.status(200).send(zipBuf);

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
