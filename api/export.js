// api/export.js
// GET /api/export?id=42          → single .docx download
// GET /api/export                → ZIP of all responses
// GET /api/export?ids=1,2,5      → ZIP of specific ids

import { sql, initSchema } from "./_db.js";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, ShadingType, BorderStyle, VerticalAlign
} from "docx";
import JSZip from "jszip";

export const config = { runtime: "nodejs", maxDuration: 30 };

// ── CHECKBOX CHARS ───────────────────────────────────────────────────────────
const TICK  = "\u2612"; // ☒
const EMPTY = "\u2610"; // ☐

function ck(answers, key, value) {
  const a = answers[key];
  if (Array.isArray(a)) return a.some(v => v.toLowerCase().startsWith(value.toLowerCase().split('.')[0]+'.') || v.toLowerCase() === value.toLowerCase()) ? TICK : EMPTY;
  if (!a) return EMPTY;
  return (a.toLowerCase().startsWith(value.toLowerCase().split('.')[0]+'.') || a.toLowerCase() === value.toLowerCase()) ? TICK : EMPTY;
}
function yn(answers, key) {
  const v = answers[key];
  return { t: v === 'TAK' ? TICK : EMPTY, n: v === 'NIE' ? TICK : EMPTY };
}

// ── STYLE HELPERS ─────────────────────────────────────────────────────────────
const FONT = "Arial";
const SZ   = 18; // 9pt in half-points

function run(text, { bold=false, sz=SZ, color=undefined } = {}) {
  return new TextRun({ text, font: FONT, size: sz, bold, color });
}
function p(children, { align=AlignmentType.LEFT, spaceBefore=20, spaceAfter=20 } = {}) {
  const runs = Array.isArray(children) ? children : [run(children)];
  return new Paragraph({
    children: runs,
    alignment: align,
    spacing: { before: spaceBefore, after: spaceAfter }
  });
}

const BDR = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const BORDERS = { top: BDR, bottom: BDR, left: BDR, right: BDR };
const NO_BDR  = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = { top: NO_BDR, bottom: NO_BDR, left: NO_BDR, right: NO_BDR };

function cell(children, { width=4500, bg=undefined, bold=false } = {}) {
  const paras = (Array.isArray(children) ? children : [children]).map(c =>
    typeof c === 'string'
      ? p(c, { spaceBefore: 20, spaceAfter: 20 })
      : c
  );
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    borders: BORDERS,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    ...(bg ? { shading: { fill: bg, type: ShadingType.CLEAR } } : {})
  });
}

function headerRow(text, cols=2, totalWidth=9000) {
  const merged = new TableCell({
    children: [p([run(text, { bold: true })], { spaceBefore: 20, spaceAfter: 20 })],
    columnSpan: cols,
    width: { size: totalWidth, type: WidthType.DXA },
    borders: BORDERS,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: { fill: "D9D9D9", type: ShadingType.CLEAR }
  });
  return new TableRow({ children: [merged] });
}

function twoCol(leftChildren, rightChildren, lw=4500, rw=4500) {
  return new TableRow({
    children: [
      cell(leftChildren, { width: lw }),
      cell(rightChildren, { width: rw }),
    ]
  });
}

function checkList(answers, key, opts, sz=SZ) {
  return opts.map(opt => {
    const box = ck(answers, key, opt);
    return p([run(`${box} ${opt}`)], { spaceBefore: 10, spaceAfter: 10 });
  });
}

// ── MAIN DOC GENERATOR ───────────────────────────────────────────────────────
function buildDoc(record) {
  const a = record.answers || {};
  const W = 9000; // usable content width in DXA (A4 narrow margins)
  const children = [];

  // ── TITLE ──────────────────────────────────────────────────────────────────
  children.push(p(
    [run("Kwestionariusz osoby w kryzysie bezdomności w ramach Ogólnopolskiego badania liczby osób w kryzysie bezdomności \u2013 rok badania: 2026*", { bold: true, sz: 18 })],
    { align: AlignmentType.CENTER, spaceBefore: 0, spaceAfter: 60 }
  ));

  // ── WSTĘP HEADER ───────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([run("WSTĘP", { bold: true })])], { width: W, bg: "D9D9D9" })] }),
    new TableRow({ children: [cell([p([run("W przypadku stwierdzenia przez ankietera zagrożenia życia lub zdrowia osoby bezdomnej należy niezwłocznie powiadomić odpowiednie służby, w tym policję \u2013 tel. 112 i 997", { bold: true })], { align: AlignmentType.CENTER })], { width: W })] }),
  ]}));

  // ── WYWIAD / ZGODA ─────────────────────────────────────────────────────────
  const { t: wT, n: wN } = yn(a, 'wywiad_dzisiaj');
  const { t: zT, n: zN } = yn(a, 'zgoda_udzial');
  const dupBox = a.duplikat === 'TAK' ? TICK : EMPTY;
  const spP = ck(a, 'sposob_wypelnienia', 'Pełny');
  const spS = ck(a, 'sposob_wypelnienia', 'Skrócony');

  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      [p([run("\u25a0 Czy był przeprowadzony z Panią/Panem taki wywiad dzisiaj?  ", { bold: true }), run(`${wT} Tak   ${wN} Nie`)])],
      [p([run("\u25a0 Czy zgadza się Pan/i na udział w badaniu?  ", { bold: true }), run(`${zT} Tak   ${zN} Nie`)])],
    ),
    twoCol(
      [p([run(`\u25a0 Jeśli tak, zakończyć i zaznaczyć duplikat: ${dupBox}`)])],
      [p([run("\u25a0 Sposób wypełnienia:  ", { bold: true }), run(`${spP} pełny / ${spS} skrócony`)])],
    ),
  ]}));

  // UWAGA box
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([
      run("UWAGA!!! ", { bold: true }),
      run("Pierwszym pytaniem, które należy zadać respondentowi jest pytanie czy w dniu dzisiejszym był badany tym wywiadem. Jeśli dana osoba już uczestniczyła w wywiadzie prosimy nie rozpoczynać wywiadu. Jeśli z osobą bezdomną z pewnych względów jest utrudniony kontakt bądź odmawia wzięcia udziału w badaniu, prosimy o wypełnienie kwestionariusza z zaznaczeniem miejsca przebywania, płci, szacowanego wieku. W przypadku dzieci (0-17 lat) wypełniamy tylko pytania 1-4 oraz miejsce przebywania."),
    ])], { width: W })] }),
  ]}));

  children.push(p(""));

  // ── MIEJSCE ────────────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("MIEJSCE PRZEPROWADZENIA BADANIA / PRZEBYWANIA OSOBY W KRYZYSIE BEZDOMNOŚCI", { bold: true })]), { width: W, bg: "D9D9D9" })] }),
    new TableRow({ children: [cell(p([run("Województwo \u2013 Pomorskie      Powiat \u2013 Gdynia      Gmina \u2013 Gdynia      Miejscowość \u2013 Gdynia", { bold: true })]), { width: W })] }),
  ]}));

  const { t: mT, n: mN } = yn(a, 'miasto_powyzej_100k');
  children.push(p([run("Czy miasto powyżej 100 tysięcy mieszkańców?  ", { bold: true }), run(`${mT} Tak   ${mN} Nie`)], { spaceBefore: 40 }));

  // Places — 2 col
  const miejsca = [
    [1,"Noclegownia"],[2,"Ogrzewalnia"],[3,"Schronisko dla osób bezdomnych"],
    [4,"Schronisko dla osób bezdomnych z usługami opiekuńczymi"],[5,"Mieszkanie wspomagane"],
    [6,"Mieszkanie treningowe"],[7,"Dom dla matek z małoletnimi dziećmi i kobiet w ciąży"],
    [8,"Ośrodek interwencji kryzysowej"],[9,"Specjalistyczny ośrodek wsparcia dla osób doznających przemocy domowej"],
    [10,"Szpital, hospicjum, ZOL, inna placówka zdrowia"],[11,"Zakład karny, areszt śledczy"],
    [12,"Izba wytrzeźwień, pogotowie socjalne"],[13,"Instytucja zdrowia psychicznego/leczenia uzależnień"],
    [14,"Inna placówka/miejsce mieszkalne"],[15,"Pustostan"],
    [16,"Domek na działce, altana działkowa"],
    [17,"Miejsce niemieszkalne: ulica, klatka schodowa, dworzec PKP/PKS, altana śmietnikowa, piwnica, itp."],
  ];
  const stored_miejsce = a.miejsce_pobytu || '';
  const leftM  = miejsca.slice(0, 8);
  const rightM = miejsca.slice(8);
  const maxRows = Math.max(leftM.length, rightM.length);
  const placeRows = [];
  for (let i = 0; i < maxRows; i++) {
    const lItem = leftM[i];
    const rItem = rightM[i];
    const lBox = lItem ? (stored_miejsce.startsWith(`${lItem[0]}.`) ? TICK : EMPTY) : '';
    const rBox = rItem ? (stored_miejsce.startsWith(`${rItem[0]}.`) ? TICK : EMPTY) : '';
    placeRows.push(twoCol(
      lItem ? [p([run(`${lBox} ${lItem[0]}. ${lItem[1]}`)])] : [p("")],
      rItem ? [p([run(`${rBox} ${rItem[0]}. ${rItem[1]}`)])] : [p("")],
    ));
  }
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: placeRows }));

  const miejsce_nazwa = a.miejsce_nazwa || '';
  children.push(p([run("Po zaznaczeniu proszę wpisać nazwę miejsca/opisać je: ", { bold: false }), run(miejsce_nazwa || "_______________________________________________")]));

  children.push(p(""));

  // ── PYTANIA HEADER ─────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("PYTANIA", { bold: true })]), { width: W, bg: "D9D9D9" })] }),
  ]}));

  // ── P1 Płeć + P2 Wiek ──────────────────────────────────────────────────────
  const kBox = ck(a, 'p1_plec', '1.1. Kobieta');
  const mBox = ck(a, 'p1_plec', '1.2. Mężczyzna');
  const wAge  = a.p2_wiek_liczba || '......';
  const wD    = ck(a, 'p2_wiek_typ', 'Wiek deklarowany');
  const wO    = ck(a, 'p2_wiek_typ', 'Wiek oszacowany');
  const wKD   = ck(a, 'p2_wiek_kategoria', 'Osoba dorosła (pow. 18 lat)');
  const wKDz  = ck(a, 'p2_wiek_kategoria', 'Dziecko (0\u201317 lat)');

  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      [p([run("1. Płeć  ", { bold: true }), run(`1.1. kobieta ${kBox}   1.2. mężczyzna ${mBox}`)])],
      [p([run("2. Wiek: ", { bold: true }), run(`${wAge} lat   ${wD} wiek deklarowany  ${wO} wiek oszacowany\n${wKD} osoba dorosła (pow. 18 lat)  ${wKDz} dziecko (0\u201317 lat)`)]),],
    ),
  ]}));

  // ── P3 Obywatelstwo ─────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("3. Obywatelstwo i dane o statusie uchodźcy", { bold: true })]), { width: W })] }),
  ]}));

  const obyw_opts   = ["Polskie","Ukraińskie","Inne z Europy (wyłączając Ukrainę)","Inne z Azji","Inne z Afryki","Inne pozostałe lub brak"];
  const status_opts = ["Uchodźcy","Ochrona tymczasowa","Stały pobyt","Nieuregulowany"];
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      [p([run("3.1. Obywatelstwo", { bold: true })]), ...checkList(a, 'p3_obywatelstwo', obyw_opts)],
      [p([run("3.2. Status cudzoziemca", { bold: true })]), ...checkList(a, 'p3_status_cudzoziemca', status_opts)],
    ),
  ]}));

  // ── P4 P5 ───────────────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      [
        p([run("4. Czy posiada Pan(i) zameldowanie na pobyt stały?", { bold: true })]),
        ...checkList(a, 'p4_zameldowanie', [
          "4.1. Tak, w gminie obecnego pobytu",
          "4.2. Tak, poza gminą obecnego pobytu",
          "4.3. Nie, ostatnie zameldowanie było w gminie obecnego pobytu",
          "4.4. Nie, ostatnie zameldowanie było poza gminą obecnego pobytu",
        ])
      ],
      [
        p([run("5. Jak długo doświadcza Pan/i bezdomności?", { bold: true })]),
        ...checkList(a, 'p5_czas_bezdomnosci', [
          "5.1. Do 3 miesięcy","5.2. Od 3 do 6 miesięcy","5.3. Od 6 do 12 miesięcy",
          "5.4. Od 12 do 24 miesięcy","5.5. Od 2 do 5 lat","5.6. Od 5 do 10 lat",
          "5.7. Od 10 lat do 20 lat","5.8. Powyżej 20 lat",
        ])
      ],
    ),
  ]}));

  // ── P6 P7 P8 ────────────────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/3, W/3, W/3], rows: [
    new TableRow({ children: [
      cell([p([run("6. Stan cywilny", { bold: true })]), ...checkList(a, 'p6_stan_cywilny', ["6.1. kawaler/panna","6.2. żonaty/zamężna","6.3. rozwiedziony/rozwiedziona","6.4. wdowiec/wdowa","6.5. w wolnym związku","6.6. w separacji"])], { width: W/3 }),
      cell([p([run("7. Wykształcenie", { bold: true })]), ...checkList(a, 'p7_wyksztalcenie', ["7.1. niepełne podstawowe","7.2. podstawowe","7.3. gimnazjalne","7.4. zawodowe","7.5. średnie (techniczne też)","7.6. wyższe","7.7. nie wiem"])], { width: W/3 }),
      cell([p([run("8. Z kim obecnie Pani/Pan gospodaruje", { bold: true })]), ...checkList(a, 'p8_gospodarstwo', ["8.1. samodzielnie/samotnie","8.2. partner/partnerka","8.3. kolega/koleżanka/znajomy/znajoma","8.4. małoletnie dzieci (0\u201317 lat)","8.5. dorosłe dzieci/członkowie dalszej rodziny","8.6. zbiorowo/w grupie"])], { width: W/3 }),
    ]})
  ]}));

  // ── P9 Dochody ──────────────────────────────────────────────────────────────
  children.push(p([run("9. Jakie źródła dochodu Pan(i) posiada? (Można zaznaczyć dowolną liczbę odpowiedzi):", { bold: true })], { spaceBefore: 40 }));
  const p9opts = ["9.1. zatrudnienie","9.2. praca na czarno","9.3. praca chroniona/zatrudnienie wspierane","9.4. zbieractwo","9.5. zasiłek z pomocy społecznej","9.6. świadczenia ZUS","9.7. żebractwo","9.8. alimenty","9.9. renta/emerytura","9.10. nie posiadam dochodu","9.11. odmowa odpowiedzi"];
  children.push(twoColCheckTable(a, 'p9_dochody', p9opts, W));

  // ── P10 Przyczyny ────────────────────────────────────────────────────────────
  children.push(p([run("10. Które wydarzenia były według Pana(i) przyczyną bezdomności? (proszę zaznaczyć maksymalnie 3):", { bold: true })], { spaceBefore: 40 }));
  const p10opts = ["10.1. konflikt rodzinny","10.2. odejście/śmierć rodzica/opiekuna w dzieciństwie","10.3. przemoc domowa","10.4. rozpad związku","10.5. zadłużenie","10.6. bezrobocie, brak pracy, utrata pracy","10.7. problemy wynikające z orientacji seksualnej","10.8. zły stan zdrowia, niepełnosprawność","10.9. eksmisja, wymeldowanie z mieszkania","10.10. uzależnienie od alkoholu","10.11. uzależnienie od narkotyków","10.12. uzależnienie od hazardu","10.13. migracja/wyjazd na stałe do innego kraju","10.14. choroba/zaburzenia psychiczne inne niż uzależnienia","10.15. opuszczenie placówki opiekuńczo-wychowawczej","10.16. opuszczenie zakładu karnego","10.17. konflikt z prawem","10.18. inna przemoc niż domowa","10.19. problemy wynikające ze zmiany wiary","10.20. odmowa odpowiedzi"];
  children.push(twoColCheckTable(a, 'p10_przyczyny', p10opts, W));

  // ── P11 P12 ──────────────────────────────────────────────────────────────────
  const p11opts = ["11.1. wsparcie finansowe","11.2. posiłek","11.3. odzież","11.4. schronienie","11.5. terapia uzależnień","11.6. opieka zdrowotna","11.7. nie korzystam"];
  const p12opts = ["12.1. żywnościowe","12.2. higieniczne (w tym dostęp do łaźni)","12.3. zdrowotne","12.4. schronienie","12.5. terapia uzależnień","12.6. wsparcie psychologiczne","12.7. pomoc prawna","12.8. pomoc w znalezieniu pracy","12.9. finansowe","12.10. mieszkaniowe","12.11. wyjście z długów","12.12. nie oczekuję pomocy"];
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      [p([run("11. Czy Pan(i) korzysta z pomocy i w jakiej postaci? (proszę zaznaczyć wszystkie formy):", { bold: true })]), ...checkList(a, 'p11_pomoc', p11opts)],
      [p([run("12. W jakich obszarach oczekuje Pan(i) wsparcia/pomocy? (maksymalnie 3 potrzeby):", { bold: true })]), ...checkList(a, 'p12_oczekiwane_wsparcie', p12opts)],
    ),
  ]}));

  // ── P13 Bezdomność ukryta ────────────────────────────────────────────────────
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("13. Pytania o tzw. bezdomność ukrytą", { bold: true })]), { width: W })] }),
  ]}));

  const { t: t1, n: n1 } = yn(a, 'p13_1_czy_pomieszkuje');
  const { t: t2, n: n2 } = yn(a, 'p13_2_czy_pomieszkiwal');
  const { t: t3, n: n3 } = yn(a, 'p13_3_czy_zna');
  const dOpts = ["Do 3 miesięcy","Od 3 do 12 miesięcy","Od 1 roku do 2 lat","Od 2 do 5 lat","Powyżej 5 lat"];

  const p13Children = [
    p([run("13.1. Czy obecnie Pan(i) pomieszkuje w domu/mieszkaniu u rodziny, znajomych, czy innych osób, tj. nie ma Pan(i) własnego miejsca zamieszkania i przebywa tymczasowo u innych osób?")]),
    p([run(`${t1} Tak    ${n1} Nie`)]),
  ];
  if (a.p13_1_czy_pomieszkuje === 'TAK') {
    p13Children.push(p([run("13.1.2. Jak długo obecnie Pan(i) pomieszkuje...?")]));
    p13Children.push(p(dOpts.map(o => run(`${ck(a,'p13_1_2_jak_dlugo',o)} ${o}   `))));
  }
  p13Children.push(p([run("13.2. Czy w przeszłości Pan(i) tymczasowo pomieszkiwał(a)...?")]));
  p13Children.push(p([run(`${t2} Tak    ${n2} Nie`)]));
  if (a.p13_2_czy_pomieszkiwal === 'TAK') {
    p13Children.push(p([run("13.2. Jak długo tymczasowo Pan(i) pomieszkiwał(a)...?")]));
    p13Children.push(p(dOpts.map(o => run(`${ck(a,'p13_2_jak_dlugo',o)} ${o}   `))));
  }
  p13Children.push(p([run("13.3. Czy zna Pan(i) inne osoby, które w ciągu ostatnich 12 miesięcy tymczasowo pomieszkiwały...?")]));
  p13Children.push(p([run(`${t3} Tak    ${n3} Nie`)]));
  if (a.p13_3_czy_zna === 'TAK') {
    const ileOpts = ["1\u20132","3\u20135","6\u201310","Więcej niż 10"];
    p13Children.push(p([run("jeśli Tak, ile Pan(i) zna takich osób:  "), ...ileOpts.map(o => run(`${ck(a,'p13_3_ile_osob',o)} ${o}   `))]));
  }
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p13Children, { width: W })] }),
  ]}));

  // ── FUNKCJA ANKIETERA ────────────────────────────────────────────────────────
  const fa_l = ["Wolontariusz","Pracownik socjalny","Pracownik placówki dla bezdomnych","Inna"];
  const fa_r = ["Pracownik gminy","Strażnik miejski / policjant","Pracownik do spraw streetworkingu (Streetworker)"];
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("FUNKCJA ANKIETERA", { bold: true })]), { width: W, bg: "D9D9D9" })] }),
  ]}));
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol(
      checkList(a, 'funkcja_ankietera', fa_l),
      checkList(a, 'funkcja_ankietera', fa_r),
    ),
  ]}));

  // ── FOOTER ───────────────────────────────────────────────────────────────────
  children.push(p([run("*Wzór kwestionariusza może ulec zmianie. W takim przypadku zostanie ona zakomunikowana w odpowiednim czasie przed badaniem.", { sz: 16 })], { spaceBefore: 80 }));
  children.push(p([run(`Wygenerowano automatycznie | ID: ${record.id} | ${(record.ts||'').slice(0,10)}`, { sz: 14, color: "999999" })]));

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 720, bottom: 720, left: 850, right: 850 }
        }
      },
      children
    }]
  });
}

// ── 2-col checkbox table helper ───────────────────────────────────────────────
function twoColCheckTable(answers, key, opts, W) {
  const mid   = Math.ceil(opts.length / 2);
  const left  = opts.slice(0, mid);
  const right = opts.slice(mid);
  const n     = Math.max(left.length, right.length);
  const rows  = [];
  for (let i = 0; i < n; i++) {
    rows.push(twoCol(
      left[i]  ? [p([run(`${ck(answers, key, left[i])} ${left[i]}`)])] : [p("")],
      right[i] ? [p([run(`${ck(answers, key, right[i])} ${right[i]}`)])] : [p("")],
    ));
  }
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows });
}

// ── HANDLER ───────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).end();

  try {
    await initSchema();

    const { id, ids } = req.query;

    // ── single docx ───────────────────────────────────────────────────────────
    if (id) {
      const rows = await sql("SELECT * FROM responses WHERE id = ?", [Number(id)]);
      if (!rows.length) return res.status(404).json({ error: "Not found" });
      const r = rows[0];
      const record = { id: r.id, ext_id: r.ext_id, ts: r.ts, created: r.created, answers: JSON.parse(r.answers) };
      const doc = buildDoc(record);
      const buf = await Packer.toBuffer(doc);
      const ts  = (record.ts || record.created || '').slice(0, 10);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="kwestionariusz_${String(record.id).padStart(4,'0')}_${ts}.docx"`);
      return res.status(200).send(Buffer.from(buf));
    }

    // ── ZIP of selected or all ─────────────────────────────────────────────────
    let rows;
    if (ids) {
      const idList = ids.split(',').map(x => Number(x.trim())).filter(x => !isNaN(x));
      if (!idList.length) return res.status(400).json({ error: "Invalid ids" });
      const placeholders = idList.map(() => '?').join(',');
      rows = await sql(`SELECT * FROM responses WHERE id IN (${placeholders}) ORDER BY id`, idList);
    } else {
      rows = await sql("SELECT * FROM responses ORDER BY id");
    }

    if (!rows.length) return res.status(404).json({ error: "No responses" });

    const zip = new JSZip();
    for (const r of rows) {
      const record = { id: r.id, ext_id: r.ext_id, ts: r.ts, created: r.created, answers: JSON.parse(r.answers) };
      const doc = buildDoc(record);
      const buf = await Packer.toBuffer(doc);
      const ts  = (record.ts || record.created || '').slice(0, 10);
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
