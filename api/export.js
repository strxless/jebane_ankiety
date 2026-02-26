// api/export.js  — 2-PAGE VERSION
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

const TICK  = "\u2612";
const EMPTY = "\u2610";

function ck(answers, key, value) {
  const a = answers[key];
  const eq = (x, y) => x.toLowerCase().trim() === y.toLowerCase().trim();
  if (Array.isArray(a)) return a.some(v => eq(v, value)) ? TICK : EMPTY;
  if (!a) return EMPTY;
  return eq(a, value) ? TICK : EMPTY;
}

function yn(answers, key) {
  const v = answers[key];
  return { t: v === 'TAK' ? TICK : EMPTY, n: v === 'NIE' ? TICK : EMPTY };
}

// Safely parse answers — accepts already-parsed object or JSON string
function parseAnswers(raw) {
  if (!raw) return {};
  if (typeof raw === 'object') return raw;
  try { return JSON.parse(raw); } catch { return {}; }
}

const FONT = "Arial";
const SZ   = 16; // 8pt — keeps doc to 2 pages

function run(text, { bold=false, sz=SZ, color=undefined } = {}) {
  return new TextRun({ text: String(text ?? ''), font: FONT, size: sz, bold, color });
}
function p(children, { align=AlignmentType.LEFT, spaceBefore=0, spaceAfter=0, line=200 } = {}) {
  const runs = Array.isArray(children) ? children : [run(children)];
  return new Paragraph({ children: runs, alignment: align, spacing: { before: spaceBefore, after: spaceAfter, line, lineRule: "auto" } });
}

const BDR = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const BORDERS = { top: BDR, bottom: BDR, left: BDR, right: BDR };

function cell(children, { width=4500, bg=undefined } = {}) {
  const paras = (Array.isArray(children) ? children : [children]).map(c => typeof c === 'string' ? p(c) : c);
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    borders: BORDERS,
    margins: { top: 30, bottom: 30, left: 80, right: 80 },
    verticalAlign: VerticalAlign.TOP,
    ...(bg ? { shading: { fill: bg, type: ShadingType.CLEAR } } : {})
  });
}

function twoCol(leftChildren, rightChildren, lw=4500, rw=4500) {
  return new TableRow({ children: [cell(leftChildren, { width: lw }), cell(rightChildren, { width: rw })] });
}

function buildDoc(record) {
  const a = parseAnswers(record.answers);
  const W = 9600;
  const children = [];

  // TITLE
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([run("Kwestionariusz osoby w kryzysie bezdomności w ramach Ogólnopolskiego badania liczby osób w kryzysie bezdomności \u2013 rok badania: 2026*", { bold: true, sz: 17 })], { align: AlignmentType.CENTER })], { width: W })] }),
  ]}));

  // WSTĘP
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([run("WSTĘP", { bold: true })])], { width: W, bg: "D9D9D9" })] }),
    new TableRow({ children: [cell([p([run("W przypadku stwierdzenia przez ankietera zagrożenia życia lub zdrowia osoby bezdomnej należy niezwłocznie powiadomić odpowiednie służby, w tym policję \u2013 tel. 112 i 997", { bold: true })], { align: AlignmentType.CENTER })], { width: W })] }),
  ]}));

  // WYWIAD / ZGODA
  const { t: wT, n: wN } = yn(a, 'wywiad_dzisiaj');
  const { t: zT, n: zN } = yn(a, 'zgoda_udzial');
  const dupBox = a.duplikat === 'TAK' ? TICK : EMPTY;
  const spP = ck(a, 'sposob_wypelnienia', 'Pełny');
  const spS = ck(a, 'sposob_wypelnienia', 'Skrócony');
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: [
    twoCol([p([run("\u25a0 Czy był przeprowadzony z Panią/Panem taki wywiad dzisiaj?  ", { bold: true }), run(`${wT} Tak   ${wN} Nie`)])], [p([run("\u25a0 Czy zgadza się Pan/i na udział w badaniu?  ", { bold: true }), run(`${zT} Tak   ${zN} Nie`)])], W/2, W/2),
    twoCol([p([run(`\u25a0 Jeśli tak, zakończyć i zaznaczyć duplikat: ${dupBox}`)])], [p([run("\u25a0 Sposób wypełnienia:  ", { bold: true }), run(`${spP} pełny / ${spS} skrócony`)])], W/2, W/2),
  ]}));

  // UWAGA
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([run("UWAGA!!! ", { bold: true }), run("Pierwszym pytaniem, które należy zadać respondentowi jest pytanie czy w dniu dzisiejszym był badany tym wywiadem. Jeśli dana osoba już uczestniczyła w wywiadzie prosimy nie rozpoczynać wywiadu. Jeśli z osobą bezdomną z pewnych względów jest utrudniony kontakt bądź odmawia wzięcia udziału w badaniu, prosimy o wypełnienie kwestionariusza z zaznaczeniem miejsca przebywania, płci, szacowanego wieku. W przypadku dzieci (0-17 lat) wypełniamy tylko pytania 1-4 oraz miejsce przebywania.")])], { width: W })] }),
  ]}));

  // MIEJSCE HEADER
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("MIEJSCE PRZEPROWADZENIA BADANIA / PRZEBYWANIA OSOBY W KRYZYSIE BEZDOMNOŚCI", { bold: true })]), { width: W, bg: "D9D9D9" })] }),
    new TableRow({ children: [cell(p([run("Województwo \u2013 Pomorskie      Powiat \u2013 Gdynia      Gmina \u2013 Gdynia      Miejscowość \u2013 Gdynia", { bold: true })]), { width: W })] }),
  ]}));

  // MIEJSCA
  const { t: mT, n: mN } = yn(a, 'miasto_powyzej_100k');
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
  const leftM = miejsca.slice(0,8), rightM = miejsca.slice(8);
  const placeRows = [
    new TableRow({ children: [new TableCell({ children: [p([run("Czy miasto powyżej 100 tysięcy mieszkańców?  ", { bold: true }), run(`${mT} Tak   ${mN} Nie`)])], columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } })] })
  ];
  for (let i = 0; i < Math.max(leftM.length, rightM.length); i++) {
    const lI = leftM[i], rI = rightM[i];
    placeRows.push(twoCol(
      lI ? [p([run(`${stored_miejsce.startsWith(`${lI[0]}.`) ? TICK : EMPTY} ${lI[0]}. ${lI[1]}`)])] : [p("")],
      rI ? [p([run(`${stored_miejsce.startsWith(`${rI[0]}.`) ? TICK : EMPTY} ${rI[0]}. ${rI[1]}`)])] : [p("")],
      W/2, W/2
    ));
  }
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2, W/2], rows: placeRows }));
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell([p([run("Po zaznaczeniu proszę wpisać nazwę miejsca/opisać je: "), run(a.miejsce_nazwa || "_______________________________")]), p([run(".................................................................................................................................................")])], { width: W })] }),
  ]}));

  // PYTANIA HEADER
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p([run("PYTANIA", { bold: true })]), { width: W, bg: "D9D9D9" })] }),
  ]}));

  // P1 + P2 — FIX: actually render the age value
  const kBox = ck(a, 'p1_plec', '1.1. Kobieta'), mBox = ck(a, 'p1_plec', '1.2. Mężczyzna');
  // *** FIX: wAge was defined but never inserted into the paragraph runs ***
  const wAge = a.p2_wiek_liczba ? String(a.p2_wiek_liczba) : '......';
  const wD = ck(a, 'p2_wiek_typ', 'Wiek deklarowany'), wO = ck(a, 'p2_wiek_typ', 'Wiek oszacowany');
  const wKD = ck(a, 'p2_wiek_kategoria', 'Osoba dorosła (pow. 18 lat)'), wKDz = ck(a, 'p2_wiek_kategoria', 'Dziecko (0\u201317 lat)');
  const p1w = Math.floor(W/4);
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [p1w, W-p1w], rows: [
    twoCol(
      [p([run("1. Płeć:", { bold: true })]), p([run(`1.1. kobieta ${kBox}`)]), p([run(`1.2. mężczyzna ${mBox}`)])],
      [
        // *** FIX: insert wAge value inline, bold, after the label ***
        p([
          run("2. Wiek: ", { bold: true }),
          run("2.1. Wiek (liczba lat): "),
          run(wAge, { bold: true }),
          run("  (może być szacowany w przypadku utrudnionego kontaktu)"),
        ]),
        p([run(`${wD} wiek deklarowany   ${wO} wiek oszacowany`)]),
        p([run(`${wKD} osoba dorosła (pow. 18 lat)   ${wKDz} dziecko (0\u201317 lat)`)]),
      ],
      p1w, W-p1w
    ),
  ]}));

  // P3 — 3 col
  const lw3 = Math.floor(W/4), mw3 = Math.floor(W/2), rw3 = W-lw3-mw3;
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [lw3, mw3, rw3], rows: [
    new TableRow({ children: [new TableCell({ children: [p([run("3. Obywatelstwo i dane o statusie uchodźcy", { bold: true })])], columnSpan: 3, width: { size: W, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } })] }),
    new TableRow({ children: [
      cell([p([run("3.1. Obywatelstwo", { bold: true })]), p([run(`${ck(a,'p3_obywatelstwo','Polskie')} polskie`)])], { width: lw3 }),
      cell([p([run("Inne \u2013 pozostałe")]), p([run(`${ck(a,'p3_obywatelstwo','Ukraińskie')} ukraińskie`)]), p([run(`${ck(a,'p3_obywatelstwo','Inne z Europy (wyłączając Ukrainę)')} inne z Europy (wyłączając Ukrainę)`)]), p([run(`${ck(a,'p3_obywatelstwo','Inne z Azji')} inne z Azji`)]), p([run(`${ck(a,'p3_obywatelstwo','Inne z Afryki')} inne z Afryki`)]), p([run(`${ck(a,'p3_obywatelstwo','Inne pozostałe lub brak')} inne pozostałe lub brak`)])], { width: mw3 }),
      cell([p([run("3.2. Status cudzoziemca", { bold: true })]), p([run(`${ck(a,'p3_status_cudzoziemca','Uchodźcy')} uchodźczy`)]), p([run(`${ck(a,'p3_status_cudzoziemca','Ochrona tymczasowa')} ochrona tymczasowa`)]), p([run(`${ck(a,'p3_status_cudzoziemca','Stały pobyt')} stały pobyt`)]), p([run(`${ck(a,'p3_status_cudzoziemca','Nieuregulowany')} nieuregulowany`)])], { width: rw3 }),
    ]})
  ]}));

  // P4 + P5 — 4 cols
  const p4w = Math.floor(W/4);
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [p4w,p4w,p4w,p4w], rows: [
    new TableRow({ children: [
      new TableCell({ children: [p([run("4. Czy posiada Pan(i) zameldowanie na pobyt stały?:", { bold: true })])], columnSpan: 2, width: { size: W/2, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } }),
      new TableCell({ children: [p([run("5. Jak długo doświadcza Pan/i bezdomności?", { bold: true })])], columnSpan: 2, width: { size: W/2, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } }),
    ]}),
    new TableRow({ children: [
      cell([p([run(`${ck(a,'p4_zameldowanie','4.1. Tak, w gminie obecnego pobytu')} 4.1. tak, w gminie obecnego pobytu`)]), p([run(`${ck(a,'p4_zameldowanie','4.3. Nie, ostatnie zameldowanie było w gminie obecnego pobytu')} 4.3. nie, ostatnie zameldowanie było w gminie obecnego pobytu`)])], { width: p4w }),
      cell([p([run(`${ck(a,'p4_zameldowanie','4.2. Tak, poza gminą obecnego pobytu')} 4.2. tak, poza gminą obecnego pobytu`)]), p([run(`${ck(a,'p4_zameldowanie','4.4. Nie, ostatnie zameldowanie było poza gminą obecnego pobytu')} 4.4. nie, ostatnie zameldowanie było poza gminą obecnego pobytu`)])], { width: p4w }),
      cell([p([run(`${ck(a,'p5_czas_bezdomnosci','5.1. Do 3 miesięcy')} 5.1. do 3 miesięcy`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.2. Od 3 do 6 miesięcy')} 5.2. od 3 do 6 miesięcy`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.3. Od 6 do 12 miesięcy')} 5.3. od 6 do 12 miesięcy`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.4. Od 12 do 24 miesięcy')} 5.4. od 12 do 24 miesięcy`)])], { width: p4w }),
      cell([p([run(`${ck(a,'p5_czas_bezdomnosci','5.5. Od 2 do 5 lat')} 5.5. od 2 do 5 lat`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.6. Od 5 do 10 lat')} 5.6. od 5 do 10 lat`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.7. Od 10 lat do 20 lat')} 5.7. od 10 lat do 20 lat`)]), p([run(`${ck(a,'p5_czas_bezdomnosci','5.8. Powyżej 20 lat')} 5.8. powyżej 20 lat`)])], { width: p4w }),
    ]})
  ]}));

  // P6 + P7 + P8 — 3 cols
  const w3 = Math.floor(W/3), w3r = W-w3*2;
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [w3,w3,w3r], rows: [
    new TableRow({ children: [
      cell([p([run("6. Stan cywilny", { bold: true })]), p([run(`6.1. kawaler/panna ${ck(a,'p6_stan_cywilny','6.1. kawaler/panna')}`)]), p([run(`6.2. żonaty/zamężna ${ck(a,'p6_stan_cywilny','6.2. żonaty/zamężna')}`)]), p([run(`6.3. rozwiedziony/rozwiedziona ${ck(a,'p6_stan_cywilny','6.3. rozwiedziony/rozwiedziona')}`)]), p([run(`6.4. wdowiec/wdowa ${ck(a,'p6_stan_cywilny','6.4. wdowiec/wdowa')}`)]), p([run(`6.5. w wolnym związku ${ck(a,'p6_stan_cywilny','6.5. w wolnym związku')}`)]), p([run(`6.6. w separacji ${ck(a,'p6_stan_cywilny','6.6. w separacji')}`)])], { width: w3 }),
      cell([p([run("7. Wykształcenie", { bold: true })]), p([run(`7.1. niepełne podstawowe ${ck(a,'p7_wyksztalcenie','7.1. niepełne podstawowe')}`)]), p([run(`7.2. podstawowe ${ck(a,'p7_wyksztalcenie','7.2. podstawowe')}`)]), p([run(`7.3. gimnazjalne ${ck(a,'p7_wyksztalcenie','7.3. gimnazjalne')}`)]), p([run(`7.4. zawodowe ${ck(a,'p7_wyksztalcenie','7.4. zawodowe')}`)]), p([run(`7.5. średnie (techniczne też) ${ck(a,'p7_wyksztalcenie','7.5. średnie (techniczne też)')}`)]), p([run(`7.6. wyższe ${ck(a,'p7_wyksztalcenie','7.6. wyższe')}`)]), p([run(`7.7. nie wiem ${ck(a,'p7_wyksztalcenie','7.7. nie wiem')}`)])], { width: w3 }),
      cell([p([run("8. Z kim obecnie Pani/Pan gospodaruje:", { bold: true })]), p([run(`8.1. samodzielnie/samotnie ${ck(a,'p8_gospodarstwo','8.1. samodzielnie/samotnie')}`)]), p([run(`8.2. partner/partnerka ${ck(a,'p8_gospodarstwo','8.2. partner/partnerka')}`)]), p([run(`8.3. kolega/koleżanka/znajomy/znajoma ${ck(a,'p8_gospodarstwo','8.3. kolega/koleżanka/znajomy/znajoma')}`)]), p([run(`8.4. małoletnie dzieci (0\u201317 lat) ${ck(a,'p8_gospodarstwo','8.4. małoletnie dzieci (0–17 lat)')}`)]), p([run(`8.5. dorosłe dzieci/członkowie dalszej rodziny ${ck(a,'p8_gospodarstwo','8.5. dorosłe dzieci/członkowie dalszej rodziny')}`)]), p([run(`8.6. zbiorowo/w grupie ${ck(a,'p8_gospodarstwo','8.6. zbiorowo/w grupie')}`)])], { width: w3r }),
    ]})
  ]}));

  // P9 — header + 2-col options
  const p9opts = ["9.1. zatrudnienie","9.2. praca \"na czarno\"","9.3. praca chroniona/zatrudnienie wspierane","9.4. zbieractwo","9.5. zasiłek z pomocy społecznej","9.6. świadczenia ZUS","9.7. żebractwo","9.8. alimenty","9.9. renta/emerytura","9.10. nie posiadam dochodu","9.11. odmowa odpowiedzi"];
  const p9mid = Math.ceil(p9opts.length/2);
  const p9rows = [new TableRow({ children: [new TableCell({ children: [p([run("9. Jakie źródła dochodu Pan(i) posiada? (Można zaznaczyć dowolną liczbę odpowiedzi):", { bold: true })])], columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } })] })];
  for (let i = 0; i < p9mid; i++) p9rows.push(twoCol([p([run(`${ck(a,'p9_dochody',p9opts[i])} ${p9opts[i]}`)])], p9opts[i+p9mid] ? [p([run(`${ck(a,'p9_dochody',p9opts[i+p9mid])} ${p9opts[i+p9mid]}`)])] : [p("")], W/2, W/2));
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2,W/2], rows: p9rows }));

  // P10 — header + 2-col options
  const p10opts = ["10.1. konflikt rodzinny","10.2. odejście/śmierć rodzica/opiekuna w dzieciństwie","10.3. przemoc domowa","10.4. rozpad związku","10.5. zadłużenie","10.6. bezrobocie, brak pracy, utrata pracy","10.7. problemy wynikające z orientacji seksualnej","10.8. zły stan zdrowia, niepełnosprawność","10.9. eksmisja, wymeldowanie z mieszkania","10.10. uzależnienie od alkoholu,","10.11. uzależnienie od narkotyków","10.12. uzależnienie od hazardu","10.13. migracja/wyjazd na stałe do innego kraju","10.14. choroba/zaburzenia psychiczne inne niż uzależnienia","10.15. opuszczenie placówki opiekuńczo-wychowawczej","10.16. opuszczenie zakładu karnego","10.17. konflikt z prawem","10.18. inna przemoc niż domowa","10.19. problemy wynikające ze zmiany wiary","10.20. odmowa odpowiedzi"];
  const p10mid = Math.ceil(p10opts.length/2);
  const p10rows = [new TableRow({ children: [new TableCell({ children: [p([run("10. Które wydarzenia były według Pana(i) przyczyną bezdomności? (proszę zaznaczyć maksymalnie 3):", { bold: true })])], columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 } })] })];
  for (let i = 0; i < p10mid; i++) p10rows.push(twoCol([p([run(`${ck(a,'p10_przyczyny',p10opts[i])} ${p10opts[i]}`)])], p10opts[i+p10mid] ? [p([run(`${ck(a,'p10_przyczyny',p10opts[i+p10mid])} ${p10opts[i+p10mid]}`)])] : [p("")], W/2, W/2));
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2,W/2], rows: p10rows }));

  // P11 + P12 — 3 cols
  const p11opts = ["11.1. wsparcie finansowe","11.2. posiłek","11.3. odzież","11.4. schronienie","11.5. terapia uzależnień","11.6. opieka zdrowotna","11.7. nie korzystam"];
  const p12optsL = ["12.1 żywnościowe","12.2. higieniczne (w tym dostęp do łaźni)","12.3. zdrowotne","12.4. schronienie","12.5. terapia uzależnień","12.6. wsparcie psychologiczne"];
  const p12optsR = ["12.7. pomoc prawna","12.8. pomoc w znalezieniu pracy","12.9. finansowe","12.10. mieszkaniowe","12.11. wyjście z długów","12.12. nie oczekuję pomocy"];
  const p11w = Math.floor(W/3), p12lw = Math.floor(W/3), p12rw = W-p11w-p12lw;
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [p11w,p12lw,p12rw], rows: [
    new TableRow({ children: [
      cell([p([run("11. Czy Pan(i) korzysta z pomocy i w jakiej postaci? (proszę zaznaczyć wszystkie formy, z których osoba korzysta):", { bold: true })]), ...p11opts.map(opt => p([run(`${ck(a,'p11_pomoc',opt)} ${opt}`)]))], { width: p11w }),
      cell([p([run("12. W jakich obszarach oczekuje Pan(i) wsparcia/pomocy? (należy zaznaczyć maksymalnie 3 potrzeby)", { bold: true })]), ...p12optsL.map(opt => p([run(`${ck(a,'p12_oczekiwane_wsparcie',opt)} ${opt}`)]))], { width: p12lw }),
      cell([p([run(" ")]), ...p12optsR.map(opt => p([run(`${ck(a,'p12_oczekiwane_wsparcie',opt)} ${opt}`)]))], { width: p12rw }),
    ]})
  ]}));

  // P13
  const { t: t1, n: n1 } = yn(a, 'p13_1_czy_pomieszkuje');
  const { t: t2, n: n2 } = yn(a, 'p13_2_czy_pomieszkiwal');
  const { t: t3, n: n3 } = yn(a, 'p13_3_czy_zna');
  const dOpts = ["Do 3 miesięcy","Od 3 do 12 miesięcy","Od 1 roku do 2 lat","Od 2 do 5 lat","Powyżej 5 lat"];
  const p13Children = [
    p([run("13. Pytania o tzw. bezdomność ukrytą", { bold: true })]),
    p([run("13.1. Czy obecnie Pan(i) pomieszkuje w domu/mieszkaniu u rodziny, znajomych, czy innych osób, tj. nie ma Pan(i) własnego miejsca zamieszkania i przebywa tymczasowo u innych osób?")]),
    p([run(`${t1} Tak,   ${n1} Nie`)]),
    p([run("13.1.2 Jak długo obecnie Pan(i) pomieszkuje w domu/mieszkaniu u rodziny, czy znajomych, innych osób nie mając własnego miejsca zamieszkania? [wypełnia się tylko w przypadku odpowiedzi \u201eTak\u201d na pytanie nr 13.1.]")]),
  ];
  dOpts.forEach((o,i) => p13Children.push(p([run(`13.1.2.${i+1} ${ck(a,'p13_1_2_jak_dlugo',o)} ${o}`)])));
  p13Children.push(p([run("13.2. Czy w przeszłości Pan(i) tymczasowo pomieszkiwał(a) w domu/mieszkaniu u rodziny, znajomych, czy innych osób, nie mając własnego miejsca zamieszkania?")]));
  p13Children.push(p([run(`${t2} Tak,   ${n2} Nie`)]));
  p13Children.push(p([run("13.2. Jak długo tymczasowo Pan(i) pomieszkiwał(a) w domu/mieszkaniu u rodziny, czy znajomych, innych osób, nie mając własnego miejsca zamieszkania? [wypełnia się tylko w przypadku odpowiedzi \u201eTak\u201d na pytanie nr 13.2.]")]));
  dOpts.forEach((o,i) => p13Children.push(p([run(`13.2.${i+1} ${ck(a,'p13_2_jak_dlugo',o)} ${o}`)])));
  p13Children.push(p([run("13.3 Czy zna Pan(i) inne osoby, które w ciągu ostatnich 12 miesięcy tymczasowo pomieszkiwały w domu/mieszkaniu u rodziny, czy znajomych, innych osób, nie mając własnego miejsca zamieszkania?")]));
  p13Children.push(p([run(`${t3} Tak,   ${n3} Nie`)]));
  const ileOpts = ["1\u20132","3\u20135","6\u201310","Więcej niż 10"];
  p13Children.push(p([run("jeśli Tak, ile Pan(i) zna takich osób:  "), ...ileOpts.map(o => run(`${ck(a,'p13_3_ile_osob',o)} ${o}   `))]));
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
    new TableRow({ children: [cell(p13Children, { width: W })] }),
  ]}));

  // FUNKCJA ANKIETERA
  const fa_l = ["Wolontariusz","Pracownik socjalny","Pracownik placówki dla bezdomnych","Inna"];
  const fa_r = ["Pracownik gminy","Strażnik miejski/policjant","Pracownik do spraw streetworkingu (Streetworker)"];
  children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W/2,W/2], rows: [
    new TableRow({ children: [new TableCell({ children: [p([run("FUNKCJA ANKIETERA", { bold: true })])], columnSpan: 2, width: { size: W, type: WidthType.DXA }, borders: BORDERS, margins: { top: 30, bottom: 30, left: 80, right: 80 }, shading: { fill: "D9D9D9", type: ShadingType.CLEAR } })] }),
    twoCol(fa_l.map(opt => p([run(`${ck(a,'funkcja_ankietera',opt)} ${opt}`)])), fa_r.map(opt => p([run(`${ck(a,'funkcja_ankietera',opt)} ${opt}`)])), W/2, W/2)
  ]}));

  // FOOTER
  children.push(p([run("*Wzór kwestionariusza może ulec zmianie. W takim przypadku zostanie ona zakomunikowana w odpowiednim czasie przed badaniem.", { sz: 14 })]));
  children.push(p([run(`Wygenerowano automatycznie | ID: ${record.id} | ${(record.ts||'').slice(0,10)}`, { sz: 12, color: "999999" })]));

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 400, bottom: 400, left: 500, right: 500 } // ~7mm margins
        }
      },
      children
    }]
  });
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).end();

  try {
    await initSchema();
    const { id, ids } = req.query;

    if (id) {
      const rows = await sql("SELECT * FROM responses WHERE id = ?", [Number(id)]);
      if (!rows.length) return res.status(404).json({ error: "Not found" });
      const r = rows[0];
      // *** FIX: answers comes back as a string from Turso — parseAnswers handles both cases ***
      const record = { id: r.id, ext_id: r.ext_id, ts: r.ts, created: r.created, answers: r.answers };
      const doc = buildDoc(record);
      const buf = await Packer.toBuffer(doc);
      const ts = (record.ts || record.created || '').slice(0, 10);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="kwestionariusz_${String(record.id).padStart(4,'0')}_${ts}.docx"`);
      return res.status(200).send(Buffer.from(buf));
    }

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
      // *** FIX: pass answers as raw string; parseAnswers() inside buildDoc handles it ***
      const record = { id: r.id, ext_id: r.ext_id, ts: r.ts, created: r.created, answers: r.answers };
      const doc = buildDoc(record);
      const buf = await Packer.toBuffer(doc);
      const ts = (record.ts || record.created || '').slice(0, 10);
      zip.file(`kwestionariusz_${String(record.id).padStart(4,'0')}_${ts}.docx`, buf);
    }

    const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const date = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="kwestionariusze_${date}.zip"`);
    return res.status(200).send(zipBuf);

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
