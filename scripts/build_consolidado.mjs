import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [inputPath, outputPath] = process.argv.slice(2);
if (!inputPath || !outputPath) {
  throw new Error("Uso: node build_consolidado.mjs <datos.json> <salida.xlsx>");
}

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));
const workbook = Workbook.create();
const summary = workbook.worksheets.add("Resumen");
const courses = workbook.worksheets.add("Cursos");
const indicators = workbook.worksheets.add("Indicadores");
const levels = workbook.worksheets.add("Resumen por nivel");
const application = workbook.worksheets.add("Aplicación");
const methodology = workbook.worksheets.add("Metodología");

const palette = {
  navy: "#0F172A",
  slate: "#334155",
  blue: "#0EA5E9",
  amber: "#F59E0B",
  green: "#10B981",
  purple: "#8B5CF6",
  red: "#EF4444",
  pale: "#F8FAFC",
  grid: "#CBD5E1",
  text: "#0F172A",
  muted: "#64748B",
  white: "#FFFFFF",
};

const areaColors = {
  Lectura: palette.blue,
  Matemática: palette.amber,
  "Ciencias Naturales": palette.green,
  "Historia y Cs. Sociales": palette.purple,
};

const areaOrder = ["Lectura", "Matemática", "Ciencias Naturales", "Historia y Cs. Sociales"];
const levelOrder = { "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, I: 9, II: 10 };
const sortedCourses = [...data.cursos].sort(
  (a, b) =>
    (levelOrder[a.nivel] ?? 99) - (levelOrder[b.nivel] ?? 99) ||
    a.letra.localeCompare(b.letra, "es") ||
    areaOrder.indexOf(a.area) - areaOrder.indexOf(b.area),
);
const comparisonMap = new Map(data.comparaciones.map((row) => [`${row.area}|${row.curso}`, row]));

function styleTitle(sheet, range, title, subtitle) {
  sheet.showGridLines = false;
  sheet.getRange(range).merge();
  const titleCell = sheet.getRange(range.split(":")[0]);
  titleCell.values = [[title]];
  titleCell.format = {
    fill: palette.navy,
    font: { bold: true, color: palette.white, size: 18 },
    verticalAlignment: "center",
  };
  titleCell.format.rowHeight = 34;
  if (subtitle) {
    const [start, end] = range.split(":");
    const startColumn = start.replace(/\d+/g, "");
    const endColumn = end.replace(/\d+/g, "");
    const subtitleRow = Number(start.replace(/\D+/g, "")) + 1;
    const subtitleRange = `${startColumn}${subtitleRow}:${endColumn}${subtitleRow}`;
    sheet.getRange(subtitleRange).merge();
    const subtitleCell = sheet.getRange(`${startColumn}${subtitleRow}`);
    subtitleCell.values = [[subtitle]];
    subtitleCell.format = {
      fill: palette.slate,
      font: { color: "#E2E8F0", size: 10 },
      verticalAlignment: "center",
      wrapText: true,
    };
    subtitleCell.format.rowHeight = 28;
  }
}

function styleHeader(range, fill = palette.slate) {
  range.format = {
    fill,
    font: { bold: true, color: palette.white },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#64748B" },
  };
  range.format.rowHeight = 26;
}

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }
}

function bestWorst(course) {
  const entries = Object.entries(course.indicadores ?? {}).filter(([, value]) => value != null);
  entries.sort((a, b) => b[1] - a[1]);
  return entries.length ? { best: entries[0], worst: entries.at(-1) } : { best: null, worst: null };
}

// Hoja Cursos: fuente normalizada para los cálculos del libro.
const courseHeaders = [
  "Área",
  "Ciclo",
  "Nivel",
  "Curso",
  "Estudiantes",
  "Diagnóstico 2026",
  "Intermedio 2026",
  "Variación",
  "Logro",
  "Indicador mayor",
  "% mayor",
  "Indicador menor",
  "% menor",
  "Fuente",
];
courses.getRange(`A1:N${sortedCourses.length + 1}`).values = [
  courseHeaders,
  ...sortedCourses.map((course) => {
    const comparison = comparisonMap.get(`${course.area}|${course.curso}`);
    const { best, worst } = bestWorst(course);
    return [
      course.area,
      course.ciclo,
      course.nivel_label,
      course.curso,
      course.estudiantes,
      comparison?.diagnostico ?? null,
      course.promedio,
      null,
      null,
      best?.[0] ?? null,
      best?.[1] ?? null,
      worst?.[0] ?? null,
      worst?.[1] ?? null,
      course.fuente,
    ];
  }),
];
courses.getRange("H2").formulas = [["=IF(OR(F2=\"\",G2=\"\"),\"\",G2-F2)"]];
courses.getRange(`H2:H${sortedCourses.length + 1}`).fillDown();
courses.getRange("I2").formulas = [["=IF(G2<50,\"NL\",IF(G2<66.7,\"PL\",\"L\"))"]];
courses.getRange(`I2:I${sortedCourses.length + 1}`).fillDown();
styleHeader(courses.getRange("A1:N1"));
courses.freezePanes.freezeRows(1);
courses.showGridLines = false;
courses.getRange(`E2:E${sortedCourses.length + 1}`).format.numberFormat = "#,##0";
courses.getRange(`F2:H${sortedCourses.length + 1}`).format.numberFormat = "0.0\"%\"";
courses.getRange(`K2:K${sortedCourses.length + 1}`).format.numberFormat = "0.0\"%\"";
courses.getRange(`M2:M${sortedCourses.length + 1}`).format.numberFormat = "0.0\"%\"";
courses.getRange(`G2:G${sortedCourses.length + 1}`).conditionalFormats.add("colorScale", {
  colors: ["#FEE2E2", "#FEF3C7", "#D1FAE5"],
  thresholds: ["min", "50%", "max"],
});
courses.getRange(`H2:H${sortedCourses.length + 1}`).conditionalFormats.add("cellIs", {
  operator: "greaterThan",
  formula: 0,
  format: { fill: "#DCFCE7", font: { color: "#166534", bold: true } },
});
courses.getRange(`H2:H${sortedCourses.length + 1}`).conditionalFormats.add("cellIs", {
  operator: "lessThan",
  formula: 0,
  format: { fill: "#FEE2E2", font: { color: "#991B1B", bold: true } },
});
const courseTable = courses.tables.add(`A1:N${sortedCourses.length + 1}`, true, "CursosIntermedio");
courseTable.style = "TableStyleMedium2";
setWidths(courses, { A: 23, B: 14, C: 15, D: 10, E: 12, F: 15, G: 15, H: 12, I: 9, J: 28, K: 10, L: 28, M: 10, N: 48 });

// Hoja Indicadores: detalle resumido sin nombres de estudiantes.
const indicatorRows = [];
for (const course of sortedCourses) {
  for (const [indicator, average] of Object.entries(course.indicadores ?? {})) {
    indicatorRows.push([
      course.area,
      course.ciclo,
      course.nivel_label,
      course.curso,
      indicator,
      average,
      null,
      course.fuente,
    ]);
  }
}
indicators.getRange(`A1:H${indicatorRows.length + 1}`).values = [
  ["Área", "Ciclo", "Nivel", "Curso", "Indicador / eje", "Promedio", "Logro", "Fuente"],
  ...indicatorRows,
];
indicators.getRange("G2").formulas = [["=IF(F2<50,\"NL\",IF(F2<66.7,\"PL\",\"L\"))"]];
indicators.getRange(`G2:G${indicatorRows.length + 1}`).fillDown();
styleHeader(indicators.getRange("A1:H1"));
indicators.freezePanes.freezeRows(1);
indicators.showGridLines = false;
indicators.getRange(`F2:F${indicatorRows.length + 1}`).format.numberFormat = "0.0\"%\"";
indicators.getRange(`F2:F${indicatorRows.length + 1}`).conditionalFormats.add("colorScale", {
  colors: ["#FEE2E2", "#FEF3C7", "#D1FAE5"],
  thresholds: ["min", "50%", "max"],
});
const indicatorTable = indicators.tables.add(`A1:H${indicatorRows.length + 1}`, true, "IndicadoresIntermedio");
indicatorTable.style = "TableStyleMedium4";
setWidths(indicators, { A: 23, B: 14, C: 15, D: 10, E: 38, F: 12, G: 9, H: 48 });

// Hoja Resumen por nivel.
const levelPairs = [];
for (const area of areaOrder) {
  const areaLevels = [...new Map(
    sortedCourses.filter((row) => row.area === area).map((row) => [row.nivel_label, row.ciclo]),
  ).entries()];
  areaLevels.sort((a, b) => {
    const findLevel = (label) => Object.entries(levelOrder).find(([key]) => label.startsWith(`${key}°`))?.[1] ?? 99;
    return findLevel(a[0]) - findLevel(b[0]);
  });
  for (const [level, cycle] of areaLevels) levelPairs.push([area, cycle, level]);
}
levels.getRange(`A1:G${levelPairs.length + 1}`).values = [
  ["Área", "Ciclo", "Nivel", "Cursos", "Rendiciones", "Promedio", "Logro"],
  ...levelPairs.map(([area, cycle, level]) => [area, cycle, level, null, null, null, null]),
];
const lastCourseRow = sortedCourses.length + 1;
for (let row = 2; row <= levelPairs.length + 1; row += 1) {
  levels.getRange(`D${row}`).formulas = [[`=COUNTIFS('Cursos'!$A$2:$A$${lastCourseRow},A${row},'Cursos'!$C$2:$C$${lastCourseRow},C${row})`]];
  levels.getRange(`E${row}`).formulas = [[`=SUMIFS('Cursos'!$E$2:$E$${lastCourseRow},'Cursos'!$A$2:$A$${lastCourseRow},A${row},'Cursos'!$C$2:$C$${lastCourseRow},C${row})`]];
  levels.getRange(`F${row}`).formulas = [[`=IF(D${row}=0,\"\",SUMIFS('Cursos'!$G$2:$G$${lastCourseRow},'Cursos'!$A$2:$A$${lastCourseRow},A${row},'Cursos'!$C$2:$C$${lastCourseRow},C${row})/D${row})`]];
  levels.getRange(`G${row}`).formulas = [[`=IF(F${row}<50,\"NL\",IF(F${row}<66.7,\"PL\",\"L\"))`]];
}
styleHeader(levels.getRange("A1:G1"));
levels.freezePanes.freezeRows(1);
levels.showGridLines = false;
levels.getRange(`D2:E${levelPairs.length + 1}`).format.numberFormat = "#,##0";
levels.getRange(`F2:F${levelPairs.length + 1}`).format.numberFormat = "0.0\"%\"";
levels.getRange(`F2:F${levelPairs.length + 1}`).conditionalFormats.add("colorScale", {
  colors: ["#FEE2E2", "#FEF3C7", "#D1FAE5"],
  thresholds: ["min", "50%", "max"],
});
const levelTable = levels.tables.add(`A1:G${levelPairs.length + 1}`, true, "ResumenNivel");
levelTable.style = "TableStyleMedium9";
setWidths(levels, { A: 23, B: 14, C: 15, D: 10, E: 13, F: 12, G: 9 });

// Hoja Aplicación: reporte de avance de la plataforma.
const appRows = data.aplicacion.map((row) => [
  row.ciclo,
  row.nivel_label,
  row.curso,
  row.area,
  row.matricula,
  row.asignados,
  row.rindieron,
  null,
  row.informe_generado,
  row.accion_pendiente,
]);
application.getRange(`A1:J${appRows.length + 1}`).values = [
  ["Ciclo", "Nivel", "Curso", "Área", "Matrícula", "Asignados", "Rindieron", "Cobertura", "Informe generado", "Acción pendiente"],
  ...appRows,
];
application.getRange("H2").formulas = [["=IF(F2=0,\"\",G2/F2)"]];
application.getRange(`H2:H${appRows.length + 1}`).fillDown();
styleHeader(application.getRange("A1:J1"));
application.freezePanes.freezeRows(1);
application.showGridLines = false;
application.getRange(`E2:G${appRows.length + 1}`).format.numberFormat = "#,##0";
application.getRange(`H2:H${appRows.length + 1}`).format.numberFormat = "0.0%";
application.getRange(`H2:H${appRows.length + 1}`).conditionalFormats.add("dataBar", {
  color: palette.blue,
  thresholds: ["min", "max"],
});
const appTable = application.tables.add(`A1:J${appRows.length + 1}`, true, "AvanceAplicacion");
appTable.style = "TableStyleMedium2";
setWidths(application, { A: 14, B: 15, C: 10, D: 42, E: 11, F: 11, G: 11, H: 12, I: 16, J: 52 });

// Hoja Resumen ejecutivo.
styleTitle(
  summary,
  "A1:H1",
  "DIA Monitoreo Intermedio 2026 — Resumen consolidado",
  "Complejo Educacional Maipú · RBD 9959 · Resultados descargados el 28-08-2026 · Sin datos personales",
);
summary.getRange("A4:B4").merge();
summary.getRange("C4:D4").merge();
summary.getRange("E4:F4").merge();
summary.getRange("G4:H4").merge();
summary.getRange("A4").values = [["Rendiciones curso-área"]];
summary.getRange("C4").values = [["Registros curso-área"]];
summary.getRange("E4").values = [["Promedio global"]];
summary.getRange("G4").values = [["Informes generados"]];
for (const cell of ["A4", "C4", "E4", "G4"]) {
  summary.getRange(cell).format = { fill: palette.slate, font: { bold: true, color: palette.white }, verticalAlignment: "center" };
}
summary.getRange("A5:B6").merge();
summary.getRange("C5:D6").merge();
summary.getRange("E5:F6").merge();
summary.getRange("G5:H6").merge();
summary.getRange("A5").formulas = [[`=SUM('Cursos'!E2:E${lastCourseRow})`]];
summary.getRange("C5").formulas = [[`=COUNTA('Cursos'!A2:A${lastCourseRow})`]];
summary.getRange("E5").formulas = [[`=AVERAGE('Cursos'!G2:G${lastCourseRow})`]];
summary.getRange("G5").formulas = [[`=COUNTIF('Aplicación'!I2:I${appRows.length + 1},\"Sí\")`]];
summary.getRange("A5:H6").format = {
  fill: "#E2E8F0",
  font: { bold: true, color: palette.text, size: 20 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: palette.grid },
};
summary.getRange("A5:D6").format.numberFormat = "#,##0";
summary.getRange("E5:F6").format.numberFormat = "0.0\"%\"";
summary.getRange("G5:H6").format.numberFormat = "#,##0";

summary.getRange("A9:E13").values = [
  ["Área", "Cursos", "Rendiciones", "Promedio", "Logro"],
  ...areaOrder.map((area) => [area, null, null, null, null]),
];
for (let row = 10; row <= 13; row += 1) {
  summary.getRange(`B${row}`).formulas = [[`=COUNTIF('Cursos'!$A$2:$A$${lastCourseRow},A${row})`]];
  summary.getRange(`C${row}`).formulas = [[`=SUMIF('Cursos'!$A$2:$A$${lastCourseRow},A${row},'Cursos'!$E$2:$E$${lastCourseRow})`]];
  summary.getRange(`D${row}`).formulas = [[`=IF(B${row}=0,\"\",SUMIF('Cursos'!$A$2:$A$${lastCourseRow},A${row},'Cursos'!$G$2:$G$${lastCourseRow})/B${row})`]];
  summary.getRange(`E${row}`).formulas = [[`=IF(D${row}<50,\"NL\",IF(D${row}<66.7,\"PL\",\"L\"))`]];
  summary.getRange(`A${row}`).format.font = { bold: true, color: areaColors[areaOrder[row - 10]] };
}
styleHeader(summary.getRange("A9:E9"));
summary.getRange("B10:C13").format.numberFormat = "#,##0";
summary.getRange("D10:D13").format.numberFormat = "0.0\"%\"";
summary.getRange("D10:D13").conditionalFormats.add("colorScale", {
  colors: ["#FEE2E2", "#FEF3C7", "#D1FAE5"],
  thresholds: ["min", "50%", "max"],
});
summary.getRange("A9:E13").format.borders = { preset: "inside", style: "thin", color: palette.grid };

summary.getRange("J26:K30").values = [["Área", "Promedio"], ...areaOrder.map((area) => [area, null])];
for (let row = 27; row <= 30; row += 1) {
  summary.getRange(`K${row}`).formulas = [[`=D${row - 17}`]];
}
summary.getRange("K27:K30").format.numberFormat = "0.0\"%\"";

summary.getRange("A16:H19").merge();
summary.getRange("A16").values = [[
  "Lectura y Matemática se comparan con el diagnóstico 2026 solo cuando existe el mismo curso en ambos períodos. Ciencias e Historia se incorporan desde el Intermedio. Los porcentajes por curso corresponden al promedio simple de los ejes/indicadores reportados; en 1° Básico Lectura se usa L=100 y NL=0.",
]];
summary.getRange("A16").format = {
  fill: "#F1F5F9",
  font: { color: palette.muted, italic: true, size: 10 },
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "outside", style: "thin", color: palette.grid },
};
summary.getRange("A16:H19").format.rowHeight = 22;
summary.showGridLines = false;
summary.freezePanes.freezeRows(2);
setWidths(summary, { A: 24, B: 13, C: 14, D: 13, E: 13, F: 13, G: 13, H: 13 });

const summaryChart = summary.charts.add("bar", summary.getRange("J26:K30"));
summaryChart.title = "Resultados DIA Intermedio 2026 por área (%)";
summaryChart.hasLegend = true;
summaryChart.yAxis = { numberFormatCode: "0\"%\"", min: 0, max: 100 };
summaryChart.setPosition("G9", "N24");

// Hoja Metodología y control.
styleTitle(
  methodology,
  "A1:F1",
  "Metodología, fuentes y controles",
  "Documento de trazabilidad del consolidado DIA Monitoreo Intermedio 2026",
);
const methodRows = [
  ["Elemento", "Detalle"],
  ["Establecimiento", data.metadata.establecimiento],
  ["RBD", data.metadata.rbd],
  ["Período", data.metadata.periodo],
  ["Fecha de las fuentes", data.metadata.fecha_fuentes],
  ["Archivos XLSX revisados", data.metadata.archivos_xlsx],
  ["Registros curso-área", data.metadata.registros_curso_area],
  ["Cursos únicos", data.metadata.cursos_unicos],
  ["Método de cálculo", data.metadata.metodologia],
  ["Privacidad", "El consolidado contiene resultados agregados por curso y no incluye nombres de estudiantes."],
  ["Diagnóstico comparable", "Se utiliza la base del dashboard diagnóstico 2026; la variación se calcula solo para cursos coincidentes de Lectura y Matemática."],
  ["Clasificación", "NL: menos de 50%; PL: 50% a menos de 66,7%; L: 66,7% o más."],
  ["Reporte de aplicación", data.metadata.archivo_monitoreo],
];
methodology.getRange(`A4:B${methodRows.length + 3}`).values = methodRows;
styleHeader(methodology.getRange("A4:B4"));
methodology.getRange(`A5:A${methodRows.length + 3}`).format.font = { bold: true, color: palette.slate };
methodology.getRange(`A4:B${methodRows.length + 3}`).format.wrapText = true;

const duplicateStart = methodRows.length + 6;
methodology.getRange(`A${duplicateStart}:F${duplicateStart}`).values = [["Área", "Curso", "Fuente primaria", "Fuente duplicada", "Coincide N", "Coincide promedio"]];
styleHeader(methodology.getRange(`A${duplicateStart}:F${duplicateStart}`), palette.green);
methodology.getRange(`A${duplicateStart + 1}:F${duplicateStart + data.control_duplicados.length}`).values = data.control_duplicados.map((row) => [
  row.area,
  row.curso,
  row.fuente_primaria,
  row.fuente_duplicada,
  row.coincide_n,
  row.coincide_promedio,
]);
methodology.showGridLines = false;
methodology.freezePanes.freezeRows(2);
setWidths(methodology, { A: 25, B: 95, C: 48, D: 56, E: 13, F: 19 });

// Verificación compacta y exportación.
const checks = [];
checks.push((await workbook.inspect({ kind: "table", range: "Resumen!A1:H19", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 8 })).ndjson);
checks.push((await workbook.inspect({ kind: "table", range: `Cursos!A1:N${Math.min(lastCourseRow, 12)}`, include: "values,formulas", tableMaxRows: 12, tableMaxCols: 14 })).ndjson);
checks.push((await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" })).ndjson);

const previewDir = path.join(path.dirname(outputPath), "previews");
await fs.mkdir(previewDir, { recursive: true });
for (const [sheetName, range] of [
  ["Resumen", "A1:N24"],
  ["Cursos", "A1:N18"],
  ["Indicadores", "A1:H18"],
  ["Resumen por nivel", "A1:G22"],
  ["Aplicación", "A1:J18"],
  ["Metodología", `A1:F${duplicateStart + data.control_duplicados.length}`],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(previewDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(checks.filter(Boolean).join("\n"));
console.log(JSON.stringify({ outputPath, sheets: 6, courseRows: sortedCourses.length, indicatorRows: indicatorRows.length, applicationRows: appRows.length }));
