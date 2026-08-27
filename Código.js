/**
 * TABLERO DE CONTROL - KPI01: EFECTIVIDAD DE BENEFICIOS OTORGADOS POR CONCURSO DE BECAS
 * PROCESO: Asignación de Beneficios
 * FICHA TÉCNICA: A-FOR-GSBC-0303 (Versión 01 - 27/08/2026)
 * ACADEMIA VONEX
 * 
 * SPREADSHEET ID: 1tQrWJ9jjudMGtC_G7EBCt-KAwDudjOB7g_0fQeIHys8
 */

const SPREADSHEET_ID = "1tQrWJ9jjudMGtC_G7EBCt-KAwDudjOB7g_0fQeIHys8";

const HOJAS = {
  CONFIG: "CONFIG_KPI",
  CONVOCATORIAS: "CONVOCATORIAS_BECAS",
  GANADORES: "GANADORES_OTORGADOS",
  MATRICULAS: "MATRICULAS_VENTAS",
  RESUMEN: "RESUMEN_CONSOLIDADO_KPI"
};

/**
 * Servidor Web App (doGet)
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile("index");
    template.spreadsheetId = SPREADSHEET_ID;
    
    return template.evaluate()
      .setTitle("Tablero de Control KPI01 - Asignación de Beneficios | Vonex")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h3>Error al cargar la aplicación: " + err.message + "</h3>");
  }
}

/**
 * Obtener la referencia al Spreadsheet
 */
function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  } catch (e) {
    console.warn("No se pudo abrir por ID, intentando getActiveSpreadsheet:", e);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Inicializar estructura de hojas y datos provisionales si no existen
 */
function inicializarBaseDatos() {
  const ss = getSpreadsheet();
  if (!ss) {
    return { success: false, message: "No se pudo acceder a la hoja de cálculo. Verifique el SPREADSHEET_ID o los permisos." };
  }

  // 1. HOJA CONFIG_KPI
  let hojaConfig = ss.getSheetByName(HOJAS.CONFIG);
  if (!hojaConfig) {
    hojaConfig = ss.insertSheet(HOJAS.CONFIG);
  }
  hojaConfig.clear();
  const configData = [
    ["PARÁMETRO", "VALOR", "DETALLE / DESCRIPCIÓN"],
    ["CÓDIGO DE FICHA", "A-FOR-GSBC-0303", "Formato oficial de Ficha de KPI"],
    ["VERSIÓN", "01", "Versión vigente"],
    ["FECHA ACTUALIZACIÓN", "27/08/2026", "Fecha de aprobación de ficha"],
    ["PROCESO", "ASIGNACIÓN DE BENEFICIOS", "Proceso de Gestión de Beneficios Vonex"],
    ["FINALIDAD DEL PROCESO", "Gestionar la asignación de beneficios de estudios mediante criterios académicos, socioeconómicos, comerciales e institucionales, promoviendo el acceso, la permanencia y la excelencia educativa de los estudiantes.", "Objetivo misional"],
    ["DUEÑO DEL PROCESO", "Jefa de Ventas", "Lidera la estrategia del proceso"],
    ["KPI CÓDIGO", "KPI01", "Identificador único del indicador"],
    ["KPI NOMBRE", "PORCENTAJE DE EFECTIVIDAD DE LOS BENEFICIOS OTORGADOS POR CONCURSO DE BECAS", "Nombre del indicador"],
    ["FINALIDAD DEL KPI", "Medir el porcentaje de beneficiarios que formalizan su matrícula luego de la publicación de resultados.", "Propósito"],
    ["FÓRMULA", '(N.° de becados matriculados del tipo "X" / N.° total de becas otorgadas del tipo "X") * 100', "Fórmula de cálculo"],
    ["META", "70%", "Meta institucional mínima (>= 70%)"],
    ["META_VALOR_NUM", "0.70", "Valor numérico para cálculos"],
    ["UNIDAD DE MEDIDA", "Porcentaje (%)", "Unidad"],
    ["PERIODICIDAD", "Por convocatoria", "Frecuencia de medición"],
    ["TIPO DE INDICADOR", "Eficacia", "Eficacia de conversión"],
    ["RESPONSABLE DEL KPI", "Director Zonal", "Responsable de asegurar el cumplimiento"],
    ["PROCESADOR DE INFORMACIÓN", "Gestión de Procesos Vonex", "Procesa y calcula KPI, genera alertas"],
    ["EVALUADO", "Directores de sede", "Responsable del llenado de formatos y ejecución en sede"],
    ["RANGO ÓPTIMO", ">= 70%", "La mayoría de los becados otorgados se matricula (Verde)"],
    ["RANGO ACEPTABLE", "50% - 69%", "Existe una conversión moderada de becados a matriculados (Ámbar)"],
    ["RANGO CRÍTICO", "< 50%", "Se evidencia una baja efectividad en la captación de los becados otorgados (Rojo)"],
    ["FUENTE DE INFORMACIÓN", "Relación de ganadores del concurso de becas proporcionada por Soporte y reporte de matrículas del área de Ventas o sistema académico.", "Fuentes cruzadas"]
  ];
  hojaConfig.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
  hojaConfig.getRange(1, 1, 1, configData[0].length).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
  hojaConfig.autoResizeColumns(1, 3);

  // 2. HOJA CONVOCATORIAS_BECAS (Espejo Intranet)
  let hojaConv = ss.getSheetByName(HOJAS.CONVOCATORIAS);
  if (!hojaConv) {
    hojaConv = ss.insertSheet(HOJAS.CONVOCATORIAS);
  }
  hojaConv.clear();
  const convHeaders = [
    ["ID_CONVOCATORIA", "TIPO_BECA", "NOMBRE_CONVOCATORIA", "DESCRIPCION_CICLO", "SEDE", "ZONA", "DIRECTOR_SEDE_EVALUADO", "DIRECTOR_ZONAL", "FECHA_INICIO", "FECHA_FIN", "ESTADO", "META_PORCENTAJE"]
  ];
  const convData = [
    ["CONV-2026-001", "CONCURSO", "CONCURSO DE BECAS UNI - 18/08/2026", "AGOSTO UNI", "Sede Central (Santa Beatriz)", "Lima Centro", "Lic. Carlos Mendoza", "Mg. Roberto Alarcón", "2026-08-19 10:00:00", "2026-08-25 23:59:00", "Activo", 0.70],
    ["CONV-2026-002", "CONCURSO", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "JULIO CEPRE SM", "Sede Los Olivos", "Lima Norte", "Lic. Patricia Valverde", "Mg. Roberto Alarcón", "2026-07-20 18:07:00", "2026-07-24 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-003", "CONCURSO", "CONCURSO DE BECAS CEPRE - 16/07/2026", "JULIO CEPRE SM", "Sede San Juan de Lurigancho", "Lima Este", "Ing. Jorge Huamán", "Mg. Roberto Alarcón", "2026-07-20 18:05:00", "2026-07-24 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-004", "CONCURSO", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "JUNIO UNI", "Sede Central (Santa Beatriz)", "Lima Centro", "Lic. Carlos Mendoza", "Mg. Roberto Alarcón", "2026-06-09 14:00:00", "2026-06-12 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-005", "CONCURSO", "CONCURSO DE BECAS UNI - 09/06/2026", "JUNIO UNI", "Sede San Juan de Miraflores", "Lima Sur", "Lic. Elena Paredes", "Mg. Roberto Alarcón", "2026-06-09 14:00:00", "2026-06-12 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-006", "CONCURSO", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "JUNIO SAN MARCOS", "Sede Los Olivos", "Lima Norte", "Lic. Patricia Valverde", "Mg. Roberto Alarcón", "2026-05-30 16:30:00", "2026-06-03 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-007", "CONCURSO", "CONCURSO DE BECAS UNMSM - 30/05/2026", "JUNIO SAN MARCOS", "Sede Huánuco", "Zona Centro / Provincias", "Mg. Fernando Vilchez", "Mg. Roberto Alarcón", "2026-05-30 16:30:00", "2026-06-03 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-008", "RATIFICACIÓN", "RATIFICACION USM ENERO-2026", "CUOTA 9 SM", "Sede Central (Santa Beatriz)", "Lima Centro", "Lic. Carlos Mendoza", "Mg. Roberto Alarcón", "2026-01-15 08:00:00", "2026-01-22 23:59:00", "Finalizado", 0.70],
    ["CONV-2026-009", "RATIFICACIÓN", "RATIFICACION UNH MAR", "CUOTA 7 UNH", "Sede Huánuco", "Zona Centro / Provincias", "Mg. Fernando Vilchez", "Mg. Roberto Alarcón", "2026-03-01 08:00:00", "2026-03-08 23:59:00", "Finalizado", 0.70]
  ];
  hojaConv.getRange(1, 1, 1, convHeaders[0].length).setValues(convHeaders).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
  hojaConv.getRange(2, 1, convData.length, convData[0].length).setValues(convData);
  hojaConv.autoResizeColumns(1, convHeaders[0].length);

  // 3. HOJA GANADORES_OTORGADOS (Fuente Soporte)
  let hojaGan = ss.getSheetByName(HOJAS.GANADORES);
  if (!hojaGan) {
    hojaGan = ss.insertSheet(HOJAS.GANADORES);
  }
  hojaGan.clear();
  const ganHeaders = [
    ["ID_BENEFICIARIO", "ID_CONVOCATORIA", "NOMBRE_CONVOCATORIA", "DNI", "ESTUDIANTE", "TELEFONO", "CORREO", "NIVEL_BENEFICIO", "DESCUENTO_PORC", "PUNTAJE_OBTENIDO", "PUESTO", "FECHA_PUBLICACION", "SEDE", "DIRECTOR_EVALUADO"]
  ];

  const ganData = [
    // CONV-001 (UNI 18/08)
    ["BEN-001", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345101", "Mateo Alejandro Ramos Quispe", "987654321", "mateo.ramos@gmail.com", "Beca Completa (100%)", "100%", "98.50", 1, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-002", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345102", "Camila Lucía Torres Vargas", "987654322", "camila.torres@gmail.com", "Beca Completa (100%)", "100%", "95.00", 2, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-003", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345103", "Sebastián Andrés Flores Rios", "987654323", "seb.flores@gmail.com", "Semi-Beca (50%)", "50%", "91.20", 3, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-004", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345104", "Valeria Sofía Morales Soto", "987654324", "valeria.morales@gmail.com", "Semi-Beca (50%)", "50%", "89.50", 4, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-005", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345105", "Diego Alonso Castillo Ruiz", "987654325", "diego.castillo@gmail.com", "Cuarto de Beca (25%)", "25%", "84.00", 5, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-006", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345106", "Luciana Belén Chávez Paz", "987654326", "luciana.chavez@gmail.com", "Descuento 15%", "15%", "79.50", 6, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-007", "CONV-2026-001", "CONCURSO DE BECAS UNI - 18/08/2026", "72345107", "Joaquín Gabriel Medina León", "987654327", "joaquin.medina@gmail.com", "Descuento 15%", "15%", "76.00", 7, "2026-08-19", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],

    // CONV-002 (CEPRE 16/07 Desc 15%)
    ["BEN-008", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456201", "Andrea Jimena Rojas Salazar", "986543210", "andrea.rojas@gmail.com", "Descuento 15%", "15%", "88.00", 1, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-009", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456202", "Bruno Emmanuel Navarro Diaz", "986543211", "bruno.navarro@gmail.com", "Descuento 15%", "15%", "85.50", 2, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-010", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456203", "Paola Nicole Gutiérrez Cruz", "986543212", "paola.gutierrez@gmail.com", "Descuento 15%", "15%", "83.00", 3, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-011", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456204", "Rodrigo Manuel Vega Silva", "986543213", "rodrigo.vega@gmail.com", "Descuento 15%", "15%", "80.50", 4, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-012", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456205", "Daniela María Herrera Luna", "986543214", "daniela.herrera@gmail.com", "Descuento 15%", "15%", "78.00", 5, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-013", "CONV-2026-002", "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026", "73456206", "Gonzalo Ignacio Paredes Ortiz", "986543215", "gonzalo.paredes@gmail.com", "Descuento 15%", "15%", "75.50", 6, "2026-07-20", "Sede Los Olivos", "Lic. Patricia Valverde"],

    // CONV-003 (CEPRE 16/07 Concurso General)
    ["BEN-014", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567301", "Renato David Aliaga Campos", "985432101", "renato.aliaga@gmail.com", "Beca Completa (100%)", "100%", "96.00", 1, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],
    ["BEN-015", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567302", "Fernanda Isabel Bravo Peña", "985432102", "fernanda.bravo@gmail.com", "Semi-Beca (50%)", "50%", "92.00", 2, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],
    ["BEN-016", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567303", "Fabricio Leonel Solis Vidal", "985432103", "fabricio.solis@gmail.com", "Semi-Beca (50%)", "50%", "89.00", 3, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],
    ["BEN-017", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567304", "Ariana Celeste Ponce Meza", "985432104", "ariana.ponce@gmail.com", "Cuarto de Beca (25%)", "25%", "85.00", 4, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],
    ["BEN-018", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567305", "Leonardo Fabian Cabrera Vera", "985432105", "leonardo.cabrera@gmail.com", "Cuarto de Beca (25%)", "25%", "81.00", 5, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],
    ["BEN-019", "CONV-2026-003", "CONCURSO DE BECAS CEPRE - 16/07/2026", "74567306", "Mia Valentina Cordero Núñez", "985432106", "mia.cordero@gmail.com", "Descuento 15%", "15%", "77.50", 6, "2026-07-20", "Sede San Juan de Lurigancho", "Ing. Jorge Huamán"],

    // CONV-004 (UNI 09/06 Desc 15%)
    ["BEN-020", "CONV-2026-004", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "75678401", "Alonso Gabriel Rios Pardo", "984321098", "alonso.rios@gmail.com", "Descuento 15%", "15%", "86.00", 1, "2026-06-09", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-021", "CONV-2026-004", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "75678402", "Ximena Sofia Benitez Mora", "984321099", "ximena.benitez@gmail.com", "Descuento 15%", "15%", "83.50", 2, "2026-06-09", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-022", "CONV-2026-004", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "75678403", "Esteban Rodrigo Yañez Lara", "984321100", "esteban.yanez@gmail.com", "Descuento 15%", "15%", "81.00", 3, "2026-06-09", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-023", "CONV-2026-004", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "75678404", "Adriana Patricia Cano Bello", "984321101", "adriana.cano@gmail.com", "Descuento 15%", "15%", "78.50", 4, "2026-06-09", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-024", "CONV-2026-004", "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026", "75678405", "Mauricio José Aguilar Fuentes", "984321102", "mauricio.aguilar@gmail.com", "Descuento 15%", "15%", "76.00", 5, "2026-06-09", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],

    // CONV-005 (UNI 09/06 Concurso General SJM)
    ["BEN-025", "CONV-2026-005", "CONCURSO DE BECAS UNI - 09/06/2026", "76789501", "Mathias Nicolas Roldan Arce", "983210987", "mathias.roldan@gmail.com", "Beca Completa (100%)", "100%", "97.00", 1, "2026-06-09", "Sede San Juan de Miraflores", "Lic. Elena Paredes"],
    ["BEN-026", "CONV-2026-005", "CONCURSO DE BECAS UNI - 09/06/2026", "76789502", "Kiara Antonella Leyva Jara", "983210988", "kiara.leyva@gmail.com", "Semi-Beca (50%)", "50%", "93.00", 2, "2026-06-09", "Sede San Juan de Miraflores", "Lic. Elena Paredes"],
    ["BEN-027", "CONV-2026-005", "CONCURSO DE BECAS UNI - 09/06/2026", "76789503", "Aaron Josue Huapaya Cárdenas", "983210989", "aaron.huapaya@gmail.com", "Cuarto de Beca (25%)", "25%", "87.00", 3, "2026-06-09", "Sede San Juan de Miraflores", "Lic. Elena Paredes"],
    ["BEN-028", "CONV-2026-005", "CONCURSO DE BECAS UNI - 09/06/2026", "76789504", "Natalia Guadalupe Soria Rios", "983210990", "natalia.soria@gmail.com", "Descuento 15%", "15%", "80.00", 4, "2026-06-09", "Sede San Juan de Miraflores", "Lic. Elena Paredes"],

    // CONV-006 (UNMSM 30/05 Desc 15% Los Olivos)
    ["BEN-029", "CONV-2026-006", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "77890601", "Ignacio Daniel Vilchez Mora", "982109876", "ignacio.vilchez@gmail.com", "Descuento 15%", "15%", "89.00", 1, "2026-05-30", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-030", "CONV-2026-006", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "77890602", "Fabiana Milagros Cuenca Gil", "982109877", "fabiana.cuenca@gmail.com", "Descuento 15%", "15%", "85.00", 2, "2026-05-30", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-031", "CONV-2026-006", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "77890603", "Cristian Alexis Godoy Bazan", "982109878", "cristian.godoy@gmail.com", "Descuento 15%", "15%", "82.00", 3, "2026-05-30", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-032", "CONV-2026-006", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "77890604", "Genesis Abigail Polo Romero", "982109879", "genesis.polo@gmail.com", "Descuento 15%", "15%", "78.00", 4, "2026-05-30", "Sede Los Olivos", "Lic. Patricia Valverde"],
    ["BEN-033", "CONV-2026-006", "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026", "77890605", "Sergio Paul Espinoza Ocaña", "982109880", "sergio.espinoza@gmail.com", "Descuento 15%", "15%", "75.00", 5, "2026-05-30", "Sede Los Olivos", "Lic. Patricia Valverde"],

    // CONV-007 (UNMSM 30/05 Huánuco)
    ["BEN-034", "CONV-2026-007", "CONCURSO DE BECAS UNMSM - 30/05/2026", "78901701", "Gael Antonio Alvarado Soto", "981098765", "gael.alvarado@gmail.com", "Beca Completa (100%)", "100%", "95.50", 1, "2026-05-30", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-035", "CONV-2026-007", "CONCURSO DE BECAS UNMSM - 30/05/2026", "78901702", "Lucero Estefani Minaya Rosas", "981098766", "lucero.minaya@gmail.com", "Semi-Beca (50%)", "50%", "91.00", 2, "2026-05-30", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-036", "CONV-2026-007", "CONCURSO DE BECAS UNMSM - 30/05/2026", "78901703", "Emilio Rafael Tello Espinoza", "981098767", "emilio.tello@gmail.com", "Cuarto de Beca (25%)", "25%", "84.50", 3, "2026-05-30", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-037", "CONV-2026-007", "CONCURSO DE BECAS UNMSM - 30/05/2026", "78901704", "Romina Jazmin Mallqui Cruz", "981098768", "romina.mallqui@gmail.com", "Descuento 15%", "15%", "79.00", 4, "2026-05-30", "Sede Huánuco", "Mg. Fernando Vilchez"],

    // CONV-008 (RATIFICACIÓN USM ENERO)
    ["BEN-038", "CONV-2026-008", "RATIFICACION USM ENERO-2026", "79012801", "Guillermo Elias Zevallos Ramos", "980987654", "guillermo.zevallos@gmail.com", "Beca Completa (100%)", "100%", "19.50", 1, "2026-01-15", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-039", "CONV-2026-008", "RATIFICACION USM ENERO-2026", "79012802", "Fiorella Paz Lozano Viteri", "980987655", "fiorella.lozano@gmail.com", "Semi-Beca (50%)", "50%", "18.80", 2, "2026-01-15", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-040", "CONV-2026-008", "RATIFICACION USM ENERO-2026", "79012803", "Victor Manuel Lujan Silva", "980987656", "victor.lujan@gmail.com", "Semi-Beca (50%)", "50%", "18.20", 3, "2026-01-15", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],
    ["BEN-041", "CONV-2026-008", "RATIFICACION USM ENERO-2026", "79012804", "Katia Solange Hurtado Soler", "980987657", "katia.hurtado@gmail.com", "Cuarto de Beca (25%)", "25%", "17.90", 4, "2026-01-15", "Sede Central (Santa Beatriz)", "Lic. Carlos Mendoza"],

    // CONV-009 (RATIFICACIÓN UNH MAR)
    ["BEN-042", "CONV-2026-009", "RATIFICACION UNH MAR", "70123901", "Ruben Dario Huaman Meza", "979876543", "ruben.huaman@gmail.com", "Beca Completa (100%)", "100%", "19.00", 1, "2026-03-01", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-043", "CONV-2026-009", "RATIFICACION UNH MAR", "70123902", "Sheyla Noemi Chavez Leon", "979876544", "sheyla.chavez@gmail.com", "Semi-Beca (50%)", "50%", "18.50", 2, "2026-03-01", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-044", "CONV-2026-009", "RATIFICACION UNH MAR", "70123903", "Oscar Felipe Palacios Rios", "979876545", "oscar.palacios@gmail.com", "Semi-Beca (50%)", "50%", "17.90", 3, "2026-03-01", "Sede Huánuco", "Mg. Fernando Vilchez"],
    ["BEN-045", "CONV-2026-009", "RATIFICACION UNH MAR", "70123904", "Dora Maritza Salgado Diaz", "979876546", "dora.salgado@gmail.com", "Cuarto de Beca (25%)", "25%", "17.40", 4, "2026-03-01", "Sede Huánuco", "Mg. Fernando Vilchez"]
  ];
  hojaGan.getRange(1, 1, 1, ganHeaders[0].length).setValues(ganHeaders).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
  hojaGan.getRange(2, 1, ganData.length, ganData[0].length).setValues(ganData);
  hojaGan.autoResizeColumns(1, ganHeaders[0].length);

  // 4. HOJA MATRICULAS_VENTAS (Fuente Ventas / Sistema Académico)
  let hojaMat = ss.getSheetByName(HOJAS.MATRICULAS);
  if (!hojaMat) {
    hojaMat = ss.insertSheet(HOJAS.MATRICULAS);
  }
  hojaMat.clear();
  const matHeaders = [
    ["ID_MATRICULA", "ID_BENEFICIARIO", "DNI", "ESTUDIANTE", "SEDE", "CICLO_MATRICULADO", "FECHA_MATRICULA", "ESTADO_MATRICULA", "ASESOR_VENTAS", "MONTO_PAGADO", "CODIGO_BOLETA", "OBSERVACIONES"]
  ];

  const matData = [
    // CONV-001 (7 ganadores -> 5 matriculados, 1 en proceso, 1 no concretado)
    ["MAT-2026-001", "BEN-001", "72345101", "Mateo Alejandro Ramos Quispe", "Sede Central (Santa Beatriz)", "Anual UNI 2026", "2026-08-20", "Matriculado", "Asesora Lucía Castro", 50.00, "B001-004521", "Matrícula formalizada Beca 100% (solo pago de carnet y material)"],
    ["MAT-2026-002", "BEN-002", "72345102", "Camila Lucía Torres Vargas", "Sede Central (Santa Beatriz)", "Anual UNI 2026", "2026-08-21", "Matriculado", "Asesora Lucía Castro", 50.00, "B001-004535", "Matrícula formalizada Beca 100%"],
    ["MAT-2026-003", "BEN-003", "72345103", "Sebastián Andrés Flores Rios", "Sede Central (Santa Beatriz)", "Semestral UNI 2026-II", "2026-08-21", "Matriculado", "Asesor Marco Peña", 320.00, "B001-004540", "Matrícula formalizada Semi-Beca 50%"],
    ["MAT-2026-004", "BEN-004", "72345104", "Valeria Sofía Morales Soto", "Sede Central (Santa Beatriz)", "Semestral UNI 2026-II", "2026-08-22", "Matriculado", "Asesor Marco Peña", 320.00, "B001-004555", "Matrícula formalizada Semi-Beca 50%"],
    ["MAT-2026-005", "BEN-005", "72345105", "Diego Alonso Castillo Ruiz", "Sede Central (Santa Beatriz)", "Semestral UNI 2026-II", "2026-08-23", "Matriculado", "Asesora Diana Solano", 480.00, "B001-004570", "Matrícula formalizada Cuarto de Beca 25%"],
    ["MAT-2026-006", "BEN-006", "72345106", "Luciana Belén Chávez Paz", "Sede Central (Santa Beatriz)", "Semestral UNI 2026-II", "", "En Proceso", "Asesora Diana Solano", 0.00, "", "Pendiente de confirmación con apoderado"],
    ["MAT-2026-007", "BEN-007", "72345107", "Joaquín Gabriel Medina León", "Sede Central (Santa Beatriz)", "Semestral UNI 2026-II", "", "No Concretado", "Asesora Diana Solano", 0.00, "", "Desistió por cruce de horarios con colegio"],

    // CONV-002 (6 ganadores -> 3 matriculados)
    ["MAT-2026-008", "BEN-008", "73456201", "Andrea Jimena Rojas Salazar", "Sede Los Olivos", "Repaso CEPRE SM", "2026-07-21", "Matriculado", "Asesora Wendy Quispe", 390.00, "B002-003112", "Matrícula formalizada Desc 15%"],
    ["MAT-2026-009", "BEN-009", "73456202", "Bruno Emmanuel Navarro Diaz", "Sede Los Olivos", "Repaso CEPRE SM", "2026-07-22", "Matriculado", "Asesora Wendy Quispe", 390.00, "B002-003120", "Matrícula formalizada Desc 15%"],
    ["MAT-2026-010", "BEN-010", "73456203", "Paola Nicole Gutiérrez Cruz", "Sede Los Olivos", "Repaso CEPRE SM", "2026-07-23", "Matriculado", "Asesor Pedro Vera", 390.00, "B002-003135", "Matrícula formalizada Desc 15%"],
    ["MAT-2026-011", "BEN-011", "73456204", "Rodrigo Manuel Vega Silva", "Sede Los Olivos", "Repaso CEPRE SM", "", "No Concretado", "Asesor Pedro Vera", 0.00, "", "Optó por modalidad virtual sin descuento"],
    ["MAT-2026-012", "BEN-012", "73456205", "Daniela María Herrera Luna", "Sede Los Olivos", "Repaso CEPRE SM", "", "No Concretado", "Asesora Wendy Quispe", 0.00, "", "No responde llamadas de seguimiento"],
    ["MAT-2026-013", "BEN-013", "73456206", "Gonzalo Ignacio Paredes Ortiz", "Sede Los Olivos", "Repaso CEPRE SM", "", "En Proceso", "Asesor Pedro Vera", 0.00, "", "Prometió regularizar pago este fin de mes"],

    // CONV-003 (6 ganadores -> 5 matriculados)
    ["MAT-2026-014", "BEN-014", "74567301", "Renato David Aliaga Campos", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "2026-07-21", "Matriculado", "Asesora Carla Rios", 50.00, "B003-002100", "Beca 100% Formalizada"],
    ["MAT-2026-015", "BEN-015", "74567302", "Fernanda Isabel Bravo Peña", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "2026-07-21", "Matriculado", "Asesora Carla Rios", 295.00, "B003-002105", "Semi Beca 50% Formalizada"],
    ["MAT-2026-016", "BEN-016", "74567303", "Fabricio Leonel Solis Vidal", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "2026-07-22", "Matriculado", "Asesora Carla Rios", 295.00, "B003-002112", "Semi Beca 50% Formalizada"],
    ["MAT-2026-017", "BEN-017", "74567304", "Ariana Celeste Ponce Meza", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "2026-07-22", "Matriculado", "Asesor Luis Tello", 442.50, "B003-002120", "Cuarto Beca 25% Formalizada"],
    ["MAT-2026-018", "BEN-018", "74567305", "Leonardo Fabian Cabrera Vera", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "2026-07-23", "Matriculado", "Asesor Luis Tello", 442.50, "B003-002131", "Cuarto Beca 25% Formalizada"],
    ["MAT-2026-019", "BEN-019", "74567306", "Mia Valentina Cordero Núñez", "Sede San Juan de Lurigancho", "Ciclo CEPRE San Marcos", "", "No Concretado", "Asesor Luis Tello", 0.00, "", "Problemas económicos familiares"],

    // CONV-004 (5 ganadores -> 2 matriculados)
    ["MAT-2026-020", "BEN-020", "75678401", "Alonso Gabriel Rios Pardo", "Sede Central (Santa Beatriz)", "Intensivo UNI", "2026-06-10", "Matriculado", "Asesor Marco Peña", 380.00, "B001-003890", "Desc 15% Formalizado"],
    ["MAT-2026-021", "BEN-021", "75678402", "Ximena Sofia Benitez Mora", "Sede Central (Santa Beatriz)", "Intensivo UNI", "2026-06-11", "Matriculado", "Asesora Lucía Castro", 380.00, "B001-003898", "Desc 15% Formalizado"],
    ["MAT-2026-022", "BEN-022", "75678403", "Esteban Rodrigo Yañez Lara", "Sede Central (Santa Beatriz)", "Intensivo UNI", "", "No Concretado", "Asesora Lucía Castro", 0.00, "", "Se inscribió en otra academia"],
    ["MAT-2026-023", "BEN-023", "75678404", "Adriana Patricia Cano Bello", "Sede Central (Santa Beatriz)", "Intensivo UNI", "", "No Concretado", "Asesor Marco Peña", 0.00, "", "No aceptó fecha de inicio"],
    ["MAT-2026-024", "BEN-024", "75678405", "Mauricio José Aguilar Fuentes", "Sede Central (Santa Beatriz)", "Intensivo UNI", "", "No Concretado", "Asesora Diana Solano", 0.00, "", "No responde"],

    // CONV-005 (4 ganadores -> 3 matriculados)
    ["MAT-2026-025", "BEN-025", "76789501", "Mathias Nicolas Roldan Arce", "Sede San Juan de Miraflores", "Semestral UNI", "2026-06-10", "Matriculado", "Asesora Rosa Mendez", 50.00, "B004-001502", "Beca 100% Formalizada"],
    ["MAT-2026-026", "BEN-026", "76789502", "Kiara Antonella Leyva Jara", "Sede San Juan de Miraflores", "Semestral UNI", "2026-06-10", "Matriculado", "Asesora Rosa Mendez", 310.00, "B004-001509", "Semi-Beca 50% Formalizada"],
    ["MAT-2026-027", "BEN-027", "76789503", "Aaron Josue Huapaya Cárdenas", "Sede San Juan de Miraflores", "Semestral UNI", "2026-06-11", "Matriculado", "Asesor Victor Paz", 465.00, "B004-001518", "Cuarto Beca 25% Formalizada"],
    ["MAT-2026-028", "BEN-028", "76789504", "Natalia Guadalupe Soria Rios", "Sede San Juan de Miraflores", "Semestral UNI", "", "En Proceso", "Asesor Victor Paz", 0.00, "", "Esperando pago en cuotas"],

    // CONV-006 (5 ganadores -> 3 matriculados)
    ["MAT-2026-029", "BEN-029", "77890601", "Ignacio Daniel Vilchez Mora", "Sede Los Olivos", "Semestral San Marcos", "2026-05-31", "Matriculado", "Asesora Wendy Quispe", 370.00, "B002-002880", "Desc 15% Formalizado"],
    ["MAT-2026-030", "BEN-030", "77890602", "Fabiana Milagros Cuenca Gil", "Sede Los Olivos", "Semestral San Marcos", "2026-06-01", "Matriculado", "Asesor Pedro Vera", 370.00, "B002-002891", "Desc 15% Formalizado"],
    ["MAT-2026-031", "BEN-031", "77890603", "Cristian Alexis Godoy Bazan", "Sede Los Olivos", "Semestral San Marcos", "2026-06-02", "Matriculado", "Asesora Wendy Quispe", 370.00, "B002-002905", "Desc 15% Formalizado"],
    ["MAT-2026-032", "BEN-032", "77890604", "Genesis Abigail Polo Romero", "Sede Los Olivos", "Semestral San Marcos", "", "No Concretado", "Asesor Pedro Vera", 0.00, "", "Desistió"],
    ["MAT-2026-033", "BEN-033", "77890605", "Sergio Paul Espinoza Ocaña", "Sede Los Olivos", "Semestral San Marcos", "", "No Concretado", "Asesor Pedro Vera", 0.00, "", "No contesta"],

    // CONV-007 (4 ganadores -> 3 matriculados)
    ["MAT-2026-034", "BEN-034", "78901701", "Gael Antonio Alvarado Soto", "Sede Huánuco", "Semestral San Marcos", "2026-05-31", "Matriculado", "Asesora Miriam Cruz", 50.00, "B005-001090", "Beca 100% Formalizada"],
    ["MAT-2026-035", "BEN-035", "78901702", "Lucero Estefani Minaya Rosas", "Sede Huánuco", "Semestral San Marcos", "2026-06-01", "Matriculado", "Asesora Miriam Cruz", 280.00, "B005-001099", "Semi-Beca 50% Formalizada"],
    ["MAT-2026-036", "BEN-036", "78901703", "Emilio Rafael Tello Espinoza", "Sede Huánuco", "Semestral San Marcos", "2026-06-02", "Matriculado", "Asesor Jorge Ortiz", 420.00, "B005-001110", "Cuarto Beca 25% Formalizada"],
    ["MAT-2026-037", "BEN-037", "78901704", "Romina Jazmin Mallqui Cruz", "Sede Huánuco", "Semestral San Marcos", "", "No Concretado", "Asesor Jorge Ortiz", 0.00, "", "Prefirió ciclo virtual"],

    // CONV-008 (4 ratificados -> 4 matriculados)
    ["MAT-2026-038", "BEN-038", "79012801", "Guillermo Elias Zevallos Ramos", "Sede Central (Santa Beatriz)", "Continuidad Cuota 9 SM", "2026-01-16", "Matriculado", "Asesora Lucía Castro", 0.00, "B001-002010", "Ratificación Beca 100% aplicada"],
    ["MAT-2026-039", "BEN-039", "79012802", "Fiorella Paz Lozano Viteri", "Sede Central (Santa Beatriz)", "Continuidad Cuota 9 SM", "2026-01-16", "Matriculado", "Asesora Lucía Castro", 250.00, "B001-002015", "Ratificación Semi-Beca 50% aplicada"],
    ["MAT-2026-040", "BEN-040", "79012803", "Victor Manuel Lujan Silva", "Sede Central (Santa Beatriz)", "Continuidad Cuota 9 SM", "2026-01-17", "Matriculado", "Asesor Marco Peña", 250.00, "B001-002022", "Ratificación Semi-Beca 50% aplicada"],
    ["MAT-2026-041", "BEN-041", "79012804", "Katia Solange Hurtado Soler", "Sede Central (Santa Beatriz)", "Continuidad Cuota 9 SM", "2026-01-18", "Matriculado", "Asesora Diana Solano", 375.00, "B001-002030", "Ratificación Cuarto Beca 25% aplicada"],

    // CONV-009 (4 ratificados -> 4 matriculados)
    ["MAT-2026-042", "BEN-042", "70123901", "Ruben Dario Huaman Meza", "Sede Huánuco", "Continuidad Cuota 7 UNH", "2026-03-02", "Matriculado", "Asesora Miriam Cruz", 0.00, "B005-000890", "Ratificación Beca 100% aplicada"],
    ["MAT-2026-043", "BEN-043", "70123902", "Sheyla Noemi Chavez Leon", "Sede Huánuco", "Continuidad Cuota 7 UNH", "2026-03-02", "Matriculado", "Asesora Miriam Cruz", 240.00, "B005-000895", "Ratificación Semi-Beca 50% aplicada"],
    ["MAT-2026-044", "BEN-044", "70123903", "Oscar Felipe Palacios Rios", "Sede Huánuco", "Continuidad Cuota 7 UNH", "2026-03-03", "Matriculado", "Asesor Jorge Ortiz", 240.00, "B005-000902", "Ratificación Semi-Beca 50% aplicada"],
    ["MAT-2026-045", "BEN-045", "70123904", "Dora Maritza Salgado Diaz", "Sede Huánuco", "Continuidad Cuota 7 UNH", "2026-03-04", "Matriculado", "Asesor Jorge Ortiz", 360.00, "B005-000910", "Ratificación Cuarto Beca 25% aplicada"]
  ];
  hojaMat.getRange(1, 1, 1, matHeaders[0].length).setValues(matHeaders).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
  hojaMat.getRange(2, 1, matData.length, matData[0].length).setValues(matData);
  hojaMat.autoResizeColumns(1, matHeaders[0].length);

  // 5. HOJA RESUMEN_CONSOLIDADO_KPI (Cálculo consolidado)
  actualizarHojaResumenConsolidado(ss);

  return {
    success: true,
    message: "Base de datos estructurada e inicializada exitosamente con las 5 hojas y datos de prueba realistas."
  };
}

/**
 * Generar y actualizar la hoja RESUMEN_CONSOLIDADO_KPI
 */
function actualizarHojaResumenConsolidado(ss) {
  if (!ss) ss = getSpreadsheet();
  let hojaRes = ss.getSheetByName(HOJAS.RESUMEN);
  if (!hojaRes) {
    hojaRes = ss.insertSheet(HOJAS.RESUMEN);
  }
  hojaRes.clear();

  const resHeaders = [
    ["ID_CONVOCATORIA", "NOMBRE_CONVOCATORIA", "TIPO_PROCESO", "SEDE", "DIRECTOR_SEDE_EVALUADO", "DIRECTOR_ZONAL", "BECAS_OTORGADAS", "BECADOS_MATRICULADOS", "EN_PROCESO", "NO_CONCRETADOS", "EFECTIVIDAD_%", "META_%", "ESTADO_SEMAFORO", "BRECHA_PUNTOS", "ESTADO_GESTION"]
  ];

  const dashboardData = calcularMetricasDashboard(ss);
  const rows = dashboardData.convocatorias.map(c => [
    c.id,
    c.nombre,
    c.tipoProceso,
    c.sede,
    c.directorSede,
    c.directorZonal,
    c.totalOtorgadas,
    c.totalMatriculados,
    c.totalEnProceso,
    c.totalNoConcretados,
    (c.efectividad * 100).toFixed(1) + "%",
    "70.0%",
    c.semaforo.nombre,
    (c.brecha * 100).toFixed(1) + "%",
    c.efectividad >= 0.70 ? "CUMPLE META" : "REQUIERE ACCIÓN"
  ]);

  hojaRes.getRange(1, 1, 1, resHeaders[0].length).setValues(resHeaders).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
  if (rows.length > 0) {
    hojaRes.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  hojaRes.autoResizeColumns(1, resHeaders[0].length);
}

/**
 * Calcula todas las métricas requeridas para el Dashboard del KPI01
 */
function calcularMetricasDashboard(ss) {
  if (!ss) ss = getSpreadsheet();
  if (!ss) return obtenerDatosFallback();

  try {
    const hojaConv = ss.getSheetByName(HOJAS.CONVOCATORIAS);
    const hojaGan = ss.getSheetByName(HOJAS.GANADORES);
    const hojaMat = ss.getSheetByName(HOJAS.MATRICULAS);

    if (!hojaConv || !hojaGan || !hojaMat) {
      inicializarBaseDatos();
      return calcularMetricasDashboard(ss);
    }

    const dataConv = hojaConv.getDataRange().getValues();
    const dataGan = hojaGan.getDataRange().getValues();
    const dataMat = hojaMat.getDataRange().getValues();

    if (dataConv.length <= 1 || dataGan.length <= 1) {
      return obtenerDatosFallback();
    }

    // Mapeo de matrículas por ID_BENEFICIARIO
    const matriculasMap = {};
    for (let i = 1; i < dataMat.length; i++) {
      const row = dataMat[i];
      const idBen = String(row[1]).trim();
      matriculasMap[idBen] = {
        idMatricula: row[0],
        idBeneficiario: idBen,
        dni: row[2],
        estudiante: row[3],
        sede: row[4],
        ciclo: row[5],
        fechaMatricula: row[6],
        estado: row[7], // "Matriculado", "En Proceso", "No Concretado"
        asesor: row[8],
        monto: row[9],
        boleta: row[10],
        observaciones: row[11]
      };
    }

    // Mapeo de ganadores
    const ganadoresList = [];
    for (let i = 1; i < dataGan.length; i++) {
      const row = dataGan[i];
      const idBen = String(row[0]).trim();
      const mat = matriculasMap[idBen] || {
        estado: "No Registrado",
        fechaMatricula: "",
        asesor: "Sin Asignar",
        monto: 0,
        boleta: "",
        observaciones: "Sin registro en sistema de ventas"
      };

      ganadoresList.push({
        idBeneficiario: idBen,
        idConvocatoria: String(row[1]).trim(),
        nombreConvocatoria: row[2],
        dni: String(row[3]),
        estudiante: row[4],
        telefono: row[5],
        correo: row[6],
        nivelBeneficio: row[7],
        descuentoPorc: row[8],
        puntaje: row[9],
        puesto: row[10],
        fechaPublicacion: row[11],
        sede: row[12],
        directorEvaluado: row[13],
        // Datos cruzados de matrícula
        estadoMatricula: mat.estado,
        fechaMatricula: mat.fechaMatricula,
        asesorVentas: mat.asesor,
        montoPagado: mat.monto,
        codigoBoleta: mat.boleta,
        observaciones: mat.observaciones
      });
    }

    // Mapeo de convocatorias
    const convocatorias = [];
    for (let i = 1; i < dataConv.length; i++) {
      const row = dataConv[i];
      const idConv = String(row[0]).trim();
      const ganadoresConv = ganadoresList.filter(g => g.idConvocatoria === idConv);
      
      const totalOtorgadas = ganadoresConv.length;
      const totalMatriculados = ganadoresConv.filter(g => g.estadoMatricula === "Matriculado").length;
      const totalEnProceso = ganadoresConv.filter(g => g.estadoMatricula === "En Proceso").length;
      const totalNoConcretados = ganadoresConv.filter(g => g.estadoMatricula === "No Concretado" || g.estadoMatricula === "No Registrado").length;
      
      const efectividad = totalOtorgadas > 0 ? (totalMatriculados / totalOtorgadas) : 0;
      const brecha = efectividad - 0.70;

      let semaforo = {
        estado: "CRITICO",
        nombre: "Crítico (< 50%)",
        color: "#ef4444",
        badgeClass: "bg-red-500/20 text-red-700 dark:text-red-300 border-red-500/30",
        icono: "fa-solid fa-triangle-exclamation"
      };

      if (efectividad >= 0.70) {
        semaforo = {
          estado: "OPTIMO",
          nombre: "Óptimo (≥ 70%)",
          color: "#10b981",
          badgeClass: "bg-emerald-500/20 text-emerald-700 dark:text-emerald-300 border-emerald-500/30",
          icono: "fa-solid fa-circle-check"
        };
      } else if (efectividad >= 0.50) {
        semaforo = {
          estado: "ACEPTABLE",
          nombre: "Aceptable (50% - 69%)",
          color: "#f59e0b",
          badgeClass: "bg-amber-500/20 text-amber-700 dark:text-amber-300 border-amber-500/30",
          icono: "fa-solid fa-circle-exclamation"
        };
      }

      convocatorias.push({
        id: idConv,
        tipoProceso: row[1],
        nombre: row[2],
        descripcionCiclo: row[3],
        sede: row[4],
        zona: row[5],
        directorSede: row[6],
        directorZonal: row[7],
        fechaInicio: row[8],
        fechaFin: row[9],
        estadoConvocatoria: row[10],
        metaPorcentaje: Number(row[11]) || 0.70,
        totalOtorgadas: totalOtorgadas,
        totalMatriculados: totalMatriculados,
        totalEnProceso: totalEnProceso,
        totalNoConcretados: totalNoConcretados,
        efectividad: efectividad,
        brecha: brecha,
        semaforo: semaforo
      });
    }

    // Métricas Globales
    const totalGlobalOtorgadas = ganadoresList.length;
    const totalGlobalMatriculados = ganadoresList.filter(g => g.estadoMatricula === "Matriculado").length;
    const totalGlobalEnProceso = ganadoresList.filter(g => g.estadoMatricula === "En Proceso").length;
    const totalGlobalNoConcretados = ganadoresList.filter(g => g.estadoMatricula === "No Concretado" || g.estadoMatricula === "No Registrado").length;
    const efectividadGlobal = totalGlobalOtorgadas > 0 ? (totalGlobalMatriculados / totalGlobalOtorgadas) : 0;
    
    // Concursos únicamente
    const ganadoresConcurso = ganadoresList.filter(g => {
      const c = convocatorias.find(conv => conv.id === g.idConvocatoria);
      return c && c.tipoProceso === "CONCURSO";
    });
    const totalConcursoOtorgadas = ganadoresConcurso.length;
    const totalConcursoMatriculados = ganadoresConcurso.filter(g => g.estadoMatricula === "Matriculado").length;
    const efectividadConcurso = totalConcursoOtorgadas > 0 ? (totalConcursoMatriculados / totalConcursoOtorgadas) : 0;

    // Desglose por Tipo de Beneficio
    const niveles = ["Beca Completa (100%)", "Semi-Beca (50%)", "Cuarto de Beca (25%)", "Descuento 15%"];
    const porTipoBeneficio = niveles.map(nivel => {
      const gNivel = ganadoresList.filter(g => g.nivelBeneficio === nivel);
      const otorgadas = gNivel.length;
      const matriculados = gNivel.filter(g => g.estadoMatricula === "Matriculado").length;
      const efect = otorgadas > 0 ? (matriculados / otorgadas) : 0;
      return {
        tipo: nivel,
        otorgadas: otorgadas,
        matriculados: matriculados,
        efectividad: efect,
        efectividadPorc: (efect * 100).toFixed(1)
      };
    });

    // Desglose por Sede / Director Evaluado
    const sedesSet = [...new Set(convocatorias.map(c => c.sede))];
    const porSede = sedesSet.map(sede => {
      const convsSede = convocatorias.filter(c => c.sede === sede);
      const gSede = ganadoresList.filter(g => g.sede === sede);
      const director = convsSede[0] ? convsSede[0].directorSede : "No asignado";
      const otorgadas = gSede.length;
      const matriculados = gSede.filter(g => g.estadoMatricula === "Matriculado").length;
      const efect = otorgadas > 0 ? (matriculados / otorgadas) : 0;
      return {
        sede: sede,
        director: director,
        totalConvocatorias: convsSede.length,
        otorgadas: otorgadas,
        matriculados: matriculados,
        efectividad: efect,
        efectividadPorc: (efect * 100).toFixed(1),
        cumpleMeta: efect >= 0.70
      };
    });

    return {
      kpiInfo: {
        codigo: "KPI01",
        nombre: "Porcentaje de Efectividad de los Beneficios Otorgados por Concurso de Becas",
        ficha: "A-FOR-GSBC-0303",
        proceso: "Asignación de Beneficios",
        duenoProceso: "Jefa de Ventas",
        responsableKPI: "Director Zonal",
        evaluado: "Directores de sede",
        procesador: "Gestión de Procesos Vonex",
        meta: 0.70,
        metaTexto: "≥ 70%",
        formula: '(N.° de becados matriculados / N.° total de becas otorgadas) * 100'
      },
      metricasGlobales: {
        totalOtorgadas: totalGlobalOtorgadas,
        totalMatriculados: totalGlobalMatriculados,
        totalEnProceso: totalGlobalEnProceso,
        totalNoConcretados: totalGlobalNoConcretados,
        efectividadGlobal: efectividadGlobal,
        efectividadGlobalPorc: (efectividadGlobal * 100).toFixed(1),
        efectividadConcurso: efectividadConcurso,
        efectividadConcursoPorc: (efectividadConcurso * 100).toFixed(1),
        metaPorcentaje: 0.70,
        brechaMeta: (efectividadGlobal - 0.70),
        brechaMetaPorc: ((efectividadGlobal - 0.70) * 100).toFixed(1),
        cumpleMeta: efectividadGlobal >= 0.70,
        estadoSemaforo: efectividadGlobal >= 0.70 ? "OPTIMO" : (efectividadGlobal >= 0.50 ? "ACEPTABLE" : "CRITICO")
      },
      convocatorias: convocatorias,
      porTipoBeneficio: porTipoBeneficio,
      porSede: porSede,
      ganadores: ganadoresList,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error("Error al calcular métricas de Google Sheet:", error);
    return obtenerDatosFallback();
  }
}

/**
 * Endpoint para el cliente Web App
 */
function obtenerDatosTablero() {
  const ss = getSpreadsheet();
  return JSON.stringify(calcularMetricasDashboard(ss));
}

/**
 * Actualizar o registrar estado de matrícula directamente desde el Dashboard
 */
function actualizarEstadoMatricula(idBeneficiario, nuevoEstado, observaciones, asesor, boleta, monto) {
  try {
    const ss = getSpreadsheet();
    if (!ss) return { success: false, message: "No se encontró el Spreadsheet." };

    const hojaMat = ss.getSheetByName(HOJAS.MATRICULAS);
    if (!hojaMat) return { success: false, message: "Hoja de matrículas no encontrada." };

    const dataMat = hojaMat.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < dataMat.length; i++) {
      if (String(dataMat[i][1]).trim() === String(idBeneficiario).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    const fechaActual = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd");

    if (rowIndex !== -1) {
      hojaMat.getRange(rowIndex, 8).setValue(nuevoEstado);
      if (nuevoEstado === "Matriculado" && !dataMat[rowIndex - 1][6]) {
        hojaMat.getRange(rowIndex, 7).setValue(fechaActual);
      }
      if (observaciones) hojaMat.getRange(rowIndex, 12).setValue(observaciones);
      if (asesor) hojaMat.getRange(rowIndex, 9).setValue(asesor);
      if (boleta) hojaMat.getRange(rowIndex, 11).setValue(boleta);
      if (monto !== undefined && monto !== null) hojaMat.getRange(rowIndex, 10).setValue(monto);
    } else {
      const hojaGan = ss.getSheetByName(HOJAS.GANADORES);
      let datosBeneficiario = null;
      if (hojaGan) {
        const dataGan = hojaGan.getDataRange().getValues();
        for (let j = 1; j < dataGan.length; j++) {
          if (String(dataGan[j][0]).trim() === String(idBeneficiario).trim()) {
            datosBeneficiario = dataGan[j];
            break;
          }
        }
      }

      const nuevoIdMat = "MAT-2026-" + String(hojaMat.getLastRow()).padStart(3, '0');
      const dni = datosBeneficiario ? datosBeneficiario[3] : "";
      const estudiante = datosBeneficiario ? datosBeneficiario[4] : "";
      const sede = datosBeneficiario ? datosBeneficiario[12] : "";

      hojaMat.appendRow([
        nuevoIdMat,
        idBeneficiario,
        dni,
        estudiante,
        sede,
        "Ciclo Regular",
        nuevoEstado === "Matriculado" ? fechaActual : "",
        nuevoEstado,
        asesor || "Asesor Asignado",
        Number(monto) || 0,
        boleta || "",
        observaciones || "Actualizado desde el Tablero"
      ]);
    }

    actualizarHojaResumenConsolidado(ss);

    return {
      success: true,
      message: "Estado de matrícula actualizado con éxito para el beneficiario " + idBeneficiario + "."
    };
  } catch (e) {
    return {
      success: false,
      message: "Error al actualizar matrícula: " + e.message
    };
  }
}

/**
 * Datos locales de respaldo en caso de que el Sheet aún no esté conectado
 */
function obtenerDatosFallback() {
  return {
    kpiInfo: {
      codigo: "KPI01",
      nombre: "Porcentaje de Efectividad de los Beneficios Otorgados por Concurso de Becas",
      ficha: "A-FOR-GSBC-0303",
      proceso: "Asignación de Beneficios",
      duenoProceso: "Jefa de Ventas",
      responsableKPI: "Director Zonal",
      evaluado: "Directores de sede",
      procesador: "Gestión de Procesos Vonex",
      meta: 0.70,
      metaTexto: "≥ 70%",
      formula: '(N.° de becados matriculados / N.° total de becas otorgadas) * 100'
    },
    metricasGlobales: {
      totalOtorgadas: 45,
      totalMatriculados: 33,
      totalEnProceso: 4,
      totalNoConcretados: 8,
      efectividadGlobal: 0.733,
      efectividadGlobalPorc: "73.3",
      efectividadConcurso: 0.676,
      efectividadConcursoPorc: "67.6",
      metaPorcentaje: 0.70,
      brechaMeta: 0.033,
      brechaMetaPorc: "+3.3",
      cumpleMeta: true,
      estadoSemaforo: "OPTIMO"
    },
    convocatorias: [
      {
        id: "CONV-2026-001",
        tipoProceso: "CONCURSO",
        nombre: "CONCURSO DE BECAS UNI - 18/08/2026",
        descripcionCiclo: "AGOSTO UNI",
        sede: "Sede Central (Santa Beatriz)",
        zona: "Lima Centro",
        directorSede: "Lic. Carlos Mendoza",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-08-19 10:00:00",
        fechaFin: "2026-08-25 23:59:00",
        estadoConvocatoria: "Activo",
        metaPorcentaje: 0.70,
        totalOtorgadas: 7,
        totalMatriculados: 5,
        totalEnProceso: 1,
        totalNoConcretados: 1,
        efectividad: 0.714,
        brecha: 0.014,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      },
      {
        id: "CONV-2026-002",
        tipoProceso: "CONCURSO",
        nombre: "DESCUENTO 15% CONCURSO DE BECAS CEPRE - 16/07/2026",
        descripcionCiclo: "JULIO CEPRE SM",
        sede: "Sede Los Olivos",
        zona: "Lima Norte",
        directorSede: "Lic. Patricia Valverde",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-07-20 18:07:00",
        fechaFin: "2026-07-24 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 6,
        totalMatriculados: 3,
        totalEnProceso: 1,
        totalNoConcretados: 2,
        efectividad: 0.500,
        brecha: -0.200,
        semaforo: { estado: "ACEPTABLE", nombre: "Aceptable (50% - 69%)", color: "#f59e0b", badgeClass: "bg-amber-500/20 text-amber-700 border-amber-500/30", icono: "fa-solid fa-circle-exclamation" }
      },
      {
        id: "CONV-2026-003",
        tipoProceso: "CONCURSO",
        nombre: "CONCURSO DE BECAS CEPRE - 16/07/2026",
        descripcionCiclo: "JULIO CEPRE SM",
        sede: "Sede San Juan de Lurigancho",
        zona: "Lima Este",
        directorSede: "Ing. Jorge Huamán",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-07-20 18:05:00",
        fechaFin: "2026-07-24 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 6,
        totalMatriculados: 5,
        totalEnProceso: 0,
        totalNoConcretados: 1,
        efectividad: 0.833,
        brecha: 0.133,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      },
      {
        id: "CONV-2026-004",
        tipoProceso: "CONCURSO",
        nombre: "DESCUENTO 15% CONCURSO DE BECAS UNI - 09/06/2026",
        descripcionCiclo: "JUNIO UNI",
        sede: "Sede Central (Santa Beatriz)",
        zona: "Lima Centro",
        directorSede: "Lic. Carlos Mendoza",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-06-09 14:00:00",
        fechaFin: "2026-06-12 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 5,
        totalMatriculados: 2,
        totalEnProceso: 0,
        totalNoConcretados: 3,
        efectividad: 0.400,
        brecha: -0.300,
        semaforo: { estado: "CRITICO", nombre: "Crítico (< 50%)", color: "#ef4444", badgeClass: "bg-red-500/20 text-red-700 border-red-500/30", icono: "fa-solid fa-triangle-exclamation" }
      },
      {
        id: "CONV-2026-005",
        tipoProceso: "CONCURSO",
        nombre: "CONCURSO DE BECAS UNI - 09/06/2026",
        descripcionCiclo: "JUNIO UNI",
        sede: "Sede San Juan de Miraflores",
        zona: "Lima Sur",
        directorSede: "Lic. Elena Paredes",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-06-09 14:00:00",
        fechaFin: "2026-06-12 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 4,
        totalMatriculados: 3,
        totalEnProceso: 1,
        totalNoConcretados: 0,
        efectividad: 0.750,
        brecha: 0.050,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      },
      {
        id: "CONV-2026-006",
        tipoProceso: "CONCURSO",
        nombre: "DESCUENTO 15% CONCURSO DE BECAS UNMSM - 30/05/2026",
        descripcionCiclo: "JUNIO SAN MARCOS",
        sede: "Sede Los Olivos",
        zona: "Lima Norte",
        directorSede: "Lic. Patricia Valverde",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-05-30 16:30:00",
        fechaFin: "2026-06-03 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 5,
        totalMatriculados: 3,
        totalEnProceso: 0,
        totalNoConcretados: 2,
        efectividad: 0.600,
        brecha: -0.100,
        semaforo: { estado: "ACEPTABLE", nombre: "Aceptable (50% - 69%)", color: "#f59e0b", badgeClass: "bg-amber-500/20 text-amber-700 border-amber-500/30", icono: "fa-solid fa-circle-exclamation" }
      },
      {
        id: "CONV-2026-007",
        tipoProceso: "CONCURSO",
        nombre: "CONCURSO DE BECAS UNMSM - 30/05/2026",
        descripcionCiclo: "JUNIO SAN MARCOS",
        sede: "Sede Huánuco",
        zona: "Zona Centro / Provincias",
        directorSede: "Mg. Fernando Vilchez",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-05-30 16:30:00",
        fechaFin: "2026-06-03 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 4,
        totalMatriculados: 3,
        totalEnProceso: 0,
        totalNoConcretados: 1,
        efectividad: 0.750,
        brecha: 0.050,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      },
      {
        id: "CONV-2026-008",
        tipoProceso: "RATIFICACIÓN",
        nombre: "RATIFICACION USM ENERO-2026",
        descripcionCiclo: "CUOTA 9 SM",
        sede: "Sede Central (Santa Beatriz)",
        zona: "Lima Centro",
        directorSede: "Lic. Carlos Mendoza",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-01-15 08:00:00",
        fechaFin: "2026-01-22 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 4,
        totalMatriculados: 4,
        totalEnProceso: 0,
        totalNoConcretados: 0,
        efectividad: 1.000,
        brecha: 0.300,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      },
      {
        id: "CONV-2026-009",
        tipoProceso: "RATIFICACIÓN",
        nombre: "RATIFICACION UNH MAR",
        descripcionCiclo: "CUOTA 7 UNH",
        sede: "Sede Huánuco",
        zona: "Zona Centro / Provincias",
        directorSede: "Mg. Fernando Vilchez",
        directorZonal: "Mg. Roberto Alarcón",
        fechaInicio: "2026-03-01 08:00:00",
        fechaFin: "2026-03-08 23:59:00",
        estadoConvocatoria: "Finalizado",
        metaPorcentaje: 0.70,
        totalOtorgadas: 4,
        totalMatriculados: 4,
        totalEnProceso: 0,
        totalNoConcretados: 0,
        efectividad: 1.000,
        brecha: 0.300,
        semaforo: { estado: "OPTIMO", nombre: "Óptimo (≥ 70%)", color: "#10b981", badgeClass: "bg-emerald-500/20 text-emerald-700 border-emerald-500/30", icono: "fa-solid fa-circle-check" }
      }
    ],
    porTipoBeneficio: [
      { tipo: "Beca Completa (100%)", otorgadas: 9, matriculados: 9, efectividad: 1.0, efectividadPorc: "100.0" },
      { tipo: "Semi-Beca (50%)", otorgadas: 12, matriculados: 10, efectividad: 0.833, efectividadPorc: "83.3" },
      { tipo: "Cuarto de Beca (25%)", otorgadas: 8, matriculados: 6, efectividad: 0.75, efectividadPorc: "75.0" },
      { tipo: "Descuento 15%", otorgadas: 16, matriculados: 8, efectividad: 0.50, efectividadPorc: "50.0" }
    ],
    porSede: [
      { sede: "Sede San Juan de Lurigancho", director: "Ing. Jorge Huamán", totalConvocatorias: 1, otorgadas: 6, matriculados: 5, efectividad: 0.833, efectividadPorc: "83.3", cumpleMeta: true },
      { sede: "Sede Huánuco", director: "Mg. Fernando Vilchez", totalConvocatorias: 2, otorgadas: 8, matriculados: 7, efectividad: 0.875, efectividadPorc: "87.5", cumpleMeta: true },
      { sede: "Sede San Juan de Miraflores", director: "Lic. Elena Paredes", totalConvocatorias: 1, otorgadas: 4, matriculados: 3, efectividad: 0.750, efectividadPorc: "75.0", cumpleMeta: true },
      { sede: "Sede Central (Santa Beatriz)", director: "Lic. Carlos Mendoza", totalConvocatorias: 3, otorgadas: 16, matriculados: 11, efectividad: 0.688, efectividadPorc: "68.8", cumpleMeta: false },
      { sede: "Sede Los Olivos", director: "Lic. Patricia Valverde", totalConvocatorias: 2, otorgadas: 11, matriculados: 6, efectividad: 0.545, efectividadPorc: "54.5", cumpleMeta: false }
    ],
    ganadores: [],
    timestamp: new Date().toISOString()
  };
}
