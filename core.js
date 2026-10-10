/* AnjA Base de Datos (celular) — motor: lectura/escritura del Excel y reconocimiento del cuestionario.
 *
 * Es una traducción fiel del programa de la PC (AnjA_Base_de_Datos.py): mismas hojas, mismos
 * encabezados, mismo código de usuario y mismo reconocimiento del texto pegado, para que la PC
 * y el celular puedan trabajar sobre el MISMO archivo Excel.
 *
 * Uso en el navegador:  const core = crearNucleo(window.PREGUNTAS, window.ExcelJS, window.JSZip);
 * Uso en Node (pruebas): const core = require('./core.js')(require('./preguntas.json'), require('exceljs'), require('jszip'));
 */
(function (raiz, fabrica) {
  if (typeof module === 'object' && module.exports) module.exports = fabrica;
  else raiz.crearNucleo = fabrica;
})(typeof self !== 'undefined' ? self : this, function crearNucleo(PREG, ExcelJS, JSZip) {
  'use strict';

  const SECCIONES = PREG.secciones;
  const FAMILIA_CAMPOS = PREG.familia;      // [[clave, etiqueta], ...]
  const SITUACION_CAMPOS = PREG.situacion;
  const NOMBRE_EXCEL = 'AnjA Base de Datos.xlsx';
  const HOJA_USUARIOS = 'Usuarios', HOJA_FAMILIA = 'Familia', HOJA_SITUACIONES = 'Situaciones';

  // ====================================================================
  //  Utilidades de texto y fechas
  // ====================================================================
  const PARTICULAS = new Set(['de', 'del', 'la', 'las', 'los', 'y', 'e', 'van', 'von', 'da', 'di']);
  const LIGADURAS = { 'ﬀ': 'ff', 'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬃ': 'ffi', 'ﬄ': 'ffl', 'ﬅ': 'ft', 'ﬆ': 'st' };
  const quitarLigaduras = (s) => String(s).replace(/[ﬀ-ﬆ]/g, (c) => LIGADURAS[c] || c);
  const sinAcentos = (s) => s.normalize('NFD').replace(/\p{Mn}/gu, '');
  const normalizar = (t) => sinAcentos(quitarLigaduras(t)).toLowerCase().trim();
  const aplanar = (campos) => campos.flatMap((c) => (c.tipo === 'fila' ? c.campos : [c]));
  const dos = (n) => String(n).padStart(2, '0');

  /** '21/06/1990' -> {anio, mes, dia}. Vacío -> null. Formato o fecha inválidos -> lanza Error. */
  function parsearFecha(texto) {
    const t = (texto || '').trim();
    if (!t) return null;
    const m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/.exec(t);
    if (!m) throw new Error('formato');
    const dia = +m[1], mes = +m[2], anio = +m[3];
    const d = new Date(0);
    d.setUTCFullYear(anio, mes - 1, dia);
    if (anio < 1 || d.getUTCFullYear() !== anio || d.getUTCMonth() !== mes - 1 || d.getUTCDate() !== dia) throw new Error('fecha');
    return { anio, mes, dia };
  }
  const fechaUTC = (f) => { const d = new Date(0); d.setUTCFullYear(f.anio, f.mes - 1, f.dia); return d; };
  const formatoFecha = (f) => `${dos(f.dia)}/${dos(f.mes)}/${f.anio}`;
  const hoy = () => { const d = new Date(); return { anio: d.getFullYear(), mes: d.getMonth() + 1, dia: d.getDate() }; };

  function calcularEdad(nac, ref) {
    const edad = ref.anio - nac.anio - ((ref.mes < nac.mes || (ref.mes === nac.mes && ref.dia < nac.dia)) ? 1 : 0);
    return edad >= 0 ? edad : null;
  }

  /** Iniciales del nombre + DD + M + AA de la fecha de nacimiento (ej.: MAF21261). */
  function generarCodigo(nombre, nacimiento) {
    const palabras = sinAcentos(nombre).toUpperCase().split(/\s+/).filter(
      (p) => p && !PARTICULAS.has(p.toLowerCase()) && /\p{L}/u.test(p[0]));
    const iniciales = palabras.map((p) => p[0]).join('') || 'X';
    if (nacimiento) return `${iniciales}${dos(nacimiento.dia)}${nacimiento.mes}${dos(nacimiento.anio % 100)}`;
    return `${iniciales}SF`;                          // SF = sin fecha de nacimiento
  }

  // ====================================================================
  //  Texto del cuestionario (pegar / copiar)
  // ====================================================================
  const ALIAS_ETIQUETAS = { q2: ['Fecha de hoy'] };
  const RE_PERSONA = /^persona\s+(\d+)\s*:?$/;
  const RE_SITUACION = /^situacion\s+(\d+)\s*:?$/;      // se compara sobre texto sin acentos
  const RE_RESPUESTA = /^r\s*:\s*(.*)$/i;

  const mapaNormalizado = (pares) => {
    const mapa = new Map();
    for (const [clave, etiqueta] of pares) { const k = normalizar(etiqueta); if (!mapa.has(k)) mapa.set(k, clave); }
    return mapa;
  };
  const camposDirectos = [];
  for (const sec of SECCIONES) {
    if (sec.especial === 'situaciones') continue;
    for (const c of aplanar(sec.campos)) if (c.tipo !== 'familia') camposDirectos.push([c.id, c.label]);
  }
  const NORM_CAMPOS = mapaNormalizado(camposDirectos);
  for (const [id, alias] of Object.entries(ALIAS_ETIQUETAS)) for (const a of alias) if (!NORM_CAMPOS.has(normalizar(a))) NORM_CAMPOS.set(normalizar(a), id);
  const NORM_FAMILIA = mapaNormalizado(FAMILIA_CAMPOS);
  const NORM_SITUACION = mapaNormalizado(SITUACION_CAMPOS);
  const FILAS = SECCIONES.flatMap((s) => s.campos.filter((c) => c.tipo === 'fila').map((c) => c.campos.map((x) => x.id)));

  const TEXTOS_FIJOS_EXTRA = [
    'Piensa en una o varias ocasiones recientes en las que se presentó tu dificultad, y describe con el mayor detalle posible cada parte.',
    'Ya casi terminas', 'Gracias por tomarte el tiempo de responder con honestidad.', 'Vaciar formulario',
    'AjnA · Psicología para la consciencia'];
  const LINEAS_TRANSPARENTES = new Set(['0 = nada    10 = máxima intensidad', '0 = nada 10 = máxima intensidad',
    '0 = nada10 = máxima intensidad'].map(normalizar));
  const LINEAS_BORDE = new Set();
  for (const sec of SECCIONES) {
    for (const c of sec.campos) if (c.tipo === 'familia') LINEAS_BORDE.add(normalizar(c.label));
    LINEAS_BORDE.add(normalizar(`${sec.letra} — ${sec.titulo}`));
    LINEAS_BORDE.add(normalizar(`${sec.letra} ${sec.titulo}`));
    if (sec.subtitulo) LINEAS_BORDE.add(normalizar(sec.subtitulo));
    for (const c of aplanar(sec.campos)) if (c.ayuda) LINEAS_BORDE.add(normalizar(c.ayuda));
  }
  for (const t of TEXTOS_FIJOS_EXTRA) LINEAS_BORDE.add(normalizar(t));
  const PREFIJOS_BORDE = ['Documento de Word descargado', 'Para enviar tu cuestionario completo',
    'o envíanos el documento por whatsapp'].map(normalizar);
  const TODAS_ETIQUETAS = new Set([...NORM_CAMPOS.keys(), ...NORM_FAMILIA.keys(), ...NORM_SITUACION.keys(), ...LINEAS_BORDE]);

  const esSeparador = (l) => l.length >= 3 && /^(=+|-+)$/.test(l);
  const esBordeFijo = (l) => { const n = normalizar(l); return LINEAS_BORDE.has(n) || PREFIJOS_BORDE.some((p) => n.startsWith(p)); };

  function textoUtil(lineas) {
    const l = [...lineas];
    while (l.length && !l[0].trim()) l.shift();
    while (l.length && !l[l.length - 1].trim()) l.pop();
    return quitarLigaduras(l.join('\n').trim());
  }

  const ETIQUETAS_EXPORTAR = { q2: 'Fecha de hoy' };

  /** Mismo texto que produce el botón «Copiar respuestas» del cuestionario web de AjnA. */
  function generarTextoCuestionario(campos, familia, situaciones) {
    const nombre = (campos.q1 || '').trim() || '(sin nombre)';
    const fecha = (campos.q2 || '').trim() || formatoFecha(hoy());
    const L = ['CUESTIONARIO DE HISTORIA CLÍNICA — AjnA', `Paciente: ${nombre}`, `Fecha: ${fecha}`, '='.repeat(50)];
    for (const sec of SECCIONES) {
      L.push('', `${sec.letra} — ${sec.titulo}`.toUpperCase(), '-'.repeat(40));
      if (sec.especial === 'situaciones') {
        situaciones.forEach((situ, i) => {
          L.push(`Situación ${i + 1}:`);
          for (const [clave, etiqueta] of SITUACION_CAMPOS) L.push(`  ${etiqueta}: ${(situ[clave] || '').trim() || '(sin responder)'}`);
          L.push('');
        });
      } else {
        for (const c of aplanar(sec.campos)) {
          if (c.tipo === 'familia') {
            L.push(c.label);
            familia.forEach((p, i) => {
              L.push(`  Persona ${i + 1}:`);
              for (const [clave, etiqueta] of FAMILIA_CAMPOS) L.push(`    ${etiqueta}: ${(p[clave] || '').trim() || '(sin responder)'}`);
            });
            L.push('');
          } else {
            L.push(ETIQUETAS_EXPORTAR[c.id] || c.label, `R: ${(campos[c.id] || '').trim() || '(sin responder)'}`, '');
          }
        }
      }
    }
    return L.join('\n');
  }

  /** Une preguntas largas que el PDF/Word partió en 2, 3 o 4 líneas. */
  function unirEtiquetasPartidas(lineas) {
    const resultado = [];
    let i = 0;
    const n = lineas.length;
    while (i < n) {
      let fusionado = false;
      for (const numPartes of [4, 3, 2]) {
        let j = i, vacios = 0;
        const partes = [];
        while (j < n && partes.length < numPartes) {
          if (lineas[j].trim()) partes.push(lineas[j].trim());
          else if (partes.length) { vacios += 1; if (vacios > 1) break; }
          else break;
          j += 1;
        }
        if (partes.length === numPartes && TODAS_ETIQUETAS.has(normalizar(partes.join(' ')))) {
          resultado.push(partes.join(' '));
          i = j; fusionado = true;
          break;
        }
      }
      if (!fusionado) { resultado.push(lineas[i]); i += 1; }
    }
    return resultado;
  }

  /** Reconoce el texto del cuestionario. Devuelve {campos, familia, situaciones, reconocidos}. */
  function analizarTextoCuestionario(texto) {
    let lineas = texto.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    lineas = lineas.filter((l) => !LINEAS_TRANSPARENTES.has(normalizar(l)));
    lineas = unirEtiquetasPartidas(lineas);
    const anclas = [];
    let personaActual = null, situacionActual = null, contexto = null;
    lineas.forEach((linea, i) => {
      const cruda = linea.trim();
      if (!cruda) return;
      if (esSeparador(cruda) || esBordeFijo(cruda)) { anclas.push([i, 'borde', null]); contexto = null; return; }
      const norm = normalizar(cruda);
      let m = RE_PERSONA.exec(norm);
      if (m) { personaActual = +m[1]; contexto = 'familia'; anclas.push([i, 'persona', personaActual]); return; }
      m = RE_SITUACION.exec(norm);
      if (m) { situacionActual = +m[1]; contexto = 'situacion'; anclas.push([i, 'situacion', situacionActual]); return; }
      if (contexto === 'familia' && NORM_FAMILIA.has(norm)) { anclas.push([i, 'campo_familia', [personaActual, NORM_FAMILIA.get(norm), null]]); return; }
      if (contexto === 'situacion' && NORM_SITUACION.has(norm)) { anclas.push([i, 'campo_situacion', [situacionActual, NORM_SITUACION.get(norm), null]]); return; }
      if (NORM_CAMPOS.has(norm)) { anclas.push([i, 'campo', NORM_CAMPOS.get(norm)]); contexto = null; return; }
      const dp = cruda.indexOf(':');
      if (dp >= 0) {
        const izq = normalizar(cruda.slice(0, dp)), der = cruda.slice(dp + 1);
        if (contexto === 'familia' && NORM_FAMILIA.has(izq)) { anclas.push([i, 'campo_familia', [personaActual, NORM_FAMILIA.get(izq), der]]); return; }
        if (contexto === 'situacion' && NORM_SITUACION.has(izq)) { anclas.push([i, 'campo_situacion', [situacionActual, NORM_SITUACION.get(izq), der]]); }
      }
    });

    const campos = {}, familia = {}, situaciones = {};
    let reconocidos = 0;
    anclas.forEach(([i, tipo, dato], indice) => {
      const fin = indice + 1 < anclas.length ? anclas[indice + 1][0] : lineas.length;
      const cuerpo = lineas.slice(i + 1, fin);
      if (tipo === 'campo') {
        const valor = [...cuerpo];
        if (valor.length) { const m = RE_RESPUESTA.exec(valor[0].trim()); if (m) valor[0] = m[1]; }
        const v = textoUtil(valor);
        reconocidos += 1;
        if (v && v !== '(sin responder)') campos[dato] = v;
      } else if (tipo === 'campo_familia' || tipo === 'campo_situacion') {
        const [grupo, clave, primera] = dato;
        const v = textoUtil((primera !== null ? [primera] : []).concat(cuerpo));
        reconocidos += 1;
        const destino = tipo === 'campo_familia' ? familia : situaciones;
        if (v && v !== '(sin responder)') (destino[grupo] = destino[grupo] || {})[clave] = v;
      }
    });
    for (const [id1, id2] of FILAS) {
      const v1 = campos[id1] || '', v2 = campos[id2] || '';
      if (!v1 && v2) {
        const partes = v2.split('\n').map((p) => p.trim()).filter(Boolean);
        if (partes.length === 2) { campos[id1] = partes[0]; campos[id2] = partes[1]; }
      }
    }
    const ordenado = (obj) => Object.keys(obj).map(Number).sort((a, b) => a - b).map((k) => obj[k]);
    return { campos, familia: ordenado(familia), situaciones: ordenado(situaciones), reconocidos };
  }

  // ====================================================================
  //  Excel
  // ====================================================================
  const COLOR_PRINCIPAL = '#5C1A43';
  const ANCHO = { texto: 26, textarea: 50, fecha: 15, numero: 10, escala: 12, familia: 26 };
  const LADO = { style: 'thin', color: { argb: 'FFD9D9D9' } };
  const BORDE = { left: LADO, right: LADO, top: LADO, bottom: LADO };
  const argb = (hex) => 'FF' + hex.replace('#', '').toUpperCase();

  const COLUMNAS_USUARIOS = [{ id: 'codigo', header: 'CÓDIGO', tipo: 'texto', color: COLOR_PRINCIPAL, ancho: 14 }];
  for (const sec of SECCIONES) {
    for (const c of aplanar(sec.campos)) {
      if (c.tipo === 'familia') continue;
      COLUMNAS_USUARIOS.push({ id: c.id, header: `${sec.titulo}\n${c.label}`, tipo: c.tipo, color: sec.color, ancho: ANCHO[c.tipo] });
    }
  }
  COLUMNAS_USUARIOS.push({ id: 'modificado', header: 'Última modificación', tipo: 'fechahora', color: COLOR_PRINCIPAL, ancho: 20 });
  const BASE_HIJAS = [
    { id: 'codigo', header: 'CÓDIGO', tipo: 'texto', color: '#7A2359', ancho: 14 },
    { id: 'usuario', header: 'Usuario', tipo: 'texto', color: '#7A2359', ancho: 32 },
    { id: 'n', header: 'N°', tipo: 'numero', color: '#7A2359', ancho: 6 }];
  const COLUMNAS_FAMILIA = BASE_HIJAS.concat(FAMILIA_CAMPOS.map(([k, e]) => ({ id: k, header: e, tipo: 'texto', color: '#7A2359', ancho: 26 })));
  const COLUMNAS_SITUACIONES = BASE_HIJAS.concat(SITUACION_CAMPOS.map(([k, e]) => ({ id: k, header: e.split(' (')[0], tipo: 'textarea', color: '#7A2359', ancho: 48 })));
  const HOJAS = [[HOJA_USUARIOS, COLUMNAS_USUARIOS], [HOJA_FAMILIA, COLUMNAS_FAMILIA], [HOJA_SITUACIONES, COLUMNAS_SITUACIONES]];

  function letraColumna(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  /** Valor de una celda de ExcelJS como texto (como lo muestra el formulario). */
  function textoValor(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return formatoFecha({ anio: v.getUTCFullYear(), mes: v.getUTCMonth() + 1, dia: v.getUTCDate() });
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join('');
      if ('result' in v) return textoValor(v.result);
      if ('text' in v) return textoValor(v.text);
      if (v.error) return String(v.error);
    }
    return String(v);
  }
  const textoCelda = (celda) => textoValor(celda.value);
  const vacio = (v) => v === null || v === undefined || v === '';

  /** Hora «de pared» local escrita como si fuera UTC (así Excel/Python la leen igual en cualquier zona). */
  function ahoraComoUTC() {
    const d = new Date();
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()));
  }
  const aMinuto = (v) => (v instanceof Date ? Math.round(v.getTime() / 60000) : null);

  function formatoCabecera(celda, color) {
    celda.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(color) } };
    celda.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
    celda.border = BORDE;
  }

  function ponerValor(celda, valor, tipo) {
    celda.value = valor === undefined ? null : valor;
    celda.font = { name: 'Arial', size: 10 };
    celda.border = BORDE;
    const centrado = ['fecha', 'numero', 'escala', 'fechahora'].includes(tipo);
    celda.alignment = { wrapText: true, vertical: 'top', horizontal: centrado ? 'center' : 'left' };
    if (tipo === 'fecha') celda.numFmt = 'DD/MM/YYYY';
    else if (tipo === 'fechahora') celda.numFmt = 'DD/MM/YYYY HH:MM';
  }

  function aExcel(col, valor) {
    if (valor === null || valor === undefined) return null;
    if (valor instanceof Date || typeof valor === 'number') return valor;
    const t = String(valor).trim();
    if (!t) return null;
    if (col.tipo === 'fecha') { const f = parsearFecha(t); return f ? fechaUTC(f) : null; }
    if (col.tipo === 'numero' || col.tipo === 'escala') return /^-?\d+$/.test(t) ? parseInt(t, 10) : t;
    return t;
  }

  function prepararHoja(wb, nombre, columnas) {
    const ws = wb.getWorksheet(nombre) || wb.addWorksheet(nombre);
    const mapa = new Map();
    const fila1 = ws.getRow(1);
    for (let c = 1; c <= ws.columnCount; c++) { const t = textoCelda(fila1.getCell(c)); if (t) mapa.set(t, c); }
    for (const col of columnas) {
      if (mapa.has(col.header)) continue;
      const c = Math.max(0, ...mapa.values()) + 1;
      const celda = fila1.getCell(c);
      celda.value = col.header;
      formatoCabecera(celda, col.color);
      ws.getColumn(c).width = col.ancho;
      mapa.set(col.header, c);
    }
    fila1.height = 66;
    ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 1, topLeftCell: 'D2', activeCell: 'A1' }];
    if (nombre === HOJA_USUARIOS && !ws.getCell('A1').note) {
      ws.getCell('A1').note = 'No cambies ni borres los títulos de las columnas: el programa los usa para encontrar y cargar los datos de cada usuario.';
    }
    return { ws, mapa };
  }

  /** Libro nuevo y vacío con las tres hojas (para cuando todavía no existe el Excel). */
  function libroNuevo() { return new ExcelJS.Workbook(); }

  /**
   * El programa de la PC (openpyxl) guarda las notas con otros nombres y rutas que los que espera
   * ExcelJS (xl/comments/comment1.xml en vez de xl/comments1.xml, etc.). Antes de abrir el archivo se
   * renombran esas partes para poder leer cualquier Excel guardado por la PC.
   */
  async function normalizarZip(datos) {
    const zip = await JSZip.loadAsync(datos);
    let cambio = false;
    const rutaAbs = (rels, destino) => (destino.startsWith('/') ? destino.slice(1)
      : 'xl/worksheets/' + destino).replace(/\/worksheets\/\.\.\//, '/');
    for (const nombre of Object.keys(zip.files)) {
      const m = /^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/.exec(nombre);
      if (!m) continue;
      const n = m[1];
      let xml = await zip.file(nombre).async('string');
      const original = xml;
      for (const [tipo, nuevaRuta, nuevoTarget] of [
        ['comments', `xl/comments${n}.xml`, `../comments${n}.xml`],
        ['vmlDrawing', `xl/drawings/vmlDrawing${n}.vml`, `../drawings/vmlDrawing${n}.vml`]]) {
        const re = new RegExp(`<Relationship\\b[^>]*Type="[^"]*/${tipo}"[^>]*>`);
        const rel = re.exec(xml);
        if (!rel) continue;
        const target = /Target="([^"]*)"/.exec(rel[0])[1];
        const vieja = rutaAbs(null, target);
        if (vieja === nuevaRuta) continue;
        const archivo = zip.file(vieja);
        if (!archivo) continue;
        zip.file(nuevaRuta, await archivo.async('uint8array'));
        zip.remove(vieja);
        xml = xml.replace(rel[0], rel[0].replace(/Target="[^"]*"/, `Target="${nuevoTarget}"`));
      }
      if (xml !== original) { zip.file(nombre, xml); cambio = true; }
    }
    return cambio ? zip.generateAsync({ type: 'uint8array' }) : datos;
  }

  /** Lee un .xlsx (ArrayBuffer/Uint8Array/Buffer) y deja listas las tres hojas. */
  async function abrirLibro(datos) {
    const wb = libroNuevo();
    if (datos) await wb.xlsx.load(await normalizarZip(datos));
    const hojas = {};
    for (const [nombre, cols] of HOJAS) hojas[nombre] = prepararHoja(wb, nombre, cols);
    return { wb, hojas };
  }
  async function guardarLibro(wb) { return wb.xlsx.writeBuffer(); }

  const textoPlano = (v) => (typeof v === 'string' ? v : textoValor(v));
  function ultimaFila(ws) { return Math.max(ws.rowCount, 1); }
  function buscarFila(ws, codigo) {
    for (let r = 2; r <= ultimaFila(ws); r++) if (textoPlano(ws.getRow(r).getCell(1).value) === codigo) return r;
    return null;
  }
  function filaLibre(ws) { let r = 2; while (!vacio(ws.getRow(r).getCell(1).value)) r++; return r; }
  function ajustarFiltro(ws) { ws.autoFilter = `A1:${letraColumna(Math.max(ws.columnCount, 1))}${ultimaFila(ws)}`; }

  function listarUsuarios(libro) {
    const { ws, mapa } = libro.hojas[HOJA_USUARIOS];
    const col = Object.fromEntries(COLUMNAS_USUARIOS.map((c) => [c.id, mapa.get(c.header)]));
    const usuarios = [];
    for (let r = 2; r <= ultimaFila(ws); r++) {
      const fila = ws.getRow(r);
      const codigo = fila.getCell(1).value;
      if (vacio(codigo)) continue;
      usuarios.push({ codigo: textoPlano(codigo), nombre: textoCelda(fila.getCell(col.q1)),
        registro: textoCelda(fila.getCell(col.q2)), edad: textoCelda(fila.getCell(col.q3)) });
    }
    return usuarios.sort((a, b) => { const x = normalizar(a.nombre), y = normalizar(b.nombre); return x < y ? -1 : x > y ? 1 : 0; });
  }

  function cargarDeExcel(libro, codigo) {
    const { ws, mapa } = libro.hojas[HOJA_USUARIOS];
    const r = buscarFila(ws, codigo);
    if (r === null) return null;
    const fila = ws.getRow(r);
    const campos = {};
    for (const c of COLUMNAS_USUARIOS) campos[c.id] = textoCelda(fila.getCell(mapa.get(c.header)));
    const mod = fila.getCell(mapa.get('Última modificación')).value;
    const leer = (hoja, columnas) => {
      const { ws: w, mapa: m } = libro.hojas[hoja];
      const registros = [];
      for (let i = 2; i <= ultimaFila(w); i++) {
        const f = w.getRow(i);
        if (textoPlano(f.getCell(1).value) !== codigo) continue;
        const orden = f.getCell(m.get('N°')).value;
        const reg = {};
        for (const c of columnas) if (!['codigo', 'usuario', 'n'].includes(c.id)) reg[c.id] = textoCelda(f.getCell(m.get(c.header)));
        registros.push([typeof orden === 'number' ? orden : 0, reg]);
      }
      return registros.sort((a, b) => a[0] - b[0]).map((x) => x[1]);
    };
    return { campos, familia: leer(HOJA_FAMILIA, COLUMNAS_FAMILIA), situaciones: leer(HOJA_SITUACIONES, COLUMNAS_SITUACIONES),
      modificado: mod instanceof Date ? mod : null };
  }

  /** Decide el código de un usuario nuevo. {codigo, existente, nombreExistente}. */
  function resolverCodigo(libro, nombre, nacimiento) {
    const existentes = new Map(listarUsuarios(libro).map((u) => [u.codigo, u.nombre]));
    const base = generarCodigo(nombre, nacimiento);
    let codigo = base, n = 1;
    while (existentes.has(codigo)) {
      if (normalizar(existentes.get(codigo)) === normalizar(nombre)) return { codigo, existente: true, nombreExistente: existentes.get(codigo) };
      n += 1; codigo = `${base}-${n}`;
    }
    return { codigo, existente: false, nombreExistente: null };
  }

  /**
   * Escribe (o actualiza) un usuario en el libro y devuelve la hora de modificación escrita (Date).
   * Si se da «modificadoEsperado» y el Excel ya tiene otra hora para ese usuario, lanza un error con
   * .conflicto = true (alguien lo cambió desde otro dispositivo) sin tocar nada.
   */
  function aplicarUsuario(libro, codigo, campos, familia, situaciones, modificadoEsperado) {
    const { ws, mapa } = libro.hojas[HOJA_USUARIOS];
    const existente = buscarFila(ws, codigo);
    if (modificadoEsperado && existente !== null) {
      const actual = ws.getRow(existente).getCell(mapa.get('Última modificación')).value;
      if (aMinuto(actual) !== aMinuto(modificadoEsperado)) {
        const e = new Error('modificado desde otro dispositivo'); e.conflicto = true; e.actual = actual instanceof Date ? actual : null; throw e;
      }
    }
    const fila = existente !== null ? existente : filaLibre(ws);
    const ahora = ahoraComoUTC();
    const valores = Object.assign({}, campos, { codigo, modificado: ahora });
    for (const col of COLUMNAS_USUARIOS) ponerValor(ws.getRow(fila).getCell(mapa.get(col.header)), aExcel(col, valores[col.id]), col.tipo);
    ajustarFiltro(ws);

    const nombre = (campos.q1 || '').trim();
    for (const [hoja, columnas, registros] of [[HOJA_FAMILIA, COLUMNAS_FAMILIA, familia], [HOJA_SITUACIONES, COLUMNAS_SITUACIONES, situaciones]]) {
      const { ws: w, mapa: m } = libro.hojas[hoja];
      for (let r = ultimaFila(w); r >= 2; r--) if (textoPlano(w.getRow(r).getCell(1).value) === codigo) w.spliceRows(r, 1);
      let n = 0;
      for (const reg of registros) {
        if (!Object.values(reg).some((v) => String(v).trim())) continue;
        n += 1;
        const f = filaLibre(w);
        const datos = Object.assign({}, reg, { codigo, usuario: nombre, n });
        for (const col of columnas) ponerValor(w.getRow(f).getCell(m.get(col.header)), aExcel(col, datos[col.id]), col.tipo);
      }
      ajustarFiltro(w);
    }
    return ahora;
  }

  return {
    SECCIONES, FAMILIA_CAMPOS, SITUACION_CAMPOS, NOMBRE_EXCEL, COLUMNAS_USUARIOS, COLUMNAS_FAMILIA, COLUMNAS_SITUACIONES,
    normalizar, parsearFecha, formatoFecha, fechaUTC, hoy, calcularEdad, generarCodigo, aplanar,
    generarTextoCuestionario, analizarTextoCuestionario,
    abrirLibro, guardarLibro, listarUsuarios, cargarDeExcel, resolverCodigo, aplicarUsuario, textoValor,
  };
});
