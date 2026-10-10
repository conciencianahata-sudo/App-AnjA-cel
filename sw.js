/* AnjA Base de Datos (celular) — pantallas y flujo de la aplicación. */
(function () {
  'use strict';

  const A = window.AnjaAlmacenes;
  const core = window.crearNucleo(window.PREGUNTAS, window.ExcelJS, window.JSZip);
  const cfg = window.ANJA_CONFIG || {};
  const almacen = (cfg.CLIENT_ID && cfg.API_KEY && cfg.APP_ID) ? new A.AlmacenDrive(cfg) : new A.AlmacenLocal();

  const estado = {
    vista: '',            // conectar | manual | lista | formulario | error | cargando
    libro: null, version: null, usuarios: [], filtro: '',
    codigo: null,         // código del usuario abierto (null = usuario nuevo)
    modificado: null,     // hora de modificación que tenía el usuario al abrirlo
    foto: '',             // formulario tal como se abrió / guardó, para saber si hay cambios
    guardando: false,
  };

  // ====================================================================
  //  Utilidades de interfaz
  // ====================================================================
  const $ = (s, r) => (r || document).querySelector(s);

  function h(tag, attrs, ...hijos) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === false || v === null || v === undefined) continue;
      if (k === 'class') n.className = v;
      else if (k === 'texto') n.textContent = v;
      else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? '' : v);
    }
    for (const c of hijos.flat(Infinity)) if (c !== null && c !== undefined && c !== false) n.append(c.nodeType ? c : document.createTextNode(c));
    return n;
  }

  function luminancia(hex) {
    const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
      .map((v) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }
  /** Texto blanco u oscuro, el que mejor se lea sobre el color de fondo. */
  function textoSobre(hex) {
    const L = luminancia(hex);
    return 1.05 / (L + 0.05) >= (L + 0.05) / (luminancia('#2E2A28') + 0.05) ? '#FFFFFF' : '#2E2A28';
  }

  let relojAviso;
  function aviso(texto, error) {
    const t = $('#toast');
    t.textContent = texto;
    t.className = 'visible' + (error ? ' error' : '');
    clearTimeout(relojAviso);
    relojAviso = setTimeout(() => { t.className = ''; }, error ? 7000 : 3200);
  }

  /** Diálogo modal. botones: [{texto, clase, valor | accion}]. Devuelve el valor elegido (undefined si se cierra). */
  function dialogo({ titulo, cuerpo, botones, hoja }) {
    return new Promise((resolver) => {
      const d = document.createElement('dialog');       // uno nuevo por diálogo: así un cierre pendiente no afecta al siguiente
      let hecho = false;
      const fin = (v) => { if (hecho) return; hecho = true; if (d.open) d.close(); d.remove(); resolver(v); };
      d.className = hoja ? 'hoja' : '';
      d.replaceChildren(
        h('div', { class: 'contenido' }, titulo ? h('h2', { texto: titulo }) : null, cuerpo),
        h('div', { class: 'botones' }, botones.map((b) => h('button', {
          type: 'button', class: 'btn ' + (b.clase || ''),
          onclick: async () => {
            if (!b.accion) { fin(b.valor); return; }      // un botón con valor false (p. ej. «Cancelar») sí cierra
            const r = await b.accion();
            if (r !== false) fin(r);                     // una acción que devuelve false deja el diálogo abierto
          },
        }, b.texto))));
      d.addEventListener('cancel', () => fin(undefined)); // tecla Esc / botón Atrás
      document.body.append(d);
      d.showModal();
    });
  }

  async function preguntar(titulo, texto, { si, no, peligro } = {}) {
    const r = await dialogo({
      titulo, cuerpo: h('p', { texto }),
      botones: [{ texto: no || 'Cancelar', valor: false }, { texto: si || 'Aceptar', clase: peligro ? 'peligro' : 'primario', valor: true }],
    });
    return r === true;
  }
  const informar = (titulo, texto) => dialogo({ titulo, cuerpo: h('p', { texto }), botones: [{ texto: 'Entendido', clase: 'primario', valor: true }] });

  function cabecera({ atras, menu, subtitulo }) {
    $('#btnAtras').hidden = !atras;
    $('#logo').hidden = !!atras;
    $('#btnMenu').hidden = !menu;
    $('#subtitulo').textContent = subtitulo || (almacen.nombre ? `Excel: ${almacen.nombre}` : 'Base de datos de usuarios');
  }
  function acciones(contenido) {
    const barra = $('#acciones');
    barra.hidden = !contenido;
    $('.inner', barra).replaceChildren(...(contenido || []));
  }
  function pintar(...nodos) { const v = $('#vista'); v.replaceChildren(...nodos); window.scrollTo(0, 0); }

  function cargando(texto) {
    estado.vista = 'cargando';
    cabecera({}); acciones(null);
    pintar(h('div', { class: 'centrado', role: 'status' }, h('div', { class: 'giro' }), h('p', { class: 'suave', texto, style: 'text-align:center' })));
  }

  function mensajeError(e) { return (e && e.message) || 'Ocurrió un error inesperado.'; }

  function vistaError(e, reintentar) {
    estado.vista = 'error';
    cabecera({ menu: almacen.tipo === 'drive' && almacen.conectado });
    acciones(null);
    const botones = [h('button', { class: 'btn primario lleno', type: 'button', onclick: reintentar, texto: 'Reintentar' })];
    if (almacen.tipo === 'drive' && (e.estado === 401 || (e.message || '').includes('Vuelve a conectar'))) {
      botones.push(h('button', { class: 'btn lleno', type: 'button', onclick: () => { almacen.desconectar(); vistaConectar(); }, texto: 'Conectar de nuevo' }));
    } else if (almacen.tipo === 'drive' && [403, 404].includes(e.estado)) {
      botones.push(h('button', { class: 'btn lleno', type: 'button', onclick: cambiarArchivo, texto: 'Cambiar archivo' }));
    }
    pintar(h('div', { class: 'centrado' }, h('h2', { texto: 'No se pudo completar' }), h('div', { class: 'error-caja', role: 'alert', texto: mensajeError(e) }), ...botones));
  }

  // ====================================================================
  //  Conexión con el Excel
  // ====================================================================
  function vistaConectar() {
    estado.vista = 'conectar';
    cabecera({ subtitulo: 'Conexión con Google Drive' });
    acciones(null);
    if (!almacen.conectado) {
      pintar(h('div', { class: 'centrado' },
        h('h2', { texto: 'Conecta tu Google Drive' }),
        h('p', { class: 'suave', texto: 'Los usuarios se guardan en un Excel de tu Drive, el mismo que usa el programa de la PC. La app solo puede ver el archivo que tú elijas, no el resto de tu Drive.' }),
        h('button', {
          class: 'btn primario lleno', type: 'button', texto: 'Conectar con Google', id: 'btnConectar',
          onclick: async (ev) => {
            ev.currentTarget.disabled = true;
            try { await almacen.conectar(); vistaConectar(); } catch (e) { aviso(mensajeError(e), true); ev.currentTarget.disabled = false; }
          },
        })));
      return;
    }
    if (!almacen.listo) {
      pintar(h('div', { class: 'centrado' },
        h('h2', { texto: 'Elige tu Excel' }),
        h('p', { class: 'suave', texto: `Busca «${A.NOMBRE_EXCEL}» en tu Drive. Es el archivo que comparte con el programa de la PC.` }),
        h('button', { class: 'btn primario lleno', type: 'button', id: 'btnElegir', texto: 'Elegir el Excel de mi Drive', onclick: cambiarArchivo }),
        h('p', { class: 'suave', texto: '¿Todavía no tienes el Excel en Drive?' }),
        h('button', { class: 'btn lleno', type: 'button', id: 'btnCrear', texto: 'Crear un Excel nuevo en mi Drive', onclick: crearExcelNuevo })));
      return;
    }
    abrirLista();     // ya conectado y con archivo recordado (p. ej. al reabrir la app): ir directo a la lista
  }

  async function cambiarArchivo() {
    try {
      const f = await almacen.elegirArchivo();
      if (!f) return;
      await abrirLista();
    } catch (e) { aviso(mensajeError(e), true); }
  }

  async function crearExcelNuevo() {
    const ok = await preguntar('Crear un Excel nuevo',
      `Se creará «${A.NOMBRE_EXCEL}» vacío en Mi unidad de tu Drive.\n\nSi ya tienes ese archivo (por ejemplo, el que usa la PC), no lo crees: elígelo con «Elegir el Excel de mi Drive».`,
      { si: 'Crear el Excel', no: 'Volver' });
    if (!ok) return;
    cargando('Creando el Excel en tu Drive…');
    try {
      const { wb } = await core.abrirLibro(null);
      await almacen.crearArchivo(await core.guardarLibro(wb), A.NOMBRE_EXCEL);
      await abrirLista();
    } catch (e) { vistaError(e, vistaConectar); }
  }

  function vistaManual() {
    estado.vista = 'manual';
    cabecera({ subtitulo: 'Modo manual' });
    acciones(null);
    const entrada = h('input', { type: 'file', accept: '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', hidden: true, id: 'archivoExcel',
      onchange: async (ev) => {
        const f = ev.target.files && ev.target.files[0];
        if (!f) return;
        try { await almacen.abrirArchivo(f); await abrirLista(); } catch (e) { vistaError(e, vistaManual); }
      } });
    pintar(h('div', { class: 'centrado' },
      h('h2', { texto: 'Abre tu Excel' }),
      h('p', { class: 'suave', texto: 'Esta app todavía no está conectada a Google Drive, así que funciona en modo manual: abres el Excel desde el teléfono, trabajas, y al guardar se descarga una copia actualizada que debes subir a tu Drive para reemplazar la anterior.' }),
      h('label', { class: 'btn primario lleno', texto: 'Abrir mi Excel' }, entrada),
      h('button', { class: 'btn lleno', type: 'button', id: 'btnVacio', texto: 'Empezar con un Excel nuevo', onclick: async () => {
        try { await abrirLista(); } catch (e) { vistaError(e, vistaManual); }
      } })));
  }

  /** Vuelve a leer el Excel (siempre la versión más reciente) y actualiza la lista de usuarios. */
  async function refrescar() {
    const { datos, version } = await almacen.leer();
    let libro;
    try { libro = await core.abrirLibro(datos); } catch (e) {
      throw new Error('No se pudo leer el archivo como Excel (.xlsx). ¿Es el archivo correcto? En Drive, elígelo de nuevo desde el menú, en «Cambiar archivo».');
    }
    estado.libro = libro; estado.version = version; estado.usuarios = core.listarUsuarios(libro);
  }

  async function abrirLista() {
    cargando('Leyendo el Excel…');
    try { await refrescar(); vistaLista(); } catch (e) { vistaError(e, abrirLista); }
  }

  // ====================================================================
  //  Lista de usuarios
  // ====================================================================
  function vistaLista() {
    estado.vista = 'lista'; estado.codigo = null; estado.modificado = null;
    cabecera({ menu: true });
    const buscar = h('input', { type: 'search', id: 'buscar', placeholder: 'Buscar por nombre o código', 'aria-label': 'Buscar usuario', autocomplete: 'off', value: estado.filtro });
    const actualizar = h('button', { class: 'icono', type: 'button', 'aria-label': 'Actualizar la lista', id: 'btnActualizar', onclick: abrirLista, texto: '↻' });
    const conteo = h('div', { class: 'conteo', 'aria-live': 'polite' });
    const ul = h('ul', { class: 'usuarios' });
    function pintarLista() {
      const q = core.normalizar(estado.filtro);
      const visibles = estado.usuarios.filter((u) => !q || core.normalizar(`${u.nombre} ${u.codigo}`).includes(q));
      conteo.textContent = estado.usuarios.length === 0 ? '' : (q ? `${visibles.length} de ${estado.usuarios.length} usuarios` : `${estado.usuarios.length} ${estado.usuarios.length === 1 ? 'usuario' : 'usuarios'}`);
      ul.replaceChildren(...visibles.map((u) => h('li', {}, h('button', { class: 'usuario', type: 'button', 'data-codigo': u.codigo, onclick: () => abrirUsuario(u.codigo) },
        h('div', { class: 'nombre', texto: u.nombre || '(sin nombre)' }),
        h('div', { class: 'meta', texto: [u.codigo, u.edad ? `${u.edad} años` : '', u.registro ? `registro ${u.registro}` : ''].filter(Boolean).join('  ·  ') })))));
      if (!visibles.length) {
        ul.replaceChildren(h('li', {}, h('p', { class: 'suave', style: 'padding:18px 4px',
          texto: estado.usuarios.length ? 'Ningún usuario coincide con la búsqueda.' : 'Todavía no hay usuarios en este Excel. Toca «Nuevo usuario» para registrar el primero.' })));
      }
    }
    buscar.addEventListener('input', () => { estado.filtro = buscar.value; pintarLista(); });
    pintarLista();
    pintar(h('div', { class: 'buscador' }, buscar, actualizar), conteo, ul);
    acciones([h('div', { class: 'estado', texto: almacen.tipo === 'drive' ? 'Conectado a Google Drive' : 'Modo manual' }),
      h('button', { class: 'btn primario', type: 'button', id: 'btnNuevo', texto: 'Nuevo usuario', onclick: () => vistaFormulario(datosVacios(), { codigo: null, modificado: null }) })]);
  }

  async function abrirUsuario(codigo) {
    cargando('Abriendo usuario…');
    try {
      await refrescar();
      const d = core.cargarDeExcel(estado.libro, codigo);
      if (!d) { vistaLista(); aviso('Ese usuario ya no está en el Excel.', true); return; }
      vistaFormulario({ campos: d.campos, familia: d.familia, situaciones: d.situaciones }, { codigo, modificado: d.modificado });
    } catch (e) { vistaError(e, () => abrirUsuario(codigo)); }
  }

  // ====================================================================
  //  Formulario
  // ====================================================================
  let refs = {};       // id de pregunta -> {get, set}
  let grupos = {};     // familia / situaciones

  const datosVacios = () => ({ campos: { q2: core.formatoFecha(core.hoy()) }, familia: [], situaciones: [] });

  function ajustarAlto(ta) { ta.style.height = 'auto'; ta.style.height = (ta.scrollHeight + 2) + 'px'; }

  function crearEntrada(c) {
    let r, nodo;
    if (c.tipo === 'textarea') {
      nodo = h('textarea', { id: c.id, rows: c.filas || 2, class: 'auto', oninput: (e) => ajustarAlto(e.target) });
      r = { get: () => nodo.value, set: (v) => { nodo.value = v || ''; ajustarAlto(nodo); } };
    } else if (c.tipo === 'escala') {
      let tocado = false;
      const salida = h('output', { for: c.id, texto: '–' });
      nodo = h('input', { type: 'range', id: c.id, min: 0, max: 10, step: 1, value: 0, 'aria-valuetext': 'sin responder' });
      const marcar = () => { tocado = true; salida.textContent = nodo.value; nodo.setAttribute('aria-valuetext', nodo.value); };
      nodo.addEventListener('input', marcar);
      const limpiar = h('button', { class: 'enlace', type: 'button', texto: 'Quitar respuesta', onclick: () => { r.set(''); nodo.dispatchEvent(new Event('change', { bubbles: true })); } });
      r = {
        get: () => (tocado ? nodo.value : ''),
        set: (v) => { const t = String(v || '').trim(); tocado = /^\d+$/.test(t); nodo.value = tocado ? t : 0; salida.textContent = tocado ? t : '–'; nodo.setAttribute('aria-valuetext', tocado ? t : 'sin responder'); },
        envoltura: h('div', { class: 'escala' }, nodo, salida,
          h('div', { class: 'pie' }, h('span', { texto: '0 = nada' }), limpiar, h('span', { texto: '10 = máxima intensidad' }))),
      };
      return { nodo: r.envoltura, r };
    } else if (c.tipo === 'fecha') {
      nodo = h('input', { type: 'text', id: c.id, inputmode: 'numeric', autocomplete: 'off', placeholder: 'DD/MM/AAAA', maxlength: 10, readonly: !!c.solo_lectura,
        oninput: (e) => {
          if (c.solo_lectura) return;
          const d = e.target.value.replace(/\D/g, '').slice(0, 8);
          e.target.value = d.length > 4 ? `${d.slice(0, 2)}/${d.slice(2, 4)}/${d.slice(4)}` : d.length > 2 ? `${d.slice(0, 2)}/${d.slice(2)}` : d;
        } });
      r = { get: () => nodo.value, set: (v) => { nodo.value = v || ''; } };
    } else {
      nodo = h('input', { type: 'text', id: c.id, autocomplete: 'off', readonly: !!c.solo_lectura, inputmode: c.tipo === 'numero' ? 'numeric' : null });
      r = { get: () => nodo.value, set: (v) => { nodo.value = v || ''; } };
    }
    return { nodo, r };
  }

  function crearCampo(c) {
    const { nodo, r } = crearEntrada(c);
    refs[c.id] = r;
    const ayuda = c.ayuda || (c.tipo === 'fecha' && !c.solo_lectura ? 'Escribe día, mes y año; las barras se agregan solas.' : '');
    return h('div', { class: 'campo' },
      h('label', { for: c.id, texto: c.label }), nodo, ayuda ? h('div', { class: 'ayuda', texto: ayuda }) : null);
  }

  /** Bloque que se puede repetir (Persona 1, 2… / Situación 1, 2…). */
  function crearGrupo({ nombreItem, campos, disposicion, textoBoton, minimo, prefijo }) {
    const lista = h('div');
    let items = [];
    const renumerar = () => items.forEach((it, i) => { it.titulo.textContent = `${nombreItem} ${i + 1}`; });
    function agregar(valores) {
      const n = items.length;
      const entradas = {};
      const titulo = h('b');
      const item = { titulo, entradas };
      const cuerpo = h('div', { class: disposicion === 'fila' ? 'fila' : '' });
      for (const [clave, etiqueta] of campos) {
        const id = `${prefijo}_${Math.random().toString(36).slice(2, 8)}_${clave}`;
        const area = disposicion !== 'fila';
        const nodo = area ? h('textarea', { id, rows: 2, class: 'auto', oninput: (e) => ajustarAlto(e.target) }) : h('input', { type: 'text', id, autocomplete: 'off', inputmode: clave === 'edad' ? 'numeric' : null });
        entradas[clave] = nodo;
        cuerpo.append(h('div', { class: 'campo' }, h('label', { for: id, texto: etiqueta }), nodo));
        if (valores && valores[clave]) nodo.value = valores[clave];
      }
      const quitar = n > 0 ? h('button', { class: 'enlace', type: 'button', texto: 'Quitar', onclick: async () => {
        if (Object.values(entradas).some((x) => x.value.trim()) && !(await preguntar('Quitar', `¿Quitar «${titulo.textContent}» con lo que escribiste?`, { si: 'Quitar', peligro: true }))) return;
        item.nodo.remove(); items = items.filter((x) => x !== item); renumerar(); notificarCambio();
      } }) : null;
      item.nodo = h('div', { class: 'item' }, h('div', { class: 'cab-item' }, titulo, quitar), cuerpo);
      lista.append(item.nodo); items.push(item); renumerar();
      requestAnimationFrame(() => item.nodo.querySelectorAll('textarea').forEach(ajustarAlto));
    }
    const nodo = h('div', {}, lista, h('button', { class: 'btn', type: 'button', style: 'margin-top:12px', texto: textoBoton, onclick: () => { agregar(); notificarCambio(); } }));
    return {
      nodo,
      obtener: () => items.map((it) => Object.fromEntries(Object.entries(it.entradas).map(([k, n]) => [k, n.value]))),
      poner: (lst) => { lista.replaceChildren(); items = []; for (let i = 0; i < Math.max(minimo, lst.length); i++) agregar(lst[i]); },
    };
  }

  function recoger() {
    const campos = {};
    for (const [id, r] of Object.entries(refs)) campos[id] = r.get();
    return { campos, familia: grupos.familia.obtener(), situaciones: grupos.situaciones.obtener() };
  }

  function poner(d, { limpio } = { limpio: true }) {
    for (const [id, r] of Object.entries(refs)) r.set((d.campos && d.campos[id]) || '');
    grupos.familia.poner(d.familia || []);
    grupos.situaciones.poner(d.situaciones || []);
    actualizarEdad();
    if (limpio) estado.foto = JSON.stringify(recoger());
    requestAnimationFrame(() => document.querySelectorAll('#vista textarea.auto').forEach(ajustarAlto));
    notificarCambio();
  }

  function actualizarEdad() {
    let edad = '';
    try {
      const nac = core.parsearFecha(refs.q4.get());
      const ref = core.parsearFecha(refs.q2.get()) || core.hoy();
      if (nac) { const e = core.calcularEdad(nac, ref); edad = e === null ? '' : String(e); }
    } catch (e) { /* fecha incompleta: sin edad */ }
    refs.q3.set(edad);
  }

  const hayCambios = () => estado.vista === 'formulario' && JSON.stringify(recoger()) !== estado.foto;

  let horaGuardado = '';
  function notificarCambio() {
    if (estado.vista !== 'formulario') return;
    const e = $('#acciones .estado');
    if (!e) return;
    const sucio = hayCambios();
    e.className = 'estado ' + (sucio ? 'sucio' : (horaGuardado ? 'limpio' : ''));
    e.textContent = sucio ? 'Cambios sin guardar' : (horaGuardado ? `Guardado a las ${horaGuardado}` : 'Sin cambios');
  }

  function vistaFormulario(datos, ctx) {
    estado.vista = 'formulario'; estado.codigo = ctx.codigo; estado.modificado = ctx.modificado; horaGuardado = '';
    refs = {}; grupos = {};
    history.pushState({ v: 'formulario' }, '');
    cabecera({ atras: true, menu: true, subtitulo: ctx.codigo ? `Editando · ${ctx.codigo}` : 'Usuario nuevo' });

    const chips = h('nav', { class: 'chips', 'aria-label': 'Secciones del cuestionario' });
    const secciones = core.SECCIONES.map((sec) => {
      const cuerpo = h('div', { class: 'cuerpo', style: `border-left-color:${sec.color}` });
      if (sec.especial === 'situaciones') {
        cuerpo.append(h('p', { class: 'suave', style: 'margin-top:12px', texto: 'Piensa en una o varias ocasiones recientes en las que se presentó la dificultad y describe cada parte con detalle.' }));
        grupos.situaciones = crearGrupo({ nombreItem: 'Situación', campos: core.SITUACION_CAMPOS, disposicion: 'columna', textoBoton: 'Agregar otra situación', minimo: 2, prefijo: 'sit' });
        cuerpo.append(grupos.situaciones.nodo);
      } else {
        for (const c of sec.campos) {
          if (c.tipo === 'familia') {
            cuerpo.append(h('div', { class: 'campo' }, h('label', { texto: c.label })));
            grupos.familia = crearGrupo({ nombreItem: 'Persona', campos: core.FAMILIA_CAMPOS, disposicion: 'fila', textoBoton: 'Agregar otra persona', minimo: 1, prefijo: 'fam' });
            cuerpo.append(grupos.familia.nodo);
          } else if (c.tipo === 'fila') {
            cuerpo.append(h('div', { class: 'fila' }, c.campos.map(crearCampo)));
          } else cuerpo.append(crearCampo(c));
        }
      }
      const nodo = h('section', { class: 'seccion', id: `sec-${sec.id}`, 'aria-labelledby': `t-${sec.id}` },
        h('div', { class: 'cab', style: `border-left-color:${sec.color}` },
          h('span', { class: 'insignia', 'aria-hidden': 'true', style: `background:${sec.color};color:${textoSobre(sec.color)}`, texto: sec.letra }),
          h('div', {}, h('h2', { id: `t-${sec.id}`, texto: sec.titulo }), sec.subtitulo ? h('div', { class: 'sub', texto: sec.subtitulo }) : null)),
        cuerpo);
      chips.append(h('button', { class: 'chip', type: 'button', 'aria-label': sec.titulo, title: sec.titulo,
        style: `background:${sec.color};color:${textoSobre(sec.color)}`, texto: sec.letra,
        onclick: () => nodo.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'start' }) }));
      return nodo;
    });

    const avisoPegar = !ctx.codigo ? h('div', { class: 'tarjeta aviso-pegar', style: 'margin-bottom:14px' },
      h('p', { texto: '¿Te llegó el cuestionario lleno? Pégalo y se llenan las preguntas solas.' }),
      h('button', { class: 'btn', type: 'button', id: 'btnPegarAviso', texto: 'Pegar cuestionario', onclick: pegarCuestionario })) : null;

    const form = h('form', { id: 'formulario', autocomplete: 'off', novalidate: true, onsubmit: (e) => e.preventDefault() }, ...secciones);
    form.addEventListener('input', notificarCambio);
    form.addEventListener('change', notificarCambio);
    form.addEventListener('input', (e) => { if (e.target.id === 'q4' || e.target.id === 'q2') actualizarEdad(); });
    pintar(...(avisoPegar ? [avisoPegar] : []), chips, form);

    acciones([h('div', { class: 'estado', 'aria-live': 'polite' }),
      h('button', { class: 'btn primario', type: 'button', id: 'btnGuardar', texto: 'Guardar', onclick: guardar })]);
    poner(datos, { limpio: true });
  }

  // ---------- guardar ----------
  async function guardar() {
    if (estado.guardando) return;
    const d = recoger();
    const nombre = (d.campos.q1 || '').trim();
    if (!nombre) { aviso('Escribe el nombre completo para poder guardar.', true); $('#q1').scrollIntoView({ block: 'center' }); $('#q1').focus(); return; }
    let nacimiento;
    try { nacimiento = core.parsearFecha(d.campos.q4); } catch (e) {
      aviso('La fecha de nacimiento debe ser real y tener el formato DD/MM/AAAA, por ejemplo 21/06/1990.', true); $('#q4').scrollIntoView({ block: 'center' }); $('#q4').focus(); return;
    }
    try { core.parsearFecha(d.campos.q2); } catch (e) { d.campos.q2 = core.formatoFecha(core.hoy()); }

    estado.guardando = true;
    const boton = $('#btnGuardar'); boton.disabled = true; boton.textContent = 'Guardando…';
    try {
      const resultado = await guardarEnAlmacen(d, nombre, nacimiento);
      if (!resultado) return;                                   // el usuario canceló
      estado.codigo = resultado.codigo; estado.modificado = resultado.modificado;
      estado.foto = JSON.stringify(recoger());
      const ahora = new Date(); horaGuardado = `${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}`;
      cabecera({ atras: true, menu: true, subtitulo: `Editando · ${resultado.codigo}` });
      notificarCambio();
      const avisoPegar = $('#btnPegarAviso'); if (avisoPegar) avisoPegar.closest('.tarjeta').remove();
      if (almacen.tipo === 'drive') aviso(`Guardado en Drive · código ${resultado.codigo}`);
      else await informar('Excel descargado', `Se descargó «${A.NOMBRE_EXCEL}» con los datos de ${nombre} (código ${resultado.codigo}).\n\nSúbelo a tu Drive para reemplazar el archivo anterior; si no, los cambios solo quedarán en este teléfono.`);
    } catch (e) {
      await informar('No se guardó', `${mensajeError(e)}\n\nLo que escribiste sigue en el formulario: puedes volver a intentarlo.`);
    } finally {
      estado.guardando = false; boton.disabled = false; boton.textContent = 'Guardar';
    }
  }

  async function guardarEnAlmacen(d, nombre, nacimiento) {
    for (let intento = 0; intento < 3; intento++) {
      const { datos, version } = await almacen.leer();                // siempre sobre la versión más reciente
      const libro = await core.abrirLibro(datos);
      let codigo = estado.codigo;
      if (!codigo) {
        const r = core.resolverCodigo(libro, nombre, nacimiento);
        codigo = r.codigo;
        if (r.existente && !(await preguntar('Usuario ya registrado',
          `Ya existe un registro de «${r.nombreExistente}» (código ${codigo}).\n\n¿Quieres actualizar sus datos con lo que hay en el formulario?`, { si: 'Actualizar sus datos', no: 'Cancelar' }))) return null;
      }
      let modificado;
      try {
        modificado = core.aplicarUsuario(libro, codigo, d.campos, d.familia, d.situaciones, estado.codigo ? estado.modificado : null);
      } catch (e) {
        if (!e.conflicto) throw e;
        const cuando = e.actual ? ` (${core.formatoFecha({ anio: e.actual.getUTCFullYear(), mes: e.actual.getUTCMonth() + 1, dia: e.actual.getUTCDate() })} ${String(e.actual.getUTCHours()).padStart(2, '0')}:${String(e.actual.getUTCMinutes()).padStart(2, '0')})` : '';
        if (!(await preguntar('Modificado desde otro dispositivo',
          `Los datos de «${nombre}» cambiaron en el Excel${cuando} después de que los abriste aquí (por ejemplo, desde la PC).\n\nSi guardas ahora, esos cambios se perderán.`,
          { si: 'Guardar de todos modos', no: 'No guardar', peligro: true }))) return null;
        modificado = core.aplicarUsuario(libro, codigo, d.campos, d.familia, d.situaciones, null);
      }
      try {
        const nuevaVersion = await almacen.escribir(await core.guardarLibro(libro.wb), version);
        estado.libro = libro; estado.version = nuevaVersion; estado.usuarios = core.listarUsuarios(libro);
        return { codigo, modificado };
      } catch (e) {
        if (!e.cambio) throw e;                                       // el archivo cambió mientras se guardaba: reintentar
      }
    }
    throw new Error('El archivo de Drive sigue cambiando (¿alguien más está guardando?). Espera unos segundos e inténtalo de nuevo.');
  }

  // ---------- pegar y copiar ----------
  async function pegarCuestionario() {
    if (hayCambios() && !(await preguntar('Reemplazar lo escrito', 'El cuestionario pegado reemplaza todo lo que hay ahora en el formulario.', { si: 'Continuar', no: 'Cancelar', peligro: true }))) return;
    const caja = h('textarea', { id: 'cajaPegar', rows: 8, placeholder: 'Pega aquí el texto del cuestionario', 'aria-label': 'Texto del cuestionario', style: 'overflow:auto' });
    const r = await dialogo({
      titulo: 'Pegar cuestionario',
      cuerpo: [h('p', { class: 'suave', texto: 'Funciona con el texto de «Copiar respuestas» del cuestionario en línea, o con el documento descargado pegado como texto.' }), caja],
      botones: [
        { texto: 'Pegar del portapapeles', accion: async () => {
          try { caja.value = await navigator.clipboard.readText(); } catch (e) { aviso('No se pudo leer el portapapeles. Mantén presionado en el cuadro y elige Pegar.', true); }
          return false;
        } },
        { texto: 'Cancelar', valor: null },
        { texto: 'Cargar en el formulario', clase: 'primario', accion: () => {
          if (!caja.value.trim()) { aviso('Pega primero el texto del cuestionario.', true); return false; }
          const res = core.analizarTextoCuestionario(caja.value);
          if (!res.reconocidos) { aviso('No se reconoció ninguna pregunta. Pega el texto completo del cuestionario.', true); return false; }
          return res;
        } },
      ],
    });
    if (!r) return;
    const campos = Object.assign({}, r.campos);
    if (!campos.q2) campos.q2 = core.formatoFecha(core.hoy());
    poner({ campos, familia: r.familia, situaciones: r.situaciones }, { limpio: false });
    window.scrollTo(0, 0);
    aviso(`Se llenaron ${r.reconocidos} preguntas. Revisa los datos y guarda.`);
  }

  async function copiarRespuestas() {
    const d = recoger();
    const texto = core.generarTextoCuestionario(d.campos, d.familia, d.situaciones);
    try { await navigator.clipboard.writeText(texto); } catch (e) {
      const t = h('textarea', { style: 'position:fixed;opacity:0' }); t.value = texto; document.body.append(t); t.select();
      let ok = false; try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
      t.remove();
      if (!ok) { aviso('No se pudo copiar automáticamente.', true); return; }
    }
    aviso('Respuestas copiadas.');
  }

  // ---------- menú ----------
  async function menu() {
    const items = [];
    if (estado.vista === 'formulario') {
      items.push({ texto: 'Pegar cuestionario', valor: pegarCuestionario });
      items.push({ texto: 'Copiar respuestas', valor: copiarRespuestas });
      if (estado.codigo) items.push({ texto: 'Nuevo usuario', valor: async () => {
        if (hayCambios() && !(await preguntar('Cambios sin guardar', 'Hay cambios sin guardar en el formulario. ¿Descartarlos?', { si: 'Descartar', peligro: true }))) return;
        vistaFormulario(datosVacios(), { codigo: null, modificado: null });
      } });
    } else {
      items.push({ texto: 'Actualizar la lista', valor: abrirLista });
      if (almacen.tipo === 'drive') {
        items.push({ texto: 'Cambiar archivo', valor: cambiarArchivo });
        items.push({ texto: 'Cerrar sesión de Google', valor: async () => {
          if (await preguntar('Cerrar sesión de Google', 'La app dejará de tener acceso a tu Excel en este teléfono. Podrás volver a conectarla cuando quieras.', { si: 'Cerrar sesión' })) { almacen.desconectar(); vistaConectar(); }
        } });
      } else items.push({ texto: 'Abrir otro Excel', valor: vistaManual });
    }
    const elegido = await dialogo({ hoja: true, cuerpo: null, botones: [...items.map((i) => ({ texto: i.texto, valor: i.valor })), { texto: 'Cerrar', valor: null }] });
    if (typeof elegido === 'function') await elegido();
  }

  // ====================================================================
  //  Arranque
  // ====================================================================
  function registrarSW() {
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(() => { /* sin modo sin conexión */ });
  }

  async function iniciar() {
    $('#btnMenu').addEventListener('click', menu);
    $('#btnAtras').addEventListener('click', () => history.back());
    window.addEventListener('popstate', async () => {
      if (estado.vista !== 'formulario') return;
      if (hayCambios() && !(await preguntar('Salir sin guardar', 'Hay cambios sin guardar en este usuario.', { si: 'Salir sin guardar', no: 'Seguir editando', peligro: true }))) {
        history.pushState({ v: 'formulario' }, ''); return;
      }
      vistaLista();
    });
    window.addEventListener('beforeunload', (e) => { if (hayCambios()) { e.preventDefault(); e.returnValue = ''; } });
    registrarSW();
    if (almacen.tipo === 'drive') {
      if (!almacen.conectado && almacen.yaConectoAntes) { cargando('Conectando con Google…'); try { await almacen.obtenerToken(false); } catch (e) { /* se pedirá con el botón */ } }
      if (!almacen.conectado || !almacen.listo) return vistaConectar();
      return abrirLista();
    }
    return vistaManual();
  }

  window.AnjaApp = { iniciar, estado, almacen, core };
  if (!window.ANJA_SIN_INICIO) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar); else iniciar();
  }
})();
