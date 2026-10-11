/* AnjA Base de Datos (celular) — dónde vive el Excel.
 *
 * Hay dos «almacenes» con la misma interfaz:
 *   - AlmacenDrive: lee y guarda el Excel directamente en Google Drive (modo normal).
 *   - AlmacenLocal: el Excel se abre desde el teléfono y se descarga al guardar (modo manual de respaldo).
 *
 * Interfaz común:  leer() -> {datos, version}   escribir(bytes, versionEsperada) -> version
 * Si el archivo cambió en Drive entre leer() y escribir(), escribir() lanza un error con .cambio = true
 * y la aplicación vuelve a leer y a aplicar el cambio sobre la versión nueva.
 */
(function (raiz) {
  'use strict';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const CLAVE_ARCHIVO = 'anja_archivo_drive';
  const CLAVE_CONECTADO = 'anja_conectado';
  const NOMBRE_EXCEL = 'AnjA Base de Datos.xlsx';

  const guardarLS = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* sin almacenamiento: se pedirá de nuevo */ } };
  const leerLS = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const borrarLS = (k) => { try { localStorage.removeItem(k); } catch (e) { /* nada */ } };
  // sessionStorage: dura solo mientras la pestaña está abierta (para la ida y vuelta a Google)
  const guardarSS = (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) { /* nada */ } };
  const leerSS = (k) => { try { return sessionStorage.getItem(k); } catch (e) { return null; } };
  const borrarSS = (k) => { try { sessionStorage.removeItem(k); } catch (e) { /* nada */ } };
  const CLAVE_ESTADO_AUTH = 'anja_auth_estado';
  const CLAVE_MODO_AUTH = 'anja_auth_modo';

  function cargarScript(src, yaCargado) {
    return new Promise((ok, mal) => {
      if (yaCargado && yaCargado()) { ok(); return; }
      if (document.querySelector(`script[src="${src}"]`) && yaCargado) {   // ya se pidió: esperar a que termine
        const t0 = Date.now();
        const espera = setInterval(() => {
          if (yaCargado()) { clearInterval(espera); ok(); }
          else if (Date.now() - t0 > 15000) { clearInterval(espera); mal(new Error('No se pudo conectar con Google. Revisa tu conexión a internet e inténtalo de nuevo.')); }
        }, 100);
        return;
      }
      const s = document.createElement('script');
      s.src = src; s.async = true;
      s.onload = () => ok();
      s.onerror = () => mal(new Error('No se pudo conectar con Google. Revisa tu conexión a internet e inténtalo de nuevo.'));
      document.head.appendChild(s);
    });
  }

  function errorLegible(estado, texto) {
    let detalle = '';
    try { detalle = JSON.parse(texto).error.message || ''; } catch (e) { /* sin detalle */ }
    const e = new Error(
      estado === 401 ? 'Google cerró la sesión. Vuelve a conectar tu cuenta.' :
      estado === 403 ? 'Google no permite usar este archivo con la app. Elige el archivo de nuevo desde «Cambiar archivo».' :
      estado === 404 ? 'No se encontró el archivo en Drive (¿lo moviste o borraste?). Elígelo de nuevo desde «Cambiar archivo».' :
      estado === 429 ? 'Google pide esperar un momento. Inténtalo de nuevo en un minuto.' :
      `Google Drive respondió con un error (${estado}). ${detalle}`.trim());
    e.estado = estado;
    return e;
  }

  class AlmacenDrive {
    constructor(cfg) {
      this.tipo = 'drive';
      this.cfg = cfg || {};
      this.token = null;
      this.venceEn = 0;
      this.errorRetorno = null;       // error devuelto por Google al volver (p. ej. interaction_required)
      this.retornoSilencioso = false; // el intento que volvió fue el silencioso (prompt=none)
      this.alRedirigir = null;        // la app lo usa para guardar lo escrito antes de salir hacia Google
      let guardado = null;
      try { guardado = JSON.parse(leerLS(CLAVE_ARCHIVO) || 'null'); } catch (e) { guardado = null; }
      this.archivo = guardado && guardado.id ? guardado : null;
      this._leerRetorno();
    }

    get configurado() { return !!(this.cfg.CLIENT_ID && this.cfg.API_KEY && this.cfg.APP_ID); }
    get listo() { return !!this.archivo; }
    get nombre() { return this.archivo ? this.archivo.nombre : ''; }
    get yaConectoAntes() { return leerLS(CLAVE_CONECTADO) === '1'; }
    get conectado() { return !!this.token && Date.now() < this.venceEn - 60000; }

    /** Si la página se abrió de vuelta desde Google, recoge el permiso que viene en la dirección (#access_token=…). */
    _leerRetorno() {
      const hash = location.hash || '';
      if (!/(^#|&)(access_token|error)=/.test(hash)) return;
      const p = new URLSearchParams(hash.replace(/^#/, ''));
      const esperado = leerSS(CLAVE_ESTADO_AUTH);
      const modo = leerSS(CLAVE_MODO_AUTH);
      borrarSS(CLAVE_ESTADO_AUTH); borrarSS(CLAVE_MODO_AUTH);
      try { history.replaceState(null, '', location.pathname + location.search); } catch (e) { /* nada */ }
      if (!esperado || p.get('state') !== esperado) return;           // no es una respuesta que pidió esta pestaña
      this.retornoSilencioso = modo === 'silencioso';
      if (p.get('error')) { this.errorRetorno = p.get('error'); return; }
      this.token = p.get('access_token');
      this.venceEn = Date.now() + (Number(p.get('expires_in')) || 3600) * 1000;
      guardarLS(CLAVE_CONECTADO, '1');
    }

    /** Dirección de esta app (debe estar registrada en Google Cloud como «URI de redireccionamiento autorizado»). */
    get direccionRetorno() { return location.origin + location.pathname.replace(/index\.html$/, ''); }

    /** Va a Google a pedir el permiso y vuelve a esta misma página. silencioso=true no muestra pantallas (falla si hace falta). */
    iniciarRedireccion(silencioso) {
      const nonce = Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, '0')).join('');
      guardarSS(CLAVE_ESTADO_AUTH, nonce); guardarSS(CLAVE_MODO_AUTH, silencioso ? 'silencioso' : 'normal');
      const q = new URLSearchParams({ client_id: this.cfg.CLIENT_ID, redirect_uri: this.direccionRetorno, response_type: 'token',
        scope: SCOPE, state: nonce, include_granted_scopes: 'true' });
      if (silencioso) q.set('prompt', 'none');
      try { if (this.alRedirigir) this.alRedirigir(); } catch (e) { /* sin borrador */ }
      location.assign('https://accounts.google.com/o/oauth2/v2/auth?' + q.toString());
      return new Promise(() => {});                                  // la página se va; no se resuelve
    }

    /** Devuelve el permiso vigente; si venció, va a Google a renovarlo (y la app restaura lo que se estaba escribiendo). */
    async obtenerToken() {
      if (this.token && Date.now() < this.venceEn - 60000) return this.token;
      return this.iniciarRedireccion(false);
    }

    async conectar() { return this.obtenerToken(); }

    desconectar() {
      const t = this.token;
      this.token = null; this.venceEn = 0; this.archivo = null;
      borrarLS(CLAVE_ARCHIVO); borrarLS(CLAVE_CONECTADO);
      try { if (t) fetch('https://oauth2.googleapis.com/revoke', { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'token=' + encodeURIComponent(t) }).catch(() => {}); } catch (e) { /* nada */ }
    }

    async _fetch(url, opciones, reintento) {
      const token = await this.obtenerToken();
      let r;
      try {
        r = await fetch(url, Object.assign({}, opciones, { headers: Object.assign({}, (opciones || {}).headers, { Authorization: `Bearer ${token}` }) }));
      } catch (e) {
        throw new Error('Sin conexión con Google Drive. Revisa tu internet e inténtalo de nuevo.');
      }
      if (r.status === 401 && reintento !== false) { this.token = null; return this._fetch(url, opciones, false); }
      if (!r.ok) throw errorLegible(r.status, await r.text());
      return r;
    }

    _guardarArchivo(id, nombre) { this.archivo = { id, nombre }; guardarLS(CLAVE_ARCHIVO, JSON.stringify(this.archivo)); }

    /** Muestra el selector de archivos de Google para elegir el Excel (da permiso solo a ese archivo). */
    async elegirArchivo() {
      const token = await this.obtenerToken();
      await cargarScript('https://apis.google.com/js/api.js', () => !!(raiz.gapi && raiz.gapi.load));
      await new Promise((ok) => raiz.gapi.load('picker', { callback: ok }));
      return new Promise((ok) => {
        const g = raiz.google.picker;
        const vista = new g.DocsView(g.ViewId.DOCS).setMimeTypes(XLSX_MIME).setIncludeFolders(true).setSelectFolderEnabled(false);
        new g.PickerBuilder()
          .addView(vista)
          .setOAuthToken(token)
          .setDeveloperKey(this.cfg.API_KEY)
          .setAppId(this.cfg.APP_ID)
          .setLocale('es')
          .setTitle(`Elige «${NOMBRE_EXCEL}»`)
          .setCallback((d) => {
            if (d.action === 'picked') { const f = d.docs[0]; this._guardarArchivo(f.id, f.name); ok(this.archivo); }
            else if (d.action === 'cancel') ok(null);
          })
          .build().setVisible(true);
      });
    }

    /** Crea un Excel nuevo en Mi unidad con los bytes dados. */
    async crearArchivo(bytes, nombre) {
      const limite = 'anja' + Date.now();
      const meta = JSON.stringify({ name: nombre || NOMBRE_EXCEL, mimeType: XLSX_MIME });
      const cuerpo = new Blob([`--${limite}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${limite}\r\nContent-Type: ${XLSX_MIME}\r\n\r\n`, bytes, `\r\n--${limite}--`]);
      const r = await this._fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,version', {
        method: 'POST', headers: { 'Content-Type': `multipart/related; boundary=${limite}` }, body: cuerpo });
      const f = await r.json();
      this._guardarArchivo(f.id, f.name);
      return this.archivo;
    }

    async _meta() {
      const r = await this._fetch(`https://www.googleapis.com/drive/v3/files/${this.archivo.id}?fields=id,name,version,mimeType,trashed`);
      const m = await r.json();
      if (m.trashed) throw new Error('El archivo está en la papelera de Drive. Restáuralo o elige otro desde «Cambiar archivo».');
      if (m.mimeType !== XLSX_MIME) throw new Error('El archivo elegido no es un Excel (.xlsx). Si es una hoja de Google, descárgala como Excel y súbela de nuevo.');
      if (m.name && m.name !== this.archivo.nombre) this._guardarArchivo(m.id, m.name);
      return m;
    }

    async leer() {
      const m = await this._meta();
      const r = await this._fetch(`https://www.googleapis.com/drive/v3/files/${this.archivo.id}?alt=media`);
      return { datos: new Uint8Array(await r.arrayBuffer()), version: m.version };
    }

    async escribir(bytes, versionEsperada) {
      const m = await this._meta();
      if (versionEsperada && m.version !== versionEsperada) {
        const e = new Error('El archivo cambió en Drive mientras guardabas.'); e.cambio = true; throw e;
      }
      const r = await this._fetch(`https://www.googleapis.com/upload/drive/v3/files/${this.archivo.id}?uploadType=media&fields=version`, {
        method: 'PATCH', headers: { 'Content-Type': XLSX_MIME }, body: new Blob([bytes], { type: XLSX_MIME }) });
      return (await r.json()).version;
    }
  }

  /** Modo manual: el Excel se abre desde el teléfono y, al guardar, se descarga la versión actualizada. */
  class AlmacenLocal {
    constructor() { this.tipo = 'local'; this.datos = null; this.version = 0; this.nombre = ''; }
    get listo() { return true; }
    async abrirArchivo(archivo) {
      this.datos = new Uint8Array(await archivo.arrayBuffer());
      this.nombre = archivo.name; this.version = 1;
    }
    async leer() { return { datos: this.datos, version: String(this.version) }; }
    async escribir(bytes) {
      this.datos = bytes; this.version += 1;
      const url = URL.createObjectURL(new Blob([bytes], { type: XLSX_MIME }));
      const a = document.createElement('a');
      a.href = url; a.download = NOMBRE_EXCEL;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      return String(this.version);
    }
  }

  raiz.AnjaAlmacenes = { AlmacenDrive, AlmacenLocal, NOMBRE_EXCEL, XLSX_MIME };
})(typeof self !== 'undefined' ? self : this);
