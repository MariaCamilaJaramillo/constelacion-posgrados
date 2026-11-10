/*
 * CONSTELACIÓN DE POSGRADOS
 * ============================================================================
 * FORMATO DE DATOS: un solo libro de Excel (.xlsx), una hoja por tabla. No hay
 * otro formato de entrada: el ejemplo, la plantilla y los datos reales usan
 * exactamente estas hojas (los IDs son texto).
 *   cursos          curso_id, nombre, creditos, tipo (disciplinar|laboratorio|icg|
 *                   proyecto|transversal), area, alcance (escuela|programa),
 *                   electiva_regla (todas|lista|ninguna), periodo_sugerido,
 *                   activo (si|no), notas,
 *                   oferta (opciones de un curso tipo «Seminarios electivos», separadas por «;» o salto de línea),
 *                   descripcion (texto libre que se muestra en el panel de detalles)
 *   programas       programa_id, nombre, nivel (maestria|enfasis|nucleo|doctorado),
 *                   padre_id (la maestría de un énfasis o núcleo), creditos_totales,
 *                   otorga_especializado, titulo_especializado,
 *                   cursos_en_especializado (4 por defecto; 3 si lleva un electivo),
 *                   electivas_requeridas (solo maestría), color_hex, notas
 *   constitucion    curso_id, programa_id, rol (disciplinar|proyecto|transversal|
 *                   laboratorio|icg|adicional_doctorado), obligatorio, periodo_sugerido, notas
 *   electivas       curso_id, programa_destino_id, accion (incluir|excluir),
 *                   etiqueta_linea, notas
 *   secuencia       curso_origen_id, curso_destino_id, tipo (prerrequisito|correquisito), notas
 *   doctorado       doctorado_id, maestria_id, creditos_homologables, creditos_propios, notas
 *   oferta_escuela  tipo (laboratorio|icg|transversal), cursos_que_se_cursan, notas
 *   Las demás hojas del libro (Léeme, diccionario, listas, resumen) se ignoran.
 *
 * LAS CUATRO RELACIONES (cada una sale de su hoja)
 *   1. Electiva (electivas + electiva_regla): del curso al punto central del
 *      programa destino (maestría, énfasis o núcleo). No une cursos con cursos.
 *   2. Secuencia (secuencia): de curso a curso. Prerrequisito con flecha;
 *      correquisito sin flecha.
 *   3. Constitución (constitucion + programas): los cursos que conforman una
 *      maestría, énfasis o núcleo comparten color y tipo de línea y se acercan.
 *   4. Doctorado (doctorado): línea punteada entre la maestría base y el doctorado.
 *
 * REGLAS DE RESOLUCIÓN
 *   - Un énfasis o núcleo es un programa con padre_id (su maestría). Un
 *     especializado no es un programa aparte: es la marca otorga_especializado.
 *   - Los cursos con alcance «escuela» pertenecen a todas las maestrías sin
 *     necesidad de filas en constitucion (punto blanco al centro). Puede haber
 *     VARIAS opciones por tipo (p. ej. 3 laboratorios): oferta_escuela dice
 *     cuántas cursa cada estudiante (laboratorio 1, transversal 2). Sin fila se
 *     cuentan todas.
 *   - El electivo es un rol, no un tipo: si electiva_regla = «todas», el curso
 *     cuenta como electivo en las demás maestrías, salvo donde haya una fila
 *     «excluir», donde ya sea obligatorio, donde su prerrequisito no esté
 *     cubierto, o si es de la escuela. «lista» y «ninguna»: solo valen las filas
 *     «incluir» de la hoja electivas. Las electivas no mueven los puntos.
 *   - Un curso obligatorio en dos maestrías (dos filas en constitucion) queda
 *     entre ambas y recibe estela hacia el grupo que le queda lejos.
 *   - Énfasis o núcleos sin cursos (por definir) se aceptan: solo generan una nota.
 *   - Las validaciones son advertencias; nunca impiden dibujar.
 *
 * CÓMO CAMBIAR UN PARÁMETRO: edita el objeto CONFIG de abajo.
 * ============================================================================
 */
const CONFIG = {
  creditosPorDefecto: 3,
  cursosPorMaestria: 12,          // la ruta de una maestría suma 12 cursos
  // --- Puntos abiertos (parámetros, no suposiciones fijas) -----------------
  laboratoriosEjemplo: 3,         // el ejemplo ofrece varios laboratorios a toda la escuela (se cursa uno)
  transversalesEjemplo: 4,        // y varias transversales (se cursan dos)
  proyecto1ElectivoEnOtras: false,// ¿el Proyecto 1 puede contar como electivo en otras maestrías?
  obligatoriosCruzados: true,     // el ejemplo permite cursos específicos obligatorios en otra maestría
  electivaPorDefecto: 'ninguna',  // valor global si electiva_en_otras está vacío: 'todas' | 'ninguna'
  doctorado: {                    // valores por defecto: la hoja doctorado manda cuando existe
    creditosTotales: 92,
    creditosHomologables: 36,     // ¿de cualquier maestría o de la base? (aquí se asume la base)
    creditosPropios: 56,          // cuántos son cursos, seminarios o tesis queda abierto: se leen de los datos
  },
  semilla: 20240611,              // semilla fija del layout
};

const TABLAS = ['cursos', 'programas', 'constitucion', 'electivas', 'secuencia', 'doctorado', 'oferta_escuela'];
const COLUMNAS = {
  cursos: ['curso_id', 'nombre', 'creditos', 'tipo', 'area', 'alcance', 'electiva_regla', 'periodo_sugerido', 'activo', 'notas', 'oferta', 'descripcion'],
  programas: ['programa_id', 'nombre', 'nivel', 'padre_id', 'creditos_totales', 'otorga_especializado', 'titulo_especializado', 'cursos_en_especializado', 'electivas_requeridas', 'color_hex', 'notas'],
  constitucion: ['curso_id', 'programa_id', 'rol', 'obligatorio', 'periodo_sugerido', 'notas'],
  electivas: ['curso_id', 'programa_destino_id', 'accion', 'etiqueta_linea', 'notas'],
  secuencia: ['curso_origen_id', 'curso_destino_id', 'tipo', 'notas'],
  doctorado: ['doctorado_id', 'maestria_id', 'creditos_homologables', 'creditos_propios', 'notas'],
  oferta_escuela: ['tipo', 'cursos_que_se_cursan', 'notas'],
};
// «electiva» no es un rol de constitucion: las electivas viven en su propia hoja
const ROLES = ['disciplinar', 'proyecto', 'transversal', 'laboratorio', 'icg', 'adicional_doctorado'];
const TIPOS = ['disciplinar', 'laboratorio', 'icg', 'proyecto', 'transversal'];
const ROL_TEXTO = {
  disciplinar: 'curso disciplinar',
  proyecto: 'proyecto, trabajo de grado o práctica',
  transversal: 'transversal',
  laboratorio: 'laboratorio',
  icg: 'ICG (compartido por toda la escuela)',
  electiva: 'electivo',
  adicional_doctorado: 'curso adicional del doctorado',
};

/* ---------------------------- utilidades ---------------------------------- */
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function norm(s) {
  return String(s == null ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}
function normClave(s) { return norm(s).replace(/[\s\-]+/g, '_'); }
function toBool(v, porDefecto) {
  const s = norm(v);
  if (s === '') return porDefecto;
  return ['si', 'sí', 's', 'yes', 'y', 'true', '1', 'x'].includes(s);
}
function txt(v) { return String(v == null ? '' : v).trim(); }

/* ---------------------------- datos de ejemplo ----------------------------- */
const MAESTRIAS_EJEMPLO = [
  { corto: 'Cultura Visual', nombre: 'Maestría en Cultura Visual', bloques: ['Imagen y archivo', 'Cine expandido', 'Fotografía y memoria', 'Pantallas y redes', 'Medios y política'] },
  { corto: 'Escrituras Críticas', nombre: 'Maestría en Escrituras Críticas', bloques: ['Ensayo y crítica', 'Narrativas del presente', 'Poéticas del cuerpo', 'Literatura y ciudad', 'Traducción y lenguas'] },
  { corto: 'Arte y Territorio', nombre: 'Maestría en Arte y Territorio', bloques: ['Paisaje y deriva', 'Arte público', 'Ecologías sensibles', 'Archivos del lugar', 'Cartografía crítica'] },
  { corto: 'Memoria y Patrimonio', nombre: 'Maestría en Memoria y Patrimonio', bloques: [], nucleo: true, otorga: true },
  { corto: 'Artes del Cuerpo', nombre: 'Maestría en Artes del Cuerpo', bloques: [], nucleo: true, otorga: false, sinBloques: true },
  { corto: 'Estudios del Sonido', nombre: 'Maestría en Estudios del Sonido', bloques: ['Escucha y archivo', 'Composición expandida', 'Paisaje sonoro', 'Voz y oralidad', 'Tecnologías del oído'] },
];
const PREFIJOS = ['Seminario de', 'Taller de', 'Teorías de', 'Prácticas de', 'Historia de', 'Cartografías de', 'Lecturas de'];
const TEMAS = ['la imagen en movimiento', 'la escritura experimental', 'los archivos vivos', 'la voz y el gesto', 'el paisaje', 'la memoria colectiva',
  'la curaduría', 'los medios sensibles', 'el cuerpo y el espacio', 'la ciudad imaginada', 'la traducción', 'el sonido', 'las formas del relato',
  'la mirada crítica', 'la materialidad', 'el patrimonio inmaterial', 'las prácticas colaborativas', 'la performatividad', 'el ensayo visual',
  'la oralidad', 'los objetos', 'el territorio', 'las pantallas', 'la edición'];
const ROMANOS = ['I', 'II', 'III'];

function generarEjemplo(P) {
  P = Object.assign({ nMaestrias: 5, conEnfasis: 3, enfasisPor: 'mixto', proyectos: 'mixto', compartidos: 5, pctElectivas: 60, doctorado: true }, P || {});
  const rng = mulberry32(CONFIG.semilla + 7);
  const T = {};
  TABLAS.forEach((t) => { T[t] = []; });
  // todas las combinaciones prefijo + tema, barajadas con la semilla: nombres únicos
  const pares = [];
  PREFIJOS.forEach((pre) => TEMAS.forEach((tema) => pares.push(pre + ' ' + tema)));
  for (let i = pares.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    const tmp = pares[i]; pares[i] = pares[j]; pares[j] = tmp;
  }
  let contador = 0;
  const nombreCurso = () => pares[contador++ % pares.length];
  const n = Math.max(2, Math.min(6, P.nMaestrias));
  const orden = [0, 1, 2, 5, 3, 4].filter((i) => i < n);
  const conEnf = orden.slice(0, Math.max(0, Math.min(n, P.conEnfasis)));
  const slots = [];

  // filas con todas las columnas de cada hoja
  const curso = (o) => Object.assign({ creditos: '3', area: '', alcance: 'programa', electiva_regla: '', periodo_sugerido: '', activo: 'si', notas: '' }, o);
  const prog = (o) => Object.assign({ padre_id: '', creditos_totales: '', otorga_especializado: '', titulo_especializado: '', cursos_en_especializado: '', electivas_requeridas: '', color_hex: '', notas: '' }, o);
  const cons = (curso_id, programa_id, rol) => T.constitucion.push({ curso_id, programa_id, rol, obligatorio: 'si', periodo_sugerido: '', notas: '' });
  const maestriaDe = (pid) => { const q = T.programas.find((p) => p.programa_id === pid); return q ? (q.padre_id || q.programa_id) : pid; };

  // cursos de toda la escuela (alcance «escuela»: no necesitan filas en constitucion)
  const NOMBRES_LAB = ['Laboratorio de escritura y edición', 'Laboratorio de imagen y sonido', 'Laboratorio de cuerpo y espacio', 'Laboratorio de archivos y memoria', 'Laboratorio de medios digitales'];
  const NOMBRES_TR = ['Transversal: lenguajes y escrituras', 'Transversal: cuerpo y espacio', 'Transversal: culturas digitales', 'Transversal: gestión y circulación', 'Transversal: ética y mediación'];
  const escuela = [];
  for (let k = 1; k <= Math.max(1, CONFIG.laboratoriosEjemplo); k++) escuela.push({ id: 'LAB' + k, nombre: NOMBRES_LAB[(k - 1) % NOMBRES_LAB.length], tipo: 'laboratorio' });
  escuela.push({ id: 'ICG1', nombre: 'ICG', tipo: 'icg' });
  for (let k = 1; k <= Math.max(2, CONFIG.transversalesEjemplo); k++) escuela.push({ id: 'TR' + k, nombre: NOMBRES_TR[(k - 1) % NOMBRES_TR.length], tipo: 'transversal' });
  escuela.forEach((e) => T.cursos.push(curso({ curso_id: e.id, nombre: e.nombre, tipo: e.tipo, area: 'Escuela', alcance: 'escuela', electiva_regla: 'ninguna' })));
  T.oferta_escuela.push(
    { tipo: 'laboratorio', cursos_que_se_cursan: '1', notas: '' },
    { tipo: 'icg', cursos_que_se_cursan: '1', notas: '' },
    { tipo: 'transversal', cursos_que_se_cursan: '2', notas: '' }
  );

  for (let m = 0; m < n; m++) {
    const def = MAESTRIAS_EJEMPLO[m];
    const id = 'M' + (m + 1);
    T.programas.push(prog({ programa_id: id, nombre: def.nombre, nivel: 'maestria', creditos_totales: '36', electivas_requeridas: '1' }));
    slots[m] = [];
    const rangoEnf = conEnf.indexOf(m);
    const tieneEnf = rangoEnf >= 0 && !def.sinBloques;
    const bloques = [];
    if (tieneEnf) {
      const nEnf = P.enfasisPor === 'mixto' ? (rangoEnf % 2 === 0 ? 3 : 5) : Number(P.enfasisPor);
      for (let b = 0; b < nEnf; b++) {
        const nom = def.bloques[b] || ('Énfasis ' + (b + 1));
        const otorga = b % 2 === 0;
        bloques.push({
          bloque_id: id + '-B' + (b + 1), nombre: nom, enfasis: true,
          otorga: otorga ? 'si' : 'no',
          titulo: otorga ? 'Especialización en ' + nom : '',
          ce: (otorga && b === 2) ? '3' : '4',
        });
      }
    } else if (!def.sinBloques) {
      bloques.push({
        bloque_id: id + '-N', nombre: 'Núcleo', enfasis: false,
        otorga: def.otorga ? 'si' : 'no',
        titulo: def.otorga ? 'Especialización en ' + def.corto : '',
        ce: '4',
      });
    }
    bloques.forEach((b) => {
      T.programas.push(prog({ programa_id: b.bloque_id, nombre: b.nombre, nivel: b.enfasis ? 'enfasis' : 'nucleo', padre_id: id, creditos_totales: '12',
        otorga_especializado: b.otorga, titulo_especializado: b.titulo, cursos_en_especializado: b.ce }));
      for (let k = 1; k <= 4; k++) {
        const cid = b.bloque_id + '-' + k;
        T.cursos.push(curso({ curso_id: cid, nombre: nombreCurso(), tipo: 'disciplinar', area: def.corto,
          electiva_regla: rng() < P.pctElectivas / 100 ? 'todas' : 'ninguna' }));
        cons(cid, b.bloque_id, 'disciplinar');
        slots[m].push({ curso: cid, bloque: b.bloque_id });
      }
    });
    if (def.sinBloques) {
      // maestría sin bloques: sus cursos disciplinares quedan dispersos (constitucion apunta a la maestría)
      for (let k = 1; k <= 4; k++) {
        const cid = id + '-D' + k;
        T.cursos.push(curso({ curso_id: cid, nombre: nombreCurso(), tipo: 'disciplinar', area: def.corto, electiva_regla: rng() < P.pctElectivas / 100 ? 'todas' : 'ninguna' }));
        cons(cid, id, 'disciplinar');
        slots[m].push({ curso: cid, bloque: '' });
      }
    }
    if (m === 0) {
      // énfasis todavía sin definir: no rompe nada, solo genera una nota
      T.programas.push(prog({ programa_id: id + '-BX', nombre: 'Curaduría (por definir)', nivel: 'enfasis', padre_id: id, creditos_totales: '12',
        otorga_especializado: 'si', titulo_especializado: 'Especialización en Curaduría', cursos_en_especializado: '4' }));
    }

    const nProy = P.proyectos === 'mixto' ? [3, 2, 3, 2, 3, 3][m] : Number(P.proyectos);
    for (let k = 1; k <= nProy; k++) {
      const cid = id + '-P' + k;
      T.cursos.push(curso({ curso_id: cid, nombre: 'Proyecto ' + ROMANOS[k - 1] + ' · ' + def.corto, tipo: 'proyecto', area: def.corto, electiva_regla: 'ninguna' }));
      cons(cid, id, 'proyecto');
      if (k > 1) T.secuencia.push({ curso_origen_id: id + '-P' + (k - 1), curso_destino_id: cid, tipo: 'prerrequisito', notas: '' });
    }
    if (nProy === 2) {
      // con 2 proyectos la maestría necesita un curso propio más para llegar a 12 cursos y 36 créditos
      const cid = id + '-X1';
      T.cursos.push(curso({ curso_id: cid, nombre: nombreCurso(), tipo: 'disciplinar', area: def.corto, electiva_regla: 'ninguna' }));
      cons(cid, id, 'disciplinar');
    }
  }

  // cursos obligatorios en dos maestrías: reemplazan un puesto de la otra maestría
  if (CONFIG.obligatoriosCruzados) {
    const movidos = new Set();
    for (let i = 0; i < P.compartidos; i++) {
      const a = (i * 2) % n;
      const b = (a + Math.floor(n / 2) + (i > 2 ? 1 : 0)) % n; // maestrías alejadas: ahí se ve la estela
      if (a === b) continue;
      const sa = slots[a].filter((s) => !movidos.has(s.curso));
      const sb = slots[b].filter((s) => !movidos.has(s.curso));
      if (!sa.length || !sb.length) continue;
      const A = sa[Math.floor(rng() * sa.length)];
      const B = sb[Math.floor(rng() * sb.length)];
      movidos.add(A.curso); movidos.add(B.curso);
      T.cursos = T.cursos.filter((c) => c.curso_id !== B.curso);
      T.constitucion = T.constitucion.filter((r) => r.curso_id !== B.curso);
      cons(A.curso, B.bloque || ('M' + (b + 1)), 'disciplinar');
      const cur = T.cursos.find((c) => c.curso_id === A.curso);
      if (cur) cur.electiva_regla = 'ninguna';
    }
  }

  // electivas por lista explícita (una de ellas también hacia un énfasis) y una exclusión
  const disc = T.cursos.filter((c) => c.tipo === 'disciplinar' && c.alcance !== 'escuela' && c.electiva_regla !== 'todas');
  for (let i = 0; i < 2 && i < disc.length && n > 2; i++) {
    const c = disc[Math.floor(rng() * disc.length)];
    c.electiva_regla = 'lista';
    const f0 = T.constitucion.find((r) => r.curso_id === c.curso_id);
    const dueno = f0 ? maestriaDe(f0.programa_id) : '';
    let puestos = 0;
    for (let m = 1; m <= n && puestos < 2; m++) {
      const pid = 'M' + m;
      if (pid === dueno) continue;
      if (T.constitucion.some((r) => r.curso_id === c.curso_id && maestriaDe(r.programa_id) === pid)) continue;
      let destino = pid;
      if (i === 0 && puestos === 1) { // el segundo destino del primer curso es un énfasis, no una maestría
        const e = T.programas.find((q) => q.padre_id === pid && q.nivel === 'enfasis' && q.programa_id.indexOf('-BX') < 0);
        if (e) destino = e.programa_id;
      }
      T.electivas.push({ curso_id: c.curso_id, programa_destino_id: destino, accion: 'incluir', etiqueta_linea: '', notas: '' });
      puestos++;
    }
  }
  const abiertas = T.cursos.filter((c) => c.electiva_regla === 'todas');
  if (abiertas.length) {
    const c = abiertas[0];
    const otro = T.programas.find((p) => p.nivel === 'maestria' && !T.constitucion.some((r) => r.curso_id === c.curso_id && maestriaDe(r.programa_id) === p.programa_id));
    if (otro) T.electivas.push({ curso_id: c.curso_id, programa_destino_id: otro.programa_id, accion: 'excluir', etiqueta_linea: '', notas: '' });
  }

  // doctorado de ejemplo: 92 = 36 homologables + 56 propios
  if (P.doctorado && n >= 2) {
    const base = 'M' + Math.min(n, 5); // se desprende de Artes del Cuerpo (M5) si existe
    T.programas.push(prog({ programa_id: 'D1', nombre: 'Doctorado en Artes y Humanidades', nivel: 'doctorado', creditos_totales: '92' }));
    T.doctorado.push({ doctorado_id: 'D1', maestria_id: base, creditos_homologables: '36', creditos_propios: '56', notas: '' });
    const dc = [
      { id: 'D1-S1', nombre: 'Seminario doctoral de teoría', cred: 8, tipo: 'disciplinar' },
      { id: 'D1-S2', nombre: 'Seminario doctoral de método', cred: 8, tipo: 'disciplinar' },
      { id: 'D1-S3', nombre: 'Seminario doctoral de escritura', cred: 8, tipo: 'disciplinar' },
      { id: 'D1-T1', nombre: 'Tesis doctoral I', cred: 12, tipo: 'proyecto' },
      { id: 'D1-T2', nombre: 'Tesis doctoral II', cred: 20, tipo: 'proyecto' },
    ];
    dc.forEach((c) => {
      T.cursos.push(curso({ curso_id: c.id, nombre: c.nombre, creditos: String(c.cred), tipo: c.tipo, area: 'Doctorado', electiva_regla: 'ninguna' }));
      cons(c.id, 'D1', 'adicional_doctorado');
    });
    T.secuencia.push({ curso_origen_id: 'D1-T1', curso_destino_id: 'D1-T2', tipo: 'prerrequisito', notas: '' });
  }
  // un correquisito de ejemplo: los dos primeros cursos disciplinares de la primera maestría se ven en el mismo periodo
  const dis = T.constitucion.filter((r) => r.rol === 'disciplinar' && maestriaDe(r.programa_id) === 'M1').slice(0, 2);
  if (dis.length === 2) T.secuencia.push({ curso_origen_id: dis[0].curso_id, curso_destino_id: dis[1].curso_id, tipo: 'correquisito', notas: 'Ejemplo de correquisito' });
  return T;
}

/* ----------------------- lectura y normalización --------------------------- */
function normalizarTablas(crudas) {
  const T = {};
  TABLAS.forEach((t) => { T[t] = []; });
  Object.keys(crudas || {}).forEach((nombre) => {
    const t = normClave(nombre);
    if (!T[t]) return;
    (crudas[nombre] || []).forEach((fila) => {
      const o = {};
      Object.keys(fila).forEach((k) => { o[normClave(k)] = txt(fila[k]); });
      if (Object.values(o).some((v) => v !== '')) T[t].push(o);
    });
  });
  return T;
}

/* -------------------- construcción del modelo y validación ----------------- */
// Paleta de marca (no se usan colores externos). Orden pensado para el contraste en cada tema.
/* ------------------- el Excel de la malla (el único formato de entrada) -------------------
 * Tres hojas pensadas para llenar a mano:
 *   programas  una fila por maestría, énfasis, bloque único (maestría sin énfasis) o doctorado
 *   cursos     una fila por curso: nombre, tipo, créditos (opcional), requiere, se cursa junto con,
 *              electivo en todas, y luego una columna por programa (la columna j es la fila j de «programas»).
 *              En cada celda: X = el curso forma parte del programa; E = cuenta ahí como electivo.
 *   reglas     columna A = clave, C = valor (créditos y cuántos cursos de escuela cursa cada estudiante)
 * Esta capa traduce la malla a las tablas internas (cursos, programas, constitucion, electivas, secuencia,
 * doctorado, oferta_escuela) que lee construirModelo; los IDs se generan solos. */
const MAX_PROG = 40;       // columnas de programas que admite la hoja «cursos»
const COL_OFERTA = 6 + 40, COL_DESC = 6 + 40 + 1, COL_PADRE = 6 + 40 + 2;   // tras las 40 columnas de programas: oferta (opciones del curso) y descripción
const COL_PROG0 = 6;       // índice (desde 0) de la primera columna de programas en «cursos» (columna G)
const NIVEL_MALLA = { maestria: 'maestria', enfasis: 'enfasis', 'bloque unico': 'nucleo', doctorado: 'doctorado' };
const NIVEL_MALLA_TEXTO = { maestria: 'maestría', enfasis: 'énfasis', nucleo: 'bloque único', doctorado: 'doctorado' };
const REGLAS_DEFECTO = {
  creditos_por_curso: 3, creditos_maestria: 36, cursos_por_bloque: 4, creditos_doctorado: 92, creditos_homologables: 36,
  electivos_por_maestria: 1, laboratorios_por_estudiante: 1, icg_por_estudiante: 1, transversales_por_estudiante: 2,
};
const ENCABEZADOS_MALLA = {
  programas: ['nombre del programa', 'nivel', 'pertenece a (maestría)', 'título de especialista (si lo da)', 'cursos del especializado (opcional)', 'electivos que pide (opcional)'],
  cursos: ['curso', 'tipo', 'créditos (opcional)', 'requiere (prerrequisito)', 'se cursa junto con', 'electivo en todas'],
  cursosExtra: ['oferta (opciones, separadas por ;)', 'descripción', 'curso padre (opcional)'],
  reglas: ['clave', 'qué es', 'valor'],
};
const REGLAS_TEXTO = {
  creditos_por_curso: 'Créditos de un curso (si la columna créditos está vacía)',
  creditos_maestria: 'Créditos de una maestría', cursos_por_bloque: 'Cursos de un énfasis o bloque único',
  creditos_doctorado: 'Créditos del doctorado', creditos_homologables: 'Créditos de la maestría que se homologan al doctorado',
  electivos_por_maestria: 'Electivos que pide cada maestría', laboratorios_por_estudiante: 'Laboratorios que cursa cada estudiante',
  icg_por_estudiante: 'Cursos ICG que cursa cada estudiante', transversales_por_estudiante: 'Transversales que cursa cada estudiante',
};

// hojas de la malla (matrices, con encabezado en la fila 0) -> { T: tablas internas, avisos: [...] }
function tablasDesdeMalla(H) {
  const avisos = [];
  const R = Object.assign({}, REGLAS_DEFECTO);
  (H.reglas || []).slice(1).forEach((f) => {
    const k = normClave(f[0]); if (!(k in R) || txt(f[2]) === '') return;
    const v = Number(txt(f[2]).replace(',', '.'));
    if (isFinite(v)) R[k] = v; else avisos.push('reglas: «' + txt(f[1] || f[0]) + '» no es un número.');
  });
  const T = {}; TABLAS.forEach((t) => { T[t] = []; });
  // programas: la posición de la fila decide su columna en «cursos»
  const pid = [], nivelDe = new Map(), maestriaPorNombre = new Map();
  const filasP = (H.programas || []).slice(1, 1 + MAX_PROG);
  filasP.forEach((f, j) => {
    const nombre = txt(f[0]); if (!nombre) { pid.push(null); return; }
    const nivel = NIVEL_MALLA[norm(f[1])];
    if (!nivel) { avisos.push('programas: «' + nombre + '» no tiene un nivel válido (maestría, énfasis, bloque único o doctorado). Se omite.'); pid.push(null); return; }
    const id = 'P' + String(j + 1).padStart(2, '0');
    pid.push(id); nivelDe.set(id, nivel);
    if (nivel === 'maestria') {
      if (maestriaPorNombre.has(norm(nombre))) avisos.push('programas: hay dos maestrías llamadas «' + nombre + '».');
      maestriaPorNombre.set(norm(nombre), id);
    }
  });
  filasP.forEach((f, j) => {
    const id = pid[j]; if (!id) return;
    const nombre = txt(f[0]), nivel = nivelDe.get(id);
    let padre = '';
    if (nivel !== 'maestria') {
      padre = maestriaPorNombre.get(norm(f[2])) || '';
      if (!padre) avisos.push('programas: «' + nombre + '» necesita en «pertenece a» el nombre exacto de una maestría.');
    }
    const titulo = txt(f[3]);
    T.programas.push({
      programa_id: id, nombre, nivel, padre_id: padre,
      creditos_totales: String(nivel === 'maestria' ? R.creditos_maestria : nivel === 'doctorado' ? R.creditos_doctorado : R.cursos_por_bloque * R.creditos_por_curso),
      otorga_especializado: nivel === 'maestria' || nivel === 'doctorado' ? '' : (titulo ? 'si' : 'no'),
      titulo_especializado: nivel === 'maestria' || nivel === 'doctorado' ? '' : titulo,
      cursos_en_especializado: nivel === 'maestria' || nivel === 'doctorado' ? '' : txt(f[4]),
      electivas_requeridas: nivel === 'maestria' ? (txt(f[5]) !== '' ? txt(f[5]) : String(R.electivos_por_maestria)) : '',
      color_hex: '', notas: '',
    });
    if (nivel === 'doctorado' && padre) T.doctorado.push({ doctorado_id: id, maestria_id: padre, creditos_homologables: String(R.creditos_homologables), creditos_propios: String(R.creditos_doctorado - R.creditos_homologables), notas: '' });
  });
  // cursos
  const idPorNombre = new Map(), filasC = [];
  let k = 0;
  // Filas de tipo «oferta»: no son cursos ni nodos; son las opciones de un curso padre. El padre es el curso de la
  // columna «oferta de» o, si está vacía, el último curso (no oferta) que aparece más arriba en la hoja.
  const ofertaDe = new Map(); let ultimo = null;
  const esOferta = (v) => ['oferta_especifica', 'oferta'].includes(normClave(v));
  // Las columnas de la hoja «cursos» se buscan por su encabezado, así se pueden añadir columnas (curso padre, descripción,
  // oferta...) en cualquier lugar antes de los programas. Los programas empiezan en la primera columna que no es conocida.
  const cab = H.cursos && H.cursos[0] ? H.cursos[0] : [];
  const hallar = (pred, otro) => { const i = cab.findIndex((x) => pred(normClave(x))); return i >= 0 ? i : otro; };
  const iNom = hallar((x) => x === 'curso', 0), iTipo = hallar((x) => x === 'tipo', 1), iCred = hallar((x) => x.startsWith('creditos'), 2);
  const iReq = hallar((x) => x.startsWith('requiere'), 3), iJunto = hallar((x) => x.startsWith('se_cursa'), 4), iElec = hallar((x) => x.startsWith('electivo'), 5);
  const iPadre = hallar((x) => x.startsWith('curso_padre'), COL_PADRE), iDesc = hallar((x) => x.startsWith('descripcion'), COL_DESC);
  const iOferta = hallar((x) => x.startsWith('oferta') && !x.startsWith('oferta_de'), COL_OFERTA);
  const conocidas = new Set([iNom, iTipo, iCred, iReq, iJunto, iElec, iPadre, iDesc, iOferta]);
  let col0 = 0; while (conocidas.has(col0)) col0++;
  if (!col0) col0 = COL_PROG0;
  const colExtra = new Set([iPadre, iDesc, iOferta]);
  (H.cursos || []).slice(1).forEach((f) => {
    const nombre = txt(f[iNom]); if (!nombre) return;
    if (esOferta(f[iTipo])) {
      const padre = txt(f[iPadre]) || ultimo;
      if (!padre) { avisos.push('cursos: la oferta «' + nombre + '» no tiene un curso padre (ponla debajo de su curso o escribe «curso padre»).'); return; }
      const clave = norm(padre);
      if (!ofertaDe.has(clave)) ofertaDe.set(clave, []);
      ofertaDe.get(clave).push({ nombre, descripcion: txt(f[iDesc]) });
      return;
    }
    ultimo = nombre;
  });
  (H.cursos || []).slice(1).forEach((f) => {
    const nombre = txt(f[iNom]); if (!nombre) return;
    if (esOferta(f[iTipo])) return;
    k++;
    const id = 'C' + String(k).padStart(3, '0');
    if (idPorNombre.has(norm(nombre))) avisos.push('cursos: el curso «' + nombre + '» está repetido; los prerrequisitos apuntan al primero.');
    else idPorNombre.set(norm(nombre), id);
    filasC.push({ f, id, nombre });
  });
  filasC.forEach(({ f, id, nombre }) => {
    let tipo = norm(f[iTipo]);
    if (!TIPOS.includes(tipo)) { if (tipo) avisos.push('cursos: «' + nombre + '» tiene un tipo desconocido («' + txt(f[iTipo]) + '»); se trata como disciplinar.'); tipo = 'disciplinar'; }
    const escuela = ['laboratorio', 'icg', 'transversal'].includes(tipo);
    let hayE = false, hayX = false;
    for (let j = 0; j < MAX_PROG; j++) {
      if (colExtra.has(col0 + j)) continue;   // columnas de texto, no de programas
      const marca = txt(f[col0 + j]).toUpperCase(); if (!marca) continue;
      const id2 = pid[j];
      if (!id2) { avisos.push('cursos: «' + nombre + '» tiene una marca en una columna sin programa.'); continue; }
      if (marca === 'X') {
        hayX = true;
        if (escuela) continue;   // los cursos de toda la escuela no se marcan
        T.constitucion.push({ curso_id: id, programa_id: id2, rol: nivelDe.get(id2) === 'doctorado' ? 'adicional_doctorado' : tipo, obligatorio: 'si', periodo_sugerido: '', notas: '' });
      } else if (marca === 'E') {
        hayE = true;
        if (!escuela) T.electivas.push({ curso_id: id, programa_destino_id: id2, accion: 'incluir', etiqueta_linea: '', notas: '' });
      } else avisos.push('cursos: «' + nombre + '» tiene la marca «' + marca + '»; solo se usan X y E.');
    }
    if (escuela && (hayX || hayE)) avisos.push('cursos: «' + nombre + '» es de toda la escuela; sus marcas se ignoran.');
    T.cursos.push({
      curso_id: id, nombre, creditos: txt(f[iCred]) || String(R.creditos_por_curso), tipo, area: '', alcance: escuela ? 'escuela' : 'programa',
      electiva_regla: escuela ? 'ninguna' : toBool(f[iElec], false) ? 'todas' : hayE ? 'lista' : 'ninguna', periodo_sugerido: '', activo: 'si', notas: '',
      oferta: [txt(f[iOferta])].concat((ofertaDe.get(norm(nombre)) || []).map((o) => o.nombre)).filter(Boolean).join('; '), descripcion: txt(f[iDesc]), oferta_desc: ofertaDe.get(norm(nombre)) || [],
    });
  });
  // secuencia: prerrequisito (requiere -> curso) y correquisito (se cursa junto con)
  filasC.forEach(({ f, id, nombre }) => {
    [[iReq, 'prerrequisito'], [iJunto, 'correquisito']].forEach(([col, tipo]) => {
      txt(f[col]).split(/[;,]/).map(txt).filter(Boolean).forEach((n) => {
        const otro = idPorNombre.get(norm(n));
        if (!otro) { avisos.push('cursos: «' + nombre + '» menciona el curso «' + n + '», que no existe en la hoja.'); return; }
        T.secuencia.push({ curso_origen_id: otro, curso_destino_id: id, tipo, notas: '' });
      });
    });
  });
  T.oferta_escuela.push(
    { tipo: 'laboratorio', cursos_que_se_cursan: String(R.laboratorios_por_estudiante), notas: '' },
    { tipo: 'icg', cursos_que_se_cursan: String(R.icg_por_estudiante), notas: '' },
    { tipo: 'transversal', cursos_que_se_cursan: String(R.transversales_por_estudiante), notas: '' }
  );
  return { T, avisos };
}

// tablas internas -> hojas de la malla (para el ejemplo descargable)
function mallaDesdeTablas(T) {
  const nomP = new Map(T.programas.map((p) => [p.programa_id, p.nombre]));
  const padreDe = (p) => p.padre_id || ((T.doctorado.find((d) => d.doctorado_id === p.programa_id) || {}).maestria_id) || '';
  const progs = T.programas.slice(0, MAX_PROG);
  const col = new Map(progs.map((p, j) => [p.programa_id, j]));
  const H = { programas: [ENCABEZADOS_MALLA.programas], cursos: [ENCABEZADOS_MALLA.cursos.concat(progs.map((p) => p.nombre), new Array(MAX_PROG - progs.length).fill(''), ENCABEZADOS_MALLA.cursosExtra)], reglas: [ENCABEZADOS_MALLA.reglas] };
  progs.forEach((p) => {
    const esBloque = p.nivel === 'enfasis' || p.nivel === 'nucleo';
    H.programas.push([p.nombre, NIVEL_MALLA_TEXTO[p.nivel] || p.nivel, p.nivel === 'maestria' ? '' : (nomP.get(padreDe(p)) || ''),
      esBloque && p.otorga_especializado === 'si' ? p.titulo_especializado : '', esBloque ? p.cursos_en_especializado : '', p.nivel === 'maestria' ? p.electivas_requeridas : '']);
  });
  const nomC = new Map(T.cursos.map((c) => [c.curso_id, c.nombre]));
  const lista = (id, tipo) => T.secuencia.filter((r) => r.curso_destino_id === id && r.tipo === tipo).map((r) => nomC.get(r.curso_origen_id)).filter(Boolean).join('; ');
  T.cursos.forEach((c) => {
    const fila = [c.nombre, c.tipo, c.creditos === String(REGLAS_DEFECTO.creditos_por_curso) ? '' : c.creditos, lista(c.curso_id, 'prerrequisito'), lista(c.curso_id, 'correquisito'), c.electiva_regla === 'todas' ? 'sí' : ''];
    const marcas = new Array(progs.length).fill('');
    T.constitucion.filter((r) => r.curso_id === c.curso_id).forEach((r) => { if (col.has(r.programa_id)) marcas[col.get(r.programa_id)] = 'X'; });
    T.electivas.filter((r) => r.curso_id === c.curso_id && r.accion !== 'excluir').forEach((r) => { if (col.has(r.programa_destino_id) && !marcas[col.get(r.programa_destino_id)]) marcas[col.get(r.programa_destino_id)] = 'E'; });
    H.cursos.push(fila.concat(marcas, new Array(MAX_PROG - progs.length).fill(''), [String(c.oferta || '').split(/[;\n]/).map(txt).filter(Boolean).join('; '), c.descripcion || '', '']));
  });
  const oferta = Object.fromEntries(T.oferta_escuela.map((o) => [o.tipo, o.cursos_que_se_cursan]));
  const dr = T.doctorado[0];
  const val = {
    creditos_por_curso: REGLAS_DEFECTO.creditos_por_curso, creditos_maestria: 36, cursos_por_bloque: 4,
    creditos_doctorado: dr ? Number(dr.creditos_homologables) + Number(dr.creditos_propios) : REGLAS_DEFECTO.creditos_doctorado,
    creditos_homologables: dr ? Number(dr.creditos_homologables) : REGLAS_DEFECTO.creditos_homologables,
    electivos_por_maestria: 1, laboratorios_por_estudiante: oferta.laboratorio || 1, icg_por_estudiante: oferta.icg || 1, transversales_por_estudiante: oferta.transversal || 2,
  };
  Object.keys(REGLAS_TEXTO).forEach((c) => H.reglas.push([c, REGLAS_TEXTO[c], val[c]]));
  return H;
}

const PALETA_OSCURO = ['#7BB9FF', '#FF57A7', '#CEFD79', '#00EA5F', '#FF7300', '#AF0022', '#2B4C00'];
const PALETA_CLARO = ['#AF0022', '#2B4C00', '#FF57A7', '#7BB9FF', '#FF7300', '#00EA5F', '#CEFD79'];
const COLOR_DOCTORADO = '#AF7300';
const COLOR_ESCUELA = { oscuro: '#EFEEE8', claro: '#070707' };

function construirModelo(T) {
  const issues = [];
  const aviso = (msg, ref) => issues.push({ nivel: 'advertencia', msg, ref: ref || '' });
  const nota = (msg, ref) => issues.push({ nivel: 'nota', msg, ref: ref || '' });
  const fila = (t, i, id) => t + ', fila ' + (i + 2) + (id ? ' (' + id + ')' : '');
  const POOL = ['laboratorio', 'icg', 'transversal'];
  const NIVELES = ['maestria', 'enfasis', 'nucleo', 'doctorado', 'otro'];

  // ---- hoja cursos
  const cursos = new Map();
  const inactivos = new Set();
  (T.cursos || []).forEach((r, i) => {
    const id = r.curso_id;
    if (!id) { aviso('Hay un curso sin curso_id.', fila('cursos', i)); return; }
    if (cursos.has(id) || inactivos.has(id)) { aviso('ID de curso duplicado: ' + id + '. Se usa la primera fila.', fila('cursos', i, id)); return; }
    if (!toBool(r.activo, true)) { inactivos.add(id); return; }
    let tipo = norm(r.tipo);
    if (!TIPOS.includes(tipo)) {
      if (tipo) aviso('Tipo de curso desconocido «' + r.tipo + '». Se trata como disciplinar.', fila('cursos', i, id));
      tipo = 'disciplinar';
    }
    let cred = parseFloat(String(r.creditos).replace(',', '.'));
    if (!isFinite(cred) || cred <= 0) {
      if (r.creditos) aviso('Créditos no válidos «' + r.creditos + '». Se usan ' + CONFIG.creditosPorDefecto + '.', fila('cursos', i, id));
      cred = CONFIG.creditosPorDefecto;
    }
    let ee = norm(r.electiva_regla);
    if (!['todas', 'lista', 'ninguna'].includes(ee)) {
      if (ee) aviso('electiva_regla desconocida «' + r.electiva_regla + '». Se usa el valor global.', fila('cursos', i, id));
      ee = '';
    }
    let al = norm(r.alcance);
    if (al !== 'escuela' && al !== 'programa') {
      if (al) aviso('alcance desconocido «' + r.alcance + '». Se trata como «programa».', fila('cursos', i, id));
      al = 'programa';
    }
    if (al === 'escuela' && !POOL.includes(tipo)) {
      aviso('«' + (r.nombre || id) + '» tiene alcance «escuela» pero su tipo no es laboratorio, icg ni transversal. Se trata como «programa».', fila('cursos', i, id));
      al = 'programa';
    }
    cursos.set(id, { id, nombre: r.nombre || id, creditos: cred, tipo, area: r.area || '', electivaEn: ee,
      rels: [], progs: [], bloques: [], electivaProgs: new Set(), electivaDest: [], prereq: [], coreq: [], escuela: false,
      alcanceEscuela: al === 'escuela', periodo: r.periodo_sugerido || '',
      // «oferta»: opciones concretas dentro de un solo curso (no suman cursos ni créditos); «descripcion»: texto libre
      oferta: String(r.oferta || '').split(/[;\n]/).map(txt).filter(Boolean), descripcion: txt(r.descripcion), ofertaDesc: Array.isArray(r.oferta_desc) ? r.oferta_desc : [] });
  });
  if (inactivos.size) nota(inactivos.size + (inactivos.size === 1 ? ' curso inactivo no se dibuja.' : ' cursos inactivos no se dibujan.'), 'cursos');

  // ---- hoja programas: maestrías y doctorados son programas; énfasis y núcleos son bloques de una maestría
  const programas = new Map();
  const bloques = new Map();
  const pendientes = [];
  (T.programas || []).forEach((r, i) => {
    const id = r.programa_id;
    if (!id) { aviso('Hay un programa sin programa_id.', fila('programas', i)); return; }
    if (programas.has(id) || bloques.has(id) || pendientes.some((x) => x.id === id)) { aviso('ID de programa duplicado: ' + id + '.', fila('programas', i, id)); return; }
    const nivel = norm(r.nivel);
    if (!NIVELES.includes(nivel)) { aviso('Nivel desconocido «' + r.nivel + '». Se omite la fila.', fila('programas', i, id)); return; }
    if (nivel === 'otro') { nota('El programa «' + (r.nombre || id) + '» tiene nivel «otro»: no se dibuja.', fila('programas', i, id)); return; }
    if (nivel === 'enfasis' || nivel === 'nucleo') { pendientes.push({ id, r, i, nivel }); return; }
    let ct = parseFloat(String(r.creditos_totales).replace(',', '.'));
    if (!isFinite(ct) || ct <= 0) ct = nivel === 'doctorado' ? CONFIG.doctorado.creditosTotales : 36;
    const er = parseInt(r.electivas_requeridas, 10);
    const hex = /^#?[0-9a-fA-F]{6}$/.test(txt(r.color_hex)) ? '#' + txt(r.color_hex).replace('#', '') : '';
    programas.set(id, { id, nombre: r.nombre || id, nivel, creditosTotales: ct, maestriaBase: '', color: '#EFEEE8',
      electivasRequeridas: isFinite(er) && er >= 0 ? er : null, colorHex: hex, homologables: null, propios: null });
  });
  pendientes.forEach(({ id, r, i, nivel }) => {
    const m = programas.get(r.padre_id);
    if (!m || m.nivel !== 'maestria') { aviso('El ' + (nivel === 'enfasis' ? 'énfasis' : 'núcleo') + ' «' + (r.nombre || id) + '» apunta a una maestría que no existe en padre_id: «' + r.padre_id + '». Se omite.', fila('programas', i, id)); return; }
    let ce = parseInt(r.cursos_en_especializado, 10);
    if (!isFinite(ce) || ce < 1) ce = 4;
    bloques.set(id, { id, maestriaId: m.id, nombre: r.nombre || id, esEnfasis: nivel === 'enfasis',
      otorgaEsp: toBool(r.otorga_especializado, false), tituloEsp: r.titulo_especializado || '', cursosEnEsp: ce, cursos: [], implicito: false });
  });
  const maestrias = [...programas.values()].filter((p) => p.nivel === 'maestria');
  const doctorados = [...programas.values()].filter((p) => p.nivel === 'doctorado');

  // ---- hoja doctorado: maestría base y créditos
  (T.doctorado || []).forEach((r, i) => {
    const ref = fila('doctorado', i, r.doctorado_id + ' ← ' + r.maestria_id);
    const d = programas.get(r.doctorado_id);
    if (!d || d.nivel !== 'doctorado') { aviso('La hoja doctorado apunta a un doctorado que no existe en programas: «' + r.doctorado_id + '».', ref); return; }
    d.maestriaBase = r.maestria_id || '';
    const hom = parseFloat(String(r.creditos_homologables).replace(',', '.'));
    const pro = parseFloat(String(r.creditos_propios).replace(',', '.'));
    d.homologables = isFinite(hom) ? hom : null;
    d.propios = isFinite(pro) ? pro : null;
  });

  // ---- hoja constitucion: qué cursos conforman qué
  const vistas = new Set();
  (T.constitucion || []).forEach((r, i) => {
    const ref = fila('constitucion', i, r.curso_id + ' → ' + r.programa_id);
    const c = cursos.get(r.curso_id);
    if (!c) { if (!inactivos.has(r.curso_id)) aviso('La relación apunta a un curso que no existe: «' + r.curso_id + '».', ref); return; }
    const b = bloques.get(r.programa_id);
    const p = b ? programas.get(b.maestriaId) : programas.get(r.programa_id);
    if (!p) { aviso('La relación apunta a un programa que no existe: «' + r.programa_id + '».', ref); return; }
    if (c.alcanceEscuela) { aviso('«' + c.nombre + '» es de toda la escuela: no necesita filas en constitucion. Se omite la fila.', ref); return; }
    const rol = norm(r.rol);
    if (rol === 'electiva') { aviso('Las electivas van en la hoja electivas, no en constitucion. Se omite la fila.', ref); return; }
    if (!ROLES.includes(rol)) { aviso('Rol desconocido «' + r.rol + '». Se omite la fila.', ref); return; }
    const clave = r.curso_id + '|' + r.programa_id + '|' + rol;
    if (vistas.has(clave)) { aviso('Relación repetida.', ref); return; }
    vistas.add(clave);
    if (rol === 'adicional_doctorado' && p.nivel !== 'doctorado') aviso('El rol adicional_doctorado solo aplica a un doctorado.', ref);
    c.rels.push({ prog: p.id, bloque: b ? b.id : null, rol, obligatorio: toBool(r.obligatorio, true), fila: i, periodo: r.periodo_sugerido || '' });
  });
  // los cursos de toda la escuela pertenecen a todas las maestrías sin filas
  cursos.forEach((c) => {
    if (!c.alcanceEscuela) return;
    maestrias.forEach((m) => c.rels.push({ prog: m.id, bloque: null, rol: c.tipo, obligatorio: true, fila: -1, periodo: '' }));
  });

  // implícitos y pertenencias
  cursos.forEach((c) => {
    c.rels.forEach((r) => {
      if (!r.obligatorio) return;
      if (!c.progs.includes(r.prog)) c.progs.push(r.prog);
      if (r.rol === 'disciplinar') {
        if (!r.bloque) {
          const bid = r.prog + '~nucleo';
          if (!bloques.has(bid)) bloques.set(bid, { id: bid, maestriaId: r.prog, nombre: 'Núcleo', esEnfasis: false, otorgaEsp: false, tituloEsp: '', cursosEnEsp: 4, cursos: [], implicito: true });
          r.bloque = bid;
        }
        if (!c.bloques.includes(r.bloque)) c.bloques.push(r.bloque);
        const b = bloques.get(r.bloque);
        if (b && !b.cursos.includes(c.id)) b.cursos.push(c.id);
      }
    });
  });
  // «bloques reales»: los implícitos (cursos disciplinares sin bloque) no forman un bloque, quedan como cursos dispersos
  cursos.forEach((c) => { c.bloquesReales = c.bloques.filter((id) => { const b = bloques.get(id); return b && !b.implicito; }); });

  // ---- hoja secuencia: prerrequisitos (con flecha) y correquisitos (sin flecha)
  (T.secuencia || []).forEach((r, i) => {
    const ref = fila('secuencia', i, r.curso_origen_id + ' → ' + r.curso_destino_id);
    const a = cursos.get(r.curso_origen_id), c = cursos.get(r.curso_destino_id);
    if (!a) { if (!inactivos.has(r.curso_origen_id)) aviso('El curso de origen «' + r.curso_origen_id + '» no existe.', ref); return; }
    if (!c) { if (!inactivos.has(r.curso_destino_id)) aviso('El curso de destino «' + r.curso_destino_id + '» no existe.', ref); return; }
    if (a === c) { aviso('Un curso no puede estar en secuencia consigo mismo.', ref); return; }
    const tipo = norm(r.tipo) || 'prerrequisito';
    if (tipo === 'prerrequisito') { if (!c.prereq.includes(a.id)) c.prereq.push(a.id); }
    else if (tipo === 'correquisito') {
      if (!c.coreq.includes(a.id)) c.coreq.push(a.id);
      if (!a.coreq.includes(c.id)) a.coreq.push(c.id);
    } else aviso('Tipo de secuencia desconocido «' + r.tipo + '». Se omite la fila.', ref);
  });
  const estado = new Map();
  const ciclo = (id, pila) => {
    if (estado.get(id) === 2) return;
    if (estado.get(id) === 1) { aviso('Prerrequisitos circulares: ' + [...pila, id].map((x) => (cursos.get(x) || {}).nombre || x).join(' → ') + '.', 'secuencia'); return; }
    estado.set(id, 1);
    (cursos.get(id).prereq || []).forEach((q) => ciclo(q, [...pila, id]));
    estado.set(id, 2);
  };
  cursos.forEach((c, id) => ciclo(id, []));

  // ---- oferta de la escuela y electivas requeridas (reglas internas por maestría)
  const reglas = new Map();
  (T.oferta_escuela || []).forEach((r, i) => {
    const ref = fila('oferta_escuela', i, r.tipo);
    const tipo = norm(r.tipo);
    const nn = parseInt(r.cursos_que_se_cursan, 10);
    if (!POOL.includes(tipo)) { aviso('Tipo de oferta desconocido «' + r.tipo + '» (laboratorio, icg o transversal).', ref); return; }
    if (!isFinite(nn) || nn < 0) { aviso('cursos_que_se_cursan no es un número.', ref); return; }
    maestrias.forEach((m) => reglas.set(m.id + '|' + tipo, nn));
  });
  maestrias.forEach((m) => { if (m.electivasRequeridas != null) reglas.set(m.id + '|electiva', m.electivasRequeridas); });

  // escuela y tipología del punto
  cursos.forEach((c) => { c.escuela = c.alcanceEscuela && maestrias.length >= 1; });
  // Tipología del curso, que define el tamaño del punto:
  //   3 = de toda la escuela · 2 = general de la maestría (núcleo, proyectos, sueltos, doctorado)
  //   1 = específico de un énfasis o de un programa especializado
  cursos.forEach((c) => {
    c.tam = c.escuela ? 3 : (c.bloquesReales.some((id) => { const b = bloques.get(id); return b.esEnfasis || b.otorgaEsp; }) ? 1 : 2);
  });

  // ofertas de toda la escuela: varias opciones por tipo, cada estudiante cursa las que pide oferta_escuela
  cursos.forEach((c) => {
    if (!c.escuela) return;
    const r = c.rels.find((x) => x.obligatorio && POOL.includes(x.rol));
    if (!r) return;
    const grupo = [...cursos.values()].filter((o) => o.rels.some((x) => x.prog === r.prog && x.rol === r.rol && x.obligatorio));
    const pide = reglas.has(r.prog + '|' + r.rol) ? reglas.get(r.prog + '|' + r.rol) : grupo.length;
    c.pool = { rol: r.rol, total: grupo.length, pide: Math.min(pide, grupo.length) };
  });

  // ---- hoja electivas: la línea va del curso al punto central del programa destino
  const excl = new Set();
  const destinoDe = (id) => {
    const b = bloques.get(id);
    if (b) return { prog: b.maestriaId, bloque: b.id, nombre: (b.esEnfasis ? 'Énfasis ' : '') + b.nombre + ' (' + (programas.get(b.maestriaId) || {}).nombre + ')' };
    const p = programas.get(id);
    return p ? { prog: p.id, bloque: null, nombre: p.nombre, nivel: p.nivel } : null;
  };
  (T.electivas || []).forEach((r, i) => {
    const ref = fila('electivas', i, r.curso_id + ' → ' + r.programa_destino_id);
    const c = cursos.get(r.curso_id);
    if (!c) { if (!inactivos.has(r.curso_id)) aviso('La electiva apunta a un curso que no existe: «' + r.curso_id + '».', ref); return; }
    const d = destinoDe(r.programa_destino_id);
    if (!d) { aviso('El destino de la electiva no existe en programas: «' + r.programa_destino_id + '».', ref); return; }
    if (d.nivel === 'doctorado') { aviso('Una electiva no puede apuntar a un doctorado.', ref); return; }
    const accion = norm(r.accion) || 'incluir';
    if (accion === 'excluir') {
      if (d.bloque) nota('«excluir» solo aplica a maestrías: se ignora para «' + d.nombre + '».', ref);
      else excl.add(c.id + '|' + d.prog);
      return;
    }
    if (accion !== 'incluir') { aviso('Acción desconocida «' + r.accion + '». Se omite la fila.', ref); return; }
    c.electivaDest.push({ prog: d.prog, bloque: d.bloque, nombre: d.nombre, etiqueta: r.etiqueta_linea || ('Electiva para ' + d.nombre) });
  });
  const obligatorioEn = (c, pid) => c.progs.includes(pid);
  const prereqCubierto = (c, pid) => c.prereq.every((q) => obligatorioEn(cursos.get(q), pid));
  cursos.forEach((c) => {
    if (c.escuela || ['laboratorio', 'icg', 'transversal'].includes(c.tipo) || c.rels.some((r) => r.rol === 'adicional_doctorado')) {
      if (c.electivaDest.length) aviso('«' + c.nombre + '» es de toda la escuela, laboratorio, ICG, transversal o del doctorado: no puede ser electivo. Se ignoran sus filas en electivas.', 'electivas, ' + c.id);
      c.electivaDest = [];
      return;
    }
    c.electivaDest = c.electivaDest.filter((d) => {
      if (excl.has(c.id + '|' + d.prog)) return false;
      if (!prereqCubierto(c, d.prog)) aviso('«' + c.nombre + '» figura como electivo en «' + d.nombre + '», pero su prerrequisito no se cumple ahí.', 'electivas, ' + c.id + ' → ' + d.prog);
      return true;
    });
    c.electivaDest.forEach((d) => { if (!obligatorioEn(c, d.prog)) c.electivaProgs.add(d.prog); });
    const modo = c.electivaEn || CONFIG.electivaPorDefecto;
    if (modo === 'todas') {
      maestrias.forEach((p) => {
        if (obligatorioEn(c, p.id) || excl.has(c.id + '|' + p.id)) return;
        if (c.tipo === 'proyecto' && !(CONFIG.proyecto1ElectivoEnOtras && c.prereq.length === 0)) return;
        if (!prereqCubierto(c, p.id)) return;
        c.electivaProgs.add(p.id);
        if (!c.electivaDest.some((d) => d.prog === p.id && !d.bloque)) c.electivaDest.push({ prog: p.id, bloque: null, nombre: p.nombre, etiqueta: 'Electiva para ' + p.nombre });
      });
    }
    if (c.electivaDest.length && !c.progs.length) aviso('El electivo «' + c.nombre + '» no tiene fila de origen: no es obligatorio en ningún programa.', 'electivas, ' + c.id);
  });

  // colores
  maestrias.forEach((p, ci) => {
    p.colores = p.colorHex ? { oscuro: p.colorHex, claro: p.colorHex } : { oscuro: PALETA_OSCURO[ci % PALETA_OSCURO.length], claro: PALETA_CLARO[ci % PALETA_CLARO.length] };
    p.color = p.colores.oscuro;
  });
  doctorados.forEach((p) => { p.colores = { oscuro: COLOR_DOCTORADO, claro: COLOR_DOCTORADO }; p.color = COLOR_DOCTORADO; });

  // validación de rutas por maestría
  const comunes = (pid) => {
    const vistos = new Set(); const res = [];
    cursos.forEach((c) => c.rels.forEach((r) => {
      if (r.prog === pid && r.obligatorio && ['laboratorio', 'icg', 'transversal', 'proyecto'].includes(r.rol) && !vistos.has(c.id)) { vistos.add(c.id); res.push(c); }
    }));
    return res;
  };
  maestrias.forEach((p) => {
    const bl = [...bloques.values()].filter((b) => b.maestriaId === p.id);
    const sueltos = bl.find((b) => b.implicito && b.cursos.length);
    const conCursos = bl.filter((b) => b.cursos.length && !b.implicito);
    const rutas = conCursos.length ? conCursos : (sueltos ? [sueltos] : []);
    bl.filter((b) => !b.cursos.length).forEach((b) => nota('El bloque «' + b.nombre + '» de «' + p.nombre + '» todavía no tiene cursos.', 'programas, ' + b.id));
    bl.filter((b) => b.cursos.length && !b.implicito && b.cursos.length !== 4).forEach((b) =>
      aviso('El bloque «' + b.nombre + '» de «' + p.nombre + '» tiene ' + b.cursos.length + ' cursos disciplinares (se esperan 4).', 'programas, ' + b.id));
    const com = comunes(p.id);
    const nElec = reglas.has(p.id + '|electiva') ? reglas.get(p.id + '|electiva') : 1;
    const esperadoN = reglas.has(p.id + '|total') ? reglas.get(p.id + '|total') : CONFIG.cursosPorMaestria;
    // laboratorio, ICG y transversal pueden tener varias opciones: cuenta lo que pide la regla; proyectos, todos
    let nCom = 0, credCom = 0;
    POOL.concat('proyecto').forEach((rol) => {
      const grupo = com.filter((c) => c.rels.some((r) => r.prog === p.id && r.rol === rol && r.obligatorio));
      if (!grupo.length) return;
      const suma = grupo.reduce((sum, c) => sum + c.creditos, 0);
      const pide = rol === 'proyecto' ? grupo.length : (reglas.has(p.id + '|' + rol) ? reglas.get(p.id + '|' + rol) : grupo.length);
      if (pide > grupo.length) aviso('«' + p.nombre + '» ofrece ' + grupo.length + ' opciones de «' + rol + '» y oferta_escuela pide ' + pide + '.', 'oferta_escuela, ' + rol);
      const k = Math.min(pide, grupo.length);
      nCom += k; credCom += k * (suma / grupo.length);
    });
    p.ruta = { cursos: 0, creditos: 0, esperadoN, esperadoC: p.creditosTotales, ok: false };
    if (!rutas.length) { aviso('«' + p.nombre + '» no tiene cursos disciplinares.', 'programas, ' + p.id); return; }
    const grupos = new Map();
    rutas.forEach((b) => {
      // la ruta de una persona: su bloque + los cursos propios de la maestría que no están en ningún bloque
      const ids = b.cursos.concat(sueltos && b !== sueltos ? sueltos.cursos : []);
      const cred = Math.round(ids.reduce((sum, id) => sum + cursos.get(id).creditos, 0) + credCom + nElec * CONFIG.creditosPorDefecto);
      const nn = ids.length + nCom + nElec;
      const k = cred + '|' + nn;
      if (!grupos.has(k)) grupos.set(k, { cred, nn, bloques: [] });
      grupos.get(k).bloques.push(b.nombre);
    });
    grupos.forEach((g) => {
      if (g.cred !== p.creditosTotales || g.nn !== esperadoN) {
        aviso('«' + p.nombre + '»' + (grupos.size > 1 || rutas.length > 1 ? ' (' + g.bloques.slice(0, 3).join(', ') + (g.bloques.length > 3 ? '…' : '') + ')' : '') +
          ' suma ' + g.cred + ' créditos y ' + g.nn + ' cursos; se esperan ' + p.creditosTotales + ' y ' + esperadoN + '.', 'programas, ' + p.id);
      }
    });
    // resumen de la ruta (la más corta): se muestra en las convenciones
    const peor = [...grupos.values()].sort((a, b) => a.nn - b.nn)[0];
    p.ruta = { cursos: peor.nn, creditos: peor.cred, esperadoN, esperadoC: p.creditosTotales, ok: [...grupos.values()].every((g) => g.cred === p.creditosTotales && g.nn === esperadoN) };
    // proyectos: la regla pide exactamente esa cantidad
    if (reglas.has(p.id + '|proyecto')) {
      const real = com.filter((c) => c.rels.some((r) => r.prog === p.id && r.rol === 'proyecto')).length;
      if (real !== reglas.get(p.id + '|proyecto')) aviso('«' + p.nombre + '» tiene ' + real + ' proyectos; la regla pide ' + reglas.get(p.id + '|proyecto') + '.', 'reglas, ' + p.id + ' / proyecto');
    }
  });

  // especializados
  bloques.forEach((b) => {
    if (!b.otorgaEsp || !b.cursos.length) return;
    const cs = b.cursos.slice(0, b.cursosEnEsp).map((id) => cursos.get(id));
    const falta = Math.max(0, 4 - b.cursosEnEsp);
    const total = cs.reduce((s, c) => s + c.creditos, 0) + falta * CONFIG.creditosPorDefecto;
    if (total !== 12) aviso('El especializado «' + (b.tituloEsp || b.nombre) + '» suma ' + total + ' créditos (se esperan 12).', 'programas, ' + b.id);
  });

  // doctorado
  doctorados.forEach((d) => {
    const base = programas.get(d.maestriaBase);
    if (!base || base.nivel !== 'maestria') aviso('El doctorado «' + d.nombre + '» no tiene una maestría base válida («' + d.maestriaBase + '»).', 'programas, ' + d.id);
    const propios = [...cursos.values()].filter((c) => c.progs.includes(d.id)).reduce((s, c) => s + c.creditos, 0);
    const hom = d.homologables != null ? d.homologables : CONFIG.doctorado.creditosHomologables;
    const esperado = d.propios != null ? d.propios : d.creditosTotales - hom;
    if (d.propios != null && d.homologables != null && d.propios + d.homologables !== d.creditosTotales) aviso('En «' + d.nombre + '», ' + d.homologables + ' homologables más ' + d.propios + ' propios no suman ' + d.creditosTotales + ' créditos.', 'doctorado, ' + d.id);
    if (propios !== esperado) aviso('El doctorado «' + d.nombre + '» suma ' + propios + ' créditos propios en constitucion; se esperan ' + esperado + ' (' + d.creditosTotales + ' menos ' + hom + ' homologables).', 'doctorado, ' + d.id);
  });

  // forma que lee la interfaz: un destino por línea de electiva (maestría o bloque) con su etiqueta
  cursos.forEach((c) => {
    c.electivaDestinos = c.electivaDest.map((d) => {
      const b = d.bloque ? bloques.get(d.bloque) : null, pr = programas.get(d.prog);
      return b
        ? { id: b.id, tipo: 'bloque', maestria: d.prog, nombre: b.nombre, etiqueta: d.etiqueta }
        : { id: d.prog, tipo: 'programa', maestria: d.prog, nombre: pr ? pr.nombre : d.prog, etiqueta: d.etiqueta };
    });
  });
  const lista = [...cursos.values()];
  const bl = [...bloques.values()];
  // índice del bloque dentro de su maestría (para el layout)
  maestrias.forEach((p) => {
    const mios = bl.filter((b) => b.maestriaId === p.id && b.cursos.length && !b.implicito);
    mios.forEach((b, k) => { b.indice = k; b.de = mios.length; });
  });
  return {
    cursos: lista,
    programas: [...programas.values()],
    maestrias, doctorados,
    bloques: bl,
    porId: new Map(lista.map((c, i) => [c.id, i])),
    programaPorId: programas,
    bloquePorId: bloques,
    issues,
  };
}

/* ---------------------------- layout por fuerzas --------------------------- */
// Dos disposiciones del mismo conjunto: «plano» (abierta, para leer los bloques en 2D) y «denso»
// (compacta y encimada, para la vista 3D). La interfaz las mezcla según la profundidad.
// gravedad: qué tanto se atraen las maestrías hacia el centro (más alto = constelaciones más cerca entre sí)
const LAYOUTS = {
  plano: { R0: 300, radioBloque: 130, dzRep: 0.25, MIN: 30, sepBloques: 190, gravedad: 0.015, resorteEscuela: 0.0008 },
  denso: { R0: 220, radioBloque: 70, dzRep: 0.6, MIN: 22, sepBloques: 0, gravedad: 0.008, resorteEscuela: 0.0012 },
};
function simular(M, modo) {
  const L = LAYOUTS[modo || 'plano'];
  const rng = mulberry32(CONFIG.semilla);
  const C = M.cursos;
  const n = C.length;
  const X = new Float64Array(n), Y = new Float64Array(n), Z = new Float64Array(n);
  const VX = new Float64Array(n), VY = new Float64Array(n), VZ = new Float64Array(n);
  const FX = new Float64Array(n), FY = new Float64Array(n), FZ = new Float64Array(n);
  const R0 = L.R0;
  const ancla = new Map();
  M.maestrias.forEach((p, i) => {
    const a = (i / M.maestrias.length) * Math.PI * 2 - Math.PI / 2;
    ancla.set(p.id, [Math.cos(a) * R0, Math.sin(a) * R0, a, Math.sin(i * 2.1 + 0.6) * 130]);
  });
  M.doctorados.forEach((d) => {
    const b = ancla.get(d.maestriaBase) || [0, 0, 0];
    const l = Math.hypot(b[0], b[1]) || 1;
    ancla.set(d.id, [b[0] / l * (l + 150), b[1] / l * (l + 150), b[2], (b[3] || 0) + 100]);
  });

  C.forEach((c, i) => {
    let x, y, z;
    if (c.escuela) {
      const a = rng() * Math.PI * 2, r = 15 + rng() * 55;
      x = Math.cos(a) * r; y = Math.sin(a) * r; z = (rng() - 0.5) * 60;
    } else {
      let ax = 0, ay = 0, az = 0, k = 0, ang = 0;
      c.progs.forEach((p) => { const a = ancla.get(p); if (a) { ax += a[0]; ay += a[1]; az += a[3] || 0; ang = a[2]; k++; } });
      if (k) { ax /= k; ay /= k; az /= k; }
      let bx = 0, by = 0, bz = 0;
      const b = c.bloquesReales.length ? M.bloquePorId.get(c.bloquesReales[0]) : null;
      if (b && b.de > 1) {
        const ba = ang + (b.indice / b.de) * Math.PI * 2;
        bx = Math.cos(ba) * L.radioBloque; by = Math.sin(ba) * L.radioBloque; bz = (b.indice % 2 ? 1 : -1) * 55;
      } else if (!b) { bx = (rng() - 0.5) * 110; by = (rng() - 0.5) * 110; bz = (rng() - 0.5) * 80; }
      x = ax + bx + (rng() - 0.5) * 40; y = ay + by + (rng() - 0.5) * 40; z = az + bz + (rng() - 0.5) * 50;
    }
    X[i] = x; Y[i] = y; Z[i] = z;
  });

  // aristas (resortes) y grupos (imanes)
  const aristas = new Map();
  // ow: 0 = el resorte mueve a los dos; 1 = solo a i; 2 = solo a j
  const poner = (i, j, k, L, ow) => {
    if (i === j) return;
    const key = i < j ? i * n + j : j * n + i;
    const prev = aristas.get(key);
    if (!prev || prev.k < k) aristas.set(key, { i, j, k, L, ow: ow || 0 });
  };
  const cruzado = (i) => !C[i].escuela && C[i].progs.length > 1;
  const grupoProg = new Map();
  M.programas.forEach((p) => grupoProg.set(p.id, []));
  C.forEach((c, i) => { if (!c.escuela) c.progs.forEach((p) => { if (grupoProg.has(p)) grupoProg.get(p).push(i); }); });
  grupoProg.forEach((mem) => { for (let a = 0; a < mem.length; a++) for (let b = a + 1; b < mem.length; b++) {
    // un curso obligatorio en dos maestrías es atraído por cada grupo, pero no los arrastra a ellos
    const ca = cruzado(mem[a]), cb = cruzado(mem[b]);
    poner(mem[a], mem[b], ca || cb ? 0.0012 : 0.006, 105, ca && !cb ? 1 : !ca && cb ? 2 : 0);
  } });
  const grupoBloque = [];
  M.bloques.forEach((b) => {
    if (b.implicito) return;
    const mem = b.cursos.map((id) => M.porId.get(id)).filter((i) => i != null);
    grupoBloque.push(mem);
    for (let a = 0; a < mem.length; a++) for (let c = a + 1; c < mem.length; c++) poner(mem[a], mem[c], 0.09, 28);
  });
  C.forEach((c, i) => c.prereq.forEach((q) => { const j = M.porId.get(q); if (j != null) poner(i, j, 0.05, 48); }));
  const esc = []; C.forEach((c, i) => { if (c.escuela) esc.push(i); });
  esc.forEach((i) => {
    C.forEach((c, j) => { if (!c.escuela) poner(i, j, L.resorteEscuela, 170); });
    esc.forEach((j) => poner(i, j, 0.012, 50));
  });
  const lista = [...aristas.values()];

  const ITER = 700;
  const MIN = L.MIN;
  const SEPARACION_BLOQUES = L.sepBloques; // distancia mínima buscada entre centros de bloques (0 = sin separación)
  for (let it = 0; it < ITER; it++) {
    const al = Math.max(0.04, 1 - it / ITER);
    FX.fill(0); FY.fill(0); FZ.fill(0);
    for (let i = 0; i < n; i++) {
      for (let j = i + 1; j < n; j++) {
        const dx = X[i] - X[j], dy = Y[i] - Y[j], dz = (Z[i] - Z[j]) * L.dzRep;
        const d2 = dx * dx + dy * dy + dz * dz + 1;
        const d = Math.sqrt(d2);
        let f = 3400 / d2;
        if (d < MIN) f += (MIN - d) * 0.4;
        f /= d;
        FX[i] += dx * f; FY[i] += dy * f; FZ[i] += dz * f * 0.6;
        FX[j] -= dx * f; FY[j] -= dy * f; FZ[j] -= dz * f * 0.6;
      }
    }
    for (let e = 0; e < lista.length; e++) {
      const a = lista[e];
      const dx = X[a.j] - X[a.i], dy = Y[a.j] - Y[a.i], dz = Z[a.j] - Z[a.i];
      const d = Math.sqrt(dx * dx + dy * dy + dz * dz) + 0.01;
      const f = a.k * (d - a.L) / d;
      if (a.ow !== 2) { FX[a.i] += dx * f; FY[a.i] += dy * f; FZ[a.i] += dz * f; }
      if (a.ow !== 1) { FX[a.j] -= dx * f; FY[a.j] -= dy * f; FZ[a.j] -= dz * f; }
    }
    // imanes hacia el centroide de cada maestría y de cada bloque
    const imanes = (mem, k) => {
      if (mem.length < 2) return;
      let cx = 0, cy = 0, cz = 0;
      // el centroide lo definen los cursos propios del grupo, no los que también pertenecen a otro
      const propios = mem.filter((i) => !cruzado(i));
      const base = propios.length >= 2 ? propios : mem;
      base.forEach((i) => { cx += X[i]; cy += Y[i]; cz += Z[i]; });
      cx /= base.length; cy /= base.length; cz /= base.length;
      mem.forEach((i) => { FX[i] += (cx - X[i]) * k; FY[i] += (cy - Y[i]) * k; FZ[i] += (cz - Z[i]) * k; });
    };
    grupoProg.forEach((mem) => imanes(mem, 0.012));
    grupoBloque.forEach((mem) => imanes(mem, 0.03));
    // los bloques se repelen entre sí en el plano: así se distinguen también en la vista 2D
    const cb = grupoBloque.filter((m) => m.length >= 2).map((mem) => {
      let cx = 0, cy = 0; mem.forEach((i) => { cx += X[i]; cy += Y[i]; });
      return { mem, cx: cx / mem.length, cy: cy / mem.length };
    });
    for (let a = 0; a < cb.length && SEPARACION_BLOQUES > 0; a++) {
      for (let b = a + 1; b < cb.length; b++) {
        const dx = cb[a].cx - cb[b].cx, dy = cb[a].cy - cb[b].cy;
        const d = Math.sqrt(dx * dx + dy * dy) + 0.01;
        if (d >= SEPARACION_BLOQUES) continue;
        const f = (SEPARACION_BLOQUES - d) * 0.03 / d;
        cb[a].mem.forEach((i) => { FX[i] += dx * f; FY[i] += dy * f; });
        cb[b].mem.forEach((i) => { FX[i] -= dx * f; FY[i] -= dy * f; });
      }
    }
    // el doctorado se desprende de su maestría base: sus cursos propios son atraídos hacia ella
    M.doctorados.forEach((d) => {
      const mem = grupoProg.get(d.id), base = grupoProg.get(d.maestriaBase);
      if (!mem || !base || !mem.length || !base.length) return;
      let bx = 0, by = 0, bz = 0;
      base.forEach((i) => { bx += X[i]; by += Y[i]; bz += Z[i]; });
      bx /= base.length; by /= base.length; bz /= base.length;
      mem.forEach((i) => {
        const dx = bx - X[i], dy = by - Y[i], dz = bz - Z[i];
        const dist = Math.hypot(dx, dy, dz) || 1;
        if (dist > 210) { const f = (dist - 210) * 0.02 / dist; FX[i] += dx * f; FY[i] += dy * f; FZ[i] += dz * f; }
      });
    });
    for (let i = 0; i < n; i++) {
      FX[i] -= X[i] * L.gravedad; FY[i] -= Y[i] * L.gravedad; FZ[i] -= Z[i] * 0.006;
      VX[i] = (VX[i] + FX[i] * al) * 0.82; VY[i] = (VY[i] + FY[i] * al) * 0.82; VZ[i] = (VZ[i] + FZ[i] * al) * 0.82;
      const v = Math.hypot(VX[i], VY[i], VZ[i]);
      if (v > 14) { const s = 14 / v; VX[i] *= s; VY[i] *= s; VZ[i] *= s; }
      X[i] += VX[i]; Y[i] += VY[i]; Z[i] += VZ[i];
    }
  }
  let mx = 0, my = 0, mz = 0;
  for (let i = 0; i < n; i++) { mx += X[i]; my += Y[i]; mz += Z[i]; }
  if (n) { mx /= n; my /= n; mz /= n; }
  let R = 1;
  for (let i = 0; i < n; i++) { X[i] -= mx; Y[i] -= my; Z[i] -= mz; R = Math.max(R, Math.hypot(X[i], Y[i])); }
  // la profundidad se reescala para que la vista 3D se note sin importar el tamaño del conjunto
  let zm = 1;
  for (let i = 0; i < n; i++) zm = Math.max(zm, Math.abs(Z[i]));
  const kz = Math.min(3, 0.45 * R / zm);
  for (let i = 0; i < n; i++) Z[i] *= kz;
  return { X, Y, Z, R };
}

/* ============================================================================
 * INTERFAZ: dibujo en canvas, interacción, carga de archivos y paneles
 * ========================================================================== */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const lienzo = $('lienzo');
  const ctx = lienzo.getContext('2d');
  const reducir = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  const ZOOM_MIN = 0.6, ZOOM_MAX = 6;
  // Radio del punto (px) según la tipología del curso: escuela > general de la maestría > específico de énfasis
  const TAM_PUNTO = { 3: 10, 2: 8, 1: 3 };
  const TAM_TEXTO = { 3: 'Toda la escuela', 2: 'General de la maestría', 1: 'Específico de un énfasis o especializado' };

  const S = {
    modelo: null, pos: null, origen: 'ejemplo',
    profundidad: 0, objetivoProf: 0.6,
    giroX: 0, giroY: 0, giroXT: 0, giroYT: 0, velX: 0, velY: 0,
    sel: null, fCon: 1, vaiven: !reducir,
    tema: 'oscuro', w: 0, h: 0, dpr: 1, escala: 1,
    gM: null, gB: null, gN: null, tM: null, tB: null, tN: null,
    sx: null, sy: null, sp: null, sz: null,
    zoom: 1, zoomT: 1, panX: 0, panY: 0, panXT: 0, panYT: 0,
    kZoom: 0.22, kEsc: 0, atr: null, atrT: null, hubMiembros: [], hubClave: null, programaEscuela: null, autoZoomAt: 0, autoZoomIdx: [], etqLineas: [],
    estilo: 'brillante', temaPrevio: null, modoUnion: 'estela', filtro: new Set(), nivel: null, convMapa: new Map(),
    lineas: [], sucio: true, fase: null, enlaces: [],
    // capas: qué se ve (maestrias | enfasis | especializados | cursos) y qué conexiones están encendidas
    capa: 'maestrias', capaBurbuja: 'maestrias', maestriaSel: null, bloqueSel: null, troncalesSel: false, gr: null, bur: new Map(),
    ab: null, vN: null, vM: null, vB: null, sx0: null, sy0: null, gcx: null, gcy: null,
    desv: { x: 0, y: 0 }, pivote: null, pivoteUlt: null, pinK: 0, ruta: null, rutaK: 0, ladoEt: new Map(), zonasEt: [], hoverB: null, hoverP: null, focoHalo: null, focoB: null, focoTipo: '', kEscG: 0, progEscG: null, hk: 0, foco: null, focoClave: '', av: null, pfx: null, pfy: null, blendHasta: 0, mostrarCon: false, lineasConstitucion: false,
  };

  /* ------------------------------ tema ------------------------------------ */
  function colorRgb(c) {
    const el = colorRgb.el || (colorRgb.el = document.createElement('canvas').getContext('2d'));
    el.fillStyle = '#070707'; el.fillStyle = c;
    const v = el.fillStyle;
    if (v[0] === '#') { const n = parseInt(v.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; }
    const m = v.match(/[\d.]+/g) || [255, 255, 255];
    return [+m[0], +m[1], +m[2]];
  }
  const anchosEtiqueta = new Map();
  const capaCalor = document.createElement('canvas'), capaCtx = capaCalor.getContext('2d');
  const sprites = new Map();
  // color de los troncales de toda la escuela: blanco -> color de la maestría en foco, según qué tan cerca van de ella
  function mezclarColor(a, b, t) {
    const A = colorRgb(a), B = colorRgb(b), q = Math.round(Math.max(0, Math.min(1, t)) * 16) / 16;   // pasos de 1/16: pocos sprites distintos
    const h = (k) => Math.round(A[k] + (B[k] - A[k]) * q).toString(16).padStart(2, '0');
    return '#' + h(0) + h(1) + h(2);
  }
  function colorEscuelaVivo() {
    const pr = S.programaEscuela && S.modelo ? S.modelo.programaPorId.get(S.programaEscuela) : null;
    if (pr) return mezclarColor(colorEscuela(), pr.color, S.kEsc);
    // en los niveles de grupos (Maestrías, Énfasis...) con una maestría elegida, los troncales también toman su color
    const pg = S.progEscG && S.modelo ? S.modelo.programaPorId.get(S.progEscG) : null;
    return pg && S.kEscG > 0.01 ? mezclarColor(colorEscuela(), pg.color, S.kEscG) : colorEscuela();
  }
  function actualizarColorTroncales() {
    const quiere = S.capa !== 'cursos' && S.maestriaSel && !S.troncalesSel;
    if (quiere) S.progEscG = S.maestriaSel;
    const meta = quiere ? 1 : 0, d = meta - S.kEscG;
    if (Math.abs(d) > 0.004) { S.kEscG += d * (reducir ? 1 : 0.1); S.sucio = true; } else if (d) { S.kEscG = meta; S.sucio = true; if (!meta) S.progEscG = null; }
  }
  function sprite(color) {
    const clave = S.tema + color;
    if (sprites.has(clave)) return sprites.get(clave);
    const c = document.createElement('canvas'); c.width = c.height = 128;
    const g = c.getContext('2d');
    const [r, gg, b] = colorRgb(color);
    const gr = g.createRadialGradient(64, 64, 0, 64, 64, 64);
    gr.addColorStop(0, 'rgba(' + r + ',' + gg + ',' + b + ',1)');
    gr.addColorStop(0.3, 'rgba(' + r + ',' + gg + ',' + b + ',0.5)');
    gr.addColorStop(0.65, 'rgba(' + r + ',' + gg + ',' + b + ',0.14)');
    gr.addColorStop(1, 'rgba(' + r + ',' + gg + ',' + b + ',0)');
    g.fillStyle = gr; g.fillRect(0, 0, 128, 128);
    sprites.set(clave, c);
    return c;
  }
  const colorEscuela = () => COLOR_ESCUELA[S.tema];
  const colorNodo = (c) => c;
  // el color de cada programa depende del tema (ambos de la paleta de marca)
  function aplicarPaleta() {
    if (!S.modelo) return;
    S.modelo.programas.forEach((p) => { if (p.colores) p.color = p.colores[S.tema]; });
  }

  /* ------------------------------ datos ----------------------------------- */
  function aplicarTablas(T, origen) {
    const modelo = construirModelo(T);
    S.modelo = modelo; S.origen = origen;
    aplicarPaleta();
    if (modelo.cursos.length) {
      S.posPlano = simular(modelo, 'plano'); S.posDenso = simular(modelo, 'denso');
      const nn = modelo.cursos.length;
      S.pos = { X: new Float64Array(nn), Y: new Float64Array(nn), Z: new Float64Array(nn), R: S.posPlano.R };
      mezclarPos(0);
    } else S.pos = null;
    const n = modelo.cursos.length;
    S.gM = new Float32Array(n).fill(1); S.gB = new Float32Array(n).fill(1); S.gN = new Float32Array(n).fill(1);
    S.tM = new Float32Array(n).fill(1); S.tB = new Float32Array(n).fill(1); S.tN = new Float32Array(n).fill(1);
    S.sx = new Float32Array(n); S.sy = new Float32Array(n); S.sp = new Float32Array(n); S.sz = new Float32Array(n);
    S.sx0 = new Float32Array(n); S.sy0 = new Float32Array(n);
    S.vN = new Float32Array(n); S.vM = new Float32Array(n); S.vB = new Float32Array(n);
    S.ab = new Float32Array(n);   // 0 = el curso va compacto con su grupo (dentro del halo), 1 = suelto, con su nombre
    S.av = new Float32Array(n).fill(1); S.pfx = new Float32Array(n); S.pfy = new Float32Array(n); S.pfOk = false;   // av: ¿se ve el punto? (en «programas especializados» solo los que aplican)
    S.capa = 'maestrias'; S.capaBurbuja = 'maestrias'; S.maestriaSel = null; S.bloqueSel = null; S.troncalesSel = false; S.gr = null; S.bur = new Map(); S.ladoEt = new Map();
    S.filtro = new Set(); S.nivel = new Int8Array(n);
    S.atr = new Float32Array(n); S.atrT = new Float32Array(n); S.hubRecuerdo = []; S.hubMiembros = []; S.hubClave = null; S.programaEscuela = null; S.autoZoomAt = 0; S.autoZoomIdx = [];
    aplicarFiltroBase(); S.gM.set(S.tM); S.gB.set(S.tB); S.gN.set(S.tN);
    S.fase = Float32Array.from({ length: n }, (_, i) => (i * 2.399) % 6.283);
    S.sel = null; S.lineas = [];
    S.enlaces = construirEnlaces(modelo);
    // cursos más conectados (por sus enlaces, electivas y programas): son los que llevan nombre a la vista
    S.grado = new Float32Array(n);
    S.enlaces.forEach((e) => { S.grado[e.a] += e.tipo === 'cadena' || e.tipo === 'correq' ? 2 : 1; S.grado[e.b] += e.tipo === 'cadena' || e.tipo === 'correq' ? 2 : 1; });
    modelo.cursos.forEach((c, i) => { S.grado[i] += c.electivaDestinos.length * 1.5 + (c.progs.length > 1 ? 3 : 0); });
    S.ordImp = Array.from({ length: n }, (_, i) => i).sort((a, b) => S.grado[b] - S.grado[a]);
    $('vacio').style.display = n ? 'none' : 'flex';
    pintarLeyenda(); pintarValidaciones(); pintarBuscador(); pintarTextoInicial();
    $('btnEjemplo').classList.toggle('oculto', origen === 'ejemplo');
    $('seccionEjemplo').classList.toggle('oculto', origen !== 'ejemplo');
    actualizarSubtitulo();
    S.sucio = true;
  }
  // Conexiones siempre visibles: malla dentro de cada bloque y cadenas de prerrequisitos
  function construirEnlaces(M) {
    const res = [], vistos = new Set();
    const poner = (a, b, tipo, prog) => {
      if (a == null || b == null || a === b) return;
      const k = (a < b ? a + '|' + b : b + '|' + a) + tipo;
      if (vistos.has(k)) return; vistos.add(k);
      res.push({ a, b, tipo, prog });
    };
    M.bloques.forEach((b) => {
      const mem = b.cursos.map((id) => M.porId.get(id));
      for (let i = 0; i < mem.length; i++) for (let j = i + 1; j < mem.length; j++) poner(mem[i], mem[j], 'bloque', b.maestriaId);
    });
    M.cursos.forEach((c, i) => c.prereq.forEach((q) => poner(M.porId.get(q), i, 'cadena', c.progs[0])));
    M.cursos.forEach((c, i) => c.coreq.forEach((q) => { const j = M.porId.get(q); if (j > i) poner(i, j, 'correq', c.progs[0]); }));
    return res;
  }
  // La descripción general de la oferta vive en el encabezado, no en el detalle
  function actualizarSubtitulo() {
    const M = S.modelo, dato = S.origen === 'ejemplo' ? 'Datos de ejemplo. ' : '';
    const reales = M.bloques.filter((b) => !b.implicito && b.cursos.length);
    const esp = reales.filter((b) => b.otorgaEsp).length, enf = reales.filter((b) => b.esEnfasis).length;
    $('subtitulo').textContent = dato + M.cursos.length + ' cursos, ' + M.maestrias.length + ' maestrías' + (M.doctorados.length ? ' y ' + M.doctorados.length + ' doctorado' : '') +
      '. Hay ' + enf + ' énfasis y ' + esp + ' bloques que otorgan un título de especialista.';
  }
  function paramsEjemplo() {
    return {
      nMaestrias: +$('cNMaestrias').value, conEnfasis: +$('cConEnfasis').value,
      enfasisPor: $('cEnfasisPor').value, proyectos: $('cProyectos').value,
      compartidos: +$('cCompartidos').value, pctElectivas: +$('cPctElectivas').value,
      doctorado: $('cDoctorado').checked,
    };
  }
  function regenerarEjemplo() {
    const p = paramsEjemplo();
    $('oNMaestrias').textContent = p.nMaestrias;
    $('oConEnfasis').textContent = Math.min(p.conEnfasis, p.nMaestrias);
    $('oCompartidos').textContent = p.compartidos;
    $('oPctElectivas').textContent = p.pctElectivas + ' %';
    aplicarTablas(generarEjemplo(p), 'ejemplo');
  }

  /* ------------------------- leyenda y textos ----------------------------- */
  function pintarLeyenda() {
    const cont = $('leyenda'), barra = $('barraNiveles'); const M = S.modelo;
    $('btnVolver').classList.toggle('oculto', !(S.maestriaSel || S.troncalesSel));   // «Volver» vive en el encabezado
    const foco = document.activeElement && (cont.contains(document.activeElement) || barra.contains(document.activeElement)) ? document.activeElement.dataset.k : null;
    cont.textContent = ''; barra.textContent = '';
    S.convMapa = new Map();
    const boton = (clave, texto, clase) => { const b = el('button', clase); b.type = 'button'; b.dataset.k = clave; b.appendChild(el('span', 'txt', texto)); return b; };

    // 1) ¿Qué quieres ver? Niveles exclusivos; solo aparecen los que aplican a estos datos
    const reales = M.bloques.filter((b) => !b.implicito && b.cursos.length);
    const niveles = [['maestrias', 'Maestrías']];
    // los énfasis ya no son un nivel aparte: se entra a ellos desde su maestría (ver seleccionarMaestria)
    if (reales.some((b) => b.otorgaEsp)) niveles.push(['especializados', 'Programas especializados']);
    niveles.push(['cursos', 'Cursos']);
    niveles.forEach(([k, t]) => {
      const b = boton('capa:' + k, t, 'nivel'); b.setAttribute('role', 'radio');
      b.setAttribute('aria-checked', S.capa === k || (k === 'maestrias' && S.capa === 'enfasis') ? 'true' : 'false');
      b.addEventListener('click', () => { if (k === 'maestrias' && S.capa === 'enfasis') volverATodas(); else fijarCapa(k); });
      barra.appendChild(b);
    });

    // 3) Leyenda jerárquica: solo lo que está encendido en este nivel (maestría y, sangrados, sus énfasis)
    const s3 = el('div', 'seccion-capas leyenda-j');
    const lista = el('ul', 'lista-leyenda');
    const propios = (id) => M.cursos.filter((c) => !c.escuela && c.progs.includes(id)).length;
    // todo plegado salvo la maestría elegida: solo ella muestra sus énfasis o programas; las demás quedan con su nombre
    const mostrarHijos = (p) => S.maestriaSel === p.id;
    M.maestrias.concat(M.doctorados).forEach((p) => {
      const li = el('li');
      const b = boton('p:' + p.id, p.nombre, 'item-ley' + ((S.maestriaSel && S.maestriaSel !== p.id) || S.troncalesSel ? ' tenue' : ''));
      b.setAttribute('aria-pressed', S.maestriaSel === p.id && !S.bloqueSel ? 'true' : 'false'); b.style.setProperty('--c', colorNodo(p.color));
      b.title = S.maestriaSel === p.id ? 'Quitar la selección' : 'Abrir esta maestría';
      const pt = el('span', 'punto'); pt.style.background = colorNodo(p.color); b.insertBefore(pt, b.firstChild);
      b.addEventListener('click', () => seleccionarMaestria(p.id));
      li.appendChild(b);
      S.convMapa.set('p:' + p.id, { texto: p.nombre, color: colorNodo(p.color) });
      if (mostrarHijos(p)) {
        const hijos = reales.filter((bl) => bl.maestriaId === p.id && (S.capa !== 'especializados' || bl.otorgaEsp));
        if (hijos.length) {
          const ul = el('ul', 'sub-leyenda');
          hijos.forEach((bl) => {
            const h = el('li', 'hijo-li');
            const cual = S.bloqueSel ? (S.bloqueSel === bl.id) : null;
            const bb = boton('b:' + bl.id, S.capa === 'especializados' ? 'Programa especializado en ' + bl.nombre : (bl.esEnfasis ? 'Énfasis en ' : '') + bl.nombre, 'item-ley hijo' + (S.maestriaSel && S.maestriaSel !== p.id ? ' tenue' : (cual === false ? ' tenue' : '')));
            bb.setAttribute('aria-pressed', cual ? 'true' : 'false'); bb.style.setProperty('--c', colorNodo(p.color));
            bb.title = cual ? 'Quitar la selección' : 'Ver solo este programa';
            const q = el('span', 'punto chico'); q.style.background = colorNodo(p.color); bb.insertBefore(q, bb.firstChild);
            bb.addEventListener('click', () => seleccionarBloque(bl.id));
            h.appendChild(bb);
            ul.appendChild(h);
          });
          li.appendChild(ul);
        }
      }
      lista.appendChild(li);
    });
    const nEsc = M.cursos.filter((c) => c.escuela).length;
    if (nEsc && S.capa !== 'especializados') {
      const li = el('li');
      const f = boton('e', 'Cursos obligatorios', 'item-ley' + (S.bloqueSel ? ' tenue' : ''));
      f.setAttribute('aria-pressed', S.troncalesSel ? 'true' : 'false'); f.style.setProperty('--c', colorEscuela()); f.title = S.troncalesSel ? 'Quitar la selección' : 'Ver solo los cursos obligatorios';
      const pt = el('span', 'punto'); pt.style.background = colorEscuela(); f.insertBefore(pt, f.firstChild);
      f.addEventListener('click', seleccionarTroncales);
      li.appendChild(f); lista.appendChild(li);
    }
    s3.appendChild(lista);
    // guía de las líneas que están encendidas
    const act = [['secuencia', 'Primero uno, luego el otro (flecha)'], ['correq', 'Se cursan juntos'], 
S.modelo.doctorados.length && ['doctorado', 'El doctorado sale de su maestría']].filter(Boolean);
    if (S.mostrarCon && act.length) {
      const NS = 'http://www.w3.org/2000/svg';
      const muestra = (tipo) => {
        const svg = document.createElementNS(NS, 'svg'); svg.setAttribute('width', '38'); svg.setAttribute('height', '12'); svg.setAttribute('viewBox', '0 0 38 12'); svg.setAttribute('aria-hidden', 'true');
        const linea = (y, dash, ancho) => { const l = document.createElementNS(NS, 'line'); l.setAttribute('x1', '1'); l.setAttribute('x2', '37'); l.setAttribute('y1', y); l.setAttribute('y2', y); l.setAttribute('stroke', 'currentColor'); l.setAttribute('stroke-width', ancho || 1.6); l.setAttribute('stroke-linecap', 'round'); if (dash) l.setAttribute('stroke-dasharray', dash); svg.appendChild(l); };
        if (tipo === 'constitucion') linea(6);
        else if (tipo === 'electiva') linea(6, '1.5 4');
        else if (tipo === 'secuencia') { linea(6); const f = document.createElementNS(NS, 'polygon'); f.setAttribute('points', '37,6 29,2.5 29,9.5'); f.setAttribute('fill', 'currentColor'); svg.appendChild(f); }
        else if (tipo === 'correq') { linea(3.5, null, 1.2); linea(8.5, null, 1.2); }
        else linea(6, '7 4');
        return svg;
      };
      const gl = el('div', 'guia-lineas');
      act.forEach(([tipo, texto]) => { const f = el('div', 'conv-tam'), caja = el('span', 'conv-linea-caja'); caja.appendChild(muestra(tipo)); f.appendChild(caja); f.appendChild(el('span', '', texto)); gl.appendChild(f); });
      s3.appendChild(gl);
    }
    cont.appendChild(s3);
    medirBarra();
    if (foco) { const nuevo = document.querySelector('#colIzq [data-k="' + foco + '"], #barraNiveles [data-k="' + foco + '"]'); if (nuevo) nuevo.focus({ preventScroll: true }); }
  }
  // Cambiar de nivel: las burbujas se dividen (o se recogen) con la misma animación en los dos sentidos
  function fijarCapa(c) {
    S.volverCapa = null;
    if (S.capa === c) return;
    S.capa = c;
    if (c !== 'cursos') { S.capaBurbuja = c; S.gr = null; }
    S.blendHasta = performance.now() + 1300;
    S.sel = null;
    aplicarFiltro();
  }
  // Elegir una maestría (desde una burbuja o la leyenda): las demás se apagan y la elegida se abre
  function seleccionarMaestria(id) {
    const ya = S.maestriaSel === id && !S.bloqueSel;
    S.maestriaSel = ya ? null : id; S.bloqueSel = null; S.troncalesSel = false; S.ruta = null; S.rutaAviso = '';
    S.filtro.clear(); if (S.maestriaSel) S.filtro.add('p:' + id);
    S.sel = null;
    // desde «Maestrías», elegir una maestría la divide en sus énfasis y pasa a ese nivel; «Volver» regresa a Maestrías
    if (S.maestriaSel && S.capa === 'maestrias') {
      S.capa = 'enfasis'; S.capaBurbuja = 'enfasis'; S.gr = null; S.volverCapa = 'maestrias'; S.blendHasta = performance.now() + 1300;
    }
    aplicarFiltro();
  }
  // Elegir un énfasis o programa especializado: solo se abre ese; el resto de su maestría se apaga
  function seleccionarBloque(bid) {
    const b = S.modelo.bloquePorId.get(bid); if (!b) return;
    const ya = S.bloqueSel === bid;
    S.bloqueSel = ya ? null : bid; S.maestriaSel = b.maestriaId; S.troncalesSel = false;
    S.filtro.clear(); S.filtro.add(ya ? 'p:' + b.maestriaId : 'b:' + bid);
    S.sel = null;
    aplicarFiltro();
  }
  // Cursos troncales (los de toda la escuela): se eligen como una maestría; solo ellos quedan abiertos
  function seleccionarTroncales() {
    const ya = S.troncalesSel;
    S.troncalesSel = !ya; S.maestriaSel = null; S.bloqueSel = null;
    S.filtro.clear(); if (S.troncalesSel) S.filtro.add('e');
    S.sel = null;
    aplicarFiltro();
  }
  function volverATodas() {
    S.maestriaSel = null; S.bloqueSel = null; S.troncalesSel = false; S.filtro.clear(); S.sel = null;
    if (S.volverCapa) { S.capa = S.volverCapa; S.capaBurbuja = S.volverCapa; S.gr = null; S.volverCapa = null; S.blendHasta = performance.now() + 1300; }
    aplicarFiltro();
  }
  function alternarFiltro(clave) {
    const estaba = S.filtro.has(clave);
    S.filtro.clear();            // selección única: tocar otro cambia el filtro; tocar el mismo lo apaga
    if (!estaba) S.filtro.add(clave);
    aplicarFiltro();
  }
  function actualizarNivel() {
    const M = S.modelo, n = M.cursos.length;
    S.nivel = new Int8Array(n);
    S.filtro.forEach((k) => {
      const id = k.slice(2);
      for (let j = 0; j < n; j++) {
        const o = M.cursos[j]; let v = 0;
        if (k === 'e') v = o.escuela ? 2 : 0;
        else if (k[0] === 'p') {
          const pr = M.programaPorId.get(id);
          if (!pr) continue;
          if (o.escuela) v = pr.nivel === 'maestria' || pr.nivel === 'doctorado' ? 1 : 0;   // el doctorado incluye su maestría base, y con ella la escuela
          else if (o.progs.includes(id)) v = 2;
          else if (pr.nivel === 'doctorado' && o.progs.includes(pr.maestriaBase)) v = 2; // el doctorado incluye toda su maestría base
        } else if (k[0] === 'b') {
          const b = M.bloquePorId.get(id);
          if (b && b.cursos.includes(o.id)) v = 2;
        }
        if (v > S.nivel[j]) S.nivel[j] = v;
      }
    });
  }
  function aplicarFiltroBase() {
    const n = S.modelo.cursos.length, f = S.filtro.size > 0;
    for (let j = 0; j < n; j++) {
      const v = S.nivel[j];
      // sin filtro ni selección todo está apagado (sin brillo de clusters); se enciende lo que se toca
      S.tM[j] = !f ? 0 : v === 2 ? 1.1 : v === 1 ? 0.5 : 0.04;
      S.tB[j] = !f ? 0 : v === 2 ? 1.6 : v === 1 ? 0.4 : 0.05;
      S.tN[j] = !f ? 0.5 : v === 2 ? 1 : v === 1 ? 0.75 : 0.1;
    }
    S.lineas = [];
  }
  // Relación de CONSTITUCIÓN: al encender un programa (maestría, énfasis, especializado o doctorado) sus cursos se
  // acercan a su punto central y se unen a él con línea continua. Los troncales de toda la escuela toman el color de la
  // maestría y se acercan a ella para dar a entender que van juntos. Después se hace zoom automático sobre el grupo.
  function configurarConstitucion() {
    const M = S.modelo;
    S.atrT.fill(0); S.hubMiembros = []; S.hubClave = null; S.autoZoomIdx = []; S.autoZoomAt = 0; S.autoZoomAt2 = 0; S.centroPropios = null;   // programaEscuela se suelta en ciclo(), cuando la escuela ya volvió
    if (S.filtro.size !== 1) { return; }
    const k = [...S.filtro][0], id = k.slice(2);
    const propios = [], escuelas = [];
    if (k[0] === 'p') {
      if (S.capa !== 'cursos') {   // en los niveles de grupos la maestría se divide en énfasis, sin atraer ni acercar; solo el centro pasa a ella
        const mem = []; M.cursos.forEach((c, i) => { if (!c.escuela && c.progs.includes(id)) mem.push(i); });
        if (mem.length) { S.centrarIdx = mem; S.autoZoomAt = performance.now() + 450; }
        return;
      }
      const pr = M.programaPorId.get(id); if (!pr) return;
      M.cursos.forEach((c, i) => {
        if (c.escuela) { if (pr.nivel === 'maestria' || pr.nivel === 'doctorado') escuelas.push(i); }
        else if (c.progs.includes(id)) propios.push(i);
      });
      if (pr.nivel === 'maestria' || pr.nivel === 'doctorado') S.programaEscuela = id;
    } else if (k[0] === 'b') {
      const b = M.bloquePorId.get(id); if (!b) return;
      b.cursos.forEach((cid) => { const i = M.porId.get(cid); if (i != null) propios.push(i); });
    } else if (k === 'e') {
      const esc = []; M.cursos.forEach((c, i) => { if (c.escuela) esc.push(i); });
      if (esc.length) { S.autoZoomIdx = esc; S.autoZoomAt = performance.now() + 450; }
      return;
    } else return;
    if (!propios.length) return;
    S.hubClave = k; S.hubMiembros = propios; S.hubRecuerdo = propios; S.centroPropios = k[0] === 'p' ? propios : null;
    propios.forEach((i) => { S.atrT[i] = 0.3; });
    escuelas.forEach((i) => { S.atrT[i] = 1; });   // los troncales de toda la escuela se acercan a la maestría elegida
    S.autoZoomIdx = propios.concat(escuelas);
    S.autoZoomAt = performance.now() + 450;
    S.autoZoomAt2 = performance.now() + 2000;   // los troncales tardan en llegar a su lado: se vuelve a centrar cuando ya están
  }
  function aplicarFiltro() {
    actualizarNivel();
    actualizarRuta();
    actualizarPivote();
    pintarLeyenda();
    configurarConstitucion();
    if (!S.filtro.size) { S.kZoom = 0.04; S.zoomT = 1; S.panXT = 0; S.panYT = 0; }   // al apagar, la vista vuelve completa
    seleccionar(S.sel);
  }
  // Ajusta zoom y desplazamiento para que el grupo encendido llene la vista (sin quedar tapado por las convenciones)
  function ajustarVista(idx, tope, foco, centroIdx) {
    if (!idx.length) return;
    S.kZoom = 0.06;
    const px = (i) => S.sx0[i] - S.desv.x, py = (i) => S.sy0[i] - S.desv.y;
    const med = (v) => { const w = v.slice().sort((a, b) => a - b); return w[Math.floor(w.length / 2)]; };
    const p90 = (v) => { const w = v.slice().sort((a, b) => a - b); return w[Math.min(w.length - 1, Math.floor(w.length * 0.9))]; };
    // centro y tamaño del grupo: la mediana y el 90 % más cercano, para que un curso compartido y lejano no corra el centro ni encoja el zoom
    let mx, my, hw, hh;
    if (foco != null) {   // el curso elegido queda en el centro y el zoom deja ver lo relacionado con él
      mx = px(foco); my = py(foco);
      hw = Math.max(30, ...idx.map((i) => Math.abs(px(i) - mx))); hh = Math.max(30, ...idx.map((i) => Math.abs(py(i) - my)));
    } else {
      const base = centroIdx && centroIdx.length ? centroIdx : idx;   // el centro es el de lo elegido; el tamaño incluye lo que lo acompaña
      mx = med(base.map(px)); my = med(base.map(py));
      hw = Math.max(30, p90(idx.map((i) => Math.abs(px(i) - mx)))); hh = Math.max(30, p90(idx.map((i) => Math.abs(py(i) - my))));
    }
    const C = centroLibre(), margen = 70;
    // alto aprovechable: entre el buscador y la barra de niveles
    let yTop = 0, yBot = S.h - 80;
    S.reservas = null; medirReservas().forEach((q) => { if (q.x0 < S.w / 2 && q.x1 > S.w / 2) { if ((q.y0 + q.y1) / 2 < S.h / 2) yTop = Math.max(yTop, q.y1); else yBot = Math.min(yBot, q.y0); } });
    const altoUtil = Math.max(160, yBot - yTop);
    const k = Math.min((S.w - margen * 2) / (2 * hw), (altoUtil - margen * 2) / (2 * hh));
    const zN = Math.max(0.8, Math.min(tope || 3.4, S.zoom * k));   // el zoom automático deja algo de contexto alrededor
    const vx = (mx - S.w / 2 - S.panX) / S.zoom, vy = (my - S.h / 2 - S.panY) / S.zoom;   // centro del grupo, sin zoom ni desplazamiento
    S.zoomT = zN;
    S.panXT = C.x - S.w / 2 - vx * zN;
    S.panYT = C.y - S.h / 2 - vy * zN;   // lo elegido queda justo en el centro de la pantalla
    limitarPan();
    S.sucio = true;
  }
  function pintarBuscador() {
    const dl = $('listaCursos'); dl.textContent = '';
    S.modelo.cursos.forEach((c) => { const o = document.createElement('option'); o.value = c.nombre; dl.appendChild(o); });
  }
  function el(tag, clase, texto) {
    const e = document.createElement(tag);
    if (clase) e.className = clase;
    if (texto != null) e.textContent = texto;
    return e;
  }
  function parrafo(etiqueta, valor) {
    const p = el('p'); p.appendChild(el('span', 'et', etiqueta + ' ')); p.appendChild(document.createTextNode(valor)); return p;
  }
  function pintarTextoInicial() {
    const M = S.modelo, t = $('texto'); t.textContent = '';
    if (!M.cursos.length) { t.appendChild(el('p', '', 'Sin cursos que mostrar.')); return; }
    const bloquesConCursos = M.bloques.filter((b) => b.cursos.length);
    const esp = bloquesConCursos.filter((b) => b.otorgaEsp).length;
    const enf = bloquesConCursos.filter((b) => b.esEnfasis).length;
    if (S.w <= 760) $('detalle').open = false;   // en pantallas chicas el detalle solo se abre con un curso
    const pm = S.maestriaSel && M.programaPorId.get(S.maestriaSel), bs = S.bloqueSel && M.bloquePorId.get(S.bloqueSel);
    $('resumenDetalle').textContent = bs ? bs.nombre : pm ? pm.nombre : 'Detalle';
    $('panel').hidden = !(bs || pm || S.troncalesSel);   // sin selección no hay nada que mostrar aquí
    if (S.troncalesSel) {
      const nE = M.cursos.filter((c) => c.escuela).length;
      $('resumenDetalle').textContent = 'Cursos obligatorios';
      t.appendChild(el('h2', '', 'Cursos obligatorios'));
      t.appendChild(el('p', '', nE + ' cursos que van con todas las maestrías: laboratorios, ICG y transversales. Cada estudiante cursa los que pide la ruta.'));
      t.appendChild(el('p', 'nota', 'Toca un curso para ver sus detalles, o «Volver a todas las maestrías» para ver todo otra vez.'));
      return;
    }
    if (bs) {
      t.appendChild(el('h2', '', bs.nombre));
      t.appendChild(el('p', '', 'Parte de ' + pm.nombre + '. ' + bs.cursos.length + ' cursos' + (bs.otorgaEsp ? ' · da el título «' + (bs.tituloEsp || 'Especialización en ' + bs.nombre) + '»' : '') + '. El resto de la maestría está apagado.'));
      t.appendChild(el('p', 'nota', 'Toca un curso para ver sus conexiones, o «Volver a todas las maestrías» para ver todo otra vez.'));
      return;
    }
    if (pm) {
      const lit = S.nivel.reduce((sum, v) => sum + (v === 2 ? 1 : 0), 0);
      t.appendChild(el('h2', '', pm.nombre));
      if (S.ruta && S.capa === 'cursos') { pintarRuta(t, pm); return; }
      t.appendChild(el('p', '', lit + ' cursos, repartidos en sus énfasis. Toca un énfasis para ver sus cursos, o elige «Cursos» para verlos todos.'));
      t.appendChild(el('p', 'nota', 'Para ver todas las maestrías otra vez, usa el botón «Volver a todas las maestrías».'));
      return;
    }
  }
  // Qué título da la combinación: maestría con énfasis, sin énfasis, y los programas especializados que completa
  function resultadoRuta(pm, completa) {
    const M = S.modelo, R = S.ruta;
    const disc = R.cats.find((c) => c.clave === 'disciplinar'), sel = disc ? disc.cand.filter((c) => R.set.has(c.i)) : [];
    const bids = [...new Set(sel.map((c) => c.bid || ''))];
    const bloque = disc && bids.length === 1 && bids[0] && sel.length === disc.req ? M.bloquePorId.get(bids[0]) : null;
    const especializados = M.bloques.filter((b) => b.maestriaId === pm.id && !b.implicito && b.otorgaEsp && b.cursos.length && b.cursos.every((cid) => { const j = M.porId.get(cid); return j != null && R.set.has(j); }));
    let principal = null;
    if (completa) principal = bloque && bloque.esEnfasis ? pm.nombre + ', énfasis en ' + bloque.nombre : pm.nombre + ' sin énfasis';
    return { principal, especializados: especializados.map((b) => 'Programa especializado en ' + b.nombre) };
  }
  // Panel «Combinación para titularse»: explica los cursos con nombre, permite armar otra a mano (tocar un curso lo agrega o quita; no deja pasarse de lo permitido) o al azar
  function pintarRuta(t, pm) {
    const M = S.modelo, R = S.ruta, esperadoN = (pm.ruta && pm.ruta.esperadoN) || 12, esperadoC = (pm.ruta && pm.ruta.esperadoC) || 36;
    let cred = 0; R.set.forEach((i) => { cred += M.cursos[i].creditos; });
    const n = R.set.size, completa = n === esperadoN && cred === esperadoC;
    t.appendChild(el('p', 'rol', n + ' de ' + esperadoN + ' cursos · ' + cred + ' de ' + esperadoC + ' créditos'));
    const barra = el('div', 'progreso'); const rel = el('div', completa ? 'lleno' : ''); rel.style.width = Math.min(100, Math.round(100 * cred / esperadoC)) + '%'; barra.appendChild(rel); t.appendChild(barra);
    // resultado: el título que da esta combinación
    const res = resultadoRuta(pm, completa), caja = el('div', 'resultado' + (completa ? ' completa' : ''));
    caja.setAttribute('aria-live', 'polite');
    caja.appendChild(el('h3', '', 'Resultado'));
    caja.appendChild(el('p', 'titulo-res', res.principal ? res.principal : 'Todavía no completa la maestría'));
    if (!res.principal) caja.appendChild(el('p', 'nota', n < esperadoN ? 'Faltan ' + (esperadoN - n) + ' cursos para llegar a ' + esperadoC + ' créditos.' : 'Revisa la combinación: no suma los créditos del título.'));
    res.especializados.forEach((x) => caja.appendChild(el('p', 'esp-res', '+ ' + x + ' (12 créditos)')));
    t.appendChild(caja);
    if (S.rutaAviso) t.appendChild(el('p', 'aviso-ruta', S.rutaAviso));
    t.appendChild(el('p', '', 'Los cursos con aro y nombre en el mapa son los de esta combinación; los demás quedan tenues y su nombre sale al pasar el cursor. Toca un curso de las listas para agregarlo o quitarlo. Puedes mezclar cursos de distintos énfasis: el énfasis no es obligatorio.'));
    const fila = el('div', 'fila-ruta');
    const rehacer = (set, aviso) => { const sc = $('panel').scrollTop; R.set = set; S.rutaAviso = aviso || ''; pintarTextoInicial(); $('panel').scrollTop = sc; S.sucio = true; };
    const bAzar = el('button', 'boton-ruta', 'Combinación al azar'); bAzar.type = 'button'; bAzar.addEventListener('click', () => rehacer(combinacion(R.cats, true)));
    const bVaciar = el('button', 'boton-ruta suave', 'Empezar de cero'); bVaciar.type = 'button'; bVaciar.addEventListener('click', () => rehacer(new Set()));
    fila.appendChild(bAzar); fila.appendChild(bVaciar); t.appendChild(fila);
    R.cats.forEach((cat) => {
      const elegidos = cat.cand.filter((c) => R.set.has(c.i)), marcados = elegidos.length;
      const sec = el('div', 'seccion'); sec.appendChild(el('h3', '', cat.t + ' · ' + marcados + ' de ' + cat.req));
      let ultimo = null; const ul = el('ul', 'lista-ruta');
      cat.cand.forEach((c) => {
        if (c.sub && c.sub !== ultimo) { ultimo = c.sub; ul.appendChild(el('li', 'sub-ruta', c.sub)); }
        const on = R.set.has(c.i);
        const motivo = on ? '' : marcados >= cat.req ? 'Ya completaste ' + cat.t + ' (' + cat.req + '). Quita uno para cambiarlo.' : '';
        const li = el('li'); const b = el('button', 'opcion-ruta' + (motivo ? ' bloqueado' : '')); b.type = 'button'; b.setAttribute('aria-pressed', on ? 'true' : 'false');
        if (motivo) { b.setAttribute('aria-disabled', 'true'); b.title = motivo; }
        b.appendChild(el('span', 'marca', on ? '✓' : (motivo ? '–' : '+'))); b.appendChild(el('span', '', M.cursos[c.i].nombre));
        b.addEventListener('click', () => {
          if (motivo) { rehacer(R.set, motivo); return; }   // no deja pasarse de lo permitido; explica por qué
          const nuevo = new Set(R.set); if (on) nuevo.delete(c.i); else nuevo.add(c.i); rehacer(nuevo);
        });
        li.appendChild(b); ul.appendChild(li);
      });
      sec.appendChild(ul); t.appendChild(sec);
    });
  }
  const nombreProg = (id) => (S.modelo.programaPorId.get(id) || {}).nombre || id;
  const ROL_PLURAL = { laboratorio: 'laboratorio', icg: 'ICG', transversal: 'transversal' };
  function chip(texto, color, opciones) {
    opciones = opciones || {};
    const e = el(opciones.curso != null ? 'button' : 'span', 'chip' + (opciones.tenue ? ' tenue' : ''));
    if (opciones.curso != null) { e.type = 'button'; e.addEventListener('click', () => { elegirCurso(opciones.curso); }); }
    if (color) { const p = el('span', 'punto'); p.style.color = color; e.appendChild(p); }
    e.appendChild(document.createTextNode(texto));
    if (opciones.titulo) e.title = opciones.titulo;
    return e;
  }
  function seccionChips(contenedor, titulo, chips) {
    if (!chips.length) return;
    const sec = el('div', 'seccion');
    sec.appendChild(el('h3', '', titulo));
    const fila = el('div', 'chips'); chips.forEach((c) => fila.appendChild(c)); sec.appendChild(fila);
    contenedor.appendChild(sec);
  }
  // Lista con punto de color (o enlace a otro curso) bajo un título; es el formato de todas las secciones del detalle
  // ir a: abren lo mismo que tocar su burbuja, pero nunca lo apagan si ya está elegido
  const irAMaestria = (id) => { if (!(S.maestriaSel === id && !S.bloqueSel)) seleccionarMaestria(id); };
  const irABloque = (bid) => { if (S.bloqueSel !== bid) seleccionarBloque(bid); };
  const irATroncales = () => { if (!S.troncalesSel) seleccionarTroncales(); };
  function seccionLista(contenedor, titulo, items) {
    if (!items.length) return;
    const sec = el('div', 'seccion');
    sec.appendChild(el('h3', '', titulo));
    const ul = el('ul', 'lista-detalle');
    items.forEach((it) => {
      const li = el('li');
      if (it.curso != null || it.ir) {
        // enlace: a otro curso o a una maestría, énfasis o programa (se abre como si se hubiera tocado su burbuja)
        const b = el('button', 'enlace'); b.type = 'button';
        if (it.color) { const pt = el('span', 'punto'); pt.style.background = it.color; b.appendChild(pt); }
        b.appendChild(el('span', 'txt', it.texto));
        b.addEventListener('click', () => (it.ir ? it.ir() : elegirCurso(it.curso)));
        li.appendChild(b);
      } else {
        if (it.color) { const pt = el('span', 'punto'); pt.style.background = it.color; li.appendChild(pt); }
        li.appendChild(el('span', '', it.texto));
      }
      ul.appendChild(li);
    });
    sec.appendChild(ul); contenedor.appendChild(sec);
  }
  // Ficha de un curso: nombre, rol y créditos, descripción, oferta, prerrequisito, a qué pertenece y dónde puede ser electivo
  function describir(i) {
    const M = S.modelo, c = M.cursos[i], t = $('texto'); t.textContent = '';
    $('panel').hidden = false;
    if (S.w <= 760) $('detalle').open = true;
    const colorProg = (pid) => colorNodo((M.programaPorId.get(pid) || { color: '#EFEEE8' }).color);
    $('resumenDetalle').textContent = c.nombre;
    t.appendChild(el('h2', '', c.nombre));
    const obligatorio = c.rels.some((r) => r.rol !== 'electiva' && r.obligatorio);
    const rol = c.escuela ? 'Obligatorio' : obligatorio ? 'Obligatorio' : (c.electivaDestinos.length || c.electivaProgs.size) ? 'Electivo' : 'Obligatorio';
    t.appendChild(el('p', 'rol', rol + ' - ' + c.creditos + 'crd'));
    if (c.pool && c.pool.total > 1) t.appendChild(el('p', 'nota', 'Una de ' + c.pool.total + ' opciones de ' + ROL_PLURAL[c.pool.rol] + '; cada estudiante cursa ' + c.pool.pide + '.'));
    // descripción: texto libre; las líneas que empiezan con -, •, * o «1.» se muestran como lista de opciones
    if (c.descripcion) {
      let ul = null;
      c.descripcion.split(/\r?\n/).map(txt).filter(Boolean).forEach((linea) => {
        const m = linea.match(/^(?:[-•*–·]|\d+[.)])\s+(.+)$/);
        if (m) { if (!ul) { ul = el('ul', 'oferta'); t.appendChild(ul); } ul.appendChild(el('li', '', m[1])); }
        else { ul = null; t.appendChild(el('p', 'descripcion', linea)); }
      });
    }
    if (c.oferta && c.oferta.length) {
      const sec = el('div', 'seccion');
      sec.appendChild(el('h3', '', 'Oferta de cursos'));
      const ul = el('ul', 'oferta');
      c.oferta.forEach((o) => {
        const li = el('li', '', o);
        const d = (c.ofertaDesc.find((x) => x.nombre === o) || {}).descripcion;
        if (d) li.appendChild(el('span', 'oferta-desc', d));
        ul.appendChild(li);
      });
      sec.appendChild(ul); t.appendChild(sec);
    }
    // prerrequisito y correquisito: enlaces a esos cursos
    const aCursos = (ids) => ids.map((q) => M.porId.get(q)).filter((j) => j != null && j !== i).map((j) => ({ texto: M.cursos[j].nombre, curso: j }));
    seccionLista(t, 'Prerrequisito', aCursos(c.prereq));
    seccionLista(t, 'Se cursa junto con', aCursos(c.coreq));
    // hace parte de: maestría, énfasis y programa especializado
    const parte = [];
    if (c.escuela) parte.push({ texto: 'Todas las maestrías', color: colorEscuela(), ir: irATroncales });
    else {
      c.progs.forEach((pid) => { const pr = M.programaPorId.get(pid); if (pr) parte.push({ texto: pr.nombre, color: colorProg(pid), ir: () => irAMaestria(pid) }); });
      c.bloquesReales.forEach((bid) => {
        const b = M.bloquePorId.get(bid); if (!b) return;
        parte.push({ texto: (b.esEnfasis ? 'Énfasis: ' : 'Núcleo: ') + b.nombre, color: colorProg(b.maestriaId), ir: () => irABloque(bid) });
        if (b.otorgaEsp) parte.push({ texto: 'Programa especializado: ' + (b.tituloEsp || 'Especialización en ' + b.nombre), color: colorProg(b.maestriaId), ir: () => irABloque(bid) });
      });
    }
    seccionLista(t, 'Hace parte de', parte);
    seccionLista(t, 'Puede ser electivo en', c.electivaDestinos.map((d) => ({ texto: (d.tipo === 'bloque' ? 'Énfasis: ' : '') + d.nombre, color: colorProg(d.maestria), ir: () => (d.tipo === 'bloque' ? irABloque(d.id) : irAMaestria(d.id)) })));
    if (!c.electivaProgs.size && !c.escuela && !['laboratorio', 'icg', 'transversal'].includes(c.tipo)) t.appendChild(el('p', 'nota', 'No está abierto como electivo en otras maestrías.'));
  }

  function pintarValidaciones() {
    const ul = $('listaAdvertencias'); ul.textContent = '';
    const iss = S.modelo.issues;
    const nAv = iss.filter((x) => x.nivel === 'advertencia').length;
    iss.slice().sort((a, b) => (a.nivel === b.nivel ? 0 : a.nivel === 'advertencia' ? -1 : 1)).forEach((x) => {
      const li = el('li', x.nivel === 'nota' ? 'nota' : '', x.msg);
      if (x.ref) li.appendChild(el('span', 'ref', x.ref));
      ul.appendChild(li);
    });
    $('resumenValidaciones').textContent = iss.length ? nAv + ' advertencias y ' + (iss.length - nAv) + ' notas. No impiden dibujar.' : 'Sin advertencias.';
    const b = $('btnAdvertencias'); b.dataset.n = nAv; b.textContent = 'Advertencias (' + nAv + ')';
    const bd = $('badgeAv'); bd.textContent = nAv ? nAv + (nAv === 1 ? ' aviso' : ' avisos') : ''; bd.hidden = !nAv;
  }

  /* ----------------------------- selección --------------------------------- */
  function seleccionar(i) {
    const M = S.modelo, n = M.cursos.length;
    S.sel = i; S.lineas = []; S.relSel = null;
    actualizarPivote();
    if (i == null) {
      aplicarFiltroBase();
      pintarTextoInicial(); S.sucio = true; return;
    }
    const c = M.cursos[i];
    const bloque = new Set();
    c.bloquesReales.forEach((bid) => { const b = M.bloquePorId.get(bid); if (b) b.cursos.forEach((id) => bloque.add(M.porId.get(id))); });
    bloque.add(i);
    const progs = new Set(c.progs), elec = c.electivaProgs;
    const conElectivas = c.electivaDestinos.length > 0;
    // Doctorado: sus créditos incluyen TODA la maestría base. Al tocar un curso del doctorado se
    // ilumina la maestría base completa; al tocar uno de la base, el doctorado queda a media luz.
    const base = new Set(), derivados = new Set();
    c.progs.forEach((pid) => {
      const pr = M.programaPorId.get(pid);
      if (pr && pr.nivel === 'doctorado' && M.programaPorId.has(pr.maestriaBase)) base.add(pr.maestriaBase);
      M.doctorados.forEach((d) => { if (d.maestriaBase === pid) derivados.add(d.id); });
    });
    for (let j = 0; j < n; j++) {
      const o = M.cursos[j];
      const enBase = o.progs.some((p) => base.has(p));
      const enDerivado = o.progs.some((p) => derivados.has(p));
      const enBloque = bloque.has(j) || enBase;
      const enProg = o.progs.some((p) => progs.has(p));
      const enElec = o.progs.some((p) => elec.has(p));
      if (conElectivas) {
        // un curso con electivas: se encienden él, su programa y los puntos centrales a los que apunta; lo demás se apaga
        const propio = enProg && !o.escuela;   // los troncales de toda la escuela también se apagan
        S.tB[j] = enBloque ? 1.7 : 0;
        S.tM[j] = enBloque ? 1.1 : propio ? 0.5 : enElec ? 0.3 : 0;
        S.tN[j] = enBloque ? 1 : propio ? 0.45 : enElec ? 0.26 : 0;
      } else {
        S.tB[j] = enBloque ? 1.7 : 0.2;
        S.tM[j] = enBloque ? 1.1 : enProg ? 0.8 : (enDerivado || enElec) ? 0.5 : 0.1;
        S.tN[j] = enBloque ? 1 : enProg ? 0.85 : (enDerivado || enElec) ? 0.65 : 0.2;
      }
      if (S.filtro.size && S.nivel[j] === 0) { S.tM[j] *= 0.5; S.tB[j] *= 0.5; S.tN[j] *= 0.6; } // lo fuera del filtro se apaga más
      if (j !== i && bloque.has(j)) S.lineas.push({ a: i, b: j, w: 0.6 });
    }
    // los cursos más conectados con el elegido (su bloque, prerrequisitos y correquisitos) se quedan a la vista junto a él
    S.relSel = new Set(bloque);
    M.cursos.forEach((o, j) => { if (o.prereq.includes(c.id) || c.prereq.includes(o.id) || c.coreq.includes(o.id) || o.coreq.includes(c.id)) S.relSel.add(j); });
    describir(i);
    centrarEn(i);
    S.sucio = true;
  }
  // Ir a un curso desde el buscador o un enlace: si está plegado en una burbuja, se pasa al nivel «Cursos»
  function elegirCurso(i) {
    if (objetivoAbre(i) !== 1) fijarCapa('cursos');
    seleccionar(i);
  }
  // Al buscar o elegir un curso, la vista se desplaza para que quede en el centro de la zona libre de la pantalla
  // mueve el centro de la vista (sin cambiar el zoom) al centro de un grupo de cursos
  function centrarVistaEn(idx) {
    if (!idx.length || !S.sx0) return;
    // la mediana (no el promedio): un curso compartido y lejano no debe arrastrar el centro
    const med = (v) => { v.sort((a, b) => a - b); return v[Math.floor(v.length / 2)]; };
    const x = med(idx.map((i) => S.sx[i])), y = med(idx.map((i) => S.sy[i]));   // donde se ven ahora (juntos en su halo)
    const C = centroLibre();
    S.kZoom = 0.1; S.panXT = S.panX + (C.x - x); S.panYT = S.panY + (C.y - y); limitarPan(); S.sucio = true;
  }
  function centrarEn(i, rapido) {
    if (!S.sx0 || !S.pos) return;
    const C = centroLibre();
    S.kZoom = rapido ? 0.3 : 0.08;
    S.panXT = S.panX + (C.x - (S.sx0[i] - S.desv.x)); S.panYT = S.panY + (C.y - (S.sy0[i] - S.desv.y));
    limitarPan(); S.sucio = true;
  }

  /* --------------------------- proyección y dibujo ------------------------- */
  // en pantallas angostas el detalle y el zoom se apilan encima de la barra de niveles, sea de una fila o de dos
  function medirBarra() { document.documentElement.style.setProperty('--barra-h', $('barraNiveles').offsetHeight + 'px'); }
  function ajustar() {
    medirBarra();
    const r = lienzo.parentElement.getBoundingClientRect();
    S.dpr = Math.min(window.devicePixelRatio || 1, 1.5);
    S.w = Math.max(1, Math.round(r.width)); S.h = Math.max(1, Math.round(r.height));
    lienzo.width = Math.round(S.w * S.dpr); lienzo.height = Math.round(S.h * S.dpr);
    S.sucio = true;
  }
  // posición mezclada entre la disposición plana (2D) y la densa (3D) según la profundidad
  function mezclarPos(d) {
    const a = S.posPlano, b = S.posDenso, P = S.pos, k = d * d * (3 - 2 * d);
    for (let i = 0; i < P.X.length; i++) {
      P.X[i] = a.X[i] + (b.X[i] - a.X[i]) * k; P.Y[i] = a.Y[i] + (b.Y[i] - a.Y[i]) * k; P.Z[i] = a.Z[i] + (b.Z[i] - a.Z[i]) * k;
    }
    P.R = a.R + (b.R - a.R) * k;
  }
  function proyectar(t) {
    const P = S.pos, n = S.modelo.cursos.length;
    const d = S.profundidad;
    mezclarPos(d);
    // al apagar el filtro se sigue usando el último grupo hasta que los cursos terminan de regresar
    const miembros = S.hubMiembros.length ? S.hubMiembros : (S.hubRecuerdo || []);
    if (miembros.length) {
      // centro del grupo encendido (en el layout) y atracción de sus cursos hacia él
      let hx = 0, hy = 0, hz = 0;
      miembros.forEach((i) => { hx += S.pos.X[i]; hy += S.pos.Y[i]; hz += S.pos.Z[i]; });
      hx /= miembros.length; hy /= miembros.length; hz /= miembros.length;
      // los troncales de toda la escuela se acercan a la maestría como un grupo aparte, a un lado y cerca, con su propia área
      // (del color de la maestría); así no se mezclan con los cursos de ella y todo se puede leer
      // radio del grupo: el percentil 80 de las distancias, para que un curso lejano (compartido con otra maestría) no lo agrande
      const ds = miembros.map((i) => Math.hypot(S.pos.X[i] - hx, S.pos.Y[i] - hy)).sort((a, b) => a - b);
      const md = ds.length ? ds[Math.floor(ds.length * 0.5)] : 0;   // la mediana: el cuerpo de la maestría, sin contar sus cursos lejanos
      const rEsc = Math.max(24, md * 0.28);
      // el lado: el más cercano a donde estaban los troncales (el centro del layout), para que no den la vuelta a la maestría;
      // a una distancia que deja un pequeño espacio entre las dos áreas
      const dl = Math.hypot(hx, hy) || 1, ux = dl < 1e-3 ? 1 : -hx / dl, uy = dl < 1e-3 ? 0 : -hy / dl;
      const dEsc = md * 0.8 + rEsc * 1.35 + 8, cx = hx + ux * dEsc, cy = hy + uy * dEsc;
      const escIdx = []; S.modelo.cursos.forEach((c, i) => { if (c.escuela) escIdx.push(i); });
      for (let i = 0; i < S.pos.X.length; i++) {
        const a = S.atr[i]; if (a < 0.002) continue;
        const esc = S.modelo.cursos[i].escuela;
        let tx = hx, ty = hy;
        if (esc) { const k = escIdx.indexOf(i), ang = (k / Math.max(1, escIdx.length)) * 6.2832 + 0.6, rr = rEsc * (k % 2 ? 0.55 : 0.9); tx = cx + Math.cos(ang) * rr; ty = cy + Math.sin(ang) * rr; }
        S.pos.X[i] += (tx - S.pos.X[i]) * a; S.pos.Y[i] += (ty - S.pos.Y[i]) * a; S.pos.Z[i] += (hz - S.pos.Z[i]) * a;
      }
    }
    const yaw = (S.giroY + (S.vaiven ? Math.sin(t * 0.00008) * 0.6 : 0)) * d;
    const pit = (S.giroX + 0.28 + (S.vaiven ? Math.sin(t * 0.00006) * 0.1 : 0)) * d;
    const cy = Math.cos(yaw), sy = Math.sin(yaw), cp = Math.cos(pit), sp = Math.sin(pit);
    const foco = Math.max(900, P.R * 2.4); // la cámara queda siempre fuera del conjunto, aunque gire por completo
    const escala = Math.min(S.w, S.h) * 0.5 / (P.R * 1.18) * (S.w < 560 ? 1.1 : 1) * S.zoom;
    S.escala = escala;
    const cx = S.w / 2 + S.panX, cyy = S.h / 2 + S.panY;
    for (let i = 0; i < n; i++) {
      let x = P.X[i], y = P.Y[i], z = P.Z[i] * d * 1.25;
      if (S.vaiven && !reducir) {
        // en 3D los cursos levitan: ondas lentas y desfasadas entre sí, sin sacudidas
        const amp = 1.2 + 6 * d;
        x += Math.sin(t * 0.00022 + S.fase[i]) * amp; y += Math.cos(t * 0.00018 + S.fase[i] * 1.3) * amp * 1.2;
        z += Math.sin(t * 0.00015 + S.fase[i] * 2.1) * 12 * d;
      }
      const x1 = x * cy + z * sy, z1 = -x * sy + z * cy;
      const y1 = y * cp - z1 * sp, z2 = y * sp + z1 * cp;
      const p = foco / (foco + z2 * (d > 0 ? 1 : 0));
      S.sx[i] = cx + x1 * p * escala; S.sy[i] = cyy + y1 * p * escala; S.sp[i] = p; S.sz[i] = z2;
    }
    fijarPivote(P, d, cy, sy, cp, sp, foco, escala, cx, cyy);
    centrarTroncales();
    plegar();
  }
  // Con algo elegido (maestría, énfasis, troncales o un curso), su centro es el PIVOTE: queda fijo en el centro de la pantalla y todo
  // gira alrededor de él, no alrededor de toda la constelación. Entra y sale con suavidad (S.pinK).
  function fijarPivote(P, d, cy, sy, cp, sp, foco, escala, cx, cyy) {
    const idx = S.pivote || S.pivoteUlt;
    const objetivo = S.pivote ? 1 : 0, dk = objetivo - S.pinK;
    if (Math.abs(dk) > 0.003) { S.pinK += reducir ? dk : dk * 0.12; S.sucio = true; } else S.pinK = objetivo;
    if (!idx || !idx.length || S.pinK <= 0.001) return;
    const med = (v) => { const w = v.slice().sort((a, b) => a - b); return w[Math.floor(w.length / 2)]; };
    const x = med(idx.map((i) => P.X[i])), y = med(idx.map((i) => P.Y[i])), z = med(idx.map((i) => P.Z[i])) * d * 1.25;
    const x1 = x * cy + z * sy, z1 = -x * sy + z * cy, y1 = y * cp - z1 * sp, z2 = y * sp + z1 * cp;
    const p = foco / (foco + z2 * (d > 0 ? 1 : 0)), px = cx + x1 * p * escala, py = cyy + y1 * p * escala;
    const C = centroLibre(), dx = (C.x - px) * S.pinK, dy = (C.y - py) * S.pinK;
    for (let i = 0; i < S.sx.length; i++) { S.sx[i] += dx; S.sy[i] += dy; }
  }
  // cursos que sirven de pivote según lo elegido
  function actualizarPivote() {
    const M = S.modelo; let idx = null;
    if (S.sel != null) idx = [S.sel];
    else if (S.bloqueSel) { const b = M.bloquePorId.get(S.bloqueSel); idx = b ? b.cursos.map((cid) => M.porId.get(cid)).filter((j) => j != null) : null; }
    else if (S.troncalesSel) { idx = []; M.cursos.forEach((c, i) => { if (c.escuela) idx.push(i); }); }
    else if (S.maestriaSel) { idx = []; M.cursos.forEach((c, i) => { if (!c.escuela && c.progs.includes(S.maestriaSel)) idx.push(i); }); }
    S.pivote = idx && idx.length ? idx : null;
    if (S.pivote) S.pivoteUlt = S.pivote;
  }
  // Centro estático: sin nada elegido, «Cursos troncales» queda en el centro de la zona libre (entre los paneles) y el resto
  // a su alrededor. Con algo elegido el centrado se suelta para que el zoom automático enfoque lo elegido.
  function centrarTroncales() {
    const n = S.modelo.cursos.length, enfoque = S.filtro.size > 0 || S.sel != null;
    let ex = 0, ey = 0, ne = 0;
    for (let i = 0; i < n; i++) if (S.modelo.cursos[i].escuela) { ex += S.sx[i]; ey += S.sy[i]; ne++; }
    if (!ne) { for (let i = 0; i < n; i++) { ex += S.sx[i]; ey += S.sy[i]; } ne = n; }
    ex /= ne; ey /= ne;
    const C = centroLibre();
    const tx = enfoque ? 0 : C.x + S.panX - ex, ty = enfoque ? 0 : C.y + S.panY - ey;
    const d = S.desv, k = reducir || enfoque ? 1 : 0.12;
    if (Math.abs(tx - d.x) > 0.3 || Math.abs(ty - d.y) > 0.3) { d.x += (tx - d.x) * k; d.y += (ty - d.y) * k; S.sucio = true; } else { d.x = tx; d.y = ty; }
    for (let i = 0; i < n; i++) { S.sx[i] += d.x; S.sy[i] += d.y; }
  }
  // Punto central al que apunta un destino de electiva: el de una maestría o el de un énfasis / especializado
  const destinoCentro = (cent, d) => (d.tipo === 'bloque' ? cent.bloques.get(d.id) : cent.get(d.id));
  // Centroides (en pantalla y en el layout) y dispersión de cada grupo; sirven para las estelas y las órbitas
  function calcularCentroides(M) {
    const cent = new Map();
    M.programas.forEach((pr) => {
      const mem = []; M.cursos.forEach((c, i) => { if (!c.escuela && c.progs.includes(pr.id)) mem.push(i); });
      if (mem.length < 3) return;
      let cx = 0, cy = 0, cz = 0, sx = 0, sy = 0;
      mem.forEach((i) => { cx += S.pos.X[i]; cy += S.pos.Y[i]; cz += S.pos.Z[i]; sx += S.sx[i]; sy += S.sy[i]; });
      const m = mem.length; cx /= m; cy /= m; cz /= m;
      // distancia media del grupo medida con sus cursos propios (sin los que también son de otra maestría)
      const propios = mem.filter((i) => M.cursos[i].progs.length === 1);
      const base = propios.length >= 2 ? propios : mem;
      let dm = 0; base.forEach((i) => { dm += Math.hypot(S.pos.X[i] - cx, S.pos.Y[i] - cy, (S.pos.Z[i] - cz) * 0.6); });
      cent.set(pr.id, { cx, cy, cz, sx: sx / m, sy: sy / m, dm: dm / base.length, mem });
    });
    // bloques reales (énfasis, especializados): su propio punto central
    cent.bloques = new Map();
    M.bloques.forEach((b) => {
      if (b.implicito) return;
      const mem = b.cursos.map((id) => M.porId.get(id)).filter((i) => i != null);
      if (mem.length < 2) return;
      let cx = 0, cy = 0, cz = 0, sx = 0, sy = 0;
      mem.forEach((i) => { cx += S.pos.X[i]; cy += S.pos.Y[i]; cz += S.pos.Z[i]; sx += S.sx[i]; sy += S.sy[i]; });
      const m = mem.length;
      cent.bloques.set(b.id, { cx: cx / m, cy: cy / m, cz: cz / m, sx: sx / m, sy: sy / m, mem });
    });
    return cent;
  }

  /* ============================ CAPAS: burbujas ↔ cursos ============================
   * La pantalla se entiende por capas. Cada curso pertenece a una burbuja según el nivel elegido (maestría, énfasis...).
   * S.ab[i] (0..1) dice cuánto está «abierto» el curso: 0 = plegado en el centro de su burbuja, 1 = suelto como punto.
   * Al pasar de burbujas a cursos S.ab sube (la burbuja se divide en sus puntos); al volver baja (los puntos se recogen).
   * Todo reutiliza las mismas posiciones del layout; solo se pliegan en pantalla. */
  const COLOR_SIN = '#8A8A8A';
  const esProgTop = (p) => p && (p.nivel === 'maestria' || p.nivel === 'doctorado');
  // Agrupa los cursos por burbuja: nivel 'maestrias' (una por maestría o doctorado) o 'enfasis' (por énfasis / bloque único)
  function calcularGrupos(nivel) {
    const M = S.modelo, n = M.cursos.length;
    const de = new Int16Array(n).fill(-1), lista = [], idx = new Map();
    const dar = (id, o) => { if (!idx.has(id)) { idx.set(id, lista.length); lista.push(Object.assign({ id, mem: [], ring: false, esc: false, maestriaId: null, padre: null, progColor: null }, o)); } return idx.get(id); };
    M.cursos.forEach((c, i) => {
      let g;
      if (c.escuela) g = dar('esc', { nombre: 'Cursos obligatorios', esc: true });
      else {
        const tp = c.progs.map((pid) => M.programaPorId.get(pid)).find(esProgTop);
        if (!tp) g = dar('sin', { nombre: 'Sin programa' });
        else {
          const b = nivel === 'enfasis' ? c.bloquesReales.map((bid) => M.bloquePorId.get(bid)).find((x) => x && !x.implicito) : null;
          if (b) g = dar('b:' + b.id, { nombre: b.nombre, titulo: b.tituloEsp || ('Especialización en ' + b.nombre), ring: !!b.otorgaEsp, esEnfasis: !!b.esEnfasis, bloque: b.nombre, maestriaId: tp.id, padre: 'p:' + tp.id, progColor: tp.id });
          else if (nivel === 'enfasis') g = dar('g:' + tp.id, { nombre: 'Cursos de ' + tp.nombre.replace(/^(Maestría|Doctorado) en /i, ''), maestriaId: tp.id, padre: 'p:' + tp.id, progColor: tp.id });
          else g = dar('p:' + tp.id, { nombre: tp.nombre, maestriaId: tp.id, progColor: tp.id });
        }
      }
      de[i] = g; lista[g].mem.push(i);
    });
    return { de, lista, nivel };
  }
  function grupos() {
    const nivel = S.capaBurbuja === 'maestrias' ? 'maestrias' : 'enfasis';
    if (!S.gr || S.gr.nivel !== nivel || S.gr.n !== S.modelo.cursos.length) { S.gr = calcularGrupos(nivel); S.gr.n = S.modelo.cursos.length; S.gcx = new Float32Array(S.gr.lista.length); S.gcy = new Float32Array(S.gr.lista.length); }
    return S.gr;
  }
  const colorGrupo = (g) => (g.esc ? colorEscuelaVivo() : g.progColor ? colorNodo((S.modelo.programaPorId.get(g.progColor) || { color: COLOR_SIN }).color) : COLOR_SIN);
  // ¿Debe estar suelto este curso? En «cursos» todos; si hay una maestría elegida, los suyos
  // los cursos de toda la escuela se abren junto con la maestría elegida (nivel 1), como siempre
  const objetivoAbre = (i) => {
    if (S.capa === 'cursos') return 1;
    if (!S.bloqueSel && !S.troncalesSel) return 0;   // elegir una maestría la divide en énfasis; no abre sus cursos
    return S.filtro.size && S.nivel[i] >= 2 ? 1 : 0;
  };
  // La apertura avanza a velocidad constante en el tiempo (≈1 s); el suavizado entra al plegar (smoothstep), así se ve la división
  function actualizarAbre() {
    const n = S.ab.length, ahora = performance.now();
    const paso = reducir ? 1 : Math.min(60, ahora - (S.tAbre || ahora)) / 1000;
    S.tAbre = ahora;
    const gr = grupos();
    for (let i = 0; i < n; i++) {
      const T = objetivoAbre(i), d = T - S.ab[i];
      if (d !== 0) { if (Math.abs(d) <= paso) S.ab[i] = T; else { S.ab[i] += d > 0 ? paso : -paso; S.sucio = true; } }
      // visibilidad: en «programas especializados» solo se ven las mini constelaciones que aplican (más lo que esté abierto)
      // con un curso elegido se ven él y los más conectados con él (su bloque, prerrequisitos y correquisitos); lo demás se desvanece
      const V = S.sel != null ? ((i === S.sel || (S.relSel && S.relSel.has(i))) ? 1 : 0) : (S.capa === 'especializados' && T !== 1 && !gr.lista[gr.de[i]].ring ? 0 : 1), dv = V - S.av[i];
      if (dv !== 0) { if (Math.abs(dv) <= paso * 1.4) S.av[i] = V; else { S.av[i] += dv > 0 ? paso * 1.4 : -paso * 1.4; S.sucio = true; } }
    }
  }
  // Factor de compactación de un grupo cuando no está abierto: la constelación se conserva, solo se junta
  const COMPACTO = () => (S.capaBurbuja === 'maestrias' ? 0.5 : 0.58);
  // junta los puntos de cada grupo alrededor del centro de su halo (el centro se mide con las posiciones sin juntar)
  function plegar() {
    const n = S.modelo.cursos.length, gr = grupos(), L = gr.lista;
    S.sx0.set(S.sx); S.sy0.set(S.sy);
    const sx = new Float32Array(L.length), sy = new Float32Array(L.length), c = new Int32Array(L.length);
    for (let i = 0; i < n; i++) { const g = gr.de[i]; sx[g] += S.sx0[i]; sy[g] += S.sy0[i]; c[g]++; }
    const px = new Float32Array(L.length), py = new Float32Array(L.length);
    for (let g = 0; g < L.length; g++) {
      S.gcx[g] = c[g] ? sx[g] / c[g] : 0; S.gcy[g] = c[g] ? sy[g] / c[g] : 0;
      const b = S.bur.get(L[g].id);   // el grupo se junta alrededor del centro de su halo (el animado), así lo uno contiene a lo otro
      px[g] = b ? b.x : S.gcx[g]; py[g] = b ? b.y : S.gcy[g];
    }
    const cm = COMPACTO(), mezcla = S.blendHasta > performance.now() && S.pfOk;
    for (let i = 0; i < n; i++) {
      const e = S.ab[i], g = gr.de[i];
      if (e < 0.999) {
        const k = e * e * (3 - 2 * e), f = cm + (1 - cm) * k;
        S.sx[i] = px[g] + (S.sx0[i] - S.gcx[g]) * f; S.sy[i] = py[g] + (S.sy0[i] - S.gcy[g]) * f;
      }
      // al cambiar de nivel los puntos viajan suavemente a su nuevo grupo, sin saltos
      if (mezcla) { S.sx[i] = S.pfx[i] + (S.sx[i] - S.pfx[i]) * 0.14; S.sy[i] = S.pfy[i] + (S.sy[i] - S.pfy[i]) * 0.14; S.sucio = true; }
      S.pfx[i] = S.sx[i]; S.pfy[i] = S.sy[i];
    }
    S.pfOk = true;
  }
  // burbujas visibles y su animación (posición, tamaño y opacidad suavizadas); una burbuja nueva nace de su burbuja madre
  // zonas ocupadas por los paneles flotantes (en coordenadas del lienzo); se miden de vez en cuando, no en cada cuadro
  function medirReservas() {
    const ahora = performance.now();
    if (S.reservas && ahora - S.tReservas < 350) return S.reservas;
    S.tReservas = ahora;
    const c = lienzo.getBoundingClientRect(), res = [];
    ['leyendaPanel', 'panel', 'barraNiveles', 'topbar'].forEach((id) => {
      const e = $(id); if (!e || e.offsetParent === null) return;
      const r = e.getBoundingClientRect();
      res.push({ x0: r.left - c.left - 8, y0: r.top - c.top - 8, x1: r.right - c.left + 8, y1: r.bottom - c.top + 8 });
    });
    S.reservas = res; return res;
  }
  // centro de la pantalla (en vertical, entre el buscador y la barra de niveles): ahí se queda «Cursos troncales»
  function centroLibre() {
    const W = S.w, H = S.h, res = medirReservas();
    let top = 0, bot = H;
    res.forEach((q) => {
      if (q.x0 < W / 2 && q.x1 > W / 2) { if ((q.y0 + q.y1) / 2 < H / 2) top = Math.max(top, q.y1); else bot = Math.min(bot, q.y0); }
    });
    // en horizontal el centro es siempre el de la pantalla: abrir o cerrar la leyenda no lo corre
    return { x: W / 2, y: (top + bot) / 2 };
  }
  function actualizarBurbujas() {
    const gr = grupos(), L = gr.lista, W = S.w, H = S.h;
    const enCursos = S.capa === 'cursos';
    const nivel = enCursos ? S.capaBurbuja : S.capa;   // en «cursos» las burbujas siguen ahí mientras los puntos terminan de abrirse o de recogerse
    if (enCursos) {
      let todo = true; for (let i = 0; i < S.ab.length; i++) if (S.ab[i] < 0.999) { todo = false; break; }
      if (todo) { S.bur.clear(); return; }
    }
    const objetivos = new Map();
    const vis = [];
    L.forEach((g, k) => { if (nivel === 'especializados' && !g.ring) return; vis.push({ g, k, abierta: !enCursos && objetivoAbre(g.mem[0]) === 1 }); });
    const cm = COMPACTO();
    const P = vis.map(({ g, k, abierta }) => {
      // el halo envuelve a los puntos del grupo (el 90 % más cercano, para que un curso compartido y lejano no lo agrande)
      const ds = g.mem.map((i) => Math.hypot(S.sx0[i] - S.gcx[k], S.sy0[i] - S.gcy[k])).sort((x, y) => x - y);
      const md90 = ds.length ? ds[Math.min(ds.length - 1, Math.floor(ds.length * 0.9))] : 0;
      let r = md90 * cm * 1.15 + 26, a = 1;
      if (abierta) { const md = ds.length ? ds[ds.length - 1] : 0; r = md * 1.15 + 26; a = 0.22; }
      else if (S.filtro.size) a = (g.maestriaId && g.maestriaId === S.maestriaSel && !S.bloqueSel && !S.troncalesSel) ? 1 : 0.35;   // los énfasis de la maestría elegida quedan vivos
      // profundidad de la burbuja: promedio de la de sus cursos (con un mínimo de relieve aunque la vista sea plana)
      let zs = 0, zl = 0; g.mem.forEach((i) => { zs += S.sz[i]; zl += S.pos.Z[i]; });
      const zEff = zs / g.mem.length + (zl / g.mem.length) * 0.5 * (1 - S.profundidad);
      const foco = Math.max(900, S.pos.R * 2.4), pz = Math.max(0.92, Math.min(1.08, foco / (foco + zEff)));
      return { g, k, abierta, r, a, x: S.gcx[k], y: S.gcy[k], fijo: abierta, ancla: false, z: zEff, pz };
    });
    // Las burbujas van donde están sus puntos: siguen al zoom y al centrado en lo elegido. Sin nada elegido ni zoom, «Cursos troncales»
    // queda en el centro de la pantalla (así lo deja el centrado de los puntos) y lo demás a su alrededor.
    const C = centroLibre();
    const iE = P.findIndex((o) => o.g.esc && !o.abierta);
    const hayAbierta = P.some((o) => o.abierta);
    // los troncales ya no se clavan al centro: flotan donde les toca en el layout, como los demás grupos
    // separación: si dos burbujas se encimarían, se apartan (las abiertas no se mueven) y se mantienen fuera de los paneles flotantes
    // en la vista normal las burbujas se mantienen dentro de la pantalla y fuera de los paneles; con zoom no: lo que no cabe sale de la pantalla
    const libre = Math.abs(S.zoom - 1) < 0.05 && !S.filtro.size && S.sel == null, reservas = libre ? medirReservas() : [];
    // con muchas burbujas (énfasis) se abre un poco el conjunto y se deja más aire para sus nombres
    const muchas = nivel !== 'maestrias', holgura = muchas ? 22 : 30;   // solo evita que se monten; las distancias son las del layout
    // acerca los grupos a «Cursos troncales» (o al centro del conjunto): la distancia sigue siendo la del layout, solo más corta;
    // la separación de abajo evita que se monten
    {
      const libres = P.filter((o) => !o.fijo && o !== P[iE]);
      if (libres.length) {
        let mx, my;
        if (iE >= 0 && !P[iE].abierta) { mx = P[iE].x; my = P[iE].y; }
        else { mx = 0; my = 0; P.forEach((o) => { mx += o.x; my += o.y; }); mx /= P.length; my /= P.length; }
        const k = muchas ? 0.92 : 0.66;
        libres.forEach((o) => { o.x = mx + (o.x - mx) * k; o.y = my + (o.y - my) * k; });
      }
    }
    for (let it = 0; it < 160; it++) {
      let mov = false;
      P.forEach((o) => {
        if (o.fijo || o.ancla) return;
        if (!libre) return;
        o.x = Math.max(o.r + 8, Math.min(W - o.r - 8, o.x)); o.y = Math.max(o.r + 8, Math.min(H - o.r - 8, o.y));
        reservas.forEach((q) => {
          const nx = Math.max(q.x0, Math.min(o.x, q.x1)), ny = Math.max(q.y0, Math.min(o.y, q.y1));
          const d = Math.hypot(o.x - nx, o.y - ny);
          if (d >= o.r) return;
          if (d > 0.01) { const f = (o.r - d) / d; o.x += (o.x - nx) * f; o.y += (o.y - ny) * f; }
          else { // el centro quedó dentro del panel: sale por el lado más cercano
            const ds = [[o.x - q.x0, -1, 0], [q.x1 - o.x, 1, 0], [o.y - q.y0, 0, -1], [q.y1 - o.y, 0, 1]].sort((a, b) => a[0] - b[0])[0];
            o.x += ds[1] * (ds[0] + o.r); o.y += ds[2] * (ds[0] + o.r);
          }
          mov = true;
        });
      });
      for (let a = 0; a < P.length; a++) for (let b = a + 1; b < P.length; b++) {
        const A = P[a], B = P[b];
        if ((A.fijo || A.ancla) && (B.fijo || B.ancla)) continue;
        let dx = B.x - A.x, dy = B.y - A.y, d = Math.hypot(dx, dy);
        // si están a distinta profundidad pueden solaparse en pantalla (una pasa por encima de la otra)
        // lo que está abierto (énfasis o maestría elegida) nunca se monta con nada: sin solape por profundidad
        const dzRef = Math.max(1, S.pos.R * 0.22), prof = A.fijo || B.fijo ? 0 : Math.min(1, Math.abs(A.z - B.z) / dzRef);
        const falta = (A.r * A.pz + B.r * B.pz + holgura) * (1 - 0.26 * prof) - d;
        if (falta <= 0) continue;
        if (d < 0.01) { const t = (a + b) * 2.4; dx = Math.cos(t); dy = Math.sin(t); d = 1; }
        const ux = dx / d, uy = dy / d, ca = (A.fijo || A.ancla) ? 0 : (B.fijo || B.ancla) ? 1 : 0.5, cb = 1 - ca;
        A.x -= ux * falta * ca; A.y -= uy * falta * ca; B.x += ux * falta * cb; B.y += uy * falta * cb; mov = true;
      }
      if (!mov) break;
    }
    P.forEach((o) => objetivos.set(o.g.id, o));
    const ke = reducir ? 1 : 0.13;
    objetivos.forEach((o, id) => {
      let s = S.bur.get(id);
      if (!s) {
        const pa = o.g.padre && S.bur.get(o.g.padre);
        s = pa ? { x: pa.x, y: pa.y, r: pa.r * 0.55, a: 0 } : { x: o.x, y: o.y, r: o.r * 0.6, a: 0 };
        S.bur.set(id, s);
      }
      s.g = o.g; s.abierta = o.abierta; s.z = o.z;
      s.pz = s.pz == null ? o.pz : s.pz + (o.pz - s.pz) * 0.2;
      const dx = o.x - s.x, dy = o.y - s.y, dr = o.r - s.r, da = o.a - s.a;
      s.x += dx * ke; s.y += dy * ke; s.r += dr * ke; s.a += da * ke;
      if (Math.abs(dx) + Math.abs(dy) + Math.abs(dr) > 0.4 || Math.abs(da) > 0.004) S.sucio = true; else { s.x = o.x; s.y = o.y; s.r = o.r; s.a = o.a; }
      // opacidad real: mientras sus puntos se abren, la burbuja se desvanece al ritmo de la apertura (y reaparece al recogerse)
      let av = 0; o.g.mem.forEach((i) => { av += S.ab[i]; }); av /= o.g.mem.length;
      s.al = o.abierta ? s.a : s.a * (1 - av);
    });
    S.bur.forEach((s, id) => {
      if (objetivos.has(id)) return;
      s.a += (0 - s.a) * ke; s.al = s.a; s.abierta = true; if (s.pz == null) s.pz = 1;
      if (s.a < 0.012) S.bur.delete(id); else S.sucio = true;
    });
  }
  // El nombre siempre va dentro de la burbuja: se parte en líneas y la burbuja nunca es más chica de lo que ese texto necesita
  const FAMILIA = '"area-variable", "Inktrap", ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif';
  const FUENTE_BURB = (px) => '700 ' + px + 'px ' + FAMILIA;
  // nombre completo: «Maestría en…», «Énfasis en…» o «Programa especializado en…»
  const nombreBurbuja = (g) => {
    if (!g.bloque) return g.nombre;
    if (S.capa === 'especializados') return 'Programa especializado en ' + g.bloque;
    return (g.esEnfasis ? 'Énfasis en ' : 'Núcleo: ') + g.bloque;
  };
  const cuentaBurbuja = (g) => g.mem.length + (g.mem.length === 1 ? ' curso' : ' cursos');
  function partirTexto(texto, maxW) {
    const lineas = []; let cur = '';
    texto.split(/\s+/).forEach((w) => { const t = cur ? cur + ' ' + w : w; if (cur && ctx.measureText(t).width > maxW) { lineas.push(cur); cur = w; } else cur = t; });
    if (cur) lineas.push(cur);
    return lineas;
  }
  // mide el bloque de texto (nombre + cuántos cursos) a un tamaño de letra y devuelve sus líneas y su media diagonal
  function bloqueTexto(g, px, maxW) {
    ctx.save(); ctx.font = FUENTE_BURB(px);
    const lineas = partirTexto(nombreBurbuja(g), maxW).concat(cuentaBurbuja(g));
    const w = Math.max(...lineas.map((l) => ctx.measureText(l).width)), h = lineas.length * (px + 4);
    ctx.restore();
    return { lineas, w, h, diag: Math.hypot(w / 2, h / 2) };
  }
  function radioNecesario(g) {
    const clave = nombreBurbuja(g) + '|' + g.mem.length;
    S.rTxt = S.rTxt || new Map();
    if (!S.rTxt.has(clave)) { const b = bloqueTexto(g, 14, 118); S.rTxt.set(clave, b.diag * 1.12 + 6); }
    return S.rTxt.get(clave);
  }
  // Área del grupo de cursos troncales cuando se acercan a una maestría: un halo del color de esa maestría, al lado de la suya
  function dibujarHaloTroncales(tecnico) {
    if (!S.programaEscuela || S.kEsc < 0.05 || S.capa !== 'cursos' || S.sel != null) return;
    const M = S.modelo, pr = M.programaPorId.get(S.programaEscuela); if (!pr) return;
    let cx = 0, cy = 0, n = 0; const idx = [];
    M.cursos.forEach((c, i) => { if (c.escuela) { idx.push(i); cx += S.sx[i]; cy += S.sy[i]; n++; } });
    if (!n) return; cx /= n; cy /= n;
    let md = 0; idx.forEach((i) => { md = Math.max(md, Math.hypot(S.sx[i] - cx, S.sy[i] - cy)); });
    const rr = md * 1.25 + 22, claro = S.tema === 'claro', rgb = colorRgb(pr.color), a = Math.min(1, S.kEsc);
    const rgba = (x) => 'rgba(' + rgb[0] + ',' + rgb[1] + ',' + rgb[2] + ',' + x + ')';
    ctx.save(); ctx.globalCompositeOperation = 'source-over'; ctx.globalAlpha = a;
    if (tecnico) { ctx.setLineDash([2, 6]); ctx.strokeStyle = claro ? '#070707' : '#EFEEE8'; ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(cx, cy, rr, 0, 6.2832); ctx.stroke(); }
    else {
      const g = ctx.createRadialGradient(cx, cy, rr * 0.1, cx, cy, rr);
      g.addColorStop(0, rgba(claro ? 0.2 : 0.3)); g.addColorStop(0.62, rgba(claro ? 0.12 : 0.18)); g.addColorStop(1, rgba(0));
      ctx.fillStyle = g; ctx.beginPath(); ctx.arc(cx, cy, rr, 0, 6.2832); ctx.fill();
      ctx.setLineDash([2, 6]); ctx.strokeStyle = rgba(0.6); ctx.lineWidth = 1; ctx.beginPath(); ctx.arc(cx, cy, rr * 0.96, 0, 6.2832); ctx.stroke();
    }
    ctx.restore();
  }
  // curso cuyo nombre está dibujado en ese punto (los de más arriba primero)
  function nombreEn(px, py) {
    for (let k = S.zonasEt.length - 1; k >= 0; k--) { const z = S.zonasEt[k]; if (px >= z.x0 && px <= z.x1 && py >= z.y0 && py <= z.y1) return z.curso; }
    return null;
  }
  function burbujaEn(px, py) {
    let mejor = null, md = 1e9;
    S.bur.forEach((s) => { if (s.al < 0.25 || s.abierta) return; const d = Math.hypot(s.x - px, s.y - py); if (d <= s.r * (s.pz || 1) && d < md) { md = d; mejor = s; } });
    return mejor;
  }
  // Línea guía: un trazo fino y recto que sale de lo que se nombra y llega al texto (recta y ortogonal, como en un esquema editorial).
  // o: { x0, y0, wb, hb } caja del texto; { tipx, tipy } de dónde sale; color; punto (true: lleva un puntito en el origen)
  function dibujarGuia(o) {
    const { x0, y0, wb, hb, tipx, tipy, color, punto } = o, x1 = x0 + wb, y1 = y0 + hb, cy = y0 + hb / 2;
    let ax, ay;
    if (tipx >= x0 && tipx <= x1) { ax = tipx; ay = tipy < y0 ? y0 : y1; }                       // el texto queda arriba o abajo: línea vertical
    else { ax = tipx < x0 ? x0 : x1; ay = tipy >= y0 && tipy <= y1 ? tipy : cy; }                // queda a un lado: línea horizontal (o hasta su altura)
    ctx.save(); ctx.setLineDash([]); ctx.lineCap = 'round'; ctx.lineWidth = 1.2; ctx.strokeStyle = color;
    ctx.beginPath(); ctx.moveTo(tipx, tipy); ctx.lineTo(ax, ay); ctx.stroke();
    if (punto) { ctx.fillStyle = color; ctx.beginPath(); ctx.arc(tipx, tipy, 2.6, 0, 6.2832); ctx.fill(); }
    ctx.restore();
  }
  // Nombre de un curso: una línea fina sale de su punto y el texto queda al final. Prueba seis posiciones (la última usada primero, para que no
  // salte) y evita pisar otros nombres ya dibujados salvo que sea obligatorio (el elegido o el que tocas con el cursor).
  function globoCurso(clave, texto, ancho, x, y, alfa, obligatoria, cajas, W, H, claro, curso) {
    if (x < -20 || x > W + 20 || y < -20 || y > H + 20) return;
    const hb = 20, wb = ancho + 8, d = 30, dy = 26;
    const opciones = [[x + d + wb / 2, y], [x - d - wb / 2, y], [x + d + wb / 2, y - dy], [x - d - wb / 2, y - dy], [x + d + wb / 2, y + dy], [x - d - wb / 2, y + dy]];
    const orden = [0, 1, 2, 3, 4, 5], previo = S.ladoEt.get(clave);
    if (previo != null && previo < 6) { orden.splice(orden.indexOf(previo), 1); orden.unshift(previo); }
    const tinta = claro ? '#070707' : '#EFEEE8', papel = claro ? '#EFEEE8' : '#070707';
    for (const k of orden) {
      const cx = Math.max(wb / 2 + 6, Math.min(W - wb / 2 - 6, opciones[k][0])), cy = Math.max(hb / 2 + 6, Math.min(H - hb / 2 - 6, opciones[k][1]));
      const caja = { x0: cx - wb / 2 - 4, y0: cy - hb / 2 - 3, x1: cx + wb / 2 + 4, y1: cy + hb / 2 + 3 };
      if (!obligatoria && cajas.some((c) => caja.x0 < c.x1 && caja.x1 > c.x0 && caja.y0 < c.y1 && caja.y1 > c.y0)) continue;
      S.ladoEt.set(clave, k); cajas.push(caja);
      if (curso != null && alfa > 0.3) S.zonasEt.push({ x0: cx - wb / 2 - 2, y0: cy - hb / 2 - 2, x1: cx + wb / 2 + 2, y1: cy + hb / 2 + 2, curso });   // el nombre también se toca
      ctx.globalAlpha = alfa;
      dibujarGuia({ x0: cx - wb / 2, y0: cy - hb / 2, wb, hb, tipx: x + (cx > x ? 8 : -8) * (Math.abs(cy - y) < 1 ? 1 : 0.8), tipy: y, color: claro ? 'rgba(7,7,7,0.7)' : 'rgba(239,238,232,0.75)', punto: false });
      const ta = ctx.textAlign; ctx.textAlign = 'center';
      ctx.lineWidth = 3.5; ctx.strokeStyle = papel; ctx.strokeText(texto, cx, cy); ctx.fillStyle = tinta; ctx.fillText(texto, cx, cy); ctx.textAlign = ta;
      return;
    }
  }
  // Halos: un área de color suave alrededor de cada grupo de puntos (maestría, énfasis o programa especializado) con su nombre afuera.
  // Los puntos de los cursos siguen ahí, juntos y conectados; el halo solo marca a qué grupo pertenecen.
  function dibujarBurbujas(tecnico) {
    if (!S.bur.size) return;
    const claro = S.tema === 'claro', tinta = claro ? '#070707' : '#EFEEE8', papel = claro ? '#EFEEE8' : '#070707';
    const arr = [...S.bur.values()].filter((b) => b.al > 0.004).sort((a, b) => (b.z || 0) - (a.z || 0));   // lejanos primero
    ctx.save();
    ctx.globalCompositeOperation = 'source-over'; ctx.setLineDash([]); ctx.lineJoin = 'round';
    if ('letterSpacing' in ctx) ctx.letterSpacing = '0px';
    ctx.font = FUENTE_BURB(16); ctx.textBaseline = 'middle'; ctx.textAlign = 'center';
    arr.forEach((s) => {
      const col = colorGrupo(s.g), rr = s.r * (s.pz || 1), rgb = colorRgb(col);
      const rgba = (a) => 'rgba(' + rgb[0] + ',' + rgb[1] + ',' + rgb[2] + ',' + a + ')';
      ctx.globalAlpha = s.al * fHalo(s);
      if (s.abierta) {   // el grupo elegido ya está abierto: solo queda un halo tenue detrás de sus cursos
        ctx.fillStyle = rgba(claro ? 0.12 : 0.18); ctx.beginPath(); ctx.arc(s.x, s.y, rr, 0, 6.2832); ctx.fill();
      } else if (tecnico) {
        ctx.setLineDash([2, 6]); ctx.strokeStyle = tinta; ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(s.x, s.y, rr, 0, 6.2832); ctx.stroke(); ctx.setLineDash([]);
      } else {
        const g = ctx.createRadialGradient(s.x, s.y, rr * 0.1, s.x, s.y, rr);
        g.addColorStop(0, rgba(claro ? 0.22 : 0.34)); g.addColorStop(0.62, rgba(claro ? 0.14 : 0.2)); g.addColorStop(1, rgba(0));
        ctx.fillStyle = g; ctx.beginPath(); ctx.arc(s.x, s.y, rr, 0, 6.2832); ctx.fill();
        if ((s === S.focoB || (S.focoHalo && S.focoHalo(s))) && S.hk > 0) { ctx.globalAlpha = s.al * 0.7 * S.hk; ctx.fill(); ctx.globalAlpha = s.al; }   // el halo señalado se enciende un poco más
        ctx.setLineDash([2, 6]); ctx.strokeStyle = rgba(0.6); ctx.lineWidth = 1; ctx.beginPath(); ctx.arc(s.x, s.y, rr * 0.96, 0, 6.2832); ctx.stroke(); ctx.setLineDash([]);
      }
      if (S.capa === 'especializados' && s.g.ring) { ctx.strokeStyle = tinta; ctx.lineWidth = 2.5; ctx.beginPath(); ctx.arc(s.x, s.y, rr + 4, 0, 6.2832); ctx.stroke(); }
    });
    // nombres: debajo del halo o, si estorban, encima o al lado (que pisen lo menos posible)
    const etiquetas = [];
    arr.forEach((s) => {
      const suyo = S.maestriaSel && !S.bloqueSel && !S.troncalesSel && (s.g.maestriaId === S.maestriaSel || s.g.esc);   // con una maestría elegida, solo sus énfasis llevan nombre
      if ((S.filtro.size && !suyo) || s.abierta || s.al < 0.15) return;   // con algo elegido, las demás cosas no llevan texto
      const rr = s.r * (s.pz || 1);
      const lineas = partirTexto(nombreBurbuja(s.g), Math.max(150, rr * 2));
      etiquetas.push({ s, rr, lineas, ancho: Math.max(...lineas.map((l) => ctx.measureText(l).width)) });
    });
    const alto = 20, centro = centroLibre(), cajas = [];
    // el nombre va debajo del halo si este queda en la mitad de abajo de la pantalla y encima si queda arriba; el lado solo cambia si se aleja
    // bastante de la mitad (así no parpadea), y la posición se suaviza
    etiquetas.sort((p, q) => p.s.y - q.s.y);
    // 1) lado y caja natural de cada nombre; si esa caja pisa otro nombre o el halo de otro grupo, el nombre pasa a ser un GLOBO
    //    (burbuja que lo recubre, con una punta hacia su círculo). Entra al instante y sale solo tras ~40 cuadros sin pisar nada: no parpadea.
    etiquetas.forEach((e) => {
      const { s, rr } = e; e.h = e.lineas.length * alto; e.ring = S.capa === 'especializados' && s.g.ring ? 8 : 0;
      if (s.lado !== 'abajo' && s.lado !== 'arriba') s.lado = s.y >= centro.y ? 'abajo' : 'arriba';
      else if (s.lado === 'abajo' && s.y < centro.y - 50) s.lado = 'arriba';
      else if (s.lado === 'arriba' && s.y > centro.y + 50) s.lado = 'abajo';
      if (s.g.esc && S.maestriaSel && !S.troncalesSel && s.y > centro.y - 50) s.lado = 'arriba';   // troncales con una maestría elegida: el nombre queda libre, no sobre sus puntos
      const ty = s.lado === 'abajo' ? s.y + rr + 8 + e.ring : s.y - rr - 8 - e.h - e.ring;
      e.nat = { x0: s.x - e.ancho / 2 - 4, x1: s.x + e.ancho / 2 + 4, y0: ty - 2, y1: ty + e.h };
    });
    const pisaHalo = (c, otro) => { const nx = Math.max(c.x0, Math.min(otro.x, c.x1)), ny = Math.max(c.y0, Math.min(otro.y, c.y1)); return Math.hypot(otro.x - nx, otro.y - ny) < otro.r * (otro.pz || 1) * 0.9; };
    etiquetas.forEach((e) => {
      const c = e.nat; let choca = false;
      for (const o of etiquetas) { if (o === e) continue; const d = o.nat; if (c.x0 < d.x1 && c.x1 > d.x0 && c.y0 < d.y1 && c.y1 > d.y0) { choca = true; break; } }
      if (!choca) for (const o of arr) { if (o === e.s || o.abierta || o.al < 0.15) continue; if (pisaHalo(c, o)) { choca = true; break; } }
      if (choca) { e.s.globo = true; e.s.libreN = 0; }
      else if (e.s.globo) { e.s.libreN = (e.s.libreN || 0) + 1; if (e.s.libreN > 40) e.s.globo = false; else S.sucio = true; }
    });
    // 2) colocación (el orden es fijo, así que no hay saltos) y dibujo
    const globos = [];
    etiquetas.forEach((e) => {
      const { s, rr, h } = e, g = !!s.globo, pad = g ? 6 : 0, hb = h + pad * 2, wb = e.ancho + (g ? 10 : 0), sep = g ? 38 : 8;
      let ty = s.lado === 'abajo' ? s.y + rr + sep + e.ring : s.y - rr - sep - hb - e.ring, tx = s.x;
      // choca con otro nombre o, si es globo (línea guía), con el halo de otro grupo
      const choque = (x, y) => cajas.some((c) => x - wb / 2 - 4 < c.x1 && x + wb / 2 + 4 > c.x0 && y - 2 < c.y1 && y + hb > c.y0)
        || (g && arr.some((o) => o !== s && !o.abierta && o.al >= 0.15 && pisaHalo({ x0: x - wb / 2 - 4, x1: x + wb / 2 + 4, y0: y - 2, y1: y + hb }, o)));
      const ajusX = (x) => Math.max(wb / 2 + 6, Math.min(S.w - wb / 2 - 6, x));
      if (g) {
        // el globo se queda dentro de la pantalla (entre el buscador y la barra de niveles) y, si pisa otro, se corre hacia un lado
        const res = medirReservas(); let yMin = 6, yMax = S.h - 6;
        res.forEach((q) => { if (q.y1 < S.h / 3) yMin = Math.max(yMin, q.y1 + 4); else if (q.y0 > S.h * 2 / 3) yMax = Math.min(yMax, q.y0 - 4); });
        const ajusY = (y) => Math.max(yMin, Math.min(yMax - hb, y));
        const y1 = ajusY(ty); let hecho = false;
        const lejos = (x, y) => Math.hypot(Math.max(Math.abs(x - s.x) - wb / 2, 0), Math.max(Math.abs(y + hb / 2 - s.y) - hb / 2, 0)) - rr > 190;
        const dir = s.lado === 'arriba' ? -1 : 1;
        for (const v of [0, 1, 2, 3, -1, -2]) for (const m of [0, 1, -1, 2, -2, 3, -3, 4, -4, 5, -5, 6, -6]) {
          if (hecho) break;
          const x2 = ajusX(s.x + m * wb * 0.6), y2 = ajusY(y1 + dir * v * (hb + 8));
          if (!choque(x2, y2) && !lejos(x2, y2)) { tx = x2; ty = y2; hecho = true; }
        }
        // sin hueco libre cerca, se queda pegado a su círculo aunque pise algo (la línea guía no pasa de ~150 px)
        if (!hecho) { ty = y1; for (let it = 0; it < 8 && choque(tx, ty); it++) ty = ajusY(ty + (s.lado === 'arriba' ? -1 : 1) * (hb + 4)); }   // con línea guía puede alejarse un poco
      } else {
        for (const c of cajas) { if (tx - wb / 2 - 4 < c.x1 && tx + wb / 2 + 4 > c.x0 && ty - 2 < c.y1 && ty + hb > c.y0) ty = c.y1 + 2; }
      }
      tx = ajusX(tx);
      if (s.ly == null || Math.abs(s.ly - ty) > 160) { s.lx = tx; s.ly = ty; }   // la primera vez (o un salto grande) aparece directo
      else { s.lx += (tx - s.lx) * 0.25; s.ly += (ty - s.ly) * 0.25; if (Math.abs(tx - s.lx) + Math.abs(ty - s.ly) > 0.5) S.sucio = true; }
      cajas.push({ x0: s.lx - wb / 2 - 4, x1: s.lx + wb / 2 + 4, y0: s.ly - 2, y1: s.ly + hb });
      if (g) { globos.push({ e, wb, hb, pad }); return; }
      ctx.globalAlpha = (s.g.esc && S.maestriaSel && !S.troncalesSel ? Math.max(s.al, 0.85) : s.al) * fHalo(s);
      e.lineas.forEach((l, k) => {
        const y = s.ly + alto / 2 + k * alto;
        ctx.lineWidth = 4; ctx.strokeStyle = papel; ctx.strokeText(l, s.lx, y);
        ctx.fillStyle = tinta; ctx.fillText(l, s.lx, y);
      });
    });
    // los globos van al final (encima de todo): burbuja redondeada, punta hacia el círculo de quien es y el texto adentro
    globos.forEach(({ e, wb, hb, pad }) => {
      const { s, rr } = e, x0 = s.lx - wb / 2, y0 = s.ly, cx = s.lx, cy = y0 + hb / 2, col = colorGrupo(s.g), rgb = colorRgb(col);
      let ux = cx - s.x, uy = cy - s.y; const dl = Math.hypot(ux, uy) || 1; ux /= dl; uy /= dl;
      ctx.globalAlpha = (s.g.esc && S.maestriaSel && !S.troncalesSel ? Math.max(s.al, 0.85) : s.al) * fHalo(s);
      dibujarGuia({ x0, y0, wb, hb, tipx: s.x + ux * rr * 0.97, tipy: s.y + uy * rr * 0.97, color: 'rgba(' + rgb[0] + ',' + rgb[1] + ',' + rgb[2] + ',0.95)', punto: true });
      e.lineas.forEach((l, k) => {
        const y = y0 + pad + alto / 2 + k * alto;
        ctx.lineWidth = 4; ctx.strokeStyle = papel; ctx.strokeText(l, s.lx, y); ctx.fillStyle = tinta; ctx.fillText(l, s.lx, y);
      });
    });
    ctx.restore();
  }

  // En el nivel Cursos solo hay nombres de cursos: los más conectados y el que se toca o se elige. Los nombres generales
  // (maestrías, énfasis, programas, troncales) son de las capas de afuera y ya no salen aquí.
  // La COMBINACIÓN para titularse (solo al elegir una maestría en general, en la vista Cursos): categorías con sus cursos posibles
  // (ICG, laboratorio, transversales, disciplinares de un énfasis, proyectos, electivo) y la selección actual (S.ruta.set).
  function categoriasRuta(id) {
    const M = S.modelo, cats = [];
    const pool = (tipo, t) => {
      const cand = []; M.cursos.forEach((c, i) => { if (c.escuela && c.tipo === tipo) cand.push({ i }); });
      if (!cand.length) return;
      const pide = (M.cursos[cand[0].i].pool && M.cursos[cand[0].i].pool.pide) || 1;
      cats.push({ clave: tipo, t, req: Math.min(pide, cand.length), cand });
    };
    pool('icg', 'ICG'); pool('laboratorio', 'Laboratorio'); pool('transversal', 'Transversales');
    const bloques = M.bloques.filter((b) => b.maestriaId === id && !b.implicito && b.cursos.length), dcand = [];
    bloques.forEach((b) => b.cursos.forEach((cid) => { const j = M.porId.get(cid); if (j != null) dcand.push({ i: j, sub: (b.esEnfasis ? 'Énfasis en ' : 'Núcleo: ') + b.nombre, bid: b.id }); }));
    M.cursos.forEach((c, i) => { if (!c.escuela && c.progs.includes(id) && !c.bloquesReales.length && c.rels.some((r) => r.prog === id && r.rol === 'disciplinar' && r.obligatorio)) dcand.push({ i, sub: 'Cursos de la maestría' }); });
    if (dcand.length) cats.push({ clave: 'disciplinar', t: 'Cursos disciplinares', req: bloques.length ? Math.max(...bloques.map((b) => b.cursos.length)) : dcand.length, cand: dcand });
    const pc = []; M.cursos.forEach((c, i) => { if (!c.escuela && c.rels.some((r) => r.prog === id && r.rol === 'proyecto' && r.obligatorio)) pc.push({ i }); });
    if (pc.length) cats.push({ clave: 'proyecto', t: 'Proyectos', req: pc.length, cand: pc });
    const ec = []; M.cursos.forEach((c, i) => { if (!c.escuela && !c.progs.includes(id) && c.electivaProgs.has(id)) ec.push({ i }); });
    if (ec.length) cats.push({ clave: 'electivo', t: 'Electivo', req: 1, cand: ec });
    return cats;
  }
  // una combinación completa: la primera, o una al azar (otro énfasis, otro laboratorio, otras transversales, otro electivo)
  function combinacion(cats, azar) {
    const set = new Set(), rnd = (n) => (azar ? Math.floor(Math.random() * n) : 0);
    cats.forEach((cat) => {
      if (cat.clave === 'disciplinar') { const subs = [...new Set(cat.cand.map((c) => c.sub))], sub = subs[rnd(subs.length)]; cat.cand.filter((c) => c.sub === sub).slice(0, cat.req).forEach((c) => set.add(c.i)); }
      else if (cat.clave === 'proyecto') cat.cand.forEach((c) => set.add(c.i));
      else { const n = cat.cand.length, ini = rnd(n); const orden = cat.cand.map((_, k) => (ini + k) % n); if (azar) orden.sort(() => Math.random() - 0.5); orden.slice(0, cat.req).forEach((k) => set.add(cat.cand[k].i)); }
    });
    return set;
  }
  // solo con una maestría elegida en general y en la vista «Cursos»; la selección propia de la persona se conserva mientras siga en esa maestría
  function actualizarRuta() {
    const pr = S.maestriaSel && S.modelo.programaPorId.get(S.maestriaSel);
    if (!(pr && pr.nivel === 'maestria' && !S.bloqueSel && !S.troncalesSel && S.capa === 'cursos')) { S.ruta = null; return; }
    if (!S.ruta || S.ruta.id !== S.maestriaSel) { const cats = categoriasRuta(S.maestriaSel); S.ruta = { id: S.maestriaSel, cats, set: combinacion(cats, false) }; }
  }
  // luz de un punto: con una ruta a la vista, los cursos de la combinación brillan y el resto de la maestría queda tenue
  // Foco de hover: lo que señalas (un curso con su secuencia, o todos los cursos de un halo) queda con su luz y el resto se apaga.
  // S.hk sube y baja con suavidad (0..1); S.foco es el conjunto de cursos que se quedan encendidos.
  function actualizarFoco() {
    let clave = '', set = null;
    if (S.sel == null && S.pos && S.modelo) {
      if (S.hoverB && S.bur.get(S.hoverB.g.id) === S.hoverB) {
        clave = 'b:' + S.hoverB.g.id;
        if (S.focoClave !== clave) { const gr = grupos(); set = new Set(); for (let i = 0; i < S.modelo.cursos.length; i++) if (gr.lista[gr.de[i]] === S.hoverB.g) set.add(i); }
      } else if (S.hoverP) {
        // un elemento del panel Programas: se enciende lo suyo en la constelación
        clave = 'pn:' + S.hoverP;
        if (S.focoClave !== clave) {
          const k = S.hoverP, id = k.slice(2), M = S.modelo; set = new Set();
          if (k[0] === 'p') { M.cursos.forEach((c, i) => { if (!c.escuela && c.progs.includes(id)) set.add(i); }); S.focoHalo = (h) => !h.g.esc && h.g.maestriaId === id; }
          else if (k[0] === 'b') { const b = M.bloquePorId.get(id); if (b) b.cursos.forEach((cid) => { const j = M.porId.get(cid); if (j != null) set.add(j); }); S.focoHalo = (h) => h.g.id === 'b:' + id; }
          else { M.cursos.forEach((c, i) => { if (c.escuela) set.add(i); }); S.focoHalo = (h) => !!h.g.esc; }
        }
      } else if (S.hoverI != null) {
        clave = 'i:' + S.hoverI;
        if (S.focoClave !== clave) { const M = S.modelo, c = M.cursos[S.hoverI]; set = new Set([S.hoverI]); M.cursos.forEach((o, j) => { if (o.prereq.includes(c.id) || c.prereq.includes(o.id) || c.coreq.includes(o.id) || o.coreq.includes(c.id)) set.add(j); }); }
      }
    }
    if (clave && S.focoClave !== clave) { S.foco = set; S.focoClave = clave; S.focoTipo = clave[0]; S.focoB = clave[0] === 'b' ? S.hoverB : null; if (clave[0] !== 'p' || clave[1] !== 'n') S.focoHalo = null; }
    else if (!clave) S.focoClave = '';
    const meta = clave ? 1 : 0, d = meta - S.hk;
    if (Math.abs(d) > 0.01) { S.hk += d * (reducir ? 1 : meta ? 0.18 : 0.035); S.sucio = true; } else if (d) { S.hk = meta; S.sucio = true; if (!meta) { S.foco = null; S.focoB = null; S.focoHalo = null; } }
  }
  const fFoco = (i) => (S.foco && S.hk > 0 && !S.foco.has(i) ? 1 - 0.74 * S.hk : 1);
  const fHalo = (s) => (S.focoHalo && S.hk > 0 ? (S.focoHalo(s) ? 1 : 1 - 0.74 * S.hk) : S.focoB && S.hk > 0 ? (s === S.focoB ? 1 : 1 - 0.74 * S.hk) : 1);   // usa S.focoB (no hoverB) para que también se apague/encienda con transición al salir
  const enCombinacion = (i) => !!(S.ruta && S.sel == null && S.capa === 'cursos' && S.ruta.set.has(i));
  const luz = (i) => (S.ruta && S.sel == null && S.capa === 'cursos'
    ? S.av[i] * (S.ruta.set.has(i) ? 1 : (S.nivel[i] === 2 ? 0.4 : Math.min(S.gN[i], 0.4)))
    : S.vN[i]);
  const importantes = () => {
    const K = S.maestriaSel ? 6 : 8, res = [];
    for (const i of S.ordImp) {
      if (res.length >= K) break;
      if (S.modelo.cursos[i].escuela || S.ab[i] < 0.5 || S.av[i] < 0.5) continue;
      if (S.filtro.size && S.nivel[i] !== 2) continue;
      res.push(i);
    }
    return res;
  };
  // Estelas de los cursos troncales a cada maestría: siempre se ven donde haya troncales. Con algo elegido solo quedan las que le tocan:
  // las de la maestría elegida, todas si se eligen los troncales, o las del curso troncal elegido.
  const lineaTroncal = (pr, elegido) => {
    if (S.sel != null) return elegido;
    if (S.bloqueSel) return false;
    if (S.troncalesSel || !S.filtro.size) return true;
    return S.maestriaSel === pr.id;
  };
  // Con un curso seleccionado, las conexiones que no tienen que ver con él se desvanecen (S.fCon: 1 -> 0)
  // y reaparecen al deseleccionar. Relacionado = el curso elegido o los de su bloque (S.tB > 1).
  // Lo mismo si se enciende algo desde las convenciones (maestría, énfasis...): solo quedan las relaciones dentro de esa selección.
  const relCon = (i) => (S.sel != null ? (i === S.sel || S.tB[i] > 1 || !!(S.relSel && S.relSel.has(i))) : S.nivel[i] === 2);
  const fCon = (rel) => (rel ? 1 : S.fCon);
  const relDoc = (d) => (S.sel != null
    ? S.modelo.cursos[S.sel].progs.some((p) => p === d.id || p === d.maestriaBase)
    : S.filtro.has('p:' + d.id) || S.filtro.has('p:' + d.maestriaBase));
  // electivas: con un curso elegido, las suyas; con una selección (maestría, énfasis...), las de los cursos de esa selección hacia donde valgan
  const relElec = (i, d) => i === S.sel || (S.sel == null && (S.filtro.size === 0 || S.nivel[i] === 2));

  function pintar(t) {
    const M = S.modelo;
    const W = S.w, H = S.h, claro = S.tema === 'claro';
    ctx.setTransform(S.dpr, 0, 0, S.dpr, 0, 0);
    ctx.globalCompositeOperation = 'source-over';
    ctx.globalAlpha = 1;
    ctx.fillStyle = claro ? '#EFEEE8' : '#070707'; ctx.fillRect(0, 0, W, H);
    if (!M.cursos.length || !S.pos) return;

    S.mostrarCon = S.capa === 'cursos' || !!S.bloqueSel || S.troncalesSel;   // conexiones (prerrequisitos, electivas, doctorado): siempre, donde aplican
    actualizarAbre();
    proyectar(t);
    actualizarBurbujas();
    if (S.autoZoomAt && performance.now() >= S.autoZoomAt) {
      S.autoZoomAt = 0;
      if (S.centrarIdx) { centrarVistaEn(S.centrarIdx); S.centrarIdx = null; }   // solo se mueve el centro a lo elegido
      else ajustarVista(S.autoZoomIdx, S.capa === 'cursos' ? 3.4 : 1.7, null, S.centroPropios);
    }
    if (S.autoZoomAt2 && performance.now() >= S.autoZoomAt2) { S.autoZoomAt2 = 0; if (S.autoZoomIdx.length) ajustarVista(S.autoZoomIdx, S.capa === 'cursos' ? 3.4 : 1.7, null, S.centroPropios); }
    const n = M.cursos.length, R = S.pos.R;
    S.etqLineas = [];

    // intensidades (suavizadas hacia el objetivo)
    for (let i = 0; i < n; i++) {
      S.gM[i] += (S.tM[i] - S.gM[i]) * 0.12; S.gB[i] += (S.tB[i] - S.gB[i]) * 0.12; S.gN[i] += (S.tN[i] - S.gN[i]) * 0.12;
      if (Math.abs(S.tM[i] - S.gM[i]) > 0.003 || Math.abs(S.tN[i] - S.gN[i]) > 0.003) S.sucio = true;
      // intensidades visibles = intensidad suavizada × cuánto está abierto el curso (plegado = apagado)
      S.vM[i] = S.gM[i] * S.av[i]; S.vB[i] = S.gB[i] * S.av[i]; S.vN[i] = S.gN[i] * S.av[i];
    }

    const objCon = (S.sel != null || S.filtro.size > 0) ? 0 : 1;
    S.fCon += (objCon - S.fCon) * 0.12;
    if (Math.abs(objCon - S.fCon) > 0.003) S.sucio = true; else S.fCon = objCon;

    if (S.estilo === 'tecnico') { pintarTecnico(t); return; }

    // ---- mapa de calor: resplandor por maestría, foco por bloque, blanco de escuela
    // El calor es difuso: se dibuja en una capa de baja resolución y luego se estira. Cuesta una fracción
    // de pintar los resplandores a pantalla completa, sobre todo con zoom o en pantallas de alta densidad.
    const kc = Math.min(0.5, 520 / W), cw = Math.ceil(W * kc), chh = Math.ceil(H * kc);
    if (capaCalor.width !== cw || capaCalor.height !== chh) { capaCalor.width = cw; capaCalor.height = chh; }
    const hc = capaCtx;
    hc.setTransform(1, 0, 0, 1, 0, 0); hc.globalAlpha = 1; hc.globalCompositeOperation = 'source-over';
    if (claro) { hc.fillStyle = '#EFEEE8'; hc.fillRect(0, 0, cw, chh); } else hc.clearRect(0, 0, cw, chh);
    hc.setTransform(kc, 0, 0, kc, 0, 0);
    hc.globalCompositeOperation = claro ? 'multiply' : 'lighter';
    let hayCalor = false;
    const rProg = R * 0.36 * S.escala, rBloque = R * 0.15 * S.escala, rEsc = R * 0.5 * S.escala;
    // con zoom los resplandores se agrandan y se suman unos con otros: la intensidad baja para que no se queme la imagen
    const kz = 1 / (1 + 0.8 * Math.max(0, S.zoom - 1));
    const aProg = (claro ? 0.2 : 0.13) * kz, aBloque = (claro ? 0.24 : 0.17) * kz, aEsc = (claro ? 0.08 : 0.06) * kz;
    for (let i = 0; i < n; i++) {
      const c = M.cursos[i], p = S.sp[i];
      if (c.escuela) {
        const aE = aEsc * S.vM[i] * fFoco(i); if (aE < 0.004) continue;
        hayCalor = true; hc.globalAlpha = aE;
        const r = rEsc * p; hc.drawImage(sprite(colorEscuelaVivo()), S.sx[i] - r, S.sy[i] - r, r * 2, r * 2);
        continue;
      }
      const k = c.progs.length || 1;
      c.progs.forEach((pid) => {
        const pr = M.programaPorId.get(pid); if (!pr) return;
        const aP = Math.min(1, aProg * S.vM[i] * fFoco(i) / Math.sqrt(k)); if (aP < 0.004) return;
        hayCalor = true; hc.globalAlpha = aP;
        const r = rProg * p; hc.drawImage(sprite(pr.color), S.sx[i] - r, S.sy[i] - r, r * 2, r * 2);
      });
      if (c.bloquesReales.length) {
        const pr = M.programaPorId.get(c.progs[0]);
        if (pr) {
          const aB = Math.min(1, aBloque * S.vB[i] * fFoco(i));
          if (aB >= 0.004) { hayCalor = true; hc.globalAlpha = aB; const r = rBloque * p; hc.drawImage(sprite(pr.color), S.sx[i] - r, S.sy[i] - r, r * 2, r * 2); }
        }
      }
    }

    if (hayCalor) {
      ctx.globalAlpha = 1; ctx.globalCompositeOperation = claro ? 'source-over' : 'lighter';
      ctx.drawImage(capaCalor, 0, 0, W, H);
    }

    // ---- estelas: cursos lejos del centroide de un grupo al que pertenecen
    const cent = calcularCentroides(M);
    // al encender una maestría, los troncales de toda la escuela toman su color (van con ella)
    const prEsc = S.programaEscuela ? M.programaPorId.get(S.programaEscuela) : null;
    dibujarBurbujas(false);
    dibujarHaloTroncales(false);
    const colEsc = colorEscuelaVivo();
    ctx.lineCap = 'round';
    M.cursos.forEach((c, i) => {
      if (c.escuela) return;
      c.progs.forEach((pid, k) => {
        const g = cent.get(pid); if (!g) return;
        const d = Math.hypot(S.pos.X[i] - g.cx, S.pos.Y[i] - g.cy, (S.pos.Z[i] - g.cz) * 0.6);
        if (d <= g.dm * 1.6 + 12) return;
        const pr = M.programaPorId.get(pid);
        const x0 = S.sx[i], y0 = S.sy[i], x1 = g.sx, y1 = g.sy;
        const dx = x1 - x0, dy = y1 - y0, L = Math.hypot(dx, dy) || 1;
        const sway = S.vaiven && !reducir ? Math.sin(t * 0.0003 + i) * 0.04 : 0;
        const off = (0.2 + sway) * L * (k % 2 ? -1 : 1);
        const qx = (x0 + x1) / 2 - dy / L * off, qy = (y0 + y1) / 2 + dx / L * off;
        const rgb = colorRgb(pr.color);
        const a = Math.min(1, S.vM[i] * 0.9) * fCon(relCon(i)) * S.ab[i];
        const gr = ctx.createLinearGradient(x0, y0, x1, y1);
        gr.addColorStop(0, 'rgba(' + rgb + ',' + (0.8 * a) + ')');
        gr.addColorStop(1, 'rgba(' + rgb + ',0)');
        ctx.globalAlpha = 1; ctx.strokeStyle = gr;
        ctx.lineWidth = 9; ctx.globalCompositeOperation = claro ? 'multiply' : 'lighter';
        ctx.beginPath(); ctx.moveTo(x0, y0); ctx.quadraticCurveTo(qx, qy, x1, y1); ctx.globalAlpha = 0.22; ctx.stroke();
        ctx.lineWidth = 2.4; ctx.globalAlpha = 1;
        ctx.beginPath(); ctx.moveTo(x0, y0); ctx.quadraticCurveTo(qx, qy, x1, y1); ctx.stroke();
      });
    });

    // Unión entre dos puntos: «estela» (curva que se desvanece, con un halo suave) o «recta».
    // Se dibuja con el modo de mezcla y el rgb que se le pasen; alfa y ancho ya vienen calculados.
    const unir = (x0, y0, x1, y1, rgb, alfa, ancho, k, dash) => {
      if (alfa < 0.012) return null; // casi invisible: no vale el costo de dibujarla
      ctx.setLineDash(dash || []);
      if (S.modoUnion === 'recta') {
        ctx.globalAlpha = alfa; ctx.strokeStyle = 'rgb(' + rgb + ')'; ctx.lineWidth = ancho; ctx.lineCap = dash ? 'round' : 'butt';
        ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.stroke();
        ctx.setLineDash([]);
        return [(x0 + x1) / 2, (y0 + y1) / 2];
      }
      const dx = x1 - x0, dy = y1 - y0, L = Math.hypot(dx, dy) || 1;
      const sway = S.vaiven && !reducir ? Math.sin(t * 0.0003 + k) * 0.03 : 0;
      const off = (0.16 + sway) * L * (k % 2 ? -1 : 1);
      const qx = (x0 + x1) / 2 - dy / L * off, qy = (y0 + y1) / 2 + dx / L * off;
      const gr = ctx.createLinearGradient(x0, y0, x1, y1);
      gr.addColorStop(0, 'rgba(' + rgb + ',1)'); gr.addColorStop(1, 'rgba(' + rgb + ',' + (dash ? 0.55 : 0) + ')');
      ctx.strokeStyle = gr; ctx.lineCap = 'round';
      ctx.beginPath(); ctx.moveTo(x0, y0); ctx.quadraticCurveTo(qx, qy, x1, y1);
      if (!dash && alfa > 0.04) { ctx.lineWidth = ancho * 4.5; ctx.globalAlpha = alfa * 0.28; ctx.stroke(); }   // halo difuminado
      ctx.lineWidth = ancho * 1.3; ctx.globalAlpha = alfa; ctx.stroke();           // núcleo de la estela
      ctx.setLineDash([]);
      return [qx, qy];
    };
    // flecha de secuencia: punta cerca del curso de destino, orientada según el final de la línea
    const flechaN = (qx, qy, x1, y1, rgb, alfa, radio) => {
      const dx = x1 - qx, dy = y1 - qy, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
      const px = x1 - ux * (radio + 3), py = y1 - uy * (radio + 3);
      ctx.globalAlpha = Math.min(1, alfa * 1.25); ctx.fillStyle = 'rgb(' + rgb + ')';
      ctx.beginPath(); ctx.moveTo(px, py);
      ctx.lineTo(px - ux * 8 - uy * 4, py - uy * 8 + ux * 4); ctx.lineTo(px - ux * 8 + uy * 4, py - uy * 8 - ux * 4);
      ctx.closePath(); ctx.fill();
    };

    // ---- doctorado: se desprende de su maestría base (vínculo siempre visible)
    ctx.globalCompositeOperation = claro ? 'multiply' : 'lighter';
    M.doctorados.forEach((d) => {
      if (!S.mostrarCon) return;   // el vínculo del doctorado con su maestría
      const gd = cent.get(d.id), gb = cent.get(d.maestriaBase), pb = M.programaPorId.get(d.maestriaBase);
      if (!gd || !gb || !pb) return;
      const media = (mem) => mem.reduce((sum, i) => sum + Math.min(S.vN[i], 1), 0) / mem.length;
      const vis = Math.sqrt(media(gd.mem) * media(gb.mem)) * fCon(relDoc(d));
      const dx = gb.sx - gd.sx, dy = gb.sy - gd.sy, L = Math.hypot(dx, dy) || 1;
      const curva = S.modoUnion === 'recta' ? 0 : 0.12;
      const qx = (gd.sx + gb.sx) / 2 - dy / L * L * curva, qy = (gd.sy + gb.sy) / 2 + dx / L * L * curva;
      const c1 = colorRgb(d.color), c2 = colorRgb(pb.color);
      const gr = ctx.createLinearGradient(gd.sx, gd.sy, gb.sx, gb.sy);
      gr.addColorStop(0, 'rgba(' + c1 + ',0.75)'); gr.addColorStop(1, 'rgba(' + c2 + ',0.75)');
      ctx.strokeStyle = gr; ctx.setLineDash([2, 5]);
      ctx.lineWidth = 1.6; ctx.globalAlpha = 0.8 * vis;
      ctx.beginPath(); ctx.moveTo(gd.sx, gd.sy); ctx.quadraticCurveTo(qx, qy, gb.sx, gb.sy); ctx.stroke();
      ctx.setLineDash([]);
    });

    // ---- relaciones siempre visibles (se atenúan con la selección)
    ctx.globalCompositeOperation = claro ? 'multiply' : 'lighter';
    (S.enlaces || []).forEach((e) => {
      if ((e.tipo === 'correq' || e.tipo === 'cadena') && !S.mostrarCon && e.a !== S.sel && e.b !== S.sel) return;   // «qué se cursa primero»
      const pr = M.programaPorId.get(e.prog);
      const rgb = colorRgb(pr ? pr.color : '#EFEEE8');
      const vis = Math.min(S.vN[e.a], S.vN[e.b], 1) * fCon(relCon(e.a) && relCon(e.b));
      const pa = Math.min(S.sp[e.a], S.sp[e.b]);
      const x0 = S.sx[e.a], y0 = S.sy[e.a], x1 = S.sx[e.b], y1 = S.sy[e.b];
      if (e.tipo === 'correq') {
        // correquisito: se cursan juntos = doble línea, sin flecha
        const alfa = 0.8 * vis * Math.min(1, pa + 0.1), L = Math.hypot(x1 - x0, y1 - y0) || 1;
        const nx = -(y1 - y0) / L * 2.4, ny = (x1 - x0) / L * 2.4;
        unir(x0 + nx, y0 + ny, x1 + nx, y1 + ny, rgb, alfa, 1.1, e.a);
        unir(x0 - nx, y0 - ny, x1 - nx, y1 - ny, rgb, alfa, 1.1, e.a);
        return;
      }
      const cadena = e.tipo === 'cadena';
      const alfa = (cadena ? 0.85 : 0.45) * vis * Math.min(1, pa + 0.1);
      const q = unir(x0, y0, x1, y1, rgb, alfa, cadena ? 1.6 : 1, e.a + e.b);
      // secuencia: prerrequisito = flecha del curso de origen al de destino
      if (cadena && q) flechaN(q[0], q[1], x1, y1, rgb, alfa, TAM_PUNTO[M.cursos[e.b].tam] * Math.pow(S.zoom, 0.35) * S.sp[e.b]);
    });

    // troncales de toda la escuela: conforman TODAS las maestrías, así que cada uno se une con línea continua al punto
    // central de cada una (el color de la línea es el de la maestría). Con una maestría encendida esa línea ya se dibuja abajo.
    M.cursos.forEach((c, i) => {
      if (!c.escuela) return;
      const elegido = i === S.sel, vis = Math.min(S.vN[i], 1);
      M.maestrias.forEach((pr, k) => {
        if (!lineaTroncal(pr, elegido)) return;   // estelas de los cursos troncales a las maestrías que conforman
        const g = cent.get(pr.id); if (!g) return;
        const alfa = (elegido ? 0.85 : 0.26) * vis * (S.hubClave ? 0.35 : 1);
        unir(S.sx[i], S.sy[i], g.sx, g.sy, colorRgb(pr.color), alfa, elegido ? 1.3 : 0.9, i + k);
      });
    });

    // constitución: línea continua de cada curso que conforma el programa encendido a su punto central
    if (S.hubClave && S.lineasConstitucion) {
      const k = S.hubClave, id = k.slice(2);
      const hub = k[0] === 'b' ? cent.bloques.get(id) : cent.get(id);
      const pr = k[0] === 'p' ? M.programaPorId.get(id) : M.programaPorId.get((M.bloquePorId.get(id) || {}).maestriaId);
      if (hub && pr) {
        const rgb = colorRgb(pr.color);
        S.autoZoomIdx.forEach((i) => unir(S.sx[i], S.sy[i], hub.sx, hub.sy, rgb, 0.75 * Math.min(S.vN[i], 1) * fCon(relCon(i)), 1.2, i));
      }
    }

    // electivas: línea punteada del curso al PUNTO CENTRAL del programa, énfasis o especializado donde vale
    // (no a otros cursos). Cada línea dice a qué apunta cuando el curso está seleccionado.
    if (S.sel != null) {   // las líneas de electiva solo salen al tocar un curso
      M.cursos.forEach((c, i) => {
        if (!c.electivaDestinos.length) return;
        const elegido = i === S.sel;
        if (!elegido) return;
        const vis = Math.min(S.vN[i], 1);
        c.electivaDestinos.forEach((d) => {
          const g = destinoCentro(cent, d), pr = M.programaPorId.get(d.maestria);
          if (!g || !pr) return;
          const q = unir(S.sx[i], S.sy[i], g.sx, g.sy, colorRgb(pr.color), (elegido ? 0.95 : 0.2) * vis * fCon(relElec(i, d)), elegido ? 1.3 : 0.9, i, [1.5, 5]);
          if (elegido && q) S.etqLineas.push({ x: (S.sx[i] + 2 * q[0] + g.sx) / 4, y: (S.sy[i] + 2 * q[1] + g.sy) / 4, texto: d.etiqueta });
        });
      });
    }

    // puntos centrales (hubs): un aro pequeño en el centro de cada maestría; más grande el del programa encendido
    // y los de los destinos de la electiva seleccionada
    ctx.globalCompositeOperation = 'source-over';
    const hubs = [];
    cent.forEach((g, pid) => hubs.push({ g, pr: M.programaPorId.get(pid), grande: S.hubClave === 'p:' + pid }));
    if (S.sel != null) M.cursos[S.sel].electivaDestinos.forEach((d) => {
      if (d.tipo === 'bloque') { const g = cent.bloques.get(d.id); if (g) hubs.push({ g, pr: M.programaPorId.get(d.maestria), grande: true }); }
      else { const h = hubs.find((x) => x.pr && x.pr.id === d.id); if (h) h.grande = true; }
    });
    if (S.hubClave && S.hubClave[0] === 'b') { const g = cent.bloques.get(S.hubClave.slice(2)); if (g) hubs.push({ g, pr: M.programaPorId.get((M.bloquePorId.get(S.hubClave.slice(2)) || {}).maestriaId), grande: true }); }
    hubs.forEach((h) => {
      if (!h.pr) return;
      const media = h.g.mem.reduce((sum, i) => sum + Math.min(S.vN[i], 1), 0) / h.g.mem.length;
      const a = Math.min(1, 0.18 + 0.7 * media + (h.grande ? 0.3 : 0));
      ctx.globalAlpha = a; ctx.strokeStyle = h.pr.color; ctx.fillStyle = h.pr.color; ctx.lineWidth = 1.3;
      const r = (h.grande ? 7.5 : 4.5) * Math.pow(S.zoom, 0.3);
      ctx.beginPath(); ctx.arc(h.g.sx, h.g.sy, r, 0, 6.2832); ctx.stroke();
      ctx.beginPath(); ctx.arc(h.g.sx, h.g.sy, Math.max(1.4, r * 0.28), 0, 6.2832); ctx.fill();
    });

    // ---- líneas de selección
    ctx.globalCompositeOperation = 'source-over';
    if (S.sel != null && S.lineas.length) {
      const a = S.sel;
      const rgbSel = colorRgb(claro ? '#070707' : '#EFEEE8');
      S.lineas.forEach((l) => unir(S.sx[a], S.sy[a], S.sx[l.b], S.sy[l.b], rgbSel, l.w * Math.min(S.vN[l.b], 1), 0.8, l.b));
    }

    // ---- nodos, de atrás hacia adelante
    const orden = Array.from({ length: n }, (_, i) => i).sort((a, b) => S.sz[b] - S.sz[a]);
    const tw = !reducir && S.vaiven;
    ctx.globalCompositeOperation = 'source-over';
    orden.forEach((i) => {
      if (S.av[i] < 0.01) return;
      const c = M.cursos[i], p = S.sp[i];
      const base = TAM_PUNTO[c.tam] * Math.pow(S.zoom, 0.35);
      const r = base * Math.pow(p, 1.1) * (S.w < 560 ? 1.1 : 1);
      const col = c.escuela ? colEsc : colorNodo((M.programaPorId.get(c.progs[0]) || { color: '#EFEEE8' }).color);
      const profundo = Math.max(0.55, Math.min(1.1, p));
      const parpadeo = tw ? 0.9 + 0.1 * Math.sin(t * 0.0009 + S.fase[i] * 3) : 1;
      const a = Math.min(1, (0.14 + 0.86 * Math.min(luz(i), 1)) * profundo * parpadeo) * S.av[i] * fFoco(i);
      // halo
      ctx.globalCompositeOperation = claro ? 'source-over' : 'lighter';
      ctx.globalAlpha = a * (claro ? 0.2 : 0.32) * Math.sqrt(kz);
      const h = r * 4.2; ctx.drawImage(sprite(col), S.sx[i] - h, S.sy[i] - h, h * 2, h * 2);
      ctx.globalCompositeOperation = 'source-over';
      ctx.globalAlpha = a;
      ctx.fillStyle = col;
      ctx.beginPath(); ctx.arc(S.sx[i], S.sy[i], r, 0, 6.2832); ctx.fill();
      // los cursos de la combinación llevan un aro y un halo más fuerte: se leen de inmediato entre los tenues
      if (enCombinacion(i)) {
        ctx.globalAlpha = Math.min(1, S.av[i]);
        ctx.strokeStyle = claro ? '#070707' : '#EFEEE8'; ctx.lineWidth = 2.2;
        ctx.beginPath(); ctx.arc(S.sx[i], S.sy[i], r + 5, 0, 6.2832); ctx.stroke();
        ctx.globalCompositeOperation = claro ? 'source-over' : 'lighter'; ctx.globalAlpha = 0.5 * Math.min(1, S.av[i]);
        const hh = r * 3.2; ctx.drawImage(sprite(col), S.sx[i] - hh, S.sy[i] - hh, hh * 2, hh * 2);
        ctx.globalCompositeOperation = 'source-over';
      }
    });

    // ---- nombres de los cursos: el curso seleccionado y los iluminados por una convención
    ctx.globalCompositeOperation = 'source-over';
    ctx.font = '700 15px ' + FAMILIA;
    ctx.textBaseline = 'middle';
    ctx.lineJoin = 'round';
    const cajas = []; S.zonasEt = [];
    const colTexto = claro ? '#070707' : '#EFEEE8', colFondo = claro ? '#EFEEE8' : '#070707';
    // intenta poner la etiqueta a la derecha o a la izquierda del punto, sin pisar otras ya dibujadas
    const etiquetar = (i, alfa, obligatoria, pref) => {
      const nombre = M.cursos[i].nombre;
      let ancho = anchosEtiqueta.get(nombre);
      if (ancho == null) { ancho = ctx.measureText(nombre).width; anchosEtiqueta.set(nombre, ancho); }
      globoCurso(i, nombre, ancho, S.sx[i], S.sy[i], alfa, obligatoria, cajas, W, H, claro, i);
    };
    if (S.sel != null) etiquetar(S.sel, 1, true);
    // los troncales de toda la escuela solo llevan nombre al seleccionarlos o al encender una maestría (van con ella); se rotulan hacia afuera del grupo
    const troncales = []; let tcx = 0, tcy = 0;
    M.cursos.forEach((c, i) => { if (c.escuela) { troncales.push(i); tcx += S.sx[i]; tcy += S.sy[i]; } });
    if (troncales.length) { tcx /= troncales.length; tcy /= troncales.length; }
    troncales.forEach((i) => {
      return;   // los nombres de los troncales salen al pasar el cursor o al elegirlos
      const alfa = Math.min(1, Math.min(S.vN[i], 1) * 1.6) * S.ab[i];
      if (alfa < 0.08) return;
      const dx = S.sx[i] - tcx, dy = S.sy[i] - tcy, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
      etiquetar(i, alfa, false, (ancho) => [S.sx[i] + ux * 13 - (ux < 0 ? ancho : 0), S.sy[i] + uy * 13]);
    });
    // cada línea de electiva dice a qué conexión apunta
    S.etqLineas.forEach((e) => {
      let ancho = anchosEtiqueta.get(e.texto);
      if (ancho == null) { ancho = ctx.measureText(e.texto).width; anchosEtiqueta.set(e.texto, ancho); }
      const tx = Math.max(6, Math.min(W - ancho - 6, e.x - ancho / 2)), ty = Math.max(10, Math.min(H - 10, e.y));
      cajas.push({ x0: tx - 4, y0: ty - 9, x1: tx + ancho + 4, y1: ty + 9 });
      ctx.globalAlpha = 1; ctx.lineWidth = 3.5; ctx.strokeStyle = colFondo; ctx.strokeText(e.texto, tx, ty);
      ctx.fillStyle = colTexto; ctx.fillText(e.texto, tx, ty);
    });
    // Pocos nombres a la vez: al elegir un curso salen el suyo y los de su bloque; el resto se lee al pasar el cursor
    if (S.sel != null) { /* con un curso elegido solo lleva nombre ese curso; sus relaciones se ven con líneas */ }
    // con un énfasis, un programa especializado o los troncales elegidos (pocos cursos) sí se leen todos sus nombres
    else if (S.bloqueSel || S.troncalesSel) orden.slice().reverse().forEach((i) => { if (S.nivel[i] === 2 && S.ab[i] > 0.4) etiquetar(i, Math.min(1, S.vN[i]) * S.ab[i], false); });
    else if (S.capa === 'cursos' && S.ruta) [...S.ruta.set].forEach((i) => etiquetar(i, 0.95 * S.ab[i], false));   // los de la combinación que da el título
    else if (S.capa === 'cursos' && S.maestriaSel) importantes().forEach((i) => etiquetar(i, 0.9 * S.ab[i], false));   // los más conectados llevan nombre
    if (S.hoverI != null && S.sel == null && S.av[S.hoverI] > 0.5 && S.ab[S.hoverI] > 0.5) etiquetar(S.hoverI, 1, true);   // y el que tocas con el cursor

    // ---- marca de selección
    if (S.sel != null) {
      const i = S.sel;
      ctx.globalAlpha = 0.9; ctx.lineWidth = 1.2;
      ctx.strokeStyle = colTexto;
      ctx.beginPath(); ctx.arc(S.sx[i], S.sy[i], 9 * S.sp[i] + 2, 0, 6.2832); ctx.stroke();
    }
    ctx.globalAlpha = 1; ctx.globalCompositeOperation = 'source-over';
  }


  /* ---------------- estilo técnico: plano, solo negro y blanco de la paleta ---------------- */
  const anchosTec = new Map();
  function pintarTecnico(t) {
    const M = S.modelo, W = S.w, H = S.h, claro = S.tema === 'claro';
    const n = M.cursos.length;
    const tinta = claro ? '#070707' : '#EFEEE8', papel = claro ? '#EFEEE8' : '#070707';
    const cent = calcularCentroides(M);
    ctx.globalCompositeOperation = 'source-over';
    ctx.lineCap = 'butt'; ctx.lineJoin = 'round';
    ctx.strokeStyle = tinta; ctx.fillStyle = tinta; ctx.setLineDash([]);
    ctx.font = '12px ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace';
    ctx.textBaseline = 'middle';
    if ('letterSpacing' in ctx) ctx.letterSpacing = '0.06em';

    // marco: marcas en el centro de cada borde y cruz punteada en el centro del lienzo
    ctx.lineWidth = 1.2; ctx.globalAlpha = 0.9;
    ctx.beginPath();
    ctx.moveTo(W / 2, 0); ctx.lineTo(W / 2, 14); ctx.moveTo(W / 2, H); ctx.lineTo(W / 2, H - 14);
    ctx.moveTo(0, H / 2); ctx.lineTo(14, H / 2); ctx.moveTo(W, H / 2); ctx.lineTo(W - 14, H / 2);
    ctx.stroke();
    ctx.setLineDash([1.5, 5]); ctx.lineWidth = 1; ctx.globalAlpha = 0.28;
    ctx.beginPath(); ctx.moveTo(14, H / 2); ctx.lineTo(W - 14, H / 2); ctx.moveTo(W / 2, 14); ctx.lineTo(W / 2, H - 14); ctx.stroke();
    ctx.setLineDash([]);

    // punto de control de una curva suave entre dos puntos (o el punto medio si las uniones son rectas)
    const ctrl = (x0, y0, x1, y1, k, f) => {
      const dx = x1 - x0, dy = y1 - y0, L = Math.hypot(dx, dy) || 1;
      if (S.modoUnion === 'recta') return [(x0 + x1) / 2, (y0 + y1) / 2];
      const sway = S.vaiven && !reducir ? Math.sin(t * 0.0003 + k) * 0.03 : 0;
      const off = (f + sway) * L * (k % 2 ? -1 : 1);
      return [(x0 + x1) / 2 - dy / L * off, (y0 + y1) / 2 + dx / L * off];
    };
    const trazo = (x0, y0, x1, y1, k, f) => {
      const [qx, qy] = ctrl(x0, y0, x1, y1, k, f);
      ctx.beginPath(); ctx.moveTo(x0, y0);
      if (S.modoUnion === 'recta') ctx.lineTo(x1, y1); else ctx.quadraticCurveTo(qx, qy, x1, y1);
      ctx.stroke();
      return [qx, qy];
    };
    const flecha = (qx, qy, x1, y1, tam) => {
      const a = Math.atan2(y1 - qy, x1 - qx);
      ctx.beginPath(); ctx.moveTo(x1, y1);
      ctx.lineTo(x1 - Math.cos(a - 0.4) * tam, y1 - Math.sin(a - 0.4) * tam);
      ctx.lineTo(x1 - Math.cos(a + 0.4) * tam, y1 - Math.sin(a + 0.4) * tam);
      ctx.closePath(); ctx.fill();
    };

    dibujarBurbujas(true);
    dibujarHaloTroncales(true);
    // órbitas: una elipse punteada alrededor de cada maestría; sólida cuando está activa
    const activos = [];
    cent.forEach((g, pid) => {
      let x0 = 1e9, x1 = -1e9, y0 = 1e9, y1 = -1e9, gn = 0;
      g.mem.forEach((i) => { x0 = Math.min(x0, S.sx[i]); x1 = Math.max(x1, S.sx[i]); y0 = Math.min(y0, S.sy[i]); y1 = Math.max(y1, S.sy[i]); gn += Math.min(S.vN[i], 1); });
      gn /= g.mem.length;
      const ex = (x1 - x0) / 2 * 1.12 + 16, ey = (y1 - y0) / 2 * 1.12 + 16, ox = (x0 + x1) / 2, oy = (y0 + y1) / 2;
      const activo = S.filtro.has('p:' + pid);
      ctx.setLineDash(activo ? [] : [2, 6]); ctx.lineWidth = activo ? 1.5 : 0.9;
      ctx.globalAlpha = Math.min(1, 0.14 + 0.55 * gn);
      ctx.beginPath(); ctx.ellipse(ox, oy, ex, ey, 0, 0, 6.2832); ctx.stroke();
      if (activo) activos.push({ nombre: M.programaPorId.get(pid).nombre, x: ox, y: oy - ey - 12 });
    });
    ctx.setLineDash([]);

    // uniones: mallas de bloque (finas), prerrequisitos (con flecha), correquisitos (doble línea) y electivas (punteadas)
    ctx.lineCap = 'round';
    (S.enlaces || []).forEach((e) => {
      if ((e.tipo === 'correq' || e.tipo === 'cadena') && !S.mostrarCon && e.a !== S.sel && e.b !== S.sel) return;   // «qué se cursa primero»
      const vis = Math.min(S.vN[e.a], S.vN[e.b], 1) * fCon(relCon(e.a) && relCon(e.b));
      const x0 = S.sx[e.a], y0 = S.sy[e.a], x1 = S.sx[e.b], y1 = S.sy[e.b];
      if (e.tipo === 'correq') {
        const alfaC = 0.85 * vis; if (alfaC < 0.03) return;
        const L = Math.hypot(x1 - x0, y1 - y0) || 1, nx = -(y1 - y0) / L * 2.4, ny = (x1 - x0) / L * 2.4;
        ctx.globalAlpha = alfaC; ctx.lineWidth = 0.9;
        trazo(x0 + nx, y0 + ny, x1 + nx, y1 + ny, e.a, 0.04); trazo(x0 - nx, y0 - ny, x1 - nx, y1 - ny, e.a, 0.04);
        return;
      }
      const cadena = e.tipo === 'cadena';
      const alfa = (cadena ? 0.9 : 0.55) * vis;
      if (alfa < 0.03) return;
      ctx.globalAlpha = alfa; ctx.lineWidth = cadena ? 1.3 : 0.8;
      const [qx, qy] = trazo(x0, y0, x1, y1, e.a + e.b, 0.12);
      if (cadena) { const dxe = x1 - qx, dye = y1 - qy, Le = Math.hypot(dxe, dye) || 1; flecha(qx, qy, x1 - dxe / Le * 9, y1 - dye / Le * 9, 6); }
    });
    // troncales de toda la escuela: línea continua al punto central de cada maestría que conforman
    M.cursos.forEach((c, i) => {
      if (!c.escuela) return;
      const elegido = i === S.sel, vis = Math.min(S.vN[i], 1);
      M.maestrias.forEach((pr, k) => {
        if (!lineaTroncal(pr, elegido)) return;   // estelas de los cursos troncales a las maestrías que conforman
        const g = cent.get(pr.id); if (!g) return;
        ctx.lineWidth = elegido ? 1.2 : 0.8; ctx.globalAlpha = (elegido ? 0.9 : 0.4) * vis * (S.hubClave ? 0.35 : 1);
        trazo(S.sx[i], S.sy[i], g.sx, g.sy, i + k, 0.1);
      });
    });
    // constitución: línea continua de cada curso del programa encendido a su punto central
    const rotulosLinea = [];
    if (S.hubClave && S.lineasConstitucion) {
      const k = S.hubClave, hub = k[0] === 'b' ? cent.bloques.get(k.slice(2)) : cent.get(k.slice(2));
      if (hub) { ctx.lineWidth = 1.1; S.autoZoomIdx.forEach((i) => { ctx.globalAlpha = 0.75 * Math.min(S.vN[i], 1) * fCon(relCon(i)); trazo(S.sx[i], S.sy[i], hub.sx, hub.sy, i, 0.1); }); }
    }
    // electivas: línea punteada al PUNTO CENTRAL del destino; la del curso seleccionado dice a qué apunta
    if (S.sel != null) {   // las líneas de electiva solo salen al tocar un curso
      ctx.setLineDash([1, 4]);
      M.cursos.forEach((c, i) => {
        if (!c.electivaDestinos.length) return;
        const elegido = i === S.sel;
        if (!elegido) return;
        const vis = Math.min(S.vN[i], 1);
        c.electivaDestinos.forEach((d) => {
          const g = destinoCentro(cent, d); if (!g) return;
          ctx.lineWidth = elegido ? 1.2 : 0.8; ctx.globalAlpha = (elegido ? 0.95 : 0.34) * vis * fCon(relElec(i, d));
          const [qx, qy] = trazo(S.sx[i], S.sy[i], g.sx, g.sy, i, 0.1);
          if (elegido) rotulosLinea.push({ x: (S.sx[i] + 2 * qx + g.sx) / 4, y: (S.sy[i] + 2 * qy + g.sy) / 4, texto: d.etiqueta });
        });
      });
      ctx.setLineDash([]);
    }
    // cursos obligatorios en dos maestrías: línea punteada con flecha hacia el grupo que les queda lejos
    ctx.setLineDash([3, 4]); ctx.lineWidth = 1;
    M.cursos.forEach((c, i) => {
      if (c.escuela) return;
      c.progs.forEach((pid, k) => {
        const g = cent.get(pid); if (!g) return;
        const d = Math.hypot(S.pos.X[i] - g.cx, S.pos.Y[i] - g.cy, (S.pos.Z[i] - g.cz) * 0.6);
        if (d <= g.dm * 1.6 + 12) return;
        ctx.globalAlpha = Math.min(1, S.vN[i]) * 0.8 * fCon(relCon(i));
        const [qx, qy] = trazo(S.sx[i], S.sy[i], g.sx, g.sy, i + k, 0.2);
        ctx.setLineDash([]); flecha(qx, qy, g.sx, g.sy, 7); ctx.setLineDash([3, 4]);
      });
    });
    // doctorado: vínculo con su maestría base
    ctx.setLineDash([7, 4]); ctx.lineWidth = 1.2;
    M.doctorados.forEach((d) => {
      if (!S.mostrarCon) return;
      const gd = cent.get(d.id), gb = cent.get(d.maestriaBase);
      if (!gd || !gb) return;
      ctx.globalAlpha = 0.7 * Math.min(1, S.vN[gd.mem[0]] + 0.3) * fCon(relDoc(d));
      trazo(gd.sx, gd.sy, gb.sx, gb.sy, 3, 0.1);
    });
    ctx.setLineDash([]);
    // líneas de la selección
    if (S.sel != null && S.lineas.length) {
      ctx.lineWidth = 0.9;
      S.lineas.forEach((l) => { ctx.globalAlpha = Math.min(1, l.w * 1.4) * Math.min(S.vN[l.b], 1); trazo(S.sx[S.sel], S.sy[S.sel], S.sx[l.b], S.sy[l.b], l.b, 0.1); });
    }

    // puntos centrales: aro con punto en el centro de cada maestría (y del énfasis / programa encendido)
    ctx.lineWidth = 1.1; ctx.setLineDash([]);
    cent.forEach((g, pid) => {
      const media = g.mem.reduce((sum, i) => sum + Math.min(S.vN[i], 1), 0) / g.mem.length, grande = S.hubClave === 'p:' + pid;
      ctx.globalAlpha = Math.min(1, 0.2 + 0.7 * media + (grande ? 0.3 : 0));
      const r = grande ? 7.5 : 4.5;
      ctx.fillStyle = papel; ctx.beginPath(); ctx.arc(g.sx, g.sy, r, 0, 6.2832); ctx.fill(); ctx.stroke();
      ctx.fillStyle = tinta; ctx.beginPath(); ctx.arc(g.sx, g.sy, 1.5, 0, 6.2832); ctx.fill();
    });
    if (S.sel != null) M.cursos[S.sel].electivaDestinos.forEach((d) => {
      if (d.tipo !== 'bloque') return;
      const g = cent.bloques.get(d.id); if (!g) return;
      ctx.globalAlpha = 1; ctx.fillStyle = papel; ctx.beginPath(); ctx.arc(g.sx, g.sy, 7, 0, 6.2832); ctx.fill(); ctx.stroke();
      ctx.fillStyle = tinta; ctx.beginPath(); ctx.arc(g.sx, g.sy, 1.8, 0, 6.2832); ctx.fill();
    });

    // puntos: toda la escuela = círculo abierto con cruz; general = punto; específico = punto pequeño
    const orden = Array.from({ length: n }, (_, i) => i).sort((a, b) => S.sz[b] - S.sz[a]);
    ctx.lineCap = 'butt';
    orden.forEach((i) => {
      if (S.av[i] < 0.01) return;
      const c = M.cursos[i], p = S.sp[i];
      const r = TAM_PUNTO[c.tam] * Math.pow(S.zoom, 0.35) * Math.pow(p, 1.1) * (S.w < 560 ? 1.1 : 1);
      const a = Math.min(1, (0.22 + 0.78 * Math.min(luz(i), 1)) * Math.max(0.6, Math.min(1.1, p))) * S.av[i] * fFoco(i);
      const x = S.sx[i], y = S.sy[i];
      ctx.globalAlpha = a;
      if (c.tam === 3) {
        ctx.fillStyle = papel; ctx.lineWidth = 1.3;
        ctx.beginPath(); ctx.arc(x, y, r, 0, 6.2832); ctx.fill(); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(x - r * 1.9, y); ctx.lineTo(x + r * 1.9, y); ctx.moveTo(x, y - r * 1.9); ctx.lineTo(x, y + r * 1.9); ctx.lineWidth = 1; ctx.stroke();
        ctx.fillStyle = tinta;
      } else {
        ctx.beginPath(); ctx.arc(x, y, r, 0, 6.2832); ctx.fill();
      }
      if (enCombinacion(i)) { ctx.globalAlpha = Math.min(1, S.av[i]); ctx.lineWidth = 1.6; ctx.beginPath(); ctx.arc(x, y, r + 5, 0, 6.2832); ctx.stroke(); }
    });

    // nombres: curso seleccionado y cursos iluminados; el programa activo lleva su nombre sobre la órbita
    const cajas = []; S.zonasEt = [];
    const rotular = (texto, x, y, alfa, obligatoria, centrado, pref, curso) => {
      texto = texto.toUpperCase();
      let ancho = anchosTec.get(texto);
      if (ancho == null) { ancho = ctx.measureText(texto).width; anchosTec.set(texto, ancho); }
      globoCurso(texto, texto, ancho, x, y, alfa, obligatoria, cajas, W, H, claro, curso);
    };
    activos.forEach((a) => rotular(a.nombre, a.x, a.y, 1, true, true));
    rotulosLinea.forEach((e) => rotular(e.texto, e.x, e.y, 1, true, true));
    if (S.sel != null) rotular(M.cursos[S.sel].nombre, S.sx[S.sel], S.sy[S.sel], 1, true, false, null, S.sel);
    // los troncales de toda la escuela solo llevan nombre al seleccionarlos o al encender una maestría
    const troncales = []; let tcx = 0, tcy = 0;
    M.cursos.forEach((c, i) => { if (c.escuela) { troncales.push(i); tcx += S.sx[i]; tcy += S.sy[i]; } });
    if (troncales.length) { tcx /= troncales.length; tcy /= troncales.length; }
    troncales.forEach((i) => {
      return;   // los nombres de los troncales salen al pasar el cursor o al elegirlos
      const alfa = Math.min(1, Math.min(S.vN[i], 1) * 1.6) * S.ab[i];
      if (alfa < 0.08) return;
      const dx = S.sx[i] - tcx, dy = S.sy[i] - tcy, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
      rotular(M.cursos[i].nombre, S.sx[i], S.sy[i], alfa, false, false, (ancho) => [S.sx[i] + ux * 13 - (ux < 0 ? ancho : 0), S.sy[i] + uy * 13]);
    });
    if (S.sel != null) { /* solo el nombre del curso elegido */ }
    else if (S.bloqueSel || S.troncalesSel) orden.slice().reverse().forEach((i) => { if (S.nivel[i] === 2 && S.ab[i] > 0.4) rotular(M.cursos[i].nombre, S.sx[i], S.sy[i], Math.min(1, S.vN[i]) * S.ab[i], false, false, null, i); });
    else if (S.capa === 'cursos' && S.ruta) [...S.ruta.set].forEach((i) => rotular(M.cursos[i].nombre, S.sx[i], S.sy[i], 0.95 * S.ab[i], false, false, null, i));
    else if (S.capa === 'cursos' && S.maestriaSel) importantes().forEach((i) => rotular(M.cursos[i].nombre, S.sx[i], S.sy[i], 0.9 * S.ab[i], false, false, null, i));
    if (S.hoverI != null && S.sel == null && S.av[S.hoverI] > 0.5 && S.ab[S.hoverI] > 0.5) rotular(M.cursos[S.hoverI].nombre, S.sx[S.hoverI], S.sy[S.hoverI], 1, true, false, null, S.hoverI);
    ctx.strokeStyle = tinta;
    if (S.sel != null) {
      const i = S.sel;
      ctx.globalAlpha = 1; ctx.lineWidth = 1.3;
      ctx.beginPath(); ctx.arc(S.sx[i], S.sy[i], 10 * S.sp[i] + 3, 0, 6.2832); ctx.stroke();
    }

    // anotaciones del marco, como en un plano: título, conteo en marcas y datos del conjunto
    ctx.globalAlpha = 0.85; ctx.fillStyle = tinta; ctx.textBaseline = 'alphabetic';
    const programas = M.maestrias.length + M.doctorados.length;
    let marcas = ''; for (let k = 0; k < programas; k++) marcas += (k && k % 5 === 0 ? ' ' : '') + 'I';
    const ancho1 = ctx.measureText('CONSTELACIÓN DE POSGRADOS').width;
    ctx.fillText('CONSTELACIÓN DE POSGRADOS', W / 2 - ancho1 / 2, 26);
    ctx.fillText(marcas, W - 16 - ctx.measureText(marcas).width, 26);
    const datos = M.cursos.length + ' CURSOS · ' + M.maestrias.length + ' MAESTRÍAS' + (M.doctorados.length ? ' · ' + M.doctorados.length + ' DOCT.' : '');
    ctx.fillText(datos, W - 64 - ctx.measureText(datos).width, H - 18);
    ctx.fillRect(16, H - 26, 34, 5); ctx.fillRect(54, H - 32, 20, 5);
    if ('letterSpacing' in ctx) ctx.letterSpacing = '0px';
    ctx.globalAlpha = 1; ctx.textBaseline = 'middle'; ctx.setLineDash([]);
  }

  // Título abajo a la derecha: lo elegido (o «Escuela de posgrados») y, si hay un curso, su ficha va en el panel de la derecha
  function actualizarTitulo() {
    const M = S.modelo; if (!M) return;
    const panel = $('panel'), quiere = S.sel != null || !!S.ruta;   // la ficha de un curso o la combinación para titularse de la maestría
    if (panel.hidden === quiere) panel.hidden = !quiere;
    const clave = [S.sel, S.maestriaSel, S.bloqueSel, S.troncalesSel, !!S.ruta].join('|');
    if (clave === S.tituloClave) return;
    S.tituloClave = clave;
    let t = 'Escuela de posgrados', st = 'Facultad de Artes y Humanidades';
    if (quiere) { t = ''; st = ''; }
    else if (S.troncalesSel) { t = 'Cursos obligatorios'; st = 'Todas las maestrías'; }
    else if (S.bloqueSel && M.bloquePorId.get(S.bloqueSel)) {
      const b = M.bloquePorId.get(S.bloqueSel), pr = M.programaPorId.get(b.maestriaId);
      t = (S.capa === 'especializados' ? 'Programa especializado en ' : b.esEnfasis ? 'Énfasis en ' : '') + b.nombre; st = pr ? pr.nombre : '';
    } else if (S.maestriaSel && M.programaPorId.get(S.maestriaSel)) {
      const pr = M.programaPorId.get(S.maestriaSel), n = M.cursos.filter((c) => !c.escuela && c.progs.includes(pr.id)).length;
      t = pr.nombre; st = n + (n === 1 ? ' curso' : ' cursos');
    }
    $('tituloSelT').textContent = t; $('tituloSelS').textContent = st; $('tituloSel').hidden = !t;
  }
  function ciclo(t) {
    actualizarTitulo();
    actualizarFoco(); actualizarColorTroncales();
    if (S.atr) {
      for (let i = 0; i < S.atr.length; i++) {
        const dif = S.atrT[i] - S.atr[i];
        if (Math.abs(dif) > 0.002) { S.atr[i] += dif * (reducir ? 1 : 0.03); S.sucio = true; } else if (dif) S.atr[i] = S.atrT[i];
      }
      let me = 0; for (let i = 0; i < S.atr.length; i++) if (S.modelo.cursos[i].escuela) me = Math.max(me, S.atr[i]);
      const ke = Math.min(1, me);
      if (Math.abs(ke - S.kEsc) > 0.004) { S.kEsc = ke; S.sucio = true; }
      if (S.programaEscuela && !S.hubClave) {   // sin maestría elegida: suelta el color cuando los troncales ya regresaron
        let mx = 0; for (let i = 0; i < S.atr.length; i++) mx = Math.max(mx, S.atr[i]);
        if (mx < 0.02) { S.programaEscuela = null; S.sucio = true; }
      }
    }
    if (!puntero && (Math.abs(S.velX) > 0.0002 || Math.abs(S.velY) > 0.0002)) {
      S.giroXT += S.velX; S.giroYT += S.velY; S.velX *= 0.94; S.velY *= 0.94; S.sucio = true;
    }
    const gdx = S.giroXT - S.giroX, gdy = S.giroYT - S.giroY;
    if (Math.abs(gdx) > 0.0005 || Math.abs(gdy) > 0.0005) {
      const kg = reducir ? 1 : 0.22;
      S.giroX += gdx * kg; S.giroY += gdy * kg; S.sucio = true;
    }
    const dz = S.zoomT - S.zoom, dpx = S.panXT - S.panX, dpy = S.panYT - S.panY;
    if (Math.abs(dz) > 0.0005 || Math.abs(dpx) > 0.2 || Math.abs(dpy) > 0.2) {
      const k = reducir ? 1 : S.kZoom;
      S.zoom += dz * k; S.panX += dpx * k; S.panY += dpy * k; S.sucio = true;
    } else if (dz || dpx || dpy) { S.zoom = S.zoomT; S.panX = S.panXT; S.panY = S.panYT; S.sucio = true; }
    // con un curso elegido, al girar la vista el curso se queda en el centro (se gira alrededor de él, no del conjunto)
    const dif = S.objetivoProf - S.profundidad;
    if (Math.abs(dif) > 0.002) { S.profundidad += reducir ? dif : dif * 0.07; S.sucio = true; }
    else if (S.profundidad !== S.objetivoProf) { S.profundidad = S.objetivoProf; S.sucio = true; }
    // en 2D (profundidad ~0) no hay nada que animar: solo se redibuja cuando algo cambia
    const anima = S.vaiven && !reducir && S.modelo && S.modelo.cursos.length && S.profundidad > 0.02;
    if (S.pos && (S.sucio || anima)) {
      // se limpia ANTES de pintar: pintar() marca «sucio» mientras algo sigue animándose (apertura, burbujas, mezcla de niveles);
      // limpiarlo después lo borraba y, en 2D (sin vaivén que mantenga el bucle vivo), la animación se congelaba a medias
      S.sucio = false;
      pintar(t);
    }
    requestAnimationFrame(ciclo);
  }

  /* ----------------------------- interacción ------------------------------- */
  /* ---- zoom: rueda / pellizco / botones; el punto bajo el cursor se queda fijo ---- */
  function limitarPan() {
    const m = Math.max(S.w, S.h) * 0.6 * S.zoomT;
    S.panXT = Math.max(-m, Math.min(m, S.panXT)); S.panYT = Math.max(-m, Math.min(m, S.panYT));
  }
  function zoomHacia(factor, px, py) {
    S.kZoom = 0.22;
    const z0 = S.zoomT, z1 = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, z0 * factor));
    if (z1 === z0) return;
    const ux = 0, uy = 0, k = z1 / z0;   // el zoom siempre se hace hacia el centro, no hacia el cursor
    S.panXT = ux - (ux - S.panXT) * k; S.panYT = uy - (uy - S.panYT) * k;
    S.zoomT = z1;
    if (reducir) { S.zoom = S.zoomT; }
    limitarPan(); S.sucio = true;
  }
  function reiniciarZoom() {
    S.kZoom = 0.06; S.zoomT = 1; S.panXT = 0; S.panYT = 0;
    S.giroXT = 0; S.giroYT = 0; S.velX = 0; S.velY = 0;
    if (reducir) { S.zoom = 1; S.panX = 0; S.panY = 0; S.giroX = 0; S.giroY = 0; }
    S.sucio = true;
  }
  lienzo.addEventListener('wheel', (e) => {
    e.preventDefault();
    const r = lienzo.getBoundingClientRect();
    const dy = e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
    zoomHacia(Math.exp(-dy * (e.ctrlKey ? 0.01 : 0.0016)), e.clientX - r.left, e.clientY - r.top);
    ocultarTip();
  }, { passive: false });
  // teléfono: la leyenda es un menú que se abre con el botón «☰ Programas» y se cierra al elegir algo, con Escape o tocando el mapa
  (function () {
    const btn = $('btnLeyenda'); if (!btn) return;
    const poner = (ab) => { document.body.classList.toggle('leyenda-abierta', ab); btn.setAttribute('aria-expanded', ab ? 'true' : 'false'); if (ab) $('leyendaPanel').open = true; setTimeout(() => { S.reservas = null; S.sucio = true; }, 30); };
    btn.addEventListener('click', () => poner(!document.body.classList.contains('leyenda-abierta')));
    $('btnCerrarLeyenda').addEventListener('click', () => { poner(false); btn.focus(); });
    $('colIzq').addEventListener('click', (e) => { if (e.target.closest('[data-k]')) poner(false); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && document.body.classList.contains('leyenda-abierta')) { poner(false); btn.focus(); } });
    lienzo.addEventListener('pointerdown', () => { if (document.body.classList.contains('leyenda-abierta')) poner(false); });
  })();
  // hover en el panel Programas: lo suyo se enciende en la constelación y lo demás se apaga
  {
    const col = $('colIzq');
    col.addEventListener('mouseover', (e) => {
      const b = e.target.closest('[data-k]'); const k = b && b.dataset.k;
      const v = k && (k === 'e' || k.startsWith('p:') || k.startsWith('b:')) ? k : null;
      if (v !== S.hoverP) { S.hoverP = v; S.sucio = true; }
    });
    col.addEventListener('mouseleave', () => { if (S.hoverP) { S.hoverP = null; S.sucio = true; } });
  }
  $('btnZoomMas').addEventListener('click', () => zoomHacia(1.4, S.w / 2, S.h / 2));
  $('btnZoomMenos').addEventListener('click', () => zoomHacia(1 / 1.4, S.w / 2, S.h / 2));
  $('btnZoomReiniciar').addEventListener('click', reiniciarZoom);

  /* ---- puntero: un dedo/mouse gira o desplaza; dos dedos hacen zoom ---- */
  const activos = new Map();
  let puntero = null, pellizco = null;
  lienzo.addEventListener('pointerdown', (e) => {
    lienzo.setPointerCapture(e.pointerId);
    activos.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (activos.size === 2) {
      const [a, b] = [...activos.values()];
      pellizco = { d: Math.hypot(a.x - b.x, a.y - b.y) || 1, mx: (a.x + b.x) / 2, my: (a.y + b.y) / 2 };
      puntero = null; ocultarTip();
      return;
    }
    const mover = false;   // la vista no se puede sacar del centro: arrastrar solo gira (con zoom también)
    S.velX = 0; S.velY = 0;
    puntero = { x: e.clientX, y: e.clientY, t: performance.now(), gx: S.giroXT, gy: S.giroYT, px: S.panXT, py: S.panYT, modo: mover ? 'mover' : 'girar', mov: false, tipo: e.pointerType };
  });
  lienzo.addEventListener('pointermove', (e) => {
    if (activos.has(e.pointerId)) activos.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (pellizco && activos.size >= 2) {
      const [a, b] = [...activos.values()];
      const d = Math.hypot(a.x - b.x, a.y - b.y) || 1;
      const r = lienzo.getBoundingClientRect();
      const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
      zoomHacia(d / pellizco.d, mx - r.left, my - r.top);
      pellizco.d = d; pellizco.mx = mx; pellizco.my = my;
      return;
    }
    if (!puntero) return;
    const dx = e.clientX - puntero.x, dy = e.clientY - puntero.y;
    if (Math.hypot(dx, dy) > 6) puntero.mov = true;
    if (!puntero.mov) return;
    if (puntero.modo === 'mover') {
      S.kZoom = 0.22; S.panXT = puntero.px + dx; S.panYT = puntero.py + dy; limitarPan();
      if (reducir) { S.panX = S.panXT; S.panY = S.panYT; }
      S.sucio = true;
    } else if (S.profundidad > 0.03) {   // gira siempre; con algo elegido, alrededor de ese pivote
      // giro libre en los dos ejes (sin límites); la velocidad del gesto se conserva como inercia al soltar
      const ny = puntero.gy + dx * 0.008, nx = puntero.gx + dy * 0.008;
      S.velY = ny - S.giroYT; S.velX = nx - S.giroXT;
      S.giroYT = ny; S.giroXT = nx; puntero.t = performance.now();
      S.sucio = true;
    }
  });
  function soltar(e) {
    activos.delete(e.pointerId);
    if (activos.size < 2) pellizco = null;
  }
  lienzo.addEventListener('pointerup', (e) => {
    soltar(e);
    const p = puntero; puntero = null;
    if (p && performance.now() - p.t > 90) { S.velX = 0; S.velY = 0; }
    if (!p || p.mov || !S.pos) return;
    const r = lienzo.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top;
    const radio = p.tipo === 'touch' ? 26 : 15;
    let mejor = null, md = radio;
    // el nombre dibujado de un curso también se toca (abre su ficha igual que su punto)
    const sobreNombre = nombreEn(px, py);
    for (let i = 0; i < S.modelo.cursos.length; i++) {
      if (S.ab[i] < 0.5) continue;   // los cursos plegados en una burbuja no se tocan
      const d = Math.hypot(S.sx[i] - px, S.sy[i] - py);
      if (d < md) { md = d; mejor = i; }
    }
    if (mejor == null && sobreNombre != null) mejor = sobreNombre;
    // lo que está resaltado por el hover también responde al clic, aunque no se toque el punto exacto
    if (mejor == null && p.tipo !== 'touch' && S.hoverI != null && S.ab[S.hoverI] > 0.5) mejor = S.hoverI;
    if (mejor == null && S.capa !== 'cursos') {
      const bur = burbujaEn(px, py) || (p.tipo !== 'touch' && S.hoverB && S.bur.get(S.hoverB.g.id) === S.hoverB ? S.hoverB : null);
      if (bur && bur.g.esc) { seleccionarTroncales(); return; }
      // la parte de una maestría que no pertenece a ningún énfasis («general»): con la maestría ya elegida, lleva a la visualización de cursos
      if (bur && bur.g.id.startsWith('g:')) {
        if (S.maestriaSel === bur.g.maestriaId) { const vc = S.volverCapa; fijarCapa('cursos'); S.volverCapa = vc; }
        else seleccionarMaestria(bur.g.maestriaId);
        return;
      }
      if (bur && bur.g.id.startsWith('b:')) { seleccionarBloque(bur.g.id.slice(2)); return; }
      if (bur && bur.g.maestriaId) { seleccionarMaestria(bur.g.maestriaId); return; }
    }
    // clic en el fondo: primero suelta el curso elegido; si no hay curso, sale de la maestría, el énfasis o los troncales (igual que el botón «Volver»)
    if (mejor == null && S.sel == null && (S.maestriaSel || S.bloqueSel || S.troncalesSel)) { volverATodas(); return; }
    const nuevo = mejor === S.sel ? null : mejor;
    seleccionar(nuevo);
    if (nuevo == null) { if (!S.filtro.size) { S.kZoom = 0.04; S.zoomT = 1; S.panXT = 0; S.panYT = 0; S.sucio = true; } }
    else { // acerca el curso, su bloque, su secuencia y los programas donde es electivo
      const Mo = S.modelo, c = Mo.cursos[nuevo], ver = new Set([nuevo]);
      c.bloquesReales.forEach((bid) => { const b = Mo.bloquePorId.get(bid); if (b) b.cursos.forEach((id) => ver.add(Mo.porId.get(id))); });
      c.electivaDestinos.forEach((d) => {
        if (d.tipo === 'bloque') { const b = Mo.bloquePorId.get(d.id); if (b) b.cursos.forEach((id) => ver.add(Mo.porId.get(id))); }
        else Mo.cursos.forEach((o, j) => { if (!o.escuela && o.progs.includes(d.id)) ver.add(j); });
      });
      Mo.cursos.forEach((o, j) => { if (o.prereq.includes(c.id) || c.prereq.includes(o.id) || c.coreq.includes(o.id)) ver.add(j); });
      ver.delete(undefined);
      if (ver.size >= 2) ajustarVista([...ver], 3.4, nuevo);
    }
  });
  lienzo.addEventListener('contextmenu', (e) => e.preventDefault());
  lienzo.addEventListener('pointercancel', (e) => { soltar(e); puntero = null; ocultarTip(); });

  /* ---- hover: dice sobre qué maestría y/o especializado estás ---- */
  const tip = $('tip');
  let tipCurso = null;
  function ocultarTip() { tip.classList.remove('visible'); tipCurso = null; if (S.hoverI != null || S.hoverB) { S.hoverI = null; S.hoverB = null; S.sucio = true; } }
  // Al pasar el cursor solo se lee el nombre del curso (se dibuja en el lienzo, ver S.hoverI); el resto de la información está en el detalle
  function mostrarTip(i) { tipCurso = i; }
  lienzo.addEventListener('pointermove', (e) => {
    if (e.pointerType === 'touch' || (puntero && puntero.mov) || !S.pos) { ocultarTip(); return; }
    const r = lienzo.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top;
    lienzo.style.cursor = (S.capa !== 'cursos' && burbujaEn(px, py)) || nombreEn(px, py) != null ? 'pointer' : '';
    // el cluster entero responde: se toma el curso más cercano dentro del radio de su resplandor
    const radio = Math.min(110, Math.max(34, S.pos.R * 0.1 * S.escala));
    let mejor = null, md = radio;
    for (let i = 0; i < S.modelo.cursos.length; i++) {
      if (S.ab[i] < 0.5) continue;
      const d = Math.hypot(S.sx[i] - px, S.sy[i] - py);
      if (d < md) { md = d; mejor = i; }
    }
    if (mejor == null) {
      // sin curso cerca: ¿hay un halo (maestría, énfasis o especializado) bajo el cursor?
      const bur = S.capa !== 'cursos' && nombreEn(px, py) == null ? burbujaEn(px, py) : null;
      ocultarTip();
      if (bur) { S.hoverB = bur; S.sucio = true; }
    } else { mostrarTip(mejor, px, py); if (S.hoverB) { S.hoverB = null; S.sucio = true; } if (S.hoverI !== mejor) { S.hoverI = mejor; S.sucio = true; } }
    if (S.hoverI != null || S.hoverB) lienzo.style.cursor = 'pointer';
  });
  lienzo.addEventListener('pointerleave', ocultarTip);
  lienzo.addEventListener('keydown', (e) => {
    if (!S.modelo || !S.modelo.cursos.length) return;
    const n = S.modelo.cursos.length;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { seleccionar(S.sel == null ? 0 : (S.sel + 1) % n); e.preventDefault(); }
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') { seleccionar(S.sel == null ? n - 1 : (S.sel - 1 + n) % n); e.preventDefault(); }
    else if (e.key === 'Escape') { if (S.sel != null) seleccionar(null); else if (S.maestriaSel || S.troncalesSel) volverATodas(); }
    else if (e.key === '+' || e.key === '=') zoomHacia(1.25, S.w / 2, S.h / 2);
    else if (e.key === '-' || e.key === '_') zoomHacia(1 / 1.25, S.w / 2, S.h / 2);
    else if (e.key === '0') reiniciarZoom();
  });
  $('buscar').addEventListener('input', (e) => {
    const v = norm(e.target.value);
    const i = S.modelo.cursos.findIndex((c) => norm(c.nombre) === v);
    if (i >= 0) elegirCurso(i);
  });

  /* ----------------------------- controles --------------------------------- */
  $('rangoProfundidad').addEventListener('input', (e) => { S.objetivoProf = e.target.value / 100; });
  $('btnVista').addEventListener('click', () => {
    S.objetivoProf = S.objetivoProf > 0.5 ? 0 : 1;
    $('rangoProfundidad').value = Math.round(S.objetivoProf * 100);
    $('btnVista').setAttribute('aria-pressed', S.objetivoProf > 0.5 ? 'true' : 'false');
  });
  // Estilo técnico: plano, solo negro y blanco de la paleta (se pasa al modo claro al entrar y se restaura al salir)
  function cambiarEstilo(tecnico) {
    if (tecnico === (S.estilo === 'tecnico')) return;
    S.estilo = tecnico ? 'tecnico' : 'brillante';
    document.documentElement.classList.toggle('estilo-tecnico', tecnico);
    $('btnEstilo').textContent = 'Estilo: ' + (tecnico ? 'técnico' : 'brillante');
    actualizarSubtitulo();
    if (tecnico) { S.temaPrevio = S.tema; if (S.tema !== 'claro') cambiarTema(); }
    else if (S.temaPrevio && S.temaPrevio !== S.tema) cambiarTema();
    S.sucio = true;
  }
  $('btnEstilo').addEventListener('click', () => cambiarEstilo(S.estilo !== 'tecnico'));
  $('btnUniones').addEventListener('click', () => {
    S.modoUnion = S.modoUnion === 'estela' ? 'recta' : 'estela';
    $('btnUniones').textContent = 'Uniones: ' + (S.modoUnion === 'estela' ? 'estelas' : 'rectas');
    S.sucio = true;
  });
  const cambiarTema = () => {
    S.tema = S.tema === 'oscuro' ? 'claro' : 'oscuro';
    document.documentElement.dataset.tema = S.tema;
    aplicarPaleta();
    $('btnTema').textContent = S.tema === 'oscuro' ? 'Modo claro' : 'Modo oscuro';
    pintarLeyenda(); S.sucio = true;
  };
  $('btnTema').addEventListener('click', cambiarTema);
  $('btnTema2').addEventListener('click', cambiarTema);
  // Vista de coordinación: todo junto (todos los cursos y sus conexiones) con los controles de cámara y estilo
  $('btnCoord').addEventListener('click', () => {
    S.maestriaSel = null; S.bloqueSel = null; S.troncalesSel = false; S.filtro.clear(); S.sel = null;
    S.capa = 'cursos'; abrirCajon(false); aplicarFiltro();
  });
  $('btnVolver').addEventListener('click', volverATodas);
  $('cVaiven').addEventListener('change', (e) => { S.vaiven = e.target.checked && !reducir; S.sucio = true; });
  if (reducir) { $('cVaiven').checked = false; $('cVaiven').disabled = true; }
  const cajon = $('cajon');
  const abrirCajon = (abrir) => { cajon.classList.toggle('abierto', abrir); cajon.setAttribute('aria-hidden', abrir ? 'false' : 'true'); if (abrir) $('btnCerrar').focus(); };
  $('btnDatos').addEventListener('click', () => abrirCajon(true));
  $('btnAdvertencias').addEventListener('click', () => { abrirCajon(true); $('seccionAdvertencias').scrollIntoView({ block: 'start' }); });
  $('btnCerrar').addEventListener('click', () => abrirCajon(false));
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && cajon.classList.contains('abierto')) abrirCajon(false); });
  ['cNMaestrias', 'cConEnfasis', 'cEnfasisPor', 'cProyectos', 'cCompartidos', 'cPctElectivas', 'cDoctorado'].forEach((id) => $(id).addEventListener('input', regenerarEjemplo));
  $('btnEjemplo').addEventListener('click', () => { $('mensajeCarga').textContent = ''; regenerarEjemplo(); });

  /* ------------------------ carga y plantillas ----------------------------- */
  function mensaje(t, error) { const m = $('mensajeCarga'); m.textContent = t; m.className = error ? 'error' : ''; }
  $('btnCargar').addEventListener('click', () => $('archivo').click());
  $('archivo').addEventListener('change', async (e) => {
    const archivo = e.target.files[0]; e.target.value = '';
    if (!archivo) return;
    try {
      const ext = archivo.name.split('.').pop().toLowerCase();
      if (ext !== 'xlsx' && ext !== 'xls') throw new Error('El archivo debe ser un libro de Excel (.xlsx) con las hojas de la plantilla.');
      if (!window.XLSX) throw new Error('No se pudo cargar la librería de Excel (¿sin conexión?).');
      const libro = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });
      const hojas = {}, crudas = {};
      libro.SheetNames.forEach((h) => {
        const k = normClave(h);
        hojas[k] = XLSX.utils.sheet_to_json(libro.Sheets[h], { header: 1, defval: '', raw: false });
        crudas[k] = XLSX.utils.sheet_to_json(libro.Sheets[h], { defval: '', raw: false });
      });
      const viejo = (hojas.cursos && hojas.cursos[0] || []).some((c) => normClave(c) === 'curso_id');
      let T, avisos = [];
      if (viejo) T = normalizarTablas(crudas);
      else { const r = tablasDesdeMalla(hojas); T = r.T; avisos = r.avisos; }
      if (!T.cursos.length) throw new Error('No encontré cursos en la hoja «cursos». Usa la plantilla de Excel.');
      if (!T.programas.length) throw new Error('La hoja «programas» está vacía.');
      aplicarTablas(T, 'archivo');
      mensaje('Cargado: ' + T.cursos.length + ' cursos, ' + T.programas.length + ' programas.' + (avisos.length ? ' Avisos (' + avisos.length + '): ' + avisos.slice(0, 4).join(' ') + (avisos.length > 4 ? ' …' : '') : ''));
    } catch (err) { mensaje(err.message || String(err), true); }
  });
  function descargar(nombre, blob) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  // Excel generado aquí mismo: con los datos del ejemplo o solo con los encabezados (plantilla vacía)
  function bajarLibro(nombre, T) {
    if (!window.XLSX) { mensaje('No se pudo cargar la librería de Excel.', true); return; }
    const H = mallaDesdeTablas(T), libro = XLSX.utils.book_new();
    ['programas', 'cursos', 'reglas'].forEach((h) => XLSX.utils.book_append_sheet(libro, XLSX.utils.aoa_to_sheet(H[h]), h));
    descargar(nombre, new Blob([XLSX.write(libro, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  $('btnEjemploXlsx').addEventListener('click', () => bajarLibro('ejemplo-constelacion.xlsx', generarEjemplo(paramsEjemplo())));
  $('btnPlantillaXlsx').addEventListener('click', () => { const v = {}; TABLAS.forEach((t) => { v[t] = []; }); bajarLibro('plantilla-constelacion.xlsx', v); });

  /* ------------------------------- arranque -------------------------------- */
  new ResizeObserver(ajustar).observe(lienzo.parentElement);
  if (window.innerWidth <= 760) $('detalle').open = false;   // en pantallas chicas el detalle solo ocupa lugar cuando hay un curso
  ajustar();
  regenerarEjemplo();
  S.profundidad = reducir ? S.objetivoProf : 0;
  requestAnimationFrame(ciclo);
  // Parámetros de URL para revisar estados sin interacción: ?sel=<n>&hover=<n>&zoom=<n>&giro=<x>,<y>&uniones=recta&filtro=p:M1,b:M2-B1&claro=1&prof=0..100
  const q = new URLSearchParams(location.search);
  if (q.has('claro')) $('btnTema').click();
  if (q.get('estilo') === 'tecnico') cambiarEstilo(true);
  if (q.get('uniones') === 'recta') $('btnUniones').click();
  if (q.has('prof')) { S.objetivoProf = S.profundidad = +q.get('prof') / 100; $('rangoProfundidad').value = q.get('prof'); }
  if (q.has('filtro')) { q.get('filtro').split(',').forEach((k) => S.filtro.add(k)); aplicarFiltro(); }
  if (q.has('giro')) { const g = q.get('giro').split(',').map(Number); S.giroXT = S.giroX = g[0] || 0; S.giroYT = S.giroY = g[1] || 0; }
  if (q.has('zoom')) { S.zoom = S.zoomT = +q.get('zoom'); }
  // capas: ?capa=maestrias|enfasis|especializados|cursos & maestria=<n> (posición de la maestría) & con=prereq,electivas,troncales,doctorado
  if (q.has('capa')) fijarCapa(q.get('capa'));
  if (q.has('maestria')) { const m = S.modelo.maestrias[+q.get('maestria')]; if (m) seleccionarMaestria(m.id); }
  pintarLeyenda();
  if (q.has('sel')) seleccionar(+q.get('sel'));
  if (q.has('hover')) setTimeout(() => { const i = +q.get('hover'); mostrarTip(i, S.sx[i], S.sy[i]); }, 800);
  // la tipografía (Typekit) llega después: se vuelven a medir los textos del lienzo cuando esté lista
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => { anchosEtiqueta.clear(); anchosTec.clear(); if (S.rTxt) S.rTxt.clear(); S.ladoEt = new Map(); S.sucio = true; });
  window.__constelacion = S; window.__pintar = pintar; // para depuración desde la consola
})();
