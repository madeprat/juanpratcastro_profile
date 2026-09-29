// Genera los CV 2026 de Juan Antonio en .docx (una página, sin emojis, apto para filtros ATS).
// Uso: node tools/generar_cv.js   (requiere el paquete npm "docx")
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, LevelFormat, ExternalHyperlink, VerticalAlign,
  TabStopType, ShadingType,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const FOTO = path.join(ROOT, "asset/photo/juanprat_cv.jpg");
const SALIDA = path.join(ROOT, "asset/cvx/2026");

const AZUL = "1F3864";
const GRIS = "555555";
const FUENTE = "Calibri";
const ANCHO = 11906 - 2 * 850; // A4 menos márgenes

// ---------- Contenido común ----------
const DATOS = {
  nombre: "Juan Antonio Pratdesaba Castro",
  telefono: "691 843 003",
  email: "juanpratcastro@gmail.com",
  ciudad: "Valencia (46025)",
  linkedin: "https://www.linkedin.com/in/juanantonio-pratdesaba-castro-41931b438/",
  disponibilidad: "Disponibilidad: miércoles y viernes (jornada completa) y de lunes a viernes a partir de las 16:00",
};

const CERLER = {
  puesto: "Ayudante de Sala",
  lugar: "Restaurante La Cabana · Cerler (Huesca)",
  fechas: "Jul. 2026 – Sep. 2026",
  intro: "Temporada de verano en una estación de montaña con alta afluencia de turismo nacional e internacional.",
  bullets: [
    "Atención directa a clientela internacional, adaptando el trato y el idioma a cada cliente (español, catalán, inglés y nociones de francés).",
    "Servicio en picos de alta demanda con tiempos ajustados, sin perder la calidad ni la atención al detalle.",
    "Gestión de incidencias y quejas en el momento, con calma y buen criterio, orientado a la satisfacción total del cliente.",
    "Coordinación constante con cocina y compañeros de sala; valorado por la dirección por su autonomía, fiabilidad y capacidad de mantener la calma en momentos de máxima presión.",
  ],
};

const EXPO = {
  puesto: "Asistente de organización",
  lugar: "ExpoITSM 2022 · itSMF España · Hotel NH Ventas, Madrid",
  fechas: "2022",
  bullets: [
    "Montaje de sala, señalética y acreditación de asistentes en un congreso nacional de tecnología.",
    "Gestión de micrófonos y turnos de palabra, y apoyo a ponentes y moderadores durante toda la jornada.",
  ],
};

const RUTAS = {
  puesto: "Rutas históricas por Valencia para visitantes internacionales",
  lugar: "Centro histórico de Valencia · actividad personal, no profesional",
  fechas: "2024 – actualidad",
  bullets: [
    "Diseño y conducción de recorridos por la Almoina, la Catedral, las Torres de Serranos y Quart y la Plaza de la Virgen para pequeños grupos.",
    "Adaptación del contenido y del idioma a la edad, procedencia e intereses de cada grupo.",
  ],
};

const DANA = {
  puesto: "Voluntariado · Emergencia DANA",
  lugar: "Paiporta y Valencia",
  fechas: "Oct. – Nov. 2024",
  bullets: [
    "Limpieza de espacios afectados, organización de donaciones y apoyo a personas mayores y familias.",
  ],
};

const FORMACION = [
  {
    titulo: "Grado en Historia",
    lugar: "Universitat de València",
    fechas: "2025 – actualidad",
    detalle: "Primer curso completado con Matrícula de Honor.",
  },
  {
    titulo: "Bachillerato de Humanidades y Ciencias Sociales",
    lugar: "Escuelas San José – Jesuitas, Valencia",
    fechas: "2025",
    detalle: "Olimpiada de Filosofía 2024: seleccionado para representar al centro con ensayo y defensa oral.",
  },
];

const IDIOMAS = [
  ["Español", "nativo"],
  ["Valenciano / catalán", "B1"],
  ["Inglés", "B1"],
  ["Francés", "A2"],
];

// ---------- Versiones ----------
const VERSIONES = [
  {
    archivo: "CV_JAPC_2026_Eventos_Hospitality",
    titular: "Atención al cliente · Eventos y hospitality",
    perfil:
      "Estudiante de Historia en la Universitat de València (primer curso con Matrícula de Honor) con experiencia real de cara al público. " +
      "Este verano, como ayudante de sala en Cerler, he trabajado con clientela internacional, plazos ajustados y equipos bajo presión. " +
      "Mantengo la calma cuando la situación se complica, cambio de idioma según el cliente y cuido la imagen y el detalle. " +
      "Busco colaborar como azafato, staff de eventos, congresos o hospitality.",
    competencias: [
      "Atención al cliente en cuatro idiomas",
      "Calma y resolución de incidencias bajo presión",
      "Trabajo en equipo y coordinación de sala",
      "Protocolo, buena presencia y trato formal",
      "Logística de eventos: acreditaciones, montaje, apoyo a ponentes",
      "Word, PowerPoint y Canva",
    ],
    experiencia: [CERLER, EXPO, RUTAS, DANA],
  },
  {
    archivo: "CV_JAPC_2026_Turismo_Idiomas",
    titular: "Turismo cultural · Atención al visitante · Idiomas",
    perfil:
      "Estudiante de Historia en la Universitat de València (primer curso con Matrícula de Honor) que disfruta explicando el pasado a quien lo visita. " +
      "Combino conocimiento del patrimonio valenciano con experiencia de trato con turistas: este verano he atendido a clientela internacional en Cerler " +
      "en su propio idioma, y desde 2024 preparo rutas históricas por el centro de Valencia. " +
      "Busco colaborar en free tours, visitas guiadas, museos, recepción de hotel o atención al visitante.",
    competencias: [
      "Comunicación oral clara y amena con grupos",
      "Adaptación del discurso a cada público e idioma",
      "Conocimiento de la historia y el patrimonio de Valencia",
      "Atención al cliente internacional",
      "Calma y resolución de imprevistos",
      "Word, PowerPoint y Canva",
    ],
    experiencia: [CERLER, RUTAS, EXPO, DANA],
  },
  {
    archivo: "CV_JAPC_2026_Clases_Universidad",
    titular: "Estudiante de Historia · Clases particulares y colaboración académica",
    perfil:
      "Estudiante de Historia en la Universitat de València con el primer curso completado con Matrícula de Honor. " +
      "Me gusta explicar temas complejos con ejemplos sencillos y hacer que la historia se entienda y se recuerde. " +
      "Tengo experiencia hablando ante grupos (rutas históricas, Olimpiada de Filosofía) y he demostrado responsabilidad y constancia trabajando en temporada alta. " +
      "Busco dar clases de Historia, Geografía o Filosofía (ESO y Bachillerato) o colaborar en proyectos de la universidad.",
    competencias: [
      "Historia, Geografía, Historia del Arte y Filosofía",
      "Explicación didáctica y técnicas de estudio",
      "Hablar en público con claridad",
      "Paciencia, constancia y responsabilidad",
      "Redacción y búsqueda de fuentes fiables",
      "Word, PowerPoint y Canva",
    ],
    experiencia: [RUTAS, CERLER, EXPO, DANA],
    formacionPrimero: true,
  },
  {
    archivo: "CV_JAPC_2026_Relaciones_Institucionales",
    titular: "Eventos institucionales · Protocolo · Atención a invitados",
    perfil:
      "Estudiante de Historia en la Universitat de València (primer curso con Matrícula de Honor), con interés por el protocolo, " +
      "las relaciones institucionales y la organización de actos. He trabajado de cara al público con clientela internacional y en " +
      "la logística de un congreso nacional, recibiendo y acreditando a asistentes y coordinando a ponentes y moderadores. " +
      "Me distinguen la discreción, el saber estar y la calma en los momentos de máxima presión. Además, me comunico en español, valenciano e inglés. " +
      "Busco incorporarme como auxiliar o azafato en actos institucionales, congresos y eventos corporativos.",
    competencias: [
      "Protocolo, saber estar y buena presencia",
      "Discreción y trato formal con ponentes e invitados",
      "Recepción, acreditación y acompañamiento de asistentes",
      "Calma y resolución de imprevistos en directo",
      "Español, valenciano e inglés en entornos formales",
      "Contexto histórico e institucional (Grado en Historia)",
    ],
    experiencia: [
      {
        ...EXPO,
        bullets: [
          "Recepción y acreditación de asistentes, y entrega de materiales en un congreso nacional de tecnología.",
          "Coordinación con ponentes y moderadores: gestión de micrófonos, turnos de palabra y tiempos de intervención.",
          "Montaje de sala y señalética, y apoyo continuo durante toda la jornada.",
        ],
      },
      {
        ...CERLER,
        bullets: [
          "Atención a clientela nacional e internacional, adaptando el registro y el idioma a cada persona, del trato cercano al más formal.",
          "Servicio en picos de alta demanda con tiempos ajustados, sin perder la calidad ni la atención al detalle.",
          "Resolución de incidencias en el momento, con discreción y buen criterio.",
          "Valorado por la dirección por su autonomía, fiabilidad y capacidad de mantener la calma en momentos de máxima presión.",
        ],
      },
      RUTAS,
      DANA,
    ],
  },
];

// ---------- Utilidades de formato ----------
const t = (text, o = {}) => new TextRun({ text, font: FUENTE, size: 19, ...o });

function seccion(titulo) {
  return new Paragraph({
    spacing: { before: 170, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 2 } },
    children: [t(titulo.toUpperCase(), { bold: true, color: AZUL, size: 21, characterSpacing: 20 })],
  });
}

function cabeceraLinea(izq, der, o = {}) {
  return new Paragraph({
    spacing: { before: o.before ?? 80, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }],
    keepNext: true,
    children: [t(izq, { bold: true, size: 20 }), t("\t" + der, { color: GRIS, size: 18 })],
  });
}

const sub = (texto) =>
  new Paragraph({ spacing: { after: 30 }, keepNext: true, children: [t(texto, { italics: true, color: GRIS, size: 18 })] });

const vineta = (texto) =>
  new Paragraph({ numbering: { reference: "vinetas", level: 0 }, spacing: { after: 20 }, children: [t(texto)] });

function bloqueExperiencia(e, i) {
  const out = [cabeceraLinea(e.puesto, e.fechas, { before: i === 0 ? 40 : 110 }), sub(e.lugar)];
  if (e.intro) out.push(new Paragraph({ spacing: { after: 30 }, children: [t(e.intro)] }));
  e.bullets.forEach((b) => out.push(vineta(b)));
  return out;
}

function bloqueFormacion() {
  const out = [seccion("Formación")];
  FORMACION.forEach((f, i) => {
    out.push(cabeceraLinea(f.titulo, f.fechas, { before: i === 0 ? 40 : 100 }));
    out.push(sub(f.lugar));
    out.push(new Paragraph({ spacing: { after: 20 }, children: [t(f.detalle)] }));
  });
  return out;
}

const SIN_BORDE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const SIN_BORDES = { top: SIN_BORDE, bottom: SIN_BORDE, left: SIN_BORDE, right: SIN_BORDE };

function cabecera(v) {
  const anchoFoto = 1700;
  const anchoTexto = ANCHO - anchoFoto;
  const contacto = [
    t(`${DATOS.telefono}  ·  ${DATOS.email}  ·  ${DATOS.ciudad}`, { size: 18 }),
  ];
  const textos = [
    new Paragraph({ children: [t(DATOS.nombre, { bold: true, size: 36, color: AZUL })] }),
    new Paragraph({ spacing: { after: 80 }, children: [t(v.titular, { size: 22, color: GRIS })] }),
    new Paragraph({ spacing: { after: 20 }, children: contacto }),
    new Paragraph({
      spacing: { after: 20 },
      children: [
        new ExternalHyperlink({
          link: DATOS.linkedin,
          children: [t("linkedin.com/in/juanantonio-pratdesaba-castro-41931b438", { size: 18, color: "0563C1", underline: {} })],
        }),
      ],
    }),
  ];
  const foto = new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [
      new ImageRun({ type: "jpg", data: fs.readFileSync(FOTO), transformation: { width: 100, height: 100 } }),
    ],
  });
  return new Table({
    width: { size: ANCHO, type: WidthType.DXA },
    columnWidths: [anchoTexto, anchoFoto],
    borders: { ...SIN_BORDES, insideHorizontal: SIN_BORDE, insideVertical: SIN_BORDE },
    rows: [
      new TableRow({
        children: [
          new TableCell({ width: { size: anchoTexto, type: WidthType.DXA }, borders: SIN_BORDES, verticalAlign: VerticalAlign.CENTER, children: textos }),
          new TableCell({ width: { size: anchoFoto, type: WidthType.DXA }, borders: SIN_BORDES, verticalAlign: VerticalAlign.CENTER, children: [foto] }),
        ],
      }),
    ],
  });
}

const franjaDisponibilidad = () =>
  new Paragraph({
    spacing: { before: 120, after: 40 },
    shading: { type: ShadingType.CLEAR, fill: "E8EEF7", color: "auto" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 6 } },
    children: [t(DATOS.disponibilidad, { bold: true, size: 19, color: AZUL })],
  });

function construir(v) {
  const hijos = [cabecera(v), franjaDisponibilidad(), seccion("Perfil"),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [t(v.perfil)] })];

  const experiencia = [seccion("Experiencia"), ...v.experiencia.flatMap(bloqueExperiencia)];
  if (v.formacionPrimero) hijos.push(...bloqueFormacion(), ...experiencia);
  else hijos.push(...experiencia, ...bloqueFormacion());

  hijos.push(seccion("Idiomas"));
  hijos.push(new Paragraph({
    children: IDIOMAS.flatMap(([id, nivel], i) => [
      t((i ? "   ·   " : "") + id + " ", { bold: true }), t(nivel),
    ]),
  }));

  hijos.push(seccion("Competencias"));
  // Dos columnas de competencias
  const mitad = Math.ceil(v.competencias.length / 2);
  const col = ANCHO / 2;
  const celda = (items) => new TableCell({
    width: { size: col, type: WidthType.DXA }, borders: SIN_BORDES,
    children: items.map(vineta),
  });
  hijos.push(new Table({
    width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [col, col],
    borders: { ...SIN_BORDES, insideHorizontal: SIN_BORDE, insideVertical: SIN_BORDE },
    rows: [new TableRow({ children: [celda(v.competencias.slice(0, mitad)), celda(v.competencias.slice(mitad))] })],
  }));

  hijos.push(new Paragraph({
    spacing: { before: 140 },
    children: [t("Referencias disponibles a petición.", { italics: true, color: GRIS, size: 18 })],
  }));

  return new Document({
    creator: DATOS.nombre,
    title: `CV ${DATOS.nombre} – ${v.titular}`,
    numbering: {
      config: [{
        reference: "vinetas",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 300, hanging: 200 } } } }],
      }],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } },
      children: hijos,
    }],
  });
}

// ---------- Versión HTML → PDF (mismo contenido, para enviar) ----------
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");

function html(v) {
  const foto = "data:image/jpeg;base64," + fs.readFileSync(FOTO).toString("base64");
  const exp = v.experiencia.map((e) => `
    <div class="item"><div class="fila"><b>${esc(e.puesto)}</b><span>${esc(e.fechas)}</span></div>
    <div class="sub">${esc(e.lugar)}</div>${e.intro ? `<p>${esc(e.intro)}</p>` : ""}
    <ul>${e.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul></div>`).join("");
  const form = FORMACION.map((f) => `
    <div class="item"><div class="fila"><b>${esc(f.titulo)}</b><span>${esc(f.fechas)}</span></div>
    <div class="sub">${esc(f.lugar)}</div><p>${esc(f.detalle)}</p></div>`).join("");
  const bExp = `<h2>Experiencia</h2>${exp}`;
  const bForm = `<h2>Formación</h2>${form}`;
  return `<!doctype html><html lang="es"><meta charset="utf-8"><style>
    @page { size: A4; margin: 12mm 15mm 10mm; }
    body { font-family: Calibri, Carlito, "Liberation Sans", Arial, sans-serif; font-size: 9.6pt; color: #222; margin: 0; line-height: 1.3; }
    header { display: flex; justify-content: space-between; align-items: center; }
    h1 { color: #${AZUL}; font-size: 18pt; margin: 0; }
    .tit { color: #${GRIS}; font-size: 11pt; margin: 2px 0 6px; }
    .contacto { font-size: 9pt; } .contacto a { color: #0563C1; }
    header img { width: 34mm; height: 34mm; object-fit: cover; border-radius: 3px; }
    .disp { background: #E8EEF7; border-left: 3px solid #${AZUL}; color: #${AZUL}; font-weight: bold; padding: 4px 8px; margin: 8px 0 2px; }
    h2 { color: #${AZUL}; font-size: 10.5pt; letter-spacing: .5px; text-transform: uppercase; border-bottom: 1px solid #${AZUL}; margin: 9px 0 4px; padding-bottom: 1px; }
    p { margin: 0 0 2px; } .perfil { text-align: justify; }
    .item { margin-top: 5px; break-inside: avoid; }
    .fila { display: flex; justify-content: space-between; } .fila span { color: #${GRIS}; font-size: 9pt; }
    .sub { color: #${GRIS}; font-style: italic; font-size: 9pt; margin-bottom: 1px; }
    ul { margin: 1px 0 0; padding-left: 14px; } li { margin-bottom: 1px; }
    .comp { columns: 2; } .ref { color: #${GRIS}; font-style: italic; font-size: 9pt; margin-top: 8px; }
  </style><body>
  <header><div><h1>${DATOS.nombre}</h1><div class="tit">${esc(v.titular)}</div>
    <div class="contacto">${DATOS.telefono} &nbsp;·&nbsp; ${DATOS.email} &nbsp;·&nbsp; ${DATOS.ciudad}<br>
    <a href="${DATOS.linkedin}">linkedin.com/in/juanantonio-pratdesaba-castro-41931b438</a></div></div>
    <img src="${foto}" alt=""></header>
  <div class="disp">${esc(DATOS.disponibilidad)}</div>
  <h2>Perfil</h2><p class="perfil">${esc(v.perfil)}</p>
  ${v.formacionPrimero ? bForm + bExp : bExp + bForm}
  <h2>Idiomas</h2><p>${IDIOMAS.map(([i, n]) => `<b>${i}</b> ${n}`).join(" &nbsp;·&nbsp; ")}</p>
  <h2>Competencias</h2><ul class="comp">${v.competencias.map((c) => `<li>${esc(c)}</li>`).join("")}</ul>
  <div class="ref">Referencias disponibles a petición.</div></body></html>`;
}

fs.mkdirSync(SALIDA, { recursive: true });
(async () => {
  let navegador = null;
  try {
    const { chromium } = require("playwright");
    navegador = await chromium.launch();
  } catch (e) {
    console.warn("Sin Playwright: solo se generan los .docx");
  }
  for (const v of VERSIONES) {
    const buf = await Packer.toBuffer(construir(v));
    fs.writeFileSync(path.join(SALIDA, v.archivo + ".docx"), buf);
    if (navegador) {
      const pagina = await navegador.newPage();
      await pagina.setContent(html(v));
      if (process.env.CV_CAPTURA) await pagina.screenshot({ path: path.join(process.env.CV_CAPTURA, v.archivo + ".png"), fullPage: true });
      await pagina.pdf({ path: path.join(SALIDA, v.archivo + ".pdf"), format: "A4", printBackground: true, preferCSSPageSize: true });
      await pagina.close();
    }
    console.log("OK", v.archivo);
  }
  if (navegador) await navegador.close();
})();
