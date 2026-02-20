const statusEl = document.getElementById("status");
const fullDemoButton = document.getElementById("generate-all");
const focusedButtons = document.querySelectorAll("button[data-demo]");

const svgImageData = `data:image/svg+xml;utf8,${encodeURIComponent(`
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#0d1b2a" />
  <rect x="24" y="24" width="592" height="312" fill="#1b263b" rx="14" />
  <circle cx="180" cy="180" r="80" fill="#e0e1dd" />
  <rect x="310" y="110" width="210" height="26" fill="#778da9" rx="6" />
  <rect x="310" y="152" width="170" height="24" fill="#415a77" rx="6" />
  <rect x="310" y="194" width="130" height="24" fill="#778da9" rx="6" />
  <text x="320" y="84" fill="#f8f9fa" font-size="28" font-family="Arial">Image Embed Demo</text>
</svg>
`)}`;

function setStatus(message, error = false) {
  statusEl.textContent = message;
  statusEl.style.color = error ? "#b91c1c" : "#14532d";
}

function createBasePresentation(title) {
  const pptx = new PptxGenJS();
  pptx.author = "PptxGenJS Demo Gallery";
  pptx.company = "GitHub Pages";
  pptx.subject = "PptxGenJS feature examples";
  pptx.title = title;
  pptx.layout = "LAYOUT_WIDE";
  return pptx;
}

function addTitleSlide(pptx, title, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: "0B132B" };
  slide.addText(title, {
    x: 0.6,
    y: 1.6,
    w: 12,
    h: 0.8,
    bold: true,
    color: "FFFFFF",
    fontSize: 38,
  });
  slide.addText(subtitle, {
    x: 0.6,
    y: 2.6,
    w: 12,
    h: 0.6,
    color: "CFE8FF",
    fontSize: 20,
  });
}

function buildTextSlide(pptx) {
  const slide = pptx.addSlide();
  slide.addText("Text formatting and layout", {
    x: 0.5,
    y: 0.3,
    w: 8,
    h: 0.5,
    bold: true,
    fontSize: 26,
  });

  slide.addText(
    [
      { text: "Rich text", options: { bold: true, color: "005F73" } },
      { text: " lets you " },
      { text: "mix styles", options: { italic: true, color: "CA6702" } },
      { text: " in one text box." },
    ],
    { x: 0.8, y: 1.2, w: 5.5, h: 0.8, fontSize: 20 }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.7,
    y: 1.1,
    w: 5.8,
    h: 2,
    radius: 0.1,
    fill: { color: "E3F2FD" },
    line: { color: "90CAF9", pt: 1 },
  });
  slide.addText("Aligned text\n• center\n• middle", {
    x: 7.1,
    y: 1.4,
    w: 5,
    h: 1.4,
    fontSize: 18,
    align: "center",
    valign: "mid",
    color: "0D47A1",
  });
}

function buildMediaSlide(pptx) {
  const slide = pptx.addSlide();
  slide.addText("Images and shapes", {
    x: 0.5,
    y: 0.3,
    w: 8,
    h: 0.5,
    bold: true,
    fontSize: 26,
  });
  slide.addImage({ data: svgImageData, x: 0.7, y: 1.1, w: 7.5, h: 4.2 });
  slide.addShape(pptx.ShapeType.chevron, {
    x: 8.6,
    y: 2.1,
    w: 3.8,
    h: 1.4,
    fill: { color: "FFB703" },
    line: { color: "FB8500", pt: 1 },
  });
  slide.addText("Shape API", {
    x: 9.15,
    y: 2.55,
    w: 2.6,
    h: 0.4,
    bold: true,
    color: "1F2937",
    align: "center",
  });
}

function buildTableSlide(pptx) {
  const slide = pptx.addSlide();
  slide.addText("Tables", {
    x: 0.5,
    y: 0.3,
    w: 4,
    h: 0.5,
    bold: true,
    fontSize: 26,
  });
  const rows = [
    [
      { text: "Feature", options: { bold: true, color: "FFFFFF" } },
      { text: "Purpose", options: { bold: true, color: "FFFFFF" } },
      { text: "Included", options: { bold: true, color: "FFFFFF" } },
    ],
    ["Text options", "Fonts, bullets, alignment", "Yes"],
    ["Media", "Images, SVG, data URIs", "Yes"],
    ["Charts", "Bar, line, pie and more", "Yes"],
    ["Masters", "Reusable slide templates", "Yes"],
  ];
  slide.addTable(rows, {
    x: 0.7,
    y: 1.2,
    w: 12,
    colW: [3, 6.5, 2.5],
    rowH: 0.6,
    fill: "F8FAFC",
    color: "1F2937",
    border: { pt: 1, color: "CBD5E1" },
    valign: "mid",
    autoFit: false,
    fontSize: 14,
    margin: 0.06,
  });
}

function buildChartSlide(pptx) {
  const slide = pptx.addSlide();
  slide.addText("Charts", {
    x: 0.5,
    y: 0.3,
    w: 4,
    h: 0.5,
    bold: true,
    fontSize: 26,
  });

  const barData = [
    { name: "Node.js", labels: ["Q1", "Q2", "Q3", "Q4"], values: [22, 31, 40, 44] },
    { name: "Browser", labels: ["Q1", "Q2", "Q3", "Q4"], values: [14, 18, 25, 29] },
  ];

  slide.addChart(pptx.ChartType.bar, barData, {
    x: 0.7,
    y: 1.2,
    w: 7.8,
    h: 4.6,
    barDir: "col",
    barGrouping: "clustered",
    catAxisLabelPos: "nextTo",
    showLegend: true,
    legendPos: "r",
    chartColors: ["3B82F6", "10B981"],
  });

  const pieData = [
    {
      name: "Usage",
      labels: ["Text", "Images", "Charts", "Tables"],
      values: [35, 20, 25, 20],
    },
  ];

  slide.addChart(pptx.ChartType.pie, pieData, {
    x: 8.7,
    y: 1.7,
    w: 3.8,
    h: 3.8,
    showLegend: false,
    showPercent: true,
    chartColors: ["1D4ED8", "0EA5E9", "8B5CF6", "F97316"],
  });
}

function buildMasterDemo(pptx) {
  pptx.defineSlideMaster({
    title: "TITLE_MASTER",
    bkgd: "0F172A",
    objects: [
      { text: { text: "PptxGenJS • Master Slide", options: { x: 0.5, y: 0.2, w: 8, h: 0.3, color: "E2E8F0", fontSize: 12 } } },
      { rect: { x: 0, y: 6.9, w: 13.333, h: 0.6, fill: { color: "1E293B" }, line: { color: "1E293B", pt: 1 } } },
    ],
  });

  const slide = pptx.addSlide("TITLE_MASTER");
  slide.addText("Master slides reduce repetition", {
    x: 0.8,
    y: 1.3,
    w: 11,
    h: 0.8,
    fontSize: 30,
    bold: true,
    color: "F8FAFC",
  });
  slide.addText("Define once, reuse across many generated slides.", {
    x: 0.8,
    y: 2.2,
    w: 9,
    h: 0.6,
    fontSize: 18,
    color: "93C5FD",
  });
}

function buildNotesSlide(pptx) {
  const slide = pptx.addSlide();
  slide.addText("Speaker Notes", {
    x: 0.5,
    y: 0.3,
    w: 4,
    h: 0.5,
    bold: true,
    fontSize: 26,
  });
  slide.addText("This slide includes speaker notes for presenters.", {
    x: 0.8,
    y: 1.4,
    w: 11,
    h: 0.8,
    fontSize: 20,
  });
  slide.addNotes(`
[Notes]
- Mention that notes are visible in Presenter View.
- Use this for scripts, reminders, and callouts.
[/Notes]
`);
}

async function writeDeck(pptx, fileName) {
  await pptx.writeFile({ fileName });
}

async function generateFullDeck() {
  const pptx = createBasePresentation("PptxGenJS full showcase");
  addTitleSlide(pptx, "PptxGenJS in the Browser", "Client-side PowerPoint generation");
  buildTextSlide(pptx);
  buildMediaSlide(pptx);
  buildTableSlide(pptx);
  buildChartSlide(pptx);
  buildMasterDemo(pptx);
  buildNotesSlide(pptx);
  await writeDeck(pptx, "pptxgenjs-full-showcase.pptx");
}

async function generateFocusedDeck(kind) {
  const titles = {
    text: "Text & layout demo",
    media: "Images and shapes demo",
    table: "Table demo",
    chart: "Chart demo",
    master: "Master slide demo",
    notes: "Speaker notes demo",
  };

  const pptx = createBasePresentation(`PptxGenJS ${titles[kind] || "demo"}`);
  addTitleSlide(pptx, "PptxGenJS Feature Demo", titles[kind] || "Feature");

  const builders = {
    text: buildTextSlide,
    media: buildMediaSlide,
    table: buildTableSlide,
    chart: buildChartSlide,
    master: buildMasterDemo,
    notes: buildNotesSlide,
  };

  builders[kind](pptx);
  await writeDeck(pptx, `pptxgenjs-${kind}-demo.pptx`);
}

fullDemoButton.addEventListener("click", async () => {
  try {
    setStatus("Generating full showcase deck...");
    await generateFullDeck();
    setStatus("Download started: pptxgenjs-full-showcase.pptx");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate full showcase deck.", true);
  }
});

focusedButtons.forEach((button) => {
  button.addEventListener("click", async () => {
    const { demo } = button.dataset;
    try {
      setStatus(`Generating ${demo} demo deck...`);
      await generateFocusedDeck(demo);
      setStatus(`Download started: pptxgenjs-${demo}-demo.pptx`);
    } catch (error) {
      console.error(error);
      setStatus(`Failed to generate ${demo} demo deck.`, true);
    }
  });
});
