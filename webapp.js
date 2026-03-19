/* global pdfjsLib, JSZip */

const state = {
  slides: [],
  timings: [],
  index: 0,
  running: false,
  direction: 1,
  loop: true,
  pure: true,
  transition: "Fade",
  timerId: null,
  sourceName: "Presentation",
};

const $ = (id) => document.getElementById(id);

const fileInput = $("fileInput");
const loadBtn = $("loadBtn");
const startBtn = $("startBtn");
const guideBtn = $("guideBtn");
const statusText = $("statusText");
const transitionSelect = $("transitionSelect");
const timingMode = $("timingMode");
const fixedField = $("fixedField");
const customField = $("customField");
const randomField = $("randomField");
const fixedSeconds = $("fixedSeconds");
const customTimes = $("customTimes");
const randomMin = $("randomMin");
const randomMax = $("randomMax");
const startSlide = $("startSlide");
const loopMode = $("loopMode");
const shuffleMode = $("shuffleMode");
const pureMode = $("pureMode");

const stageShell = $("stageShell");
const slideLayer = $("slideLayer");
const overlay = $("overlay");
const overlayTitle = $("overlayTitle");
const overlayClock = $("overlayClock");
const overlayInfo = $("overlayInfo");
const overlayProgress = $("overlayProgress");

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

function setStatus(text) {
  statusText.textContent = text;
}

function updateTimingFields() {
  const mode = timingMode.value;
  fixedField.classList.toggle("hidden", mode !== "fixed");
  customField.classList.toggle("hidden", mode !== "custom");
  randomField.classList.toggle("hidden", mode !== "random");
}

function parseTimings(slideCount) {
  const mode = timingMode.value;

  if (mode === "custom") {
    const raw = customTimes.value
      .split(",")
      .map((v) => Number(v.trim()))
      .filter((v) => Number.isFinite(v) && v > 0);
    if (!raw.length) {
      throw new Error("Custom timing is empty. Example: 10,20,15");
    }
    while (raw.length < slideCount) raw.push(raw[raw.length - 1]);
    return raw.slice(0, slideCount);
  }

  if (mode === "random") {
    const min = Number(randomMin.value);
    const max = Number(randomMax.value);
    if (!(min > 0) || !(max > 0) || min > max) {
      throw new Error("Invalid random min/max.");
    }
    return Array.from({ length: slideCount }, () => +(Math.random() * (max - min) + min).toFixed(2));
  }

  const fixed = Number(fixedSeconds.value);
  if (!(fixed > 0)) {
    throw new Error("Fixed timing must be greater than 0.");
  }
  return Array(slideCount).fill(fixed);
}

function getMime(ext) {
  const map = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    webp: "image/webp",
  };
  return map[ext.toLowerCase()] || "application/octet-stream";
}

function sortSlideNames(names) {
  return names.sort((a, b) => {
    const na = Number(a.match(/slide(\d+)\.xml$/i)?.[1] || 0);
    const nb = Number(b.match(/slide(\d+)\.xml$/i)?.[1] || 0);
    return na - nb;
  });
}

function normalizePath(fromDir, target) {
  const stack = fromDir.split("/").filter(Boolean);
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function textNodesFromXml(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...doc.querySelectorAll("a\\:t, t")]
    .map((n) => n.textContent?.trim() || "")
    .filter(Boolean);
}

async function parsePptx(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const allNames = Object.keys(zip.files);
  const slideXmlNames = sortSlideNames(
    allNames.filter((n) => n.startsWith("ppt/slides/slide") && n.endsWith(".xml"))
  );

  if (!slideXmlNames.length) {
    throw new Error("No slides found in PPTX.");
  }

  const slides = [];
  for (const slideXmlName of slideXmlNames) {
    const xmlText = await zip.file(slideXmlName).async("string");
    const texts = textNodesFromXml(xmlText);

    const relPath = slideXmlName.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const relMap = {};
    if (zip.file(relPath)) {
      const relText = await zip.file(relPath).async("string");
      const relDoc = new DOMParser().parseFromString(relText, "application/xml");
      for (const rel of [...relDoc.querySelectorAll("Relationship")]) {
        const id = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        if (id && target) relMap[id] = target;
      }
    }

    const slideDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    const blips = [...slideDoc.querySelectorAll("a\\:blip, blip")];
    const images = [];

    for (const blip of blips) {
      const rid = blip.getAttribute("r:embed") || blip.getAttribute("embed");
      if (!rid || !relMap[rid]) continue;
      const target = normalizePath("ppt/slides", relMap[rid]);
      const mediaFile = zip.file(target);
      if (!mediaFile) continue;
      const ext = target.split(".").pop() || "png";
      const base64 = await mediaFile.async("base64");
      images.push(`data:${getMime(ext)};base64,${base64}`);
    }

    slides.push({
      kind: "pptx",
      title: `Slide ${slides.length + 1}`,
      texts,
      images,
    });
  }

  return slides;
}

async function parsePdf(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const slides = [];

  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.8 });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d", { alpha: false });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    slides.push({ kind: "image", src: canvas.toDataURL("image/png") });
  }

  return slides;
}

function renderSlide(slide) {
  const wrap = document.createElement("div");
  wrap.className = "slide-node";

  if (slide.kind === "image") {
    const img = document.createElement("img");
    img.src = slide.src;
    img.alt = "Slide";
    wrap.appendChild(img);
    return wrap;
  }

  const card = document.createElement("article");
  card.className = "slide-html";
  const title = document.createElement("h2");
  title.textContent = slide.title || "Slide";
  card.appendChild(title);

  const list = document.createElement("ul");
  if (slide.texts.length) {
    for (const t of slide.texts.slice(0, 30)) {
      const li = document.createElement("li");
      li.textContent = t;
      list.appendChild(li);
    }
  } else {
    const li = document.createElement("li");
    li.textContent = "No text detected on this slide.";
    list.appendChild(li);
  }
  card.appendChild(list);

  if (slide.images.length) {
    const media = document.createElement("div");
    media.className = "slide-media";
    for (const src of slide.images.slice(0, 12)) {
      const img = document.createElement("img");
      img.src = src;
      img.alt = "Slide media";
      media.appendChild(img);
    }
    card.appendChild(media);
  }

  wrap.appendChild(card);
  return wrap;
}

function transitionClass(name) {
  if (name === "Fade") return "fade-in";
  if (name === "Slide Left") return "slide-left";
  if (name === "Zoom In") return "zoom-in";
  return "";
}

function updateOverlay() {
  const secs = state.timings[state.index] ?? 0;
  const running = state.running ? "Running" : "Paused";
  overlayTitle.textContent = state.sourceName;
  overlayClock.textContent = new Date().toLocaleTimeString();
  overlayInfo.textContent = `Slide ${state.index + 1}/${state.slides.length} | ${secs.toFixed(1)}s | ${running}`;
  overlayProgress.max = Math.max(1, state.slides.length);
  overlayProgress.value = state.index + 1;
}

function showSlide(index, animate = true) {
  state.index = Math.max(0, Math.min(index, state.slides.length - 1));
  slideLayer.innerHTML = "";
  const node = renderSlide(state.slides[state.index]);
  if (animate) node.classList.add(transitionClass(state.transition));
  slideLayer.appendChild(node);
  updateOverlay();
}

function nextIndex(step) {
  const target = state.index + step;
  if (target < 0) return state.loop ? state.slides.length - 1 : 0;
  if (target >= state.slides.length) return state.loop ? 0 : state.slides.length - 1;
  return target;
}

function schedule() {
  clearTimeout(state.timerId);
  if (!state.running) return;
  const ms = Math.max(500, (state.timings[state.index] || 1) * 1000);
  state.timerId = setTimeout(() => {
    if (!state.running) return;
    const newIndex = nextIndex(state.direction);
    if (newIndex === state.index && !state.loop) {
      state.running = false;
      updateOverlay();
      return;
    }
    showSlide(newIndex, true);
    schedule();
  }, ms);
}

function toggleGuide() {
  overlay.classList.toggle("hidden");
}

async function loadDeck() {
  const file = fileInput.files?.[0];
  if (!file) {
    throw new Error("Choose a PDF or PPTX file first.");
  }

  setStatus("Loading slides...");
  const ext = file.name.toLowerCase().split(".").pop();
  let slides;
  if (ext === "pdf") {
    slides = await parsePdf(file);
  } else if (ext === "pptx") {
    slides = await parsePptx(file);
  } else {
    throw new Error("Only PDF and PPTX are supported.");
  }

  let timings = parseTimings(slides.length);

  if (shuffleMode.checked) {
    const zipped = slides.map((s, i) => [s, timings[i]]);
    for (let i = zipped.length - 1; i > 0; i -= 1) {
      const j = Math.floor(Math.random() * (i + 1));
      [zipped[i], zipped[j]] = [zipped[j], zipped[i]];
    }
    slides = zipped.map((x) => x[0]);
    timings = zipped.map((x) => x[1]);
  }

  state.slides = slides;
  state.timings = timings;
  state.transition = transitionSelect.value;
  state.loop = loopMode.checked;
  state.pure = pureMode.checked;
  state.sourceName = file.name;
  state.direction = 1;

  const start = Math.max(1, Math.min(Number(startSlide.value) || 1, slides.length));
  showSlide(start - 1, false);

  overlay.classList.toggle("hidden", state.pure);
  setStatus(`Loaded ${slides.length} slides from ${file.name}.`);
  startBtn.disabled = false;
}

function startPresentation() {
  if (!state.slides.length) return;
  state.running = true;
  showSlide(state.index, false);
  schedule();
  stageShell.requestFullscreen?.();
}

function saveSnapshot() {
  const node = slideLayer.querySelector("img");
  if (!node) return;
  const a = document.createElement("a");
  a.href = node.src;
  a.download = `slide-${state.index + 1}.png`;
  a.click();
}

function initEvents() {
  timingMode.addEventListener("change", updateTimingFields);

  loadBtn.addEventListener("click", async () => {
    try {
      await loadDeck();
    } catch (err) {
      setStatus(err.message || "Failed to load file.");
    }
  });

  startBtn.addEventListener("click", startPresentation);
  guideBtn.addEventListener("click", toggleGuide);

  document.addEventListener("keydown", (e) => {
    if (!state.slides.length) return;
    const key = e.key.toLowerCase();

    if (key === "arrowright") {
      showSlide(nextIndex(1), true);
      schedule();
    } else if (key === "arrowleft") {
      showSlide(nextIndex(-1), true);
      schedule();
    } else if (key === " ") {
      state.running = !state.running;
      schedule();
      updateOverlay();
      e.preventDefault();
    } else if (key === "b") {
      stageShell.classList.toggle("hidden-content");
      if (stageShell.classList.contains("hidden-content")) {
        slideLayer.innerHTML = "";
      } else {
        showSlide(state.index, false);
      }
    } else if (key === "g") {
      toggleGuide();
    } else if (key === "j") {
      const result = Number(prompt(`Jump to slide (1-${state.slides.length})`, state.index + 1));
      if (Number.isFinite(result)) showSlide(result - 1, false);
    } else if (key === "s") {
      saveSnapshot();
    } else if (key === "r") {
      state.direction *= -1;
      updateOverlay();
    } else if (key === "+" || key === "=") {
      state.timings[state.index] = Math.max(0.5, +(state.timings[state.index] + 1).toFixed(2));
      updateOverlay();
      schedule();
    } else if (key === "-") {
      state.timings[state.index] = Math.max(0.5, +(state.timings[state.index] - 1).toFixed(2));
      updateOverlay();
      schedule();
    }
  });

  setInterval(() => {
    if (!overlay.classList.contains("hidden")) updateOverlay();
  }, 1000);
}

updateTimingFields();
initEvents();
