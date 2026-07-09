// Minimal SVG chart kit (no dependencies) following a fixed mark spec:
//   bars <=24px, 4px rounded data-end + square baseline, 2px lines with a
//   ~10% area wash, hairline solid gridlines, markers with a 2px surface
//   ring, selective direct labels, and hover tooltips on every plot.
// Colors are CSS custom properties so light/dark swap automatically.

const NS = "http://www.w3.org/2000/svg";

function el(tag, attrs = {}, parent = null) {
  const node = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
  if (parent) parent.appendChild(node);
  return node;
}

// ------------------------------------------------------------- tooltip ----

let tipEl = null;
function tip() {
  if (!tipEl) {
    tipEl = document.createElement("div");
    tipEl.className = "chart-tip";
    tipEl.setAttribute("role", "status");
    document.body.appendChild(tipEl);
  }
  return tipEl;
}
export function showTip(html, x, y) {
  const t = tip();
  t.innerHTML = html;
  t.style.display = "block";
  const r = t.getBoundingClientRect();
  const left = Math.min(Math.max(8, x + 14), window.innerWidth - r.width - 8);
  const top = y - r.height - 12 < 8 ? y + 16 : y - r.height - 12;
  t.style.left = `${left}px`;
  t.style.top = `${top}px`;
}
export function hideTip() {
  if (tipEl) tipEl.style.display = "none";
}

// --------------------------------------------------------------- scales ----

// Clean axis scale: pick a step from {1, 2, 2.5, 5}×10^k so every gridline
// lands on a round number, then take the smallest multiple that covers v.
function niceScale(v) {
  if (v <= 0) return { max: 4, step: 1 };
  const pow = 10 ** Math.floor(Math.log10(v));
  for (const m of [0.2, 0.25, 0.5, 1, 2, 2.5, 5, 10]) {
    const step = m * pow;
    if (step * 5 >= v) return { step, max: Math.ceil(v / step - 1e-9) * step };
  }
  return { step: 10 * pow, max: 10 * pow };
}

function frame(container, height) {
  container.innerHTML = "";
  const width = Math.max(220, container.clientWidth || 320);
  const svg = el("svg", {
    viewBox: `0 0 ${width} ${height}`,
    width: "100%",
    height,
    role: "img",
  });
  container.appendChild(svg);
  return { svg, width };
}

function yAxis(svg, { left, right, top, bottom, height, width, max, step, fmt }) {
  const plotH = height - top - bottom;
  const ticks = Math.round(max / step);
  for (let i = 0; i <= ticks; i++) {
    const val = step * i;
    const y = height - bottom - (plotH * i) / ticks;
    if (i > 0)
      el("line", {
        x1: left, x2: width - right, y1: y, y2: y,
        class: "grid",
      }, svg);
    el("text", {
      x: left - 6, y: y + 3.5, "text-anchor": "end", class: "tick",
    }, svg).textContent = fmt(val);
  }
  el("line", {
    x1: left, x2: width - right,
    y1: height - bottom, y2: height - bottom,
    class: "axis",
  }, svg);
  return { plotH };
}

// Bar path: 4px rounded top corners, square at the baseline.
function barPath(x, y, w, h, r = 4) {
  if (h <= 0) return "";
  const rr = Math.min(r, w / 2, h);
  return `M${x},${y + h} L${x},${y + rr} Q${x},${y} ${x + rr},${y}
          L${x + w - rr},${y} Q${x + w},${y} ${x + w},${y + rr}
          L${x + w},${y + h} Z`;
}

// -------------------------------------------------------- column chart ----

export function columnChart(container, { labels, values, fmt, fmtAxis, tipLabel }) {
  const height = 240;
  const { svg, width } = frame(container, height);
  const left = 46, right = 10, top = 18, bottom = 26;
  const { max, step } = niceScale(Math.max(...values, 1));
  const { plotH } = yAxis(svg, { left, right, top, bottom, height, width, max, step, fmt: fmtAxis });

  const n = labels.length;
  const band = (width - left - right) / n;
  const barW = Math.min(24, Math.max(6, band * 0.55));
  const maxIdx = values.indexOf(Math.max(...values));
  const labelEvery = band < 34 ? 2 : 1;

  labels.forEach((lab, i) => {
    const h = (values[i] / max) * plotH;
    const x = left + band * i + (band - barW) / 2;
    const y = height - bottom - h;
    el("path", { d: barPath(x, y, barW, h), class: "bar", "data-i": i }, svg);

    // direct label on the tallest column only; ticks carry the rest
    if (i === maxIdx && values[i] > 0) {
      el("text", {
        x: x + barW / 2, y: y - 6, "text-anchor": "middle", class: "val",
      }, svg).textContent = fmt(values[i]);
    }
    if (i % labelEvery === 0) {
      el("text", {
        x: left + band * i + band / 2, y: height - bottom + 16,
        "text-anchor": "middle", class: "tick",
      }, svg).textContent = lab;
    }
    // full-band hover target (bigger than the mark)
    const hit = el("rect", {
      x: left + band * i, y: top, width: band, height: plotH + bottom - 4,
      fill: "transparent",
    }, svg);
    hit.addEventListener("pointerenter", (ev) => {
      svg.querySelectorAll(".bar").forEach((b) => b.classList.toggle("dim", b.dataset.i != i));
      showTip(
        `<b>${tipLabel ? tipLabel(i) : lab}</b><br>${fmt(values[i])}`,
        ev.clientX, ev.clientY
      );
    });
    hit.addEventListener("pointermove", (ev) =>
      showTip(
        `<b>${tipLabel ? tipLabel(i) : lab}</b><br>${fmt(values[i])}`,
        ev.clientX, ev.clientY
      )
    );
  });
  svg.addEventListener("pointerleave", () => {
    hideTip();
    svg.querySelectorAll(".bar").forEach((b) => b.classList.remove("dim"));
  });
}

// ---------------------------------------------------------- line chart ----

export function lineChart(container, { labels, values, fmt, fmtEnd, fmtAxis, fmtLabel }) {
  const height = 240;
  const { svg, width } = frame(container, height);
  const left = 46, right = 62, top = 18, bottom = 26;
  const { max, step } = niceScale(Math.max(...values, 1));
  const { plotH } = yAxis(svg, { left, right, top, bottom, height, width, max, step, fmt: fmtAxis });

  const n = values.length;
  const px = (i) => left + ((width - left - right) * i) / Math.max(1, n - 1);
  const py = (v) => height - bottom - (v / max) * plotH;

  const pts = values.map((v, i) => `${px(i)},${py(v)}`);
  el("path", {
    d: `M${pts.join(" L")} L${px(n - 1)},${height - bottom} L${px(0)},${height - bottom} Z`,
    class: "area",
  }, svg);
  el("path", { d: `M${pts.join(" L")}`, class: "line" }, svg);

  // x labels: first / middle / last
  [0, Math.floor((n - 1) / 2), n - 1].forEach((i, k) => {
    el("text", {
      x: px(i), y: height - bottom + 16,
      "text-anchor": k === 0 ? "start" : k === 2 ? "end" : "middle",
      class: "tick",
    }, svg).textContent = fmtLabel ? fmtLabel(labels[i]) : labels[i];
  });

  // end marker (>=8px with 2px surface ring) + end label
  el("circle", { cx: px(n - 1), cy: py(values[n - 1]), r: 4.5, class: "dot" }, svg);
  el("text", {
    x: px(n - 1) + 8, y: py(values[n - 1]) + 4, class: "val",
  }, svg).textContent = (fmtEnd || fmt)(values[n - 1]);

  // crosshair + tooltip
  const cross = el("line", { y1: top, y2: height - bottom, class: "crosshair", style: "display:none" }, svg);
  const hoverDot = el("circle", { r: 4.5, class: "dot", style: "display:none" }, svg);
  const hit = el("rect", { x: left, y: top, width: width - left - right, height: plotH, fill: "transparent" }, svg);
  hit.addEventListener("pointermove", (ev) => {
    const box = svg.getBoundingClientRect();
    const sx = ((ev.clientX - box.left) / box.width) * width;
    const i = Math.round(((sx - left) / (width - left - right)) * (n - 1));
    const ci = Math.max(0, Math.min(n - 1, i));
    cross.setAttribute("x1", px(ci));
    cross.setAttribute("x2", px(ci));
    cross.style.display = "";
    hoverDot.setAttribute("cx", px(ci));
    hoverDot.setAttribute("cy", py(values[ci]));
    hoverDot.style.display = "";
    showTip(
      `<b>${fmtLabel ? fmtLabel(labels[ci]) : labels[ci]}</b><br>${fmt(values[ci])}`,
      ev.clientX, ev.clientY
    );
  });
  hit.addEventListener("pointerleave", () => {
    cross.style.display = "none";
    hoverDot.style.display = "none";
    hideTip();
  });
}

// ----------------------------------------------- horizontal bar chart ----

export function hbarChart(container, { items, fmt, colorFor }) {
  const rowH = 30, labelW = 132, valueW = 76;
  const height = items.length * rowH + 6;
  const { svg, width } = frame(container, height);
  const max = Math.max(...items.map((it) => it.value), 1);
  const plotW = width - labelW - valueW - 10;

  items.forEach((it, i) => {
    const y = i * rowH + 6;
    const w = Math.max(2, (it.value / max) * plotW);
    el("text", {
      x: labelW, y: y + 12, "text-anchor": "end", class: "cat-label",
    }, svg).textContent = it.label.length > 20 ? it.label.slice(0, 19) + "…" : it.label;

    // rounded right data-end, square at the left baseline
    const bar = el("path", {
      d: `M${labelW + 10},${y} L${labelW + 10 + w - 4},${y} Q${labelW + 10 + w},${y} ${labelW + 10 + w},${y + 4}
          L${labelW + 10 + w},${y + 12} Q${labelW + 10 + w},${y + 16} ${labelW + 10 + w - 4},${y + 16}
          L${labelW + 10},${y + 16} Z`,
      fill: colorFor ? colorFor(it.label, i) : "var(--series-1)",
    }, svg);
    el("text", {
      x: labelW + 10 + w + 8, y: y + 12, class: "val",
    }, svg).textContent = fmt(it.value);

    const hit = el("rect", {
      x: 0, y: y - 6, width, height: rowH, fill: "transparent",
    }, svg);
    hit.addEventListener("pointerenter", (ev) =>
      showTip(`<b>${it.label}</b><br>${fmt(it.value)}${it.extra ? `<br><span class="tip-sub">${it.extra}</span>` : ""}`, ev.clientX, ev.clientY));
    hit.addEventListener("pointermove", (ev) =>
      showTip(`<b>${it.label}</b><br>${fmt(it.value)}${it.extra ? `<br><span class="tip-sub">${it.extra}</span>` : ""}`, ev.clientX, ev.clientY));
    hit.addEventListener("pointerleave", hideTip);
    bar.style.pointerEvents = "none";
  });
  el("line", {
    x1: labelW + 10, x2: labelW + 10, y1: 0, y2: height, class: "axis",
  }, svg);
}

// ------------------------------------------------------------ sparkline ----

export function sparkline(container, values) {
  container.innerHTML = "";
  const w = 96, h = 28;
  const svg = el("svg", { viewBox: `0 0 ${w} ${h}`, width: w, height: h, "aria-hidden": "true" });
  container.appendChild(svg);
  const max = Math.max(...values, 1);
  const px = (i) => 2 + ((w - 8) * i) / Math.max(1, values.length - 1);
  const py = (v) => h - 3 - (v / max) * (h - 8);
  const pts = values.map((v, i) => `${px(i)},${py(v)}`).join(" L");
  el("path", { d: `M${pts}`, class: "spark" }, svg);
  el("circle", {
    cx: px(values.length - 1), cy: py(values[values.length - 1]), r: 3, class: "dot",
  }, svg);
}
