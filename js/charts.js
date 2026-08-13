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

// Clean axis scale over a value range that may include negatives (credits and
// refunds). Steps come from {1, 2, 2.5, 5}×10^k so every gridline lands on a
// round number; the domain always includes zero so the baseline is meaningful.
function niceDomain(values) {
  const lo = Math.min(0, ...values);
  const hi = Math.max(0, ...values);
  if (lo === 0 && hi === 0) return { min: 0, max: 4, step: 1 }; // nothing to plot yet
  const span = hi - lo;
  const pow = 10 ** Math.floor(Math.log10(span));
  let step = 10 * pow;
  for (const m of [0.2, 0.25, 0.5, 1, 2, 2.5, 5, 10]) {
    if (m * pow * 5 >= span) { step = m * pow; break; }
  }
  const min = Math.floor(lo / step + 1e-9) * step;
  let max = Math.ceil(hi / step - 1e-9) * step;
  if (max <= min) max = min + step;
  return { min, max, step };
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

// Draws gridlines + tick labels across the domain and returns the value→y
// scale. The zero line gets the axis treatment wherever it falls, so negative
// marks hang below it.
function yAxis(svg, { left, right, top, bottom, height, width, min, max, step, fmt }) {
  const plotH = height - top - bottom;
  const y = (v) => height - bottom - ((v - min) / (max - min)) * plotH;
  const ticks = Math.max(1, Math.round((max - min) / step));
  for (let i = 0; i <= ticks; i++) {
    const val = min + step * i;
    const yy = y(val);
    if (Math.abs(val) > 1e-9)
      el("line", {
        x1: left, x2: width - right, y1: yy, y2: yy,
        class: "grid",
      }, svg);
    el("text", {
      x: left - 6, y: yy + 3.5, "text-anchor": "end", class: "tick",
    }, svg).textContent = fmt(val);
  }
  const zeroY = y(0);
  el("line", {
    x1: left, x2: width - right, y1: zeroY, y2: zeroY, class: "axis",
  }, svg);
  return { plotH, y, zeroY };
}

// Bar path: 4px rounded data-end, square where it meets the baseline. Bars
// below zero are rounded on their bottom edge instead.
function barPath(x, y, w, h, roundTop = true) {
  if (h <= 0.5) return `M${x},${y} L${x + w},${y} L${x + w},${y + 1} L${x},${y + 1} Z`;
  const rr = Math.min(4, w / 2, h);
  if (roundTop) {
    return `M${x},${y + h} L${x},${y + rr} Q${x},${y} ${x + rr},${y}
            L${x + w - rr},${y} Q${x + w},${y} ${x + w},${y + rr}
            L${x + w},${y + h} Z`;
  }
  return `M${x},${y} L${x},${y + h - rr} Q${x},${y + h} ${x + rr},${y + h}
          L${x + w - rr},${y + h} Q${x + w},${y + h} ${x + w},${y + h - rr}
          L${x + w},${y} Z`;
}

// -------------------------------------------------------- column chart ----

export function columnChart(container, { labels, values, fmt, fmtAxis, tipLabel }) {
  const height = 240;
  const { svg, width } = frame(container, height);
  const left = 46, right = 10, top = 18, bottom = 26;
  const { min, max, step } = niceDomain(values);
  const { plotH, y: yv, zeroY } = yAxis(svg, {
    left, right, top, bottom, height, width, min, max, step, fmt: fmtAxis,
  });

  const n = labels.length;
  const band = (width - left - right) / n;
  const barW = Math.min(24, Math.max(6, band * 0.55));
  const maxIdx = values.indexOf(Math.max(...values));
  const minIdx = values.indexOf(Math.min(...values));
  const labelEvery = band < 34 ? 2 : 1;

  labels.forEach((lab, i) => {
    const v = values[i];
    const up = v >= 0;
    const vy = yv(v);
    const yTop = Math.min(vy, zeroY);
    const h = Math.abs(vy - zeroY);
    const x = left + band * i + (band - barW) / 2;
    el("path", {
      d: barPath(x, yTop, barW, h, up),
      class: `bar${up ? "" : " neg"}`,
      "data-i": i,
    }, svg);

    // Direct labels on the extremes only; ticks carry the rest. A label under
    // a credit bar is dropped when it would land on the month labels — the
    // axis ticks and tooltip still carry the value.
    const labelY = up ? vy - 6 : vy + 14;
    const fits = up || labelY <= height - bottom - 4;
    if (((i === maxIdx && v > 0) || (i === minIdx && v < 0)) && fits) {
      el("text", {
        x: x + barW / 2, y: labelY, "text-anchor": "middle", class: "val",
      }, svg).textContent = fmt(v);
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
  const { min, max, step } = niceDomain(values);
  const { plotH, y: py, zeroY } = yAxis(svg, {
    left, right, top, bottom, height, width, min, max, step, fmt: fmtAxis,
  });

  const n = values.length;
  const px = (i) => left + ((width - left - right) * i) / Math.max(1, n - 1);

  const pts = values.map((v, i) => `${px(i)},${py(v)}`);
  el("path", {
    d: `M${pts.join(" L")} L${px(n - 1)},${zeroY} L${px(0)},${zeroY} Z`,
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
  const rowH = 30, labelW = 132;
  const height = items.length * rowH + 6;
  const { svg, width } = frame(container, height);
  const values = items.map((it) => it.value);
  const lo = Math.min(0, ...values);
  const hi = Math.max(0, ...values);
  const span = hi - lo || 1;
  const hasNeg = lo < 0;
  // credits need a wider value column (minus sign) and it is right-aligned,
  // so long amounts can never be clipped at the edge
  const valueW = hasNeg ? 96 : 76;
  const x0 = labelW + 10;
  const plotW = width - labelW - valueW - 10;
  const x = (v) => x0 + ((v - lo) / span) * plotW;
  const zeroX = x(0);

  items.forEach((it, i) => {
    const y = i * rowH + 6;
    const up = it.value >= 0;
    const vx = x(it.value);
    const barL = Math.min(vx, zeroX);
    const w = Math.max(2, Math.abs(vx - zeroX));
    el("text", {
      x: labelW, y: y + 12, "text-anchor": "end", class: "cat-label",
    }, svg).textContent = it.label.length > 20 ? it.label.slice(0, 19) + "…" : it.label;

    // rounded data-end (right for positive, left for a credit), square at zero
    const r = 4;
    const bar = el("path", {
      d: up
        ? `M${barL},${y} L${barL + w - r},${y} Q${barL + w},${y} ${barL + w},${y + r}
           L${barL + w},${y + 12} Q${barL + w},${y + 16} ${barL + w - r},${y + 16}
           L${barL},${y + 16} Z`
        : `M${barL + w},${y} L${barL + r},${y} Q${barL},${y} ${barL},${y + r}
           L${barL},${y + 12} Q${barL},${y + 16} ${barL + r},${y + 16}
           L${barL + w},${y + 16} Z`,
      fill: up ? (colorFor ? colorFor(it.label, i) : "var(--series-1)") : "var(--good)",
    }, svg);
    // With credits in play the bars start at different x, so values get their
    // own right-hand column instead of trailing each bar end.
    el("text", {
      x: hasNeg ? width - 2 : barL + w + 8,
      y: y + 12, class: "val",
      ...(hasNeg ? { "text-anchor": "end" } : {}),
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
    x1: zeroX, x2: zeroX, y1: 0, y2: height, class: "axis",
  }, svg);
}

// ------------------------------------------------------------ sparkline ----

export function sparkline(container, values) {
  container.innerHTML = "";
  const w = 96, h = 28;
  const svg = el("svg", { viewBox: `0 0 ${w} ${h}`, width: w, height: h, "aria-hidden": "true" });
  container.appendChild(svg);
  const lo = Math.min(0, ...values);
  const span = Math.max(...values, 0) - lo || 1;
  const px = (i) => 2 + ((w - 8) * i) / Math.max(1, values.length - 1);
  const py = (v) => h - 3 - ((v - lo) / span) * (h - 8);
  const pts = values.map((v, i) => `${px(i)},${py(v)}`).join(" L");
  el("path", { d: `M${pts}`, class: "spark" }, svg);
  el("circle", {
    cx: px(values.length - 1), cy: py(values[values.length - 1]), r: 3, class: "dot",
  }, svg);
}
