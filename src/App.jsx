import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const SF     = "-apple-system,'SF Pro Display','SF Pro Text',BlinkMacSystemFont,'Helvetica Neue',sans-serif";
const SFMono = "'SF Mono','SFMono-Regular',ui-monospace,Menlo,monospace";

// ─── PARSER ──────────────────────────────────────────────────────────────────

function parseControlLimit(raw) {
  if (!raw || typeof raw !== "string") return { min: null, max: null };
  const s = raw.replace(/,/g, "").trim();
  const range = s.match(/([\d.]+)\s*[-–]\s*([\d.]+)/);
  if (range) return { min: parseFloat(range[1]), max: parseFloat(range[2]) };
  const lt = s.match(/^<\s*([\d.]+)/);
  if (lt) return { min: null, max: parseFloat(lt[1]) };
  const gt = s.match(/^>\s*([\d.]+)/);
  if (gt) return { min: parseFloat(gt[1]), max: null };
  return { min: null, max: null };
}

function tryNum(v) {
  if (v === null || v === undefined || v === "" || v === "-") return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

const isTextParam = name => /appearance|visual|colour|color/i.test(name);
const isTextOk    = val  => /clear/i.test(String(val ?? ""));

function parseSheet(sheet, sheetName) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // ── Detect format ──────────────────────────────────────────
  let isCOA = false; // Certificate of Analysis (Ambu, Adventist)
  for (let i = 0; i < Math.min(25, rows.length); i++) {
    const s = rows[i].map(c => c != null ? String(c) : "").join(" ");
    if (/CERTIFICATE OF ANALYSIS/i.test(s)) { isCOA = true; break; }
  }

  // ── Extract metadata ───────────────────────────────────────
  let reportMonth = sheetName, company = "", attention = "", cc = "", systemLabel = "";
  for (let i = 0; i < Math.min(30, rows.length); i++) {
    const r = rows[i];
    const rowStr = r.map(c => c != null ? String(c) : "").join(" ");
    // CWTAR: "REPORT FOR THE MONTH OF : January 2025"
    if (/REPORT FOR THE MONTH OF/i.test(rowStr)) {
      const parts = r.filter(c => c != null && String(c).trim() !== "");
      const last = parts[parts.length - 1];
      if (last && !/MONTH/i.test(String(last))) reportMonth = String(last);
    }
    // COA: "DATE : 3rd January 2025" (use as reportMonth)
    if (isCOA && /^\s*DATE\s*$/.test(String(r[1] || "")) && r[5]) {
      reportMonth = String(r[5]).trim();
    }
    if (/^\s*TO\s*$/.test(String(r[1] || "")))    company   = String(r[5] || "");
    if (/ATTENTION/i.test(String(r[1] || "")))     attention = String(r[5] || "");
    if (/^\s*CC\.?\s*$/.test(String(r[1] || ""))) cc        = String(r[5] || "");
    // System label: row 17 in CWTAR (e.g. "1.3 : B.BRAUN...")
    if (/^\d+\.\d+\s*:/.test(String(r[1] || "")) && !systemLabel)
      systemLabel = String(r[1]).trim().slice(0, 80);
  }

  // ── Find header row (has "TYPE OF PARAMETER") ─────────────
  let headerRowIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const s = rows[i].map(c => c != null ? String(c) : "").join(" ");
    if (/TYPE OF PARAMETER/i.test(s)) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) return null;

  const headerRow     = rows[headerRowIdx]     || [];
  const headerRowNext = rows[headerRowIdx + 1] || [];

  // ── Find CONTROL LIMIT column ──────────────────────────────
  let controlLimitCol = -1;
  for (let c = (headerRow.length || 0) - 1; c >= 0; c--) {
    const v = String(headerRow[c] || headerRowNext[c] || "").trim();
    if (/CONTROL|LIMIT/i.test(v)) { controlLimitCol = c; break; }
  }
  if (controlLimitCol === -1) controlLimitCol = (headerRow.length || 1) - 1;

  let cityWaterCol = -1;
  let sampleCols   = [];
  let sampleDates  = [];
  let dataRowStart = headerRowIdx + 1;

  if (isCOA) {
    // ── COA format ─────────────────────────────────────────
    // City water = M/U column (in header row)
    for (let c = 0; c < headerRow.length; c++) {
      const v = String(headerRow[c] || "").trim();
      if (/^M\/U/i.test(v)) { cityWaterCol = c; break; }
    }
    // Sample cols = between cityWaterCol+1 and controlLimitCol (exclusive)
    for (let c = (cityWaterCol >= 0 ? cityWaterCol + 1 : 6); c < controlLimitCol; c++) {
      sampleCols.push(c);
    }
    // Sample labels = the column header text (e.g. "B1", "B1 Process", "COOLING TOWER(750RT)")
    sampleDates = sampleCols.map(c => String(headerRow[c] || "").trim() || `S${c}`);
    dataRowStart = headerRowIdx + 1;

  } else {
    // ── CWTAR format (existing logic) ─────────────────────
    const dateRow = rows[headerRowIdx - 1] || [];
    // Detect City Water column from 2-row header span
    for (let c = 0; c < Math.max(headerRow.length || 0, headerRowNext.length || 0); c++) {
      const a = String(headerRow[c]     ?? "").toLowerCase();
      const b = String(headerRowNext[c] ?? "").toLowerCase();
      if (a.includes("city") || b.includes("city") || a.includes("water") || b.includes("water")) {
        cityWaterCol = c; break;
      }
    }
    const startScanCol = cityWaterCol >= 0 ? cityWaterCol + 1 : 6;
    for (let c = startScanCol; c < dateRow.length; c++) {
      if (dateRow[c] != null && String(dateRow[c]).trim() !== "") sampleCols.push(c);
    }
    sampleCols = sampleCols.filter(c => c !== controlLimitCol);
    sampleDates = sampleCols.map((c, si) => {
      const v = dateRow[c];
      if (!v) return `S${si+1}`;
      if (typeof v === "number") {
        const d = XLSX.SSF.parse_date_code(v);
        if (d) return `${d.d}/${d.m}`;
      }
      return String(v).replace(/\d{4}/, "").replace(/^[-/]|[-/]$/g, "").trim() || `S${si+1}`;
    });
    dataRowStart = headerRowIdx + 3; // skip 2-row header + blank row
  }

  // ── Parse data rows ────────────────────────────────────────
  const parameters = [];
  for (let i = dataRowStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c == null)) continue;
    const paramName = row[2] != null ? String(row[2]).trim() : null;
    if (!paramName || paramName === "") continue;
    if (/Kuritex.*=.*Corrosion|M\/U\s*=\s*MAKE/i.test(paramName)) continue;
    if (/^\*\s*M\/U/i.test(String(row[1] || ""))) continue;
    if (/REMARKS/i.test(String(row[1] || ""))) break;
    if (/END OF REPORT|Standard Method/i.test(paramName)) break;

    const cityWater = cityWaterCol >= 0 && row[cityWaterCol] != null
      ? String(row[cityWaterCol]).trim() : null;
    const limitRaw = row[controlLimitCol] != null ? String(row[controlLimitCol]).trim() : null;
    const { min, max } = parseControlLimit(limitRaw);
    const isText   = isTextParam(paramName);
    const methodNum = row[1] != null ? String(row[1]).trim() : null;

    // Keep full structure — null = no data, shown as "—"
    const samples = isText
      ? sampleCols.map(c => row[c] != null ? String(row[c]).trim() : null)
      : sampleCols.map(c => tryNum(row[c]));

    // Only skip if absolutely no data at all
    const hasAnyData = samples.some(v => v !== null) || (cityWater && cityWater !== "-");
    if (!hasAnyData) continue;

    parameters.push({ name: paramName, methodNum, cityWater, samples, limitRaw, min, max, isText });
  }

  const remarks = [];
  for (let i = dataRowStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const cell = String(row[1] || "").trim();
    if (cell.length > 30 && !/Method|Standard|HANNA|HACH|Merck|BKG|M\/U/i.test(cell)) remarks.push(cell);
  }

  return { reportMonth, company, attention, cc, systemLabel, parameters, sampleDates, isCOA };
}

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: false });
  const result = { sheets: {}, sheetOrder: [], meta: {} };
  for (const name of wb.SheetNames) {
    const parsed = parseSheet(wb.Sheets[name], name);
    if (parsed && parsed.parameters.length > 0) {
      result.sheets[name] = parsed;
      result.sheetOrder.push(name);
      if (!result.meta.company)
        result.meta = { company: parsed.company, attention: parsed.attention, cc: parsed.cc, systemLabel: parsed.systemLabel };
    }
  }
  return result;
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

const avgNums = arr => {
  const n = arr.filter(v => typeof v === "number" && !isNaN(v));
  return n.length ? n.reduce((a,b) => a+b, 0) / n.length : null;
};
const isInRange = (val, min, max) => {
  if (val === null || val === undefined) return true;
  if (min !== null && val < min) return false;
  if (max !== null && val > max) return false;
  return true;
};

function buildFlatSeries(paramName, isText, min, max, sheets, sheetOrder) {
  const pts = [];
  for (const sh of sheetOrder) {
    const p = sheets[sh].parameters.find(x => x.name === paramName);
    if (!p) continue;
    const dates = sheets[sh].sampleDates || [];
    p.samples.forEach((v, si) => {
      const dl = dates[si] || `${si+1}`;
      const numVal = (!isText && typeof v === "number") ? v : null;
      const ok = v === null ? null : (isText ? isTextOk(v) : isInRange(v, min, max));
      pts.push({ label: `${dl}(${sh})`, shortLabel: dl, value: numVal, rawVal: v, month: sh, date: dl, ok });
    });
  }
  return pts;
}

function buildMonthlyAvgs(paramName, sheets, sheetOrder) {
  return sheetOrder.map(sh => {
    const p = sheets[sh].parameters.find(x => x.name === paramName);
    const nums = (p?.samples || []).filter(v => typeof v === "number");
    return { label: sh, avg: nums.length ? +avgNums(nums).toFixed(4) : null };
  });
}

const PARAM_COLORS = [
  "#00d4ff","#ff6b35","#7fff7f","#ffd700","#ff69b4",
  "#87ceeb","#dda0dd","#98ff98","#f4a460","#40e0d0",
  "#ff9966","#c8a2c8","#b5e853","#ffeaa7","#fd79a8",
];

function getDecimals(name, sample) {
  if (/pH/i.test(name)) return 2;
  if (/chlorine|iron|COC|cycle/i.test(name)) return 2;
  if (/bacteria/i.test(name)) return 0;
  return (typeof sample === "number" && sample < 5) ? 2 : 0;
}

// ─── PURE SVG LINE CHART ──────────────────────────────────────────────────────

function SvgLineChart({ flatData, color, min, max, decimals, limitRaw, width = 560, height = 160, compact = false }) {
  const [tooltip, setTooltip] = useState(null);
  const PAD = { top: 10, right: 12, bottom: compact ? 28 : 36, left: 42 };
  const W = width - PAD.left - PAD.right;
  const H = height - PAD.top - PAD.bottom;

  const numPts = flatData.filter(d => d.value != null);
  if (!numPts.length) return <div style={{ color:"#555", fontSize:12, padding:20, textAlign:"center" }}>No numeric data</div>;

  const vals = numPts.map(d => d.value);
  const dataMin = Math.min(...vals);
  const dataMax = Math.max(...vals);
  const pad = Math.max((dataMax - dataMin) * 0.25, 0.5);
  const yMin = min != null ? Math.min(min, dataMin) - pad : dataMin - pad;
  const yMax = max != null ? Math.max(max, dataMax) + pad : dataMax + pad;
  const yRange = yMax - yMin || 1;

  const xStep = W / Math.max(flatData.length - 1, 1);

  const toX = i => PAD.left + i * xStep;
  const toY = v => PAD.top + H - ((v - yMin) / yRange) * H;
  const toRefY = v => PAD.top + H - ((v - yMin) / yRange) * H;

  // Build polyline path — skip nulls
  let okPath = "", badPath = "";
  flatData.forEach((d, i) => {
    if (d.value == null) return;
    const x = toX(i), y = toY(d.value);
    if (d.ok) okPath += `${okPath ? "L" : "M"}${x.toFixed(1)},${y.toFixed(1)} `;
    else badPath += `M${x.toFixed(1)},${y.toFixed(1)} `;
  });

  // Build one connected path for the line
  let linePath = "";
  flatData.forEach((d, i) => {
    if (d.value == null) return;
    const x = toX(i), y = toY(d.value);
    linePath += `${linePath ? "L" : "M"}${x.toFixed(1)},${y.toFixed(1)} `;
  });

  // Y axis ticks
  const yTicks = 4;
  const yTickVals = Array.from({ length: yTicks + 1 }, (_, i) => yMin + (yRange / yTicks) * i);

  // X axis — show every Nth label
  const xLabelStep = Math.max(1, Math.ceil(flatData.length / (compact ? 8 : 14)));

  return (
    <svg width={width} height={height} style={{ display:"block", overflow:"visible" }}
      onMouseLeave={() => setTooltip(null)}>
      {/* Grid lines */}
      {yTickVals.map((v, i) => (
        <line key={i}
          x1={PAD.left} y1={toRefY(v).toFixed(1)}
          x2={PAD.left + W} y2={toRefY(v).toFixed(1)}
          stroke="rgba(255,255,255,0.06)" strokeWidth={1} />
      ))}
      {/* Y axis labels */}
      {yTickVals.map((v, i) => (
        <text key={i} x={PAD.left - 5} y={toRefY(v) + 3}
          textAnchor="end" fontSize={9} fill="#444" fontFamily={SFMono}>
          {v.toFixed(decimals > 1 ? 1 : 0)}
        </text>
      ))}
      {/* Control limit lines */}
      {min != null && isFinite(toRefY(min)) && (
        <line x1={PAD.left} y1={toRefY(min).toFixed(1)} x2={PAD.left+W} y2={toRefY(min).toFixed(1)}
          stroke="#ff4444" strokeWidth={1} strokeDasharray="4 3" opacity={0.55} />
      )}
      {max != null && isFinite(toRefY(max)) && (
        <line x1={PAD.left} y1={toRefY(max).toFixed(1)} x2={PAD.left+W} y2={toRefY(max).toFixed(1)}
          stroke="#ff4444" strokeWidth={1} strokeDasharray="4 3" opacity={0.55} />
      )}
      {/* Main line */}
      {linePath && <path d={linePath} fill="none" stroke={color} strokeWidth={1.5} opacity={0.4} />}
      {/* Dots + X labels */}
      {flatData.map((d, i) => {
        if (d.value == null) return null;
        const x = toX(i), y = toY(d.value);
        const showLabel = !compact && i % xLabelStep === 0;
        return (
          <g key={i}>
            {showLabel && (
              <text x={x} y={PAD.top + H + 14}
                textAnchor="middle" fontSize={compact ? 7 : 7.5} fill="#444" fontFamily={SFMono}
                transform={`rotate(-30, ${x}, ${PAD.top + H + 14})`}>
                {d.label}
              </text>
            )}
            <circle cx={x} cy={y} r={4}
              fill={d.ok === null ? "#555" : d.ok ? color : "#ff3333"}
              stroke={d.ok === false ? "#ff0000" : "none"}
              strokeWidth={d.ok === false ? 1.5 : 0}
              style={{ cursor:"pointer" }}
              onMouseEnter={(e) => setTooltip({ i, x, y, d })}
            />
          </g>
        );
      })}
      {/* Tooltip */}
      {tooltip && (() => {
        const { x, y, d } = tooltip;
        const tw = 130, th = 56;
        const tx = Math.min(x + 10, width - tw - 4);
        const ty = Math.max(y - th - 6, PAD.top);
        return (
          <g>
            <rect x={tx} y={ty} width={tw} height={th} rx={6}
              fill="#0d1117" stroke={d.ok ? color : "#ff4444"} strokeWidth={1} opacity={0.97} />
            <text x={tx+8} y={ty+14} fontSize={9} fill="#888" fontFamily={SF}>{d.label}</text>
            <text x={tx+8} y={ty+30} fontSize={14} fontWeight="bold"
              fill={d.ok ? color : "#ff5252"} fontFamily={SFMono}>
              {d.value.toFixed(decimals)}
            </text>
            <text x={tx+8} y={ty+46} fontSize={9} fill="#555" fontFamily={SF}>Limit: {limitRaw || "—"}</text>
          </g>
        );
      })()}
    </svg>
  );
}

// ─── CHART POPUP MODAL ────────────────────────────────────────────────────────

function ChartModal({ param, color, decimals, flatData, onClose }) {
  useEffect(() => {
    const onKey = e => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  const outCount = flatData.filter(d => d.value != null && !d.ok).length;

  return (
    <div style={{
      position:"fixed", inset:0, zIndex:1000,
      background:"rgba(0,0,0,0.75)", backdropFilter:"blur(4px)",
      display:"flex", alignItems:"center", justifyContent:"center",
      padding:24,
    }} onClick={onClose}>
      <div style={{
        background:"#111418", border:"1px solid rgba(255,255,255,0.1)",
        borderRadius:16, padding:28, width:"min(900px,95vw)", maxHeight:"85vh",
        overflowY:"auto",
      }} onClick={e => e.stopPropagation()}>
        {/* Header */}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
          <div>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ color:"#fff", fontWeight:700, fontSize:16, fontFamily:SF }}>{param.name}</span>
              {outCount > 0 && (
                <span style={{ fontSize:11, color:"#ff5252", background:"rgba(255,82,82,0.12)", border:"1px solid rgba(255,82,82,0.3)", borderRadius:999, padding:"2px 9px" }}>
                  {outCount} out of range
                </span>
              )}
            </div>
            <div style={{ color:"#555", fontSize:11, marginTop:4, fontFamily:SF }}>
              {flatData.length} individual readings · Limit: {param.limitRaw || "—"}
            </div>
          </div>
          <button onClick={onClose} style={{
            background:"rgba(255,255,255,0.07)", border:"1px solid rgba(255,255,255,0.1)",
            borderRadius:8, color:"#888", cursor:"pointer", padding:"6px 12px",
            fontSize:13, fontFamily:SF,
          }}>✕ Close</button>
        </div>

        {/* Legend */}
        <div style={{ display:"flex", gap:16, marginBottom:16, fontSize:11, color:"#666", fontFamily:SF }}>
          <span><span style={{ display:"inline-block", width:8, height:8, borderRadius:"50%", background:color, marginRight:5 }} />Within limit</span>
          <span><span style={{ display:"inline-block", width:8, height:8, borderRadius:"50%", background:"#ff4444", marginRight:5 }} />Out of range</span>
          <span><span style={{ display:"inline-block", width:18, height:2, background:"#ff4444", opacity:0.5, marginRight:5, verticalAlign:"middle" }} />Control limit</span>
        </div>

        {/* Full-width chart */}
        <div style={{ width:"100%", overflowX:"auto" }}>
          <SvgLineChart
            flatData={flatData}
            color={color}
            min={param.min} max={param.max}
            decimals={decimals}
            limitRaw={param.limitRaw}
            width={Math.max(flatData.length * 22 + 80, 800)}
            height={260}
          />
        </div>

        {/* Data table */}
        <div style={{ marginTop:20, overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11, fontFamily:SFMono }}>
            <thead>
              <tr style={{ borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
                <th style={{ textAlign:"left", padding:"7px 10px", color:"#555", fontWeight:500 }}>Date</th>
                <th style={{ textAlign:"left", padding:"7px 10px", color:"#555", fontWeight:500 }}>Month</th>
                <th style={{ textAlign:"right", padding:"7px 10px", color:"#555", fontWeight:500 }}>Value</th>
                <th style={{ textAlign:"center", padding:"7px 10px", color:"#555", fontWeight:500 }}>Status</th>
              </tr>
            </thead>
            <tbody>
              {flatData.map((d, i) => (
                <tr key={i} style={{ borderBottom:"1px solid rgba(255,255,255,0.04)", background:!d.ok&&d.value!=null?"rgba(255,82,82,0.05)":"transparent" }}>
                  <td style={{ padding:"6px 10px", color:"#888" }}>{d.date}</td>
                  <td style={{ padding:"6px 10px", color:"#888" }}>{d.month}</td>
                  <td style={{ padding:"6px 10px", textAlign:"right", color:d.ok?(color||"#ccc"):"#ff5252", fontWeight:600 }}>
                    {d.value != null ? d.value.toFixed(decimals) : "—"}
                  </td>
                  <td style={{ padding:"6px 10px", textAlign:"center" }}>
                    {d.value != null ? (d.ok
                      ? <span style={{ color:"#00e676", fontSize:10 }}>✓ OK</span>
                      : <span style={{ color:"#ff5252", fontSize:10 }}>⚠ OUT</span>
                    ) : <span style={{ color:"#444" }}>—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── PARAM CHART CARD ─────────────────────────────────────────────────────────

function ParamChartCard({ param, color, decimals, flatData }) {
  const [modalOpen, setModalOpen] = useState(false);
  const containerRef = useRef(null);
  const [cardWidth, setCardWidth] = useState(500);

  useEffect(() => {
    if (!containerRef.current) return;
    const ro = new ResizeObserver(entries => {
      const w = entries[0]?.contentRect?.width;
      if (w && w > 50) setCardWidth(Math.floor(w) - 32);
    });
    ro.observe(containerRef.current);
    return () => ro.disconnect();
  }, []);

  if (!flatData || !flatData.length) return null;
  const nums = flatData.filter(d => d.value != null);
  if (!nums.length) return null;

  const outCount = flatData.filter(d => d.value != null && !d.ok).length;

  return (
    <>
      <div ref={containerRef}
        onClick={() => setModalOpen(true)}
        style={{
          background:"rgba(255,255,255,0.025)", border:`1px solid ${outCount > 0 ? "rgba(255,82,82,0.25)" : "rgba(255,255,255,0.06)"}`,
          borderRadius:12, padding:16, cursor:"pointer",
          transition:"transform 0.15s, box-shadow 0.15s",
          fontFamily:SF,
        }}
        onMouseEnter={e => { e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow=`0 8px 24px rgba(0,0,0,0.3), 0 0 16px ${color}18`; }}
        onMouseLeave={e => { e.currentTarget.style.transform="translateY(0)"; e.currentTarget.style.boxShadow="none"; }}>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ color:"#ddd", fontWeight:600, fontSize:12 }}>{param.name}</span>
            {outCount > 0 && (
              <span style={{ fontSize:10, color:"#ff5252", background:"rgba(255,82,82,0.12)", border:"1px solid rgba(255,82,82,0.3)", borderRadius:999, padding:"1px 7px" }}>
                {outCount} out of range
              </span>
            )}
          </div>
          <div style={{ display:"flex", gap:12, alignItems:"center", fontSize:10, color:"#555" }}>
            <span>{nums.length} readings</span>
            <span style={{ color:"#333", fontSize:10 }}>🔍 click to expand</span>
          </div>
        </div>

        <SvgLineChart
          flatData={flatData}
          color={color}
          min={param.min} max={param.max}
          decimals={decimals}
          limitRaw={param.limitRaw}
          width={cardWidth}
          height={140}
          compact
        />

        <div style={{ fontSize:10, color:"#444", marginTop:6, fontFamily:SFMono }}>
          Limit: {param.limitRaw || "—"}
        </div>
      </div>

      {modalOpen && (
        <ChartModal
          param={param}
          color={color}
          decimals={decimals}
          flatData={flatData}
          onClose={() => setModalOpen(false)}
        />
      )}
    </>
  );
}

// ─── STATUS BADGE ─────────────────────────────────────────────────────────────

function StatusBadge({ ok }) {
  return (
    <span style={{
      display:"inline-flex", alignItems:"center", gap:3,
      padding:"2px 8px", borderRadius:999, fontSize:10, fontWeight:600, fontFamily:SF,
      background: ok ? "rgba(0,255,120,0.12)" : "rgba(255,80,80,0.15)",
      color:       ok ? "#00e676" : "#ff5252",
      border:      `1px solid ${ok ? "rgba(0,230,118,0.3)" : "rgba(255,82,82,0.3)"}`,
    }}>{ok ? "✓ OK" : "⚠ OUT"}</span>
  );
}

function RemarksBar({ allOk }) {
  return (
    <div style={{
      display:"flex", alignItems:"flex-start", gap:12, padding:"13px 18px",
      borderRadius:10, marginBottom:22, fontFamily:SF,
      background: allOk ? "rgba(0,230,118,0.07)" : "rgba(255,82,82,0.07)",
      border: `1px solid ${allOk ? "rgba(0,230,118,0.22)" : "rgba(255,82,82,0.22)"}`,
    }}>
      <span style={{ fontSize:15, marginTop:1 }}>{allOk ? "✅" : "⚠️"}</span>
      <span style={{ fontSize:13, color: allOk ? "#a8f0c6" : "#ffaaaa", lineHeight:1.6 }}>
        {allOk
          ? "The current water sample shown that all the parameter is still under our acceptable control limit."
          : "One or more parameters have readings outside the acceptable control limit. Please review the highlighted values below."}
      </span>
    </div>
  );
}

function KPICard({ param, color, decimals, monthlyAvgs, latestSamples }) {
  if (!monthlyAvgs.length) return null;
  if (param.isText) {
    const lastVal = latestSamples?.[latestSamples.length-1] ?? "—";
    const ok = isTextOk(lastVal);
    return (
      <div style={{ background:"rgba(255,255,255,0.035)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:12, padding:"15px 16px", position:"relative", overflow:"hidden", fontFamily:SF }}
        onMouseEnter={e => { e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow=`0 8px 28px rgba(0,0,0,0.3)`; }}
        onMouseLeave={e => { e.currentTarget.style.transform=""; e.currentTarget.style.boxShadow=""; }}>
        <div style={{ position:"absolute", top:0, left:0, right:0, height:2, background:color, opacity:0.65 }} />
        <div style={{ fontSize:10, color:"#666", textTransform:"uppercase", marginBottom:3 }}>Visual</div>
        <div style={{ fontSize:11, color:"#aaa", marginBottom:8 }}>{param.name}</div>
        <div style={{ fontSize:20, fontWeight:600, color:ok?"#00e676":"#ff5252", marginBottom:8 }}>{String(lastVal).toUpperCase()}</div>
        <StatusBadge ok={ok} />
        <div style={{ fontSize:10, color:"#444", marginTop:5 }}>Limit: Clear</div>
      </div>
    );
  }
  const latest = monthlyAvgs[monthlyAvgs.length-1]?.avg;
  const first  = monthlyAvgs[0]?.avg;
  const trend  = (latest != null && first != null) ? latest - first : 0;
  const ok     = isInRange(latest, param.min, param.max);
  return (
    <div style={{ background:"rgba(255,255,255,0.035)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:12, padding:"15px 16px", position:"relative", overflow:"hidden", fontFamily:SF, transition:"transform 0.2s,box-shadow 0.2s", cursor:"default" }}
      onMouseEnter={e => { e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow=`0 8px 28px rgba(0,0,0,0.3),0 0 18px ${color}22`; }}
      onMouseLeave={e => { e.currentTarget.style.transform=""; e.currentTarget.style.boxShadow=""; }}>
      <div style={{ position:"absolute", top:0, left:0, right:0, height:2, background:color, opacity:0.65 }} />
      <div style={{ fontSize:10, color:"#666", textTransform:"uppercase", marginBottom:3 }}>
        {param.name.replace(/, ppm|,ppm/gi,"").replace(/as CaCO3|as CI|as Si|as Fe/gi,"").trim()}
      </div>
      <div style={{ display:"flex", alignItems:"baseline", gap:5, margin:"6px 0" }}>
        <span style={{ fontSize:26, fontWeight:700, color, fontFamily:SFMono, letterSpacing:"-0.5px" }}>
          {latest != null ? latest.toFixed(decimals) : "—"}
        </span>
        <span style={{ fontSize:11, color:"#555" }}>{param.limitRaw?.includes("ppm") ? "ppm" : ""}</span>
      </div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        <StatusBadge ok={ok} />
        <span style={{ fontSize:11, fontFamily:SFMono, color: trend>0?"#ff6b6b":trend<0?"#6bffb8":"#888" }}>
          {trend > 0 ? "▲" : trend < 0 ? "▼" : "—"} {Math.abs(trend).toFixed(decimals)}
        </span>
      </div>
      <div style={{ fontSize:10, color:"#444", marginTop:5, fontFamily:SFMono }}>Limit: {param.limitRaw || "—"}</div>
    </div>
  );
}

// ─── UPLOAD SCREEN ────────────────────────────────────────────────────────────

function UploadScreen({ onLoad, existing }) {
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState("");
  const inputRef = useRef();

  const process = useCallback(async (file) => {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) { setError("Please upload an .xlsx or .xls file."); return; }
    setLoading(true); setError("");
    try {
      const buf  = await file.arrayBuffer();
      const data = parseWorkbook(new Uint8Array(buf));
      if (!data.sheetOrder.length) { setError("No recognisable monthly data sheets found."); setLoading(false); return; }
      onLoad(data, file.name);
    } catch(e) { setError("Parse error: " + e.message); }
    setLoading(false);
  }, [onLoad]);

  return (
    <div style={{ minHeight:"100vh", background:"#080b0f", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:SF, padding:32 }}>
      <div style={{ fontSize:11, color:"#555", letterSpacing:"0.15em", textTransform:"uppercase", marginBottom:12 }}>Water Analysis Dashboard</div>
      <div style={{ fontSize:26, fontWeight:700, color:"#fff", marginBottom:6 }}>Upload Report File</div>
      <div style={{ fontSize:13, color:"#555", marginBottom:40 }}>Supports any Cooling Tower Water Analysis .xlsx report</div>
      <div
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={e => { e.preventDefault(); setDragging(false); process(e.dataTransfer.files[0]); }}
        onClick={() => inputRef.current?.click()}
        style={{
          width:"100%", maxWidth:480, cursor:"pointer", textAlign:"center",
          border:`2px dashed ${dragging ? "#00d4ff" : "rgba(255,255,255,0.12)"}`,
          borderRadius:16, padding:"52px 32px",
          background: dragging ? "rgba(0,212,255,0.06)" : "rgba(255,255,255,0.02)",
          transition:"all 0.2s",
        }}>
        <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={e => process(e.target.files[0])} />
        <div style={{ fontSize:36, marginBottom:14 }}>{loading ? "⏳" : "📂"}</div>
        <div style={{ color: dragging ? "#00d4ff" : "#888", fontSize:14, marginBottom:6 }}>
          {loading ? "Parsing file…" : dragging ? "Drop to load" : "Drag & drop your .xlsx file here"}
        </div>
        <div style={{ color:"#444", fontSize:12 }}>or click to browse</div>
        {error && <div style={{ color:"#ff5252", fontSize:12, marginTop:14 }}>{error}</div>}
      </div>
      {existing && (
        <div style={{ marginTop:20, fontSize:12, color:"#555", cursor:"pointer", textDecoration:"underline", textUnderlineOffset:3 }}
          onClick={() => onLoad(existing.data, existing.name)}>← Back to {existing.name}</div>
      )}
      <div style={{ marginTop:40, fontSize:11, color:"#333", textAlign:"center", maxWidth:400, lineHeight:1.8 }}>
        Expected format: multi-sheet workbook with monthly tabs.<br/>
        Each sheet should contain parameter rows with sample values and control limits.
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function App() {
  const [fileData, setFileData]           = useState(null);
  const [fileName, setFileName]           = useState("");
  const [showUpload, setShowUpload]       = useState(false);
  const [activeTab, setActiveTab]         = useState("overview");
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [mounted, setMounted]             = useState(false);

  useEffect(() => setMounted(true), []);

  const handleLoad = useCallback((data, name) => {
    setFileData(data); setFileName(name);
    setSelectedSheet(data.sheetOrder[data.sheetOrder.length-1]);
    setShowUpload(false); setActiveTab("overview");
  }, []);

  if (!fileData || showUpload)
    return <UploadScreen onLoad={handleLoad} existing={fileData ? { data:fileData, name:fileName } : null} />;

  const { sheets, sheetOrder, meta } = fileData;

  const allParamNames = [];
  const seenNames = new Set();
  for (const sh of sheetOrder)
    for (const p of sheets[sh].parameters)
      if (!seenNames.has(p.name)) { seenNames.add(p.name); allParamNames.push(p.name); }

  const paramMeta = allParamNames.map((name, idx) => {
    let min=null, max=null, limitRaw=null, isText=false;
    for (const sh of sheetOrder) {
      const p = sheets[sh].parameters.find(x => x.name === name);
      if (p) { isText = p.isText||false; if (p.limitRaw) { limitRaw=p.limitRaw; const l=parseControlLimit(p.limitRaw); min=l.min; max=l.max; } }
    }
    const allSamples = sheetOrder.flatMap(sh => sheets[sh].parameters.find(x=>x.name===name)?.samples||[]);
    const nums = allSamples.filter(v => typeof v==="number");
    const dec  = isText ? 0 : getDecimals(name, nums[0]);
    const flatData    = buildFlatSeries(name, isText, min, max, sheets, sheetOrder);
    const monthlyAvgs = buildMonthlyAvgs(name, sheets, sheetOrder);
    return { name, min, max, limitRaw, isText, color:PARAM_COLORS[idx%PARAM_COLORS.length], decimals:dec, flatData, monthlyAvgs };
  });

  const currentSheetData = sheets[selectedSheet];

  const latestSheetAllOk = paramMeta.every(p => {
    const found = currentSheetData?.parameters.find(x=>x.name===p.name);
    if (!found) return true;
    if (p.isText) return found.samples.every(v => isTextOk(v));
    return found.samples.filter(v=>typeof v==="number").every(v=>isInRange(v,p.min,p.max));
  });

  const complMatrix = paramMeta.map(p => ({
    ...p,
    bySheet: sheetOrder.map(sh => {
      const found = sheets[sh].parameters.find(x=>x.name===p.name);
      if (!found) return { sh, ok:null };
      if (p.isText) return { sh, ok:found.samples.every(v=>v!=null&&isTextOk(v)), val:found.samples[0] };
      const nums = found.samples.filter(v=>typeof v==="number");
      if (!nums.length) return { sh, ok:null };
      return { sh, ok:isInRange(avgNums(nums),p.min,p.max), val:avgNums(nums) };
    }),
  }));

  return (
    <div style={{ minHeight:"100vh", background:"#080b0f", color:"#e0e0e0", fontFamily:SF, opacity:mounted?1:0, transition:"opacity 0.5s" }}>

      {/* Header */}
      <div style={{ borderBottom:"1px solid rgba(255,255,255,0.07)", padding:"16px 28px", display:"flex", alignItems:"center", justifyContent:"space-between", background:"rgba(0,0,0,0.35)", position:"sticky", top:0, zIndex:100 }}>
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ width:34, height:34, borderRadius:8, background:"linear-gradient(135deg,#00d4ff,#0066cc)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:17 }}>🧪</div>
          <div>
            <div style={{ fontWeight:700, fontSize:14, color:"#fff" }}>Cooling Tower Water Analysis</div>
            <div style={{ fontSize:11, color:"#444", maxWidth:420, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
              {meta.systemLabel||fileName} · {sheetOrder.length} month{sheetOrder.length>1?"s":""}
            </div>
          </div>
        </div>
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>
          {[{id:"overview",l:"Overview"},{id:"trends",l:"Trends"},{id:"monthly",l:"Monthly Detail"}].map(t => (
            <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
              padding:"7px 16px", borderRadius:7, border:"none", cursor:"pointer",
              fontSize:13, fontFamily:SF, fontWeight:500,
              background: activeTab===t.id ? "rgba(0,212,255,0.14)" : "transparent",
              color:       activeTab===t.id ? "#00d4ff" : "#555",
              borderBottom: activeTab===t.id ? "2px solid #00d4ff" : "2px solid transparent",
              transition:"all 0.15s",
            }}>{t.l}</button>
          ))}
          <div style={{ width:1, height:20, background:"rgba(255,255,255,0.1)", margin:"0 6px" }} />
          <button onClick={() => setShowUpload(true)} style={{
            padding:"7px 16px", borderRadius:7, border:"1px solid rgba(0,212,255,0.3)",
            background:"rgba(0,212,255,0.08)", color:"#00d4ff", cursor:"pointer",
            fontSize:13, fontFamily:SF, fontWeight:500,
          }}>⬆ Load File</button>
        </div>
      </div>

      <div style={{ padding:"24px 28px", maxWidth:1440, margin:"0 auto" }}>

        {/* ══ OVERVIEW ══ */}
        {activeTab === "overview" && (
          <div>
            <RemarksBar allOk={latestSheetAllOk} />
            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:24 }}>
              {[
                { label:"Months Loaded",   value:`${sheetOrder.length}`, sub:sheetOrder.join(" · "), color:"#00d4ff" },
                { label:"Parameters",      value:`${paramMeta.length}`,  sub:"Physical & Chemical",  color:"#7fff7f" },
                { label:"Full Compliance", value:`${complMatrix.filter(p=>p.bySheet.every(s=>s.ok!==false)).length}/${paramMeta.length}`, sub:"params at 100% pass", color:"#ffd700" },
                { label:"Latest Sheet",    value:selectedSheet,          sub:currentSheetData?.reportMonth||"", color:"#dda0dd" },
              ].map((s,i) => (
                <div key={i} style={{ background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.07)", borderRadius:12, padding:18, position:"relative", overflow:"hidden" }}>
                  <div style={{ position:"absolute", bottom:-8, right:-8, width:52, height:52, borderRadius:"50%", background:s.color, opacity:0.07 }} />
                  <div style={{ fontSize:11, color:"#555", textTransform:"uppercase", marginBottom:7 }}>{s.label}</div>
                  <div style={{ fontSize:22, fontWeight:700, color:s.color, fontFamily:SFMono }}>{s.value}</div>
                  <div style={{ fontSize:11, color:"#555", marginTop:4, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{s.sub}</div>
                </div>
              ))}
            </div>

            <div style={{ fontSize:11, color:"#555", textTransform:"uppercase", marginBottom:12 }}>
              {selectedSheet} — Latest Parameter Readings vs Control Limits
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:10, marginBottom:24 }}>
              {paramMeta.map(p => {
                const sp = currentSheetData?.parameters.find(x=>x.name===p.name);
                return <KPICard key={p.name} param={p} color={p.color} decimals={p.decimals} monthlyAvgs={p.monthlyAvgs} latestSamples={sp?.samples||[]} />;
              })}
            </div>

            <div style={{ fontSize:11, color:"#555", textTransform:"uppercase", marginBottom:12 }}>Compliance Matrix — Pass / Fail by Month</div>
            <div style={{ background:"rgba(255,255,255,0.025)", border:"1px solid rgba(255,255,255,0.06)", borderRadius:12, overflow:"auto" }}>
              <div style={{ display:"grid", gridTemplateColumns:`200px repeat(${sheetOrder.length},1fr)`, borderBottom:"1px solid rgba(255,255,255,0.07)", background:"rgba(255,255,255,0.03)", minWidth:600 }}>
                <div style={{ padding:"10px 14px", fontSize:11, color:"#555" }}>PARAMETER</div>
                {sheetOrder.map(sh => <div key={sh} style={{ padding:"10px 4px", fontSize:11, color:"#555", textAlign:"center" }}>{sh}</div>)}
              </div>
              {complMatrix.map((p,pi) => (
                <div key={p.name} style={{ display:"grid", gridTemplateColumns:`200px repeat(${sheetOrder.length},1fr)`, borderBottom:pi<complMatrix.length-1?"1px solid rgba(255,255,255,0.04)":"none", background:pi%2?"rgba(255,255,255,0.01)":"transparent", minWidth:600 }}>
                  <div style={{ padding:"9px 14px", fontSize:12, color:"#aaa", display:"flex", alignItems:"center", gap:7 }}>
                    <span style={{ display:"inline-block", width:7, height:7, borderRadius:"50%", background:p.color, flexShrink:0 }} />
                    <span style={{ overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:155 }} title={p.name}>
                      {p.name.replace(/, ppm|,ppm/gi,"").replace(/as CaCO3|as CI|as Si|as Fe/gi,"").trim()}
                    </span>
                  </div>
                  {p.bySheet.map(({sh,ok}) => (
                    <div key={sh} style={{ display:"flex", alignItems:"center", justifyContent:"center", padding:"9px 4px" }}>
                      <div style={{
                        width:22, height:22, borderRadius:4,
                        background: ok===null?"rgba(255,255,255,0.04)":ok?"rgba(0,230,118,0.15)":"rgba(255,82,82,0.2)",
                        border:`1px solid ${ok===null?"rgba(255,255,255,0.08)":ok?"rgba(0,230,118,0.3)":"rgba(255,82,82,0.4)"}`,
                        display:"flex", alignItems:"center", justifyContent:"center",
                        fontSize:10, color:ok===null?"#333":ok?"#00e676":"#ff5252",
                      }}>{ok===null?"·":ok?"✓":"✗"}</div>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══ TRENDS ══ */}
        {activeTab === "trends" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18 }}>
              <div style={{ fontSize:11, color:"#555", textTransform:"uppercase" }}>
                All Individual Sample Readings · {sheetOrder.length} months · Click any chart to expand
              </div>
              <div style={{ display:"flex", gap:16, fontSize:11, color:"#666" }}>
                <span><span style={{ display:"inline-block", width:8, height:8, borderRadius:"50%", background:"#00d4ff", marginRight:5 }} />Within limit</span>
                <span><span style={{ display:"inline-block", width:8, height:8, borderRadius:"50%", background:"#ff4444", marginRight:5 }} />Out of range</span>
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
              {paramMeta.filter(p=>!p.isText).map(p => (
                <ParamChartCard key={p.name} param={p} color={p.color} decimals={p.decimals} flatData={p.flatData} />
              ))}
            </div>
            {/* Text params */}
            {paramMeta.filter(p=>p.isText).map(p => (
              <div key={p.name} style={{ background:"rgba(255,255,255,0.025)", border:"1px solid rgba(255,255,255,0.06)", borderRadius:12, padding:"18px 16px", marginTop:14 }}>
                <div style={{ fontSize:13, fontWeight:600, color:"#ddd", marginBottom:12 }}>{p.name}</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
                  {p.flatData.map((d,i) => (
                    <div key={i} style={{
                      padding:"5px 12px", borderRadius:6, fontSize:11, fontWeight:500,
                      background:   isTextOk(d.rawVal)?"rgba(0,230,118,0.12)":"rgba(255,82,82,0.15)",
                      border:`1px solid ${isTextOk(d.rawVal)?"rgba(0,230,118,0.3)":"rgba(255,82,82,0.4)"}`,
                      color:        isTextOk(d.rawVal)?"#00e676":"#ff5252",
                    }}>{d.label}: {String(d.rawVal)}</div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* ══ MONTHLY DETAIL ══ */}
        {activeTab === "monthly" && (
          <div>
            <RemarksBar allOk={latestSheetAllOk} />
            <div style={{ display:"flex", gap:7, marginBottom:20, flexWrap:"wrap" }}>
              {sheetOrder.map(sh => (
                <button key={sh} onClick={()=>setSelectedSheet(sh)} style={{
                  padding:"7px 14px", borderRadius:7,
                  border:`1px solid ${selectedSheet===sh?"#00d4ff":"rgba(255,255,255,0.1)"}`,
                  background:selectedSheet===sh?"rgba(0,212,255,0.12)":"transparent",
                  color:selectedSheet===sh?"#00d4ff":"#666",
                  cursor:"pointer", fontSize:12, fontFamily:SF, fontWeight:500,
                }}>{sh}</button>
              ))}
            </div>

            {currentSheetData && (() => {
              const numSamples = currentSheetData.parameters[0]?.samples.length || 0;
              const colTemplate = `70px 220px 90px repeat(${numSamples}, 90px) 120px`;
              const minW = 70 + 220 + 90 + (numSamples * 90) + 120;
              return (
                <div>
                  <div style={{ fontSize:11, color:"#555", textTransform:"uppercase", marginBottom:12 }}>
                    {currentSheetData.reportMonth} — {numSamples} Sample{numSamples!==1?"s":""}
                    {currentSheetData.sampleDates?.length ? ` · ${currentSheetData.sampleDates.join(" · ")}` : ""}
                  </div>
                  <div style={{ background:"rgba(255,255,255,0.025)", border:"1px solid rgba(255,255,255,0.06)", borderRadius:12, overflow:"auto" }}>
                    {/* Header row — matches Excel column order */}
                    <div style={{ display:"grid", gridTemplateColumns:colTemplate, borderBottom:"1px solid rgba(255,255,255,0.1)", background:"rgba(255,255,255,0.04)", minWidth:minW }}>
                      <div style={{ padding:"11px 8px", fontSize:10, color:"#555", textAlign:"center", letterSpacing:"0.04em" }}>METHOD #</div>
                      <div style={{ padding:"11px 14px", fontSize:10, color:"#555", letterSpacing:"0.04em" }}>TYPE OF PARAMETER</div>
                      <div style={{ padding:"11px 6px", fontSize:10, color:"#6baaff", textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.07)", borderRight:"1px solid rgba(255,255,255,0.07)", background:"rgba(107,170,255,0.04)" }}>CITY WATER</div>
                      {Array.from({length:numSamples}).map((_,si) => (
                        <div key={si} style={{ padding:"11px 6px", fontSize:10, color:"#ccc", textAlign:"center", fontWeight:600 }}>
                          {currentSheetData.sampleDates?.[si] || `S${si+1}`}
                        </div>
                      ))}
                      <div style={{ padding:"11px 8px", fontSize:10, color:"#555", textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.07)" }}>CONTROL LIMIT</div>
                    </div>

                    {/* Data rows */}
                    {currentSheetData.parameters.map((p, pi) => {
                      const meta2 = paramMeta.find(x => x.name === p.name) || {};
                      const { min, max, decimals=2, color="#888", isText=false } = meta2;
                      const cityWaterNum = tryNum(p.cityWater);
                      const hasCW = p.cityWater != null && p.cityWater !== "" && p.cityWater !== "-";

                      return (
                        <div key={pi} style={{ display:"grid", gridTemplateColumns:colTemplate, borderBottom:pi<currentSheetData.parameters.length-1?"1px solid rgba(255,255,255,0.04)":"none", background:pi%2?"rgba(255,255,255,0.015)":"transparent", minWidth:minW }}>
                          {/* Method # */}
                          <div style={{ padding:"11px 8px", textAlign:"center", fontSize:11, color:"#444", fontFamily:SFMono }}>
                            {p.methodNum || "—"}
                          </div>
                          {/* Parameter name */}
                          <div style={{ padding:"11px 14px", display:"flex", alignItems:"center", gap:7 }}>
                            <span style={{ display:"inline-block", width:7, height:7, borderRadius:"50%", background:color, flexShrink:0 }} />
                            <span style={{ fontSize:12, color:"#ccc", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={p.name}>{p.name}</span>
                          </div>
                          {/* City Water */}
                          <div style={{ padding:"11px 6px", textAlign:"center", fontSize:12, fontWeight:500, borderLeft:"1px solid rgba(255,255,255,0.06)", borderRight:"1px solid rgba(255,255,255,0.06)", background:"rgba(107,170,255,0.03)", color: hasCW ? "#6baaff" : "#333" }}>
                            {hasCW
                              ? (isText ? p.cityWater : cityWaterNum!=null ? cityWaterNum.toFixed(decimals) : p.cityWater)
                              : "—"}
                          </div>
                          {/* Sample readings */}
                          {Array.from({length:numSamples}).map((_,si) => {
                            const v  = p.samples[si] ?? null;
                            const ok = v===null ? null : (isText ? isTextOk(v) : isInRange(v,min,max));
                            const display = v===null ? "—" : isText ? String(v) : typeof v==="number" ? v.toFixed(decimals) : String(v);
                            return (
                              <div key={si} style={{
                                padding:"11px 6px", textAlign:"center", fontSize:12, fontWeight:500,
                                color:      v===null?"#333": ok?(isText?"#00e676":"#ddd"):"#ff6b6b",
                                background: v!==null&&ok===false?"rgba(255,100,100,0.07)":"transparent",
                              }}>{display}</div>
                            );
                          })}
                          {/* Control limit */}
                          <div style={{ padding:"11px 8px", textAlign:"center", fontSize:11, color:"#555", borderLeft:"1px solid rgba(255,255,255,0.06)", fontFamily:SFMono }}>
                            {isText ? "Clear" : (p.limitRaw || "—")}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <div style={{ marginTop:10, fontSize:11, color:"#6baaff", opacity:0.6, fontFamily:SF }}>
                    <span style={{ display:"inline-block", width:8, height:8, borderRadius:2, background:"#6baaff", marginRight:6, opacity:0.6 }} />
                    City Water values shown for reference — used to assess treatment efficiency
                  </div>
                </div>
              );
            })()}
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{ borderTop:"1px solid rgba(255,255,255,0.05)", padding:"14px 28px", display:"flex", justifyContent:"space-between", alignItems:"center", fontSize:11, color:"#383838", marginTop:28, flexWrap:"wrap", gap:8, fontFamily:SF }}>
        <span>{meta.company||fileName}{meta.attention?` · Attn: ${meta.attention}`:""}{meta.cc?` · CC: ${meta.cc}`:""}</span>
        <span style={{ cursor:"pointer", color:"#444", textDecoration:"underline", textUnderlineOffset:3 }} onClick={()=>setShowUpload(true)}>Load different file</span>
      </div>
    </div>
  );
}