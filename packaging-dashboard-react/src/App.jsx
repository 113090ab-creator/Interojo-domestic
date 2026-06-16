import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

const STATUS_ORDER = ["전체", "미착수", "부족", "진행중", "완료", "초과", "요청外 포장"];

const STATUS_COLORS = {
  미착수: "#6b7280",
  부족: "#d9480f",
  진행중: "#f59f00",
  완료: "#2b8a3e",
  초과: "#5f3dc4",
  "요청外 포장": "#0c8599",
};

const REQUEST_FIELDS = [
  { key: "code", label: "제품코드", required: true },
  { key: "name", label: "제품명", required: true },
  { key: "qty", label: "요청수량", required: true },
  { key: "date", label: "요청일자", required: false },
  { key: "customer", label: "거래처/요청부서", required: false },
];

const PACKING_FIELDS = [
  { key: "code", label: "제품코드", required: true },
  { key: "name", label: "제품명", required: true },
  { key: "qty", label: "포장수량", required: false },
  { key: "packCount", label: "팩수량(대체)", required: false },
  { key: "packUnit", label: "포장단위(대체)", required: false },
  { key: "date", label: "포장일자", required: false },
  { key: "lot", label: "LOT", required: false },
];

function normalizeHeader(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[_-]/g, "");
}

function guessColumn(headers, keywords) {
  const normalizedHeaders = headers.map((header) => ({
    raw: header,
    key: normalizeHeader(header),
  }));
  for (const keyword of keywords) {
    const target = normalizeHeader(keyword);
    const found = normalizedHeaders.find((item) => item.key.includes(target));
    if (found) return found.raw;
  }
  return "";
}

function parseNumber(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value).trim();
  if (!text) return 0;
  const cleaned = text.replace(/,/g, "").replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") return 0;
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function normalizeCode(value) {
  const text = String(value ?? "").trim().toUpperCase();
  if (!text) return "";
  return text.replace(/\.0+$/, "");
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function toDateKey(value) {
  if (value === null || value === undefined || value === "") return "";

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const raw = String(value).trim();
  if (!raw) return "";

  const digits = raw.replace(/[^\d]/g, "");
  if (digits.length === 8) {
    return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  }

  const candidate = raw.replace(/[./]/g, "-");
  const date = new Date(candidate);
  if (!Number.isNaN(date.getTime())) {
    return date.toISOString().slice(0, 10);
  }

  return "";
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR");
}

function formatPercent(value) {
  return `${Number(value || 0).toLocaleString("ko-KR", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })}%`;
}

function classifyStatus(requestQty, packedQty, progressPct) {
  if (requestQty <= 0 && packedQty > 0) return "요청外 포장";
  if (requestQty > 0 && packedQty > requestQty) return "초과";
  if (requestQty > 0 && packedQty <= 0) return "미착수";
  if (requestQty > 0 && progressPct < 80) return "부족";
  if (requestQty > 0 && progressPct < 100) return "진행중";
  if (requestQty > 0 && progressPct >= 100) return "완료";
  return "미착수";
}

function getRequestDefaults(headers) {
  return {
    code: guessColumn(headers, ["제품코드", "생산코드", "품목코드", "코드", "productcode"]),
    name: guessColumn(headers, ["제품명", "품명", "itemname", "품목명"]),
    qty: guessColumn(headers, ["요청수량", "출고요청수량", "주문수량", "요청량", "수량"]),
    date: guessColumn(headers, ["요청일자", "출고요청일", "요청일", "date"]),
    customer: guessColumn(headers, ["거래처", "요청부서", "부서", "고객", "customer"]),
  };
}

function getPackingDefaults(headers) {
  return {
    code: guessColumn(headers, ["제품코드", "생산코드", "품목코드", "코드", "productcode"]),
    name: guessColumn(headers, ["제품명", "품명", "itemname", "품목명"]),
    qty: guessColumn(headers, ["포장수량", "포장실적", "실적수량", "수량"]),
    packCount: guessColumn(headers, ["팩수량", "packqty", "pack_count"]),
    packUnit: guessColumn(headers, ["포장단위", "입수", "unit", "pack_unit"]),
    date: guessColumn(headers, ["포장일자", "포장일", "일자", "date"]),
    lot: guessColumn(headers, ["lot", "로트"]),
  };
}

async function readTabularFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  let workbook;

  if (ext === "csv") {
    const text = await file.text();
    workbook = XLSX.read(text, { type: "string" });
  } else {
    const buf = await file.arrayBuffer();
    workbook = XLSX.read(buf, { type: "array" });
  }

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("첫 번째 시트를 찾지 못했습니다.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const headerRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headersFromFirstRow = Array.isArray(headerRows?.[0]) ? headerRows[0] : [];
  const headers = headersFromFirstRow
    .map((v) => String(v ?? "").trim())
    .filter((v) => !!v);

  if (headers.length === 0 && rows.length > 0) {
    return { rows, headers: Object.keys(rows[0]), sheetName: firstSheetName };
  }
  return { rows, headers, sheetName: firstSheetName };
}

function MappingPanel({ title, headers, fields, mapping, onChange, onAutoMap }) {
  if (!headers.length) return null;

  return (
    <div className="mapping-card">
      <div className="section-header">
        <h3>{title}</h3>
        <button type="button" className="btn btn-light" onClick={onAutoMap}>
          자동 매핑
        </button>
      </div>
      <div className="mapping-grid">
        {fields.map((field) => (
          <label key={field.key} className="field">
            <span>
              {field.label}
              {field.required ? <strong className="required"> *</strong> : null}
            </span>
            <select
              value={mapping[field.key] ?? ""}
              onChange={(e) => onChange(field.key, e.target.value)}
            >
              <option value="">선택 안함</option>
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>
        ))}
      </div>
    </div>
  );
}

function App() {
  const [requestMeta, setRequestMeta] = useState({
    rows: [],
    headers: [],
    fileName: "",
    sheetName: "",
  });
  const [packingMeta, setPackingMeta] = useState({
    rows: [],
    headers: [],
    fileName: "",
    sheetName: "",
  });

  const [requestMapping, setRequestMapping] = useState({});
  const [packingMapping, setPackingMapping] = useState({});
  const [loadingType, setLoadingType] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  const [searchText, setSearchText] = useState("");
  const [statusFilter, setStatusFilter] = useState("전체");
  const [sortConfig, setSortConfig] = useState({
    key: "shortageQty",
    direction: "desc",
  });

  useEffect(() => {
    if (requestMeta.headers.length > 0) {
      setRequestMapping(getRequestDefaults(requestMeta.headers));
    }
  }, [requestMeta.headers]);

  useEffect(() => {
    if (packingMeta.headers.length > 0) {
      setPackingMapping(getPackingDefaults(packingMeta.headers));
    }
  }, [packingMeta.headers]);

  const isMappingValid =
    requestMeta.rows.length > 0 &&
    packingMeta.rows.length > 0 &&
    !!requestMapping.code &&
    !!requestMapping.name &&
    !!requestMapping.qty &&
    !!packingMapping.code &&
    !!packingMapping.name &&
    (!!packingMapping.qty || (!!packingMapping.packCount && !!packingMapping.packUnit));

  const dashboard = useMemo(() => {
    if (!isMappingValid) return null;

    const requestAgg = new Map();
    const packingAgg = new Map();

    for (const row of requestMeta.rows) {
      const productCode = normalizeCode(row[requestMapping.code]);
      if (!productCode) continue;

      const requestQty = parseNumber(row[requestMapping.qty]);
      const productName = normalizeText(row[requestMapping.name]);
      const dateValue = requestMapping.date ? toDateKey(row[requestMapping.date]) : "";
      const requester = requestMapping.customer ? normalizeText(row[requestMapping.customer]) : "";

      if (!requestAgg.has(productCode)) {
        requestAgg.set(productCode, {
          productCode,
          productName,
          requestQty: 0,
          requestDates: new Set(),
          requesters: new Set(),
        });
      }

      const item = requestAgg.get(productCode);
      item.requestQty += requestQty;
      if (!item.productName && productName) item.productName = productName;
      if (dateValue) item.requestDates.add(dateValue);
      if (requester) item.requesters.add(requester);
    }

    for (const row of packingMeta.rows) {
      const productCode = normalizeCode(row[packingMapping.code]);
      if (!productCode) continue;

      const qtyFromMain = packingMapping.qty ? row[packingMapping.qty] : "";
      const mainQtyHasValue = String(qtyFromMain ?? "").trim() !== "";
      const mainQty = parseNumber(qtyFromMain);
      const altQty =
        parseNumber(row[packingMapping.packCount]) * parseNumber(row[packingMapping.packUnit]);
      const packedQty = mainQtyHasValue ? mainQty : altQty;

      const productName = normalizeText(row[packingMapping.name]);
      const dateValue = packingMapping.date ? toDateKey(row[packingMapping.date]) : "";
      const lotValue = packingMapping.lot ? normalizeText(row[packingMapping.lot]) : "";

      if (!packingAgg.has(productCode)) {
        packingAgg.set(productCode, {
          productCode,
          productName,
          packedQty: 0,
          packingDates: new Set(),
          lots: new Set(),
        });
      }

      const item = packingAgg.get(productCode);
      item.packedQty += packedQty;
      if (!item.productName && productName) item.productName = productName;
      if (dateValue) item.packingDates.add(dateValue);
      if (lotValue) item.lots.add(lotValue);
    }

    const allCodes = new Set([...requestAgg.keys(), ...packingAgg.keys()]);
    const productRows = [];

    for (const code of allCodes) {
      const req = requestAgg.get(code);
      const pack = packingAgg.get(code);

      const requestQty = req?.requestQty ?? 0;
      const packedQty = pack?.packedQty ?? 0;
      const shortageQty = Math.max(requestQty - packedQty, 0);
      const progressPct = requestQty > 0 ? (packedQty / requestQty) * 100 : 0;
      const status = classifyStatus(requestQty, packedQty, progressPct);

      productRows.push({
        productCode: code,
        productName: req?.productName || pack?.productName || "-",
        requestQty,
        packedQty,
        shortageQty,
        progressPct,
        status,
        requestDates: req?.requestDates ? [...req.requestDates].sort() : [],
        packingDates: pack?.packingDates ? [...pack.packingDates].sort() : [],
        lots: pack?.lots ? [...pack.lots].sort() : [],
      });
    }

    const totalRequestQty = productRows.reduce((sum, row) => sum + row.requestQty, 0);
    const totalPackedQty = productRows.reduce((sum, row) => sum + row.packedQty, 0);
    const totalShortageQty = productRows.reduce((sum, row) => sum + row.shortageQty, 0);
    const packedForRequest = productRows.reduce(
      (sum, row) => sum + (row.requestQty > 0 ? row.packedQty : 0),
      0
    );

    const kpis = {
      totalRequestQty,
      totalPackedQty,
      totalShortageQty,
      overallProgressPct: totalRequestQty > 0 ? (packedForRequest / totalRequestQty) * 100 : 0,
      lackItemCount: productRows.filter((row) => row.status === "부족").length,
      excessItemCount: productRows.filter((row) => row.status === "초과").length,
    };

    const priorityRows = productRows
      .filter((row) => row.shortageQty > 0)
      .sort((a, b) => b.shortageQty - a.shortageQty);

    const excessRows = productRows
      .filter((row) => row.status === "초과")
      .sort((a, b) => b.packedQty - b.requestQty - (a.packedQty - a.requestQty));

    const outsideRows = productRows
      .filter((row) => row.status === "요청外 포장")
      .sort((a, b) => b.packedQty - a.packedQty);

    const trendMap = new Map();
    if (packingMapping.date) {
      for (const row of packingMeta.rows) {
        const dateKey = toDateKey(row[packingMapping.date]);
        if (!dateKey) continue;

        const qtyFromMain = packingMapping.qty ? row[packingMapping.qty] : "";
        const mainQtyHasValue = String(qtyFromMain ?? "").trim() !== "";
        const mainQty = parseNumber(qtyFromMain);
        const altQty =
          parseNumber(row[packingMapping.packCount]) * parseNumber(row[packingMapping.packUnit]);
        const packedQty = mainQtyHasValue ? mainQty : altQty;

        trendMap.set(dateKey, (trendMap.get(dateKey) ?? 0) + packedQty);
      }
    }

    const trendData = [...trendMap.entries()]
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([date, qty]) => ({ date, qty }));

    const riskChartData = priorityRows.slice(0, 20).map((row) => ({
      label: `${row.productCode} | ${row.productName}`,
      요청수량: row.requestQty,
      포장완료수량: row.packedQty,
    }));

    return {
      productRows,
      priorityRows,
      excessRows,
      outsideRows,
      trendData,
      riskChartData,
      kpis,
    };
  }, [
    isMappingValid,
    packingMapping,
    packingMeta.rows,
    requestMapping,
    requestMeta.rows,
  ]);

  const visibleRows = useMemo(() => {
    if (!dashboard) return [];

    const keyword = searchText.trim().toLowerCase();
    let rows = dashboard.productRows;

    if (statusFilter !== "전체") {
      rows = rows.filter((row) => row.status === statusFilter);
    }

    if (keyword) {
      rows = rows.filter(
        (row) =>
          row.productCode.toLowerCase().includes(keyword) ||
          row.productName.toLowerCase().includes(keyword)
      );
    }

    const sorted = [...rows].sort((a, b) => {
      const { key, direction } = sortConfig;
      const order = direction === "asc" ? 1 : -1;

      const aVal = a[key];
      const bVal = b[key];

      if (typeof aVal === "number" && typeof bVal === "number") {
        return (aVal - bVal) * order;
      }
      return String(aVal).localeCompare(String(bVal), "ko-KR") * order;
    });

    return sorted;
  }, [dashboard, searchText, sortConfig, statusFilter]);

  const handleSort = (key) => {
    setSortConfig((prev) => {
      if (prev.key === key) {
        return { key, direction: prev.direction === "asc" ? "desc" : "asc" };
      }
      return { key, direction: "desc" };
    });
  };

  const handleUpload = async (type, file) => {
    if (!file) return;
    setErrorMessage("");
    setLoadingType(type);
    try {
      const parsed = await readTabularFile(file);
      const nextState = {
        rows: parsed.rows,
        headers: parsed.headers,
        fileName: file.name,
        sheetName: parsed.sheetName,
      };
      if (type === "request") setRequestMeta(nextState);
      if (type === "packing") setPackingMeta(nextState);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "파일 읽기에 실패했습니다.");
    } finally {
      setLoadingType("");
    }
  };

  return (
    <div className="app">
      <header className="hero">
        <h1>국내제품 포장현황 대시보드</h1>
        <p>제품코드 기준으로 요청물량 대비 포장진도율, 부족/초과/요청外를 통합 관리합니다.</p>
      </header>

      <section className="card upload-card">
        <h2>1) 데이터 업로드</h2>
        <div className="upload-grid">
          <label className="upload-box">
            <span className="upload-title">국내 요청물량 파일</span>
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              onChange={(e) => handleUpload("request", e.target.files?.[0])}
            />
            {requestMeta.fileName ? (
              <small>
                {requestMeta.fileName} / {requestMeta.sheetName} / {formatNumber(requestMeta.rows.length)}
                건
              </small>
            ) : (
              <small>CSV 또는 Excel 업로드</small>
            )}
          </label>

          <label className="upload-box">
            <span className="upload-title">포장실적 파일</span>
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              onChange={(e) => handleUpload("packing", e.target.files?.[0])}
            />
            {packingMeta.fileName ? (
              <small>
                {packingMeta.fileName} / {packingMeta.sheetName} / {formatNumber(packingMeta.rows.length)}
                건
              </small>
            ) : (
              <small>CSV 또는 Excel 업로드</small>
            )}
          </label>
        </div>

        {loadingType ? <p className="notice">파일을 읽는 중입니다: {loadingType}</p> : null}
        {errorMessage ? <p className="error">{errorMessage}</p> : null}
      </section>

      <section className="card">
        <h2>2) 컬럼 매핑</h2>
        <p className="muted">
          필수 컬럼만 맞추면 바로 계산됩니다. 포장수량 컬럼이 없으면 팩수량 × 포장단위로 자동 계산합니다.
        </p>

        <MappingPanel
          title="국내 요청물량 매핑"
          headers={requestMeta.headers}
          fields={REQUEST_FIELDS}
          mapping={requestMapping}
          onChange={(key, value) => setRequestMapping((prev) => ({ ...prev, [key]: value }))}
          onAutoMap={() => setRequestMapping(getRequestDefaults(requestMeta.headers))}
        />

        <MappingPanel
          title="포장실적 매핑"
          headers={packingMeta.headers}
          fields={PACKING_FIELDS}
          mapping={packingMapping}
          onChange={(key, value) => setPackingMapping((prev) => ({ ...prev, [key]: value }))}
          onAutoMap={() => setPackingMapping(getPackingDefaults(packingMeta.headers))}
        />

        {!isMappingValid ? (
          <p className="warning">
            대시보드 계산 조건: 요청(제품코드/제품명/요청수량) + 포장(제품코드/제품명/포장수량 또는 팩수량·포장단위)
          </p>
        ) : null}
      </section>

      {dashboard ? (
        <>
          <section className="kpi-grid">
            <div className="kpi-card">
              <span>총 요청수량</span>
              <strong>{formatNumber(dashboard.kpis.totalRequestQty)}</strong>
            </div>
            <div className="kpi-card">
              <span>총 포장완료수량</span>
              <strong>{formatNumber(dashboard.kpis.totalPackedQty)}</strong>
            </div>
            <div className="kpi-card">
              <span>총 미포장수량</span>
              <strong>{formatNumber(dashboard.kpis.totalShortageQty)}</strong>
            </div>
            <div className="kpi-card">
              <span>전체 포장진행률</span>
              <strong>{formatPercent(dashboard.kpis.overallProgressPct)}</strong>
            </div>
            <div className="kpi-card">
              <span>부족 품목 수</span>
              <strong>{formatNumber(dashboard.kpis.lackItemCount)}</strong>
            </div>
            <div className="kpi-card">
              <span>초과 포장 품목 수</span>
              <strong>{formatNumber(dashboard.kpis.excessItemCount)}</strong>
            </div>
          </section>

          <section className="card">
            <h2>3) 부족 우선순위 TOP 20 (요청 vs 포장)</h2>
            <div className="chart-wrap">
              <ResponsiveContainer width="100%" height={460}>
                <BarChart
                  data={dashboard.riskChartData}
                  layout="vertical"
                  margin={{ top: 8, right: 16, left: 8, bottom: 8 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis type="number" />
                  <YAxis
                    type="category"
                    dataKey="label"
                    width={320}
                    tick={{ fontSize: 12 }}
                  />
                  <Tooltip
                    formatter={(value) => formatNumber(value)}
                    contentStyle={{ borderRadius: 12, border: "1px solid #cbd5e1" }}
                  />
                  <Legend />
                  <Bar dataKey="요청수량" fill="#f97316" />
                  <Bar dataKey="포장완료수량" fill="#0ea5e9" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section className="card">
            <div className="section-header">
              <h2>4) 제품별 통합 현황</h2>
              <div className="toolbar">
                <input
                  type="text"
                  placeholder="제품코드/제품명 검색"
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                />
                <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                  {STATUS_ORDER.map((status) => (
                    <option key={status} value={status}>
                      {status}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th onClick={() => handleSort("productCode")}>제품코드</th>
                    <th onClick={() => handleSort("productName")}>제품명</th>
                    <th className="num" onClick={() => handleSort("requestQty")}>
                      요청수량
                    </th>
                    <th className="num" onClick={() => handleSort("packedQty")}>
                      포장완료수량
                    </th>
                    <th className="num" onClick={() => handleSort("shortageQty")}>
                      부족수량
                    </th>
                    <th className="num" onClick={() => handleSort("progressPct")}>
                      포장진행률
                    </th>
                    <th onClick={() => handleSort("status")}>상태</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((row) => (
                    <tr key={row.productCode}>
                      <td>{row.productCode}</td>
                      <td>{row.productName}</td>
                      <td className="num">{formatNumber(row.requestQty)}</td>
                      <td className="num">{formatNumber(row.packedQty)}</td>
                      <td className="num">{formatNumber(row.shortageQty)}</td>
                      <td className="num">{formatPercent(row.progressPct)}</td>
                      <td>
                        <span
                          className="status-badge"
                          style={{ backgroundColor: STATUS_COLORS[row.status] ?? "#64748b" }}
                        >
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="split-grid">
            <div className="card">
              <h2>5) 부족수량 우선순위</h2>
              <div className="table-wrap compact">
                <table>
                  <thead>
                    <tr>
                      <th>제품코드</th>
                      <th>제품명</th>
                      <th className="num">부족수량</th>
                      <th className="num">포장진행률</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dashboard.priorityRows.slice(0, 30).map((row) => (
                      <tr key={`priority-${row.productCode}`}>
                        <td>{row.productCode}</td>
                        <td>{row.productName}</td>
                        <td className="num">{formatNumber(row.shortageQty)}</td>
                        <td className="num">{formatPercent(row.progressPct)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="card">
              <h2>6) 요청外 포장 / 초과 포장</h2>
              <h3 className="sub-title">요청外 포장</h3>
              <div className="table-wrap compact">
                <table>
                  <thead>
                    <tr>
                      <th>제품코드</th>
                      <th>제품명</th>
                      <th className="num">포장수량</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dashboard.outsideRows.length === 0 ? (
                      <tr>
                        <td colSpan={3} className="empty">
                          없음
                        </td>
                      </tr>
                    ) : (
                      dashboard.outsideRows.map((row) => (
                        <tr key={`outside-${row.productCode}`}>
                          <td>{row.productCode}</td>
                          <td>{row.productName}</td>
                          <td className="num">{formatNumber(row.packedQty)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              <h3 className="sub-title">초과 포장</h3>
              <div className="table-wrap compact">
                <table>
                  <thead>
                    <tr>
                      <th>제품코드</th>
                      <th>제품명</th>
                      <th className="num">요청수량</th>
                      <th className="num">포장수량</th>
                      <th className="num">초과수량</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dashboard.excessRows.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="empty">
                          없음
                        </td>
                      </tr>
                    ) : (
                      dashboard.excessRows.map((row) => (
                        <tr key={`excess-${row.productCode}`}>
                          <td>{row.productCode}</td>
                          <td>{row.productName}</td>
                          <td className="num">{formatNumber(row.requestQty)}</td>
                          <td className="num">{formatNumber(row.packedQty)}</td>
                          <td className="num">{formatNumber(row.packedQty - row.requestQty)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>

          <section className="card">
            <h2>7) 일자별 포장수량 추이</h2>
            {dashboard.trendData.length === 0 ? (
              <p className="muted">포장일자 매핑 컬럼이 없거나 유효한 날짜 데이터가 없습니다.</p>
            ) : (
              <div className="chart-wrap">
                <ResponsiveContainer width="100%" height={320}>
                  <LineChart data={dashboard.trendData} margin={{ top: 8, right: 16, left: 0, bottom: 8 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" />
                    <YAxis />
                    <Tooltip
                      formatter={(value) => formatNumber(value)}
                      contentStyle={{ borderRadius: 12, border: "1px solid #cbd5e1" }}
                    />
                    <Line type="monotone" dataKey="qty" name="포장수량" stroke="#2563eb" strokeWidth={3} dot={false} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            )}
          </section>
        </>
      ) : null}
    </div>
  );
}

export default App;
