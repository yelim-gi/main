
import { useEffect, useMemo, useState } from "react";
import { supabase } from "./supabase";
import { createPortal } from "react-dom";


const ADMIN_EMAIL = "qzwxec88888@gmail.com";
import * as XLSX from "xlsx";
import XlsxPopulate from "xlsx-populate/browser/xlsx-populate";
import "./App.css";

const PRICE_RANGES = [
  "전체", "0~5000", "5000~10000", "10000~15000", "15000~20000",
  "20000~25000", "25000~30000", "30000~35000", "35000+",
];

const TABS = ["대시보드", "정산매입매출", "AI사입입고분석", "재고관리", "수동박스", "주문관리", "택배접수", "라방주문"];

function nowString() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function toNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  const cleaned = String(v).replaceAll(",", "").replaceAll("원", "").replaceAll("%", "").trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function toInt(v) {
  return Math.round(toNum(v));
}

function money(v) {
  return `${toInt(v).toLocaleString()}원`;
}

function normalizeColName(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[ \t\n\r()[\]{}_\-·./]/g, "");
}

function splitMultiValues(value) {
  if (!value) return [];
  let s = String(value).trim();
  ["\n", "\r", "/", "\\", ",", "，", "、", "·", "ㆍ", "|", "&", "+"].forEach((sep) => {
    s = s.replaceAll(sep, ",");
  });
  const out = [];
  s.split(",").forEach((v) => {
    const t = v.trim().replace(/\s+/g, " ");
    if (t && !out.includes(t)) out.push(t);
  });
  return out;
}

function valueMatchesSelected(value, selected) {
  if (!selected || selected.length === 0) return true;
  const tokens = splitMultiValues(value);
  return selected.some((v) => tokens.includes(v));
}

function inPriceRange(value, label) {
  const price = toNum(value);
  if (label === "전체") return true;
  if (label.endsWith("+")) return price >= Number(label.replace("+", ""));
  const [min, max] = label.split("~").map(Number);
  return price >= min && price < max;
}

function calcFinance(items, salePrice, feeRate) {
  const wholesaleSum = items.reduce((sum, p) => sum + toInt(p.wholesale), 0);
  const retailSum = items.reduce((sum, p) => sum + toInt(p.retail), 0);
  const sale = toInt(salePrice);
  const feeAmount = Math.round((sale * Number(feeRate || 0)) / 100);
  const netAmount = sale - feeAmount;
  const profit = netAmount - wholesaleSum;
  const margin = sale > 0 ? (profit / sale) * 100 : 0;
  return { wholesaleSum, retailSum, feeAmount, netAmount, profit, margin };
}

function pickCol(row, names) {
  const keys = Object.keys(row || {});
  const direct = {};
  const norm = {};
  keys.forEach((k) => {
    direct[String(k).trim()] = k;
    norm[normalizeColName(k)] = k;
  });
  for (const name of names) {
    if (direct[name]) return row[direct[name]];
    const nk = normalizeColName(name);
    if (norm[nk]) return row[norm[nk]];
  }
  return "";
}


function buildRecommendationCheck({ type, saleTotal, retailTarget, bodyRetailSum, totalRetailSum, margin, minMargin, maxMargin, giftName, zeroStockNames = [] }) {
  const checks = [];
  const retailOk = bodyRetailSum >= retailTarget;
  const marginOk = margin >= minMargin && margin <= maxMargin;
  const giftOk = type !== "소확행" || !!giftName;

  checks.push(retailOk ? "✅ 소비자가 조건 통과" : "❌ 소비자가 부족");
  checks.push(marginOk ? "✅ 마진 조건 통과" : "❌ 마진 범위 밖");
  if (type === "소확행") checks.push(giftOk ? "🎁 랜덤선물 포함" : "❌ 랜덤선물 없음");
  if (zeroStockNames.length > 0) checks.push(`⚠️ 마지막 재고 포함 ${zeroStockNames.length}개`);
  else checks.push("✅ 마지막 재고 없음");

  return checks.join(" / ");
}

function productFromExcelRow(row) {
  const name = pickCol(row, ["상품명", "상품 이름", "제품명", "품명", "name"]);
  let retail = pickCol(row, ["개별가격", "소비자가", "판매가", "정가", "retail"]);
  if (!retail && name) {
    const m = String(name).match(/^\s*(\d+)/);
    if (m) retail = m[1];
  }
  return {
    name: String(name || "").trim(),
    char1: String(pickCol(row, ["캐릭터(1)", "캐릭터1", "캐릭터 1", "캐릭터①", "캐릭터대분류", "대분류캐릭터", "대표캐릭터", "브랜드", "char1"]) || "").trim(),
    char2: String(pickCol(row, ["캐릭터(2)", "캐릭터2", "캐릭터 2", "캐릭터②", "캐릭터소분류", "소분류캐릭터", "세부캐릭터", "상세캐릭터", "캐릭터명", "캐릭터", "char2"]) || "").trim(),
    category: String(pickCol(row, ["카테고리", "분류", "category"]) || "").trim(),
    stock: toInt(pickCol(row, ["현재재고", "재고", "수량", "stock"])),
    wholesale: toInt(pickCol(row, ["도매가", "원가", "매입가", "wholesale"])),
    retail: toInt(retail),
    hidden: false,
  };
}

function scoreProductForStyle(p, style) {
  const r = toInt(p.retail);
  if (style === "자잘자잘") return 100000 - r;
  if (style === "큼직큼직") return r;
  if (style === "믹스") return 50000 - Math.abs(r - 12000);
  return Math.random() * 1000;
}


function isGiftCandidate(p) {
  const name = String(p.name || "");
  const cat = String(p.category || "");
  const r = toInt(p.retail);
  return (
    name.includes("소확행") ||
    name.includes("랜덤선물") ||
    cat.includes("소확행") ||
    cat.includes("선물") ||
    (r >= 3000 && r <= 6000)
  );
}

function isWithinMargin(fin, targetMargin) {
  const min = Number(targetMargin || 0);
  const max = min + 5;
  return fin.margin >= min && fin.margin <= max;
}

function retailSumOf(items) {
  return items.reduce((s, p) => s + toInt(p.retail), 0);
}

function bodyItemsOf(items) {
  return items.filter((p) => p._tag !== "랜덤선물");
}

function findClosestRetailProduct(pool, gap, usedIds = new Set()) {
  return pool
    .filter((p) => !usedIds.has(p.id))
    .sort((a, b) => Math.abs(toInt(a.retail) - gap) - Math.abs(toInt(b.retail) - gap))[0];
}

function productCharacters(p) {
  return [...splitMultiValues(p.char1), ...splitMultiValues(p.char2)].filter(Boolean);
}

function hasSharedCharacter(a, b) {
  const aa = new Set(Array.isArray(a) ? a : productCharacters(a));
  const bb = Array.isArray(b) ? b : productCharacters(b);
  return Array.from(bb).some((x) => aa.has(x));
}



function productMatchesPreferredChars(p, pref1, pref2) {
  const chars = productCharacters(p);
  const c1Ok = !pref1?.length || splitMultiValues(p.char1).some((x) => pref1.includes(x));
  const c2Ok = !pref2?.length || splitMultiValues(p.char2).some((x) => pref2.includes(x));
  return c1Ok && c2Ok;
}



function v48MatchesSelectedCharacter(p, selected) {
  if (!selected || selected.length === 0) return true;
  const chars = [...splitMultiValues(p.char1), ...splitMultiValues(p.char2)];
  return chars.some((c) => selected.includes(c));
}

function v48SelectedChars(a = [], b = []) {
  return Array.from(new Set([...(a || []), ...(b || [])])).filter(Boolean);
}


function productRetailValue(p) {
  return toInt(p.retail || p.retail_price || p.consumer_price || p.retailPrice || p.consumerPrice || 0);
}

function productWholesaleValue(p) {
  return toInt(p.wholesale || p.wholesale_price || p.cost || p.wholesalePrice || 0);
}

function compactText(v, max = 42) {
  const s = String(v || "");
  return s.length > max ? s.slice(0, max) + "…" : s;
}

function MultiCheckFilter({ label, options, selected, setSelected }) {
  const [open, setOpen] = useState(false);
  const [kw, setKw] = useState("");

  const shown = options.filter((v) => v !== "전체" && (!kw.trim() || String(v).toLowerCase().includes(kw.trim().toLowerCase())));
  const text = selected.length === 0 ? `${label}: 전체` : selected.length === 1 ? `${label}: ${selected[0]}` : `${label}: ${selected.length}개 선택`;

  function toggle(v) {
    setSelected(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]);
  }

  function selectShown() {
    setSelected(Array.from(new Set([...selected, ...shown])));
  }

  function clearShown() {
    setSelected(selected.filter((x) => !shown.includes(x)));
  }

  const modal = (
    <div className="modalOverlay fixedCharacterModalOverlay" onMouseDown={(e) => {
      if (e.target.classList.contains("modalOverlay")) setOpen(false);
    }}>
      <div className="multiModal fixedCharacterModal" onMouseDown={(e) => e.stopPropagation()}>
        <div className="modalTitle">
          <strong>{label} 선택 ({options.filter((v) => v !== "전체").length}개)</strong>
          <button type="button" className="closeBtn" onClick={() => setOpen(false)}>닫기</button>
        </div>

        <div className="modalSearchRow">
          <label>검색</label>
          <input value={kw} onChange={(e) => setKw(e.target.value)} placeholder={`${label} 검색`} autoFocus />
        </div>

        <div className="selectedPreview">
          {selected.length === 0 ? "선택: 전체" : "선택: " + selected.join(", ")}
        </div>

        <div className="multiActions">
          <button type="button" onClick={selectShown}>현재 검색 전체선택</button>
          <button type="button" onClick={clearShown}>현재 검색 해제</button>
          <button type="button" onClick={() => setSelected([])}>전체 해제</button>
        </div>

        <div className="modalCheckList">
          {shown.map((v) => (
            <label key={v} className="modalCheckItem">
              <input type="checkbox" checked={selected.includes(v)} onChange={() => toggle(v)} />
              <span>{v}</span>
            </label>
          ))}
          {shown.length === 0 && <div className="emptySmall">목록 없음</div>}
        </div>

        <div className="modalBottom">
          <button type="button" onClick={() => setOpen(false)}>적용</button>
        </div>
      </div>
    </div>
  );

  return (
    <>
      <button type="button" className="multiBtn" onClick={() => setOpen(true)} title={selected.length ? selected.join(", ") : "전체"}>
        {text}
      </button>
      {open && createPortal(modal, document.body)}
    </>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState("대시보드");
  const [geminiOpen, setGeminiOpen] = useState(false);
  const [geminiInput, setGeminiInput] = useState("");
  const [geminiMessages, setGeminiMessages] = useState([
    { role: "assistant", text: "안녕하세요. 여깁니다유 운영 비서예요. 고객 요청사항 분석, 랜덤박스 조합, 상품 검색, 상품명/카테고리 수정, 택배접수 정리 등을 도와드릴게요." }
  ]);
  const [geminiLoading, setGeminiLoading] = useState(false);
  const [geminiActionDraft, setGeminiActionDraft] = useState(null);
  const [todoItems, setTodoItems] = useState([]);
  const [newTodoText, setNewTodoText] = useState("");
  const [quickTodo, setQuickTodo] = useState("");
  const [financeYear, setFinanceYear] = useState(String(new Date().getFullYear()));
  const [financeMonth, setFinanceMonth] = useState("전체");
  const [shippingPasteText, setShippingPasteText] = useState("");
  const [shippingRows, setShippingRows] = useState([]);
  const [v48ManualStrictCharsOnly, setV48ManualStrictCharsOnly] = useState(false);
  const [v48ScoopStrictCharsOnly, setV48ScoopStrictCharsOnly] = useState(false);

  const [authUser, setAuthUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [products, setProducts] = useState([]);
  const [orders, setOrders] = useState([]);
  const [orderItems, setOrderItems] = useState([]);
  const [materials, setMaterials] = useState([]);

  const [selectedProductId, setSelectedProductId] = useState(null);
  const [bulkSelectedProductIds, setBulkSelectedProductIds] = useState([]);
  const [bulkEditForm, setBulkEditForm] = useState(null);
  const [selectedOrderId, setSelectedOrderId] = useState(null);
  const [selectedOrderItemsOpen, setSelectedOrderItemsOpen] = useState(false);
  const [selectedMaterialId, setSelectedMaterialId] = useState(null);
  const [isShipping, setIsShipping] = useState(false);
  const [isImportingExcel, setIsImportingExcel] = useState(false);

  const [search, setSearch] = useState("");
  const [searchInput, setSearchInput] = useState("");
  const [char1Selected, setChar1Selected] = useState([]);
  const [char2Selected, setChar2Selected] = useState([]);
  const [categoryFilter, setCategoryFilter] = useState("전체");
  const [priceFilter, setPriceFilter] = useState("전체");
  const [hiddenOnly, setHiddenOnly] = useState(false);
  const [excludeLowStock, setExcludeLowStock] = useState(false);
  const [productSort, setProductSort] = useState("기본순");

  const [composeItems, setComposeItems] = useState([]);
  const [salePrice, setSalePrice] = useState("39900");
  const [feeRate, setFeeRate] = useState("6.4");
  const [defaultSale, setDefaultSale] = useState("39900");
  const [defaultFee, setDefaultFee] = useState("3.63");
  const [customer, setCustomer] = useState("");
  const [memo, setMemo] = useState("");
  const [reorder, setReorder] = useState(false);

  const [materialName, setMaterialName] = useState("");
  const [materialAmount, setMaterialAmount] = useState("");

  const [orderSearchCustomer, setOrderSearchCustomer] = useState("");
  const [orderSearchDate, setOrderSearchDate] = useState("");
  const [orderReorderOnly, setOrderReorderOnly] = useState(false);

  const [editProductForm, setEditProductForm] = useState(null);
  const [productForm, setProductForm] = useState({
    name: "", char1: "", char2: "", category: "", stock: "", wholesale: "", retail: "", hidden: false,
  });

  const [aiImportFileName, setAiImportFileName] = useState("");
  const [aiImportLoading, setAiImportLoading] = useState(false);
  const [aiImportRawText, setAiImportRawText] = useState("");
  const [aiImportRows, setAiImportRows] = useState([]);
  const [aiImportMode, setAiImportMode] = useState("add");
  const [aiImportPasteText, setAiImportPasteText] = useState("");
  const [manualCandidateCount, setManualCandidateCount] = useState("12");
  const [manualCharStrategy, setManualCharStrategy] = useState("골고루");
  const [manualAiRequest, setManualAiRequest] = useState("");
  const [manualPresetName, setManualPresetName] = useState("");
  const [manualSavedPresets, setManualSavedPresets] = useState(() => {
    try { return JSON.parse(localStorage.getItem("manual_ai_presets") || "[]"); } catch { return []; }
  });
  const [manualAiLoading, setManualAiLoading] = useState(false);
  const [manualAiMemo, setManualAiMemo] = useState("");

  const [scoopAiRequest, setScoopAiRequest] = useState("");
  const [scoopPresetName, setScoopPresetName] = useState("");
  const [scoopSavedPresets, setScoopSavedPresets] = useState(() => {
    try { return JSON.parse(localStorage.getItem("scoop_ai_presets") || "[]"); } catch { return []; }
  });
  const [scoopAiLoading, setScoopAiLoading] = useState(false);
  const [scoopAiMemo, setScoopAiMemo] = useState("");
  const [manualAiRevision, setManualAiRevision] = useState("");
  const [scoopAiRevision, setScoopAiRevision] = useState("");



  const [manualType, setManualType] = useState("프리미엄박스");
  const [manualBoxCount, setManualBoxCount] = useState("1");
  const [manualTargetMargin, setManualTargetMargin] = useState("20");
  const [manualRetailExtra, setManualRetailExtra] = useState("0");
  const [manualHiddenDiscount, setManualHiddenDiscount] = useState("5");
  const [manualStyle, setManualStyle] = useState("선택안함");
  const [manualPrefChar1, setManualPrefChar1] = useState([]);
  const [manualPrefChar2, setManualPrefChar2] = useState([]);
  const [manualRecommendations, setManualRecommendations] = useState([]);
  const [selectedManualIndex, setSelectedManualIndex] = useState(null);
  const [manualCustomer, setManualCustomer] = useState("");
  const [manualMemo, setManualMemo] = useState("");
  const [manualReorder, setManualReorder] = useState(false);

  const [scoopGroupCount, setScoopGroupCount] = useState("6");
  const [scoopMode, setScoopMode] = useState("상품 수 균등");
  const [scoopPrice, setScoopPrice] = useState("전체");
  const [scoopRetailLimit, setScoopRetailLimit] = useState("");
  const [scoopChar1Selected, setScoopChar1Selected] = useState([]);
  const [scoopChar2Selected, setScoopChar2Selected] = useState([]);
  const [scoopGroups, setScoopGroups] = useState([]);
  const [scoopRecommendations, setScoopRecommendations] = useState([]);
  const [selectedScoopIndex, setSelectedScoopIndex] = useState(null);
  const [scoopCustomer, setScoopCustomer] = useState("");
  const [scoopMemo, setScoopMemo] = useState("");
  const [scoopReorder, setScoopReorder] = useState(false);
  const [scoopAnalysisText, setScoopAnalysisText] = useState("카테고리 자동 분석을 누르면 분석 결과가 표시됩니다.");
  const [scoopCategoryStats, setScoopCategoryStats] = useState([]);
  const [scoopExcludedCount, setScoopExcludedCount] = useState(0);
  const [scoopTargetMargin, setScoopTargetMargin] = useState("20");
  const [scoopRecType, setScoopRecType] = useState("전체 보기");
  const [scoopRecSort, setScoopRecSort] = useState("추천순");
  const [scoopSelectedCategories, setScoopSelectedCategories] = useState([]);
  const [scoopGapScope, setScoopGapScope] = useState("same");
  const [selectedOrderItems, setSelectedOrderItems] = useState([]);


  const [liveSessions, setLiveSessions] = useState([]);
  const [liveMembers, setLiveMembers] = useState([]);
  const [liveOrders, setLiveOrders] = useState([]);
  const [selectedLiveSessionId, setSelectedLiveSessionId] = useState(null);
  const [liveProductSearch, setLiveProductSearch] = useState("");
  const [liveStatusFilter, setLiveStatusFilter] = useState("전체");
  const [livePaymentFilter, setLivePaymentFilter] = useState("전체");
  const [liveOrderSearch, setLiveOrderSearch] = useState("");
  const [liveDueOnly, setLiveDueOnly] = useState(false);
  const [liveDataLoading, setLiveDataLoading] = useState(false);
  const [liveMemberSearch, setLiveMemberSearch] = useState("");
  const [liveMemberLookupSearch, setLiveMemberLookupSearch] = useState("");
  const [selectedLiveMemberId, setSelectedLiveMemberId] = useState("");
  const [liveNewSession, setLiveNewSession] = useState({
    title: "", date: new Date().toISOString().slice(0, 10), keepDays: "7", shippingFee: "4000", cardFeeRate: "3",
    bankName: "", accountNumber: "", accountHolder: "여깁니다유",
    notice: "입금 확인 순서대로 포장 후 출고됩니다.\n킵 상품은 킵 기간 만료 후 자동 출고됩니다.\n본 정산서는 여깁니다유 라이브 구매 확인용이며 외부 공유를 금합니다."
  });
  const [liveMemberForm, setLiveMemberForm] = useState({ name: "", phone: "", postalCode: "", baseAddress: "", detailAddress: "", address: "", points: "0", memo: "" });
  const [liveOrderForm, setLiveOrderForm] = useState({ buyer: "", phone: "", postalCode: "", baseAddress: "", detailAddress: "", address: "", paymentMethod: "계좌이체", status: "미입금", trackingNo: "", memo: "", shippingApply: true, cardApply: false, boxWeight: "2", boxVolume: "60", household: "생활용품", deliveryMessage: "", points: "0", usedPoints: 0 });
  const [liveCart, setLiveCart] = useState([]);
  const [editingLiveOrderId, setEditingLiveOrderId] = useState("");


  useEffect(() => {
    if (!selectedLiveSessionId && liveSessions.length > 0) setSelectedLiveSessionId(liveSessions[0].id);
  }, [liveSessions, selectedLiveSessionId]);

  useEffect(() => {
    let mounted = true;

    async function initAuth() {
      const { data } = await supabase.auth.getSession();
      if (!mounted) return;
      setAuthUser(data?.session?.user || null);
      setAuthLoading(false);
    }

    initAuth();

    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setAuthUser(session?.user || null);
      setAuthLoading(false);
    });

    return () => {
      mounted = false;
      sub?.subscription?.unsubscribe?.();
    };
  }, []);

  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem("ai_common_box_settings") || "{}");
      if (saved.salePrice) setSalePrice(saved.salePrice);
      if (saved.feeRate) setFeeRate(saved.feeRate);
      if (saved.manualTargetMargin) setManualTargetMargin(saved.manualTargetMargin);
      if (saved.manualRetailExtra) setManualRetailExtra(saved.manualRetailExtra);
      if (saved.scoopTargetMargin) setScoopTargetMargin(saved.scoopTargetMargin);
      if (saved.manualPrefChar1) setManualPrefChar1(saved.manualPrefChar1);
      if (saved.manualPrefChar2) setManualPrefChar2(saved.manualPrefChar2);
      if (saved.scoopChar1Selected) setScoopChar1Selected(saved.scoopChar1Selected);
      if (saved.scoopChar2Selected) setScoopChar2Selected(saved.scoopChar2Selected);
    } catch {}
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem("ai_common_box_settings", JSON.stringify({
        salePrice,
        feeRate,
        manualTargetMargin,
        manualRetailExtra,
        scoopTargetMargin,
        manualPrefChar1,
        manualPrefChar2,
        scoopChar1Selected,
        scoopChar2Selected,
      }));
    } catch {}
  }, [salePrice, feeRate, manualTargetMargin, manualRetailExtra, scoopTargetMargin, manualPrefChar1, manualPrefChar2, scoopChar1Selected, scoopChar2Selected]);

  useEffect(() => {
    if (!authUser) return;

    loadAll();

    const channels = [
      supabase.channel("products-live").on("postgres_changes", { event: "*", schema: "public", table: "products" }, getProducts).subscribe(),
      supabase.channel("orders-live").on("postgres_changes", { event: "*", schema: "public", table: "orders" }, getOrders).subscribe(),
      supabase.channel("order-items-live").on("postgres_changes", { event: "*", schema: "public", table: "order_items" }, getOrderItems).subscribe(),
      supabase.channel("materials-live").on("postgres_changes", { event: "*", schema: "public", table: "materials" }, getMaterials).subscribe(),
      supabase.channel("settings-live").on("postgres_changes", { event: "*", schema: "public", table: "settings" }, getSettings).subscribe(),
      supabase.channel("live-sessions-db").on("postgres_changes", { event: "*", schema: "public", table: "live_sessions" }, getLiveSessions).subscribe(),
      supabase.channel("live-members-db").on("postgres_changes", { event: "*", schema: "public", table: "live_members" }, getLiveMembers).subscribe(),
      supabase.channel("live-orders-db").on("postgres_changes", { event: "*", schema: "public", table: "live_orders" }, getLiveOrders).subscribe(),
    ];

    return () => channels.forEach((c) => supabase.removeChannel(c));
  }, [authUser]);

  async function loadAll() {
    await Promise.all([getSettings(), getProducts(), getOrders(), getOrderItems(), getMaterials(), getLiveSessions(), getLiveMembers(), getLiveOrders()]);
  }

  async function writeAudit(action, detail) {
    try {
      await supabase.from("audit_logs").insert([{ action, detail: String(detail || "") }]);
    } catch (e) {
      console.log("audit log failed", e);
    }
  }

  async function createInventoryBackup(reason = "manual") {
    const { data, error } = await supabase.from("products").select("*").order("id", { ascending: true });
    if (error) {
      alert("백업 생성 실패: " + error.message);
      return false;
    }
    const payload = JSON.stringify(data || []);
    const { error: insertError } = await supabase.from("inventory_backups").insert([{ reason, data: payload }]);
    if (insertError) {
      alert("백업 저장 실패: " + insertError.message);
      return false;
    }
    await writeAudit("inventory_backup", `${reason} / ${data?.length || 0} items`);
    return true;
  }

  async function restoreLatestInventoryBackup() {
    const ok = window.confirm("가장 최근 재고 백업으로 복구할까요? 현재 재고는 복구 전에 다시 백업됩니다.");
    if (!ok) return;

    const { data: backups, error } = await supabase
      .from("inventory_backups")
      .select("*")
      .order("id", { ascending: false })
      .limit(1);

    if (error || !backups || backups.length === 0) {
      alert("복구할 백업이 없어요.");
      return;
    }

    const beforeOk = await createInventoryBackup("restore_before_backup");
    if (!beforeOk) return;

    let rows = [];
    try {
      rows = JSON.parse(backups[0].data || "[]");
    } catch {
      alert("백업 데이터가 손상됐어요.");
      return;
    }

    const restoreRows = rows.map((p) => ({
      name: p.name,
      char1: p.char1,
      char2: p.char2,
      category: p.category,
      stock: toInt(p.stock),
      wholesale: toInt(p.wholesale),
      retail: toInt(p.retail),
      hidden: !!p.hidden,
    }));

    const { error: delErr } = await supabase.from("products").delete().neq("id", 0);
    if (delErr) return alert("현재 재고 삭제 실패: " + delErr.message);

    if (restoreRows.length > 0) {
      const { error: insErr } = await supabase.from("products").insert(restoreRows);
      if (insErr) return alert("백업 복구 실패: " + insErr.message);
    }

    await writeAudit("inventory_restore", `backup_id=${backups[0].id} / rows=${restoreRows.length}`);
    alert("최근 백업으로 재고를 복구했어요.");
    getProducts();
  }

  function downloadCurrentInventoryBackupFile() {
    const data = products.map((p) => ({
      상품명: p.name,
      캐릭터1: p.char1,
      캐릭터2: p.char2,
      카테고리: p.category,
      재고: p.stock,
      도매가: p.wholesale,
      소비자가: p.retail,
      히든: p.hidden ? "Y" : "",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "재고백업");
    XLSX.writeFile(wb, `재고백업_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  async function getProducts() {
    const { data, error } = await supabase.from("products").select("*").order("id", { ascending: false });
    if (error) return alert("재고 불러오기 실패: " + error.message);
    setProducts(data || []);
  }

  async function getOrders() {
    const { data, error } = await supabase.from("orders").select("*").order("id", { ascending: false });
    if (error) return console.log(error);
    setOrders(data || []);
  }

  async function getOrderItems() {
    const { data, error } = await supabase.from("order_items").select("*").order("id", { ascending: true });
    if (error) return console.log(error);
    setOrderItems(data || []);
  }

  async function getMaterials() {
    const { data, error } = await supabase.from("materials").select("*").order("id", { ascending: false });
    if (error) return console.log(error);
    setMaterials(data || []);
  }

  async function getSettings() {
    const { data, error } = await supabase.from("settings").select("*");
    if (error) return console.log(error);
    const map = {};
    (data || []).forEach((r) => { map[r.k] = r.v; });

    const s = map.default_sale || "39900";
    const f = map.default_fee || "3.63";
    setDefaultSale(s);
    setDefaultFee(f);
    setSalePrice((prev) => prev || s);
    setFeeRate((prev) => prev || f);
    setManualType(map.manual_type || "프리미엄박스");
    setManualBoxCount(map.manual_box_count || "1");
    setManualTargetMargin(map.manual_target_margin || "20");
    setManualRetailExtra(map.manual_retail_extra || "0");
    setManualHiddenDiscount(map.manual_hidden_discount || "5");
    setManualStyle(map.manual_style || "선택안함");
    setScoopGroupCount(map.scoop_groups || "6");
    setScoopMode(map.scoop_mode || "상품 수 균등");
    setScoopRetailLimit(map.scoop_retail_limit || "");
    setScoopTargetMargin(map.scoop_target_margin || "20");
    setScoopRecType(map.scoop_rec_type || "전체 보기");
    setScoopRecSort(map.scoop_rec_sort || "추천순");
  }

  async function saveSettings() {
    const payload = [
      { k: "default_sale", v: String(defaultSale || "39900") },
      { k: "default_fee", v: String(defaultFee || "3.63") },
      { k: "manual_type", v: String(manualType) },
      { k: "manual_box_count", v: String(manualBoxCount) },
      { k: "manual_target_margin", v: String(manualTargetMargin) },
      { k: "manual_retail_extra", v: String(manualRetailExtra) },
      { k: "manual_hidden_discount", v: String(manualHiddenDiscount) },
      { k: "manual_style", v: String(manualStyle) },
      { k: "scoop_groups", v: String(scoopGroupCount || "6") },
      { k: "scoop_mode", v: String(scoopMode || "상품 수 균등") },
      { k: "scoop_retail_limit", v: String(scoopRetailLimit || "") },
      { k: "scoop_target_margin", v: String(scoopTargetMargin || "20") },
      { k: "scoop_rec_type", v: String(scoopRecType || "전체 보기") },
      { k: "scoop_rec_sort", v: String(scoopRecSort || "추천순") },
    ];
    const { error } = await supabase.from("settings").upsert(payload);
    if (error) return alert("설정 저장 실패: " + error.message);
    setSalePrice(defaultSale);
    setFeeRate(defaultFee);
    alert("설정 저장 완료!");
  }

  const char1Options = useMemo(() => ["전체", ...Array.from(new Set(products.flatMap((p) => splitMultiValues(p.char1)))).sort()], [products]);
  const char2Options = useMemo(() => ["전체", ...Array.from(new Set(products.flatMap((p) => splitMultiValues(p.char2)))).sort()], [products]);
  const categoryOptions = useMemo(() => ["전체", ...Array.from(new Set(products.map((p) => p.category).filter(Boolean))).sort()], [products]);

  const filteredProducts = useMemo(() => {
    const kw = search.trim().toLowerCase();
    const rows = products.filter((p) => {
      const keyword =
        !kw ||
        String(p.name || "").toLowerCase().includes(kw) ||
        String(p.char1 || "").toLowerCase().includes(kw) ||
        String(p.char2 || "").toLowerCase().includes(kw) ||
        String(p.category || "").toLowerCase().includes(kw);
      const c1 = valueMatchesSelected(p.char1, char1Selected);
      const c2 = valueMatchesSelected(p.char2, char2Selected);
      const cat = categoryFilter === "전체" || p.category === categoryFilter;
      const price = inPriceRange(p.retail, priceFilter);
      const hidden = !hiddenOnly || p.hidden === true;
      const lowStock = !excludeLowStock || toInt(p.stock) >= 2;
      return keyword && c1 && c2 && cat && price && hidden && lowStock;
    });

    const sorted = [...rows];
    if (productSort === "도매가 낮은순") sorted.sort((a, b) => toInt(a.wholesale) - toInt(b.wholesale));
    if (productSort === "도매가 높은순") sorted.sort((a, b) => toInt(b.wholesale) - toInt(a.wholesale));
    if (productSort === "소비자가 낮은순") sorted.sort((a, b) => toInt(a.retail) - toInt(b.retail));
    if (productSort === "소비자가 높은순") sorted.sort((a, b) => toInt(b.retail) - toInt(a.retail));
    if (productSort === "재고 많은순") sorted.sort((a, b) => toInt(b.stock) - toInt(a.stock));
    if (productSort === "재고 적은순") sorted.sort((a, b) => toInt(a.stock) - toInt(b.stock));
    if (productSort === "상품명순") sorted.sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "ko"));
    return sorted;
  }, [products, search, char1Selected, char2Selected, categoryFilter, priceFilter, hiddenOnly, excludeLowStock, productSort]);

  const totalStock = products.reduce((s, p) => s + toInt(p.stock), 0);
  const totalWholesale = products.reduce((s, p) => s + toInt(p.stock) * toInt(p.wholesale), 0);
  const totalMaterials = materials.reduce((s, m) => s + toInt(m.amount), 0);
  const completedOrders = orders.filter((o) => o.status === "출고완료");
  const totalSales = completedOrders.reduce((s, o) => s + toInt(o.sale_price), 0);
  const totalNet = completedOrders.reduce((s, o) => s + toInt(o.net_amount), 0);
  const totalProfit = completedOrders.reduce((s, o) => s + toInt(o.profit), 0);
  const finance = calcFinance(composeItems, salePrice, feeRate);

  const filteredOrders = useMemo(() => {
    return orders.filter((o) => {
      if (o.deleted_at) return false;
      const customerOk = !orderSearchCustomer.trim() || String(o.customer || "").toLowerCase().includes(orderSearchCustomer.trim().toLowerCase());
      const dateOk = !orderSearchDate.trim() || String(o.created_at || "").slice(0, 10) === orderSearchDate.trim();
      const reorderOk = !orderReorderOnly || toInt(o.reorder) === 1;
      return customerOk && dateOk && reorderOk;
    });
  }, [orders, orderSearchCustomer, orderSearchDate, orderReorderOnly]);


  const trashOrders = useMemo(() => {
    return orders.filter((o) => o.deleted_at);
  }, [orders]);

  function daysLeftForTrash(order) {
    if (!order.deleted_at) return "-";
    const deleted = new Date(order.deleted_at);
    if (Number.isNaN(deleted.getTime())) return "-";
    const ms = Date.now() - deleted.getTime();
    const daysPassed = Math.floor(ms / (1000 * 60 * 60 * 24));
    return Math.max(0, 30 - daysPassed);
  }

  const pendingOrders = filteredOrders.filter((o) => o.status !== "출고완료" && o.status !== "취소");
  const shippedOrders = filteredOrders.filter((o) => o.status === "출고완료");
  const canceledOrders = filteredOrders.filter((o) => o.status === "취소");

  function resetFilters() {
    setSearch("");
    setSearchInput("");
    setChar1Selected([]);
    setChar2Selected([]);
    setCategoryFilter("전체");
    setPriceFilter("전체");
    setHiddenOnly(false);
    setExcludeLowStock(false);
    setProductSort("기본순");
  }

  async function addProduct() {
    if (!productForm.name.trim()) return alert("상품명을 입력해줘.");
    const { error } = await supabase.from("products").insert([{
      name: productForm.name,
      char1: productForm.char1,
      char2: productForm.char2,
      category: productForm.category,
      stock: toInt(productForm.stock),
      wholesale: toInt(productForm.wholesale),
      retail: toInt(productForm.retail),
      hidden: !!productForm.hidden,
    }]);
    if (error) return alert("상품 저장 실패: " + error.message);
    setProductForm({ name: "", char1: "", char2: "", category: "", stock: "", wholesale: "", retail: "", hidden: false });
    getProducts();
  }

  async function deleteProduct(id) {
    if (!id) return alert("삭제할 상품을 선택해줘.");
    if (!window.confirm("선택한 상품을 삭제할까?")) return;
    const { error } = await supabase.from("products").delete().eq("id", id);
    if (error) return alert("상품 삭제 실패: " + error.message);
    setSelectedProductId(null);
    getProducts();
  }

  async function addMaterial() {
    if (!materialName.trim()) return alert("재료비명을 입력해줘.");
    if (toInt(materialAmount) <= 0) return alert("금액을 입력해줘.");
    const { error } = await supabase.from("materials").insert([{ name: materialName, amount: toInt(materialAmount) }]);
    if (error) return alert("재료비 저장 실패: " + error.message);
    setMaterialName("");
    setMaterialAmount("");
    setSelectedMaterialId(null);
    getMaterials();
  }

  async function deleteMaterial() {
    if (!selectedMaterialId) return alert("삭제할 재료비를 선택해줘.");
    if (!window.confirm("선택한 재료비를 삭제할까?")) return;
    const { error } = await supabase.from("materials").delete().eq("id", selectedMaterialId);
    if (error) return alert("재료비 삭제 실패: " + error.message);
    setSelectedMaterialId(null);
    getMaterials();
  }

  async function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const mode = window.prompt("엑셀 불러오기 방식\n1 = 기존 재고 전체 삭제 후 교체\n2 = 기존 재고 유지하고 추가", "2");
    if (mode !== "1" && mode !== "2") {
      e.target.value = "";
      return;
    }

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const workbook = XLSX.read(new Uint8Array(evt.target.result), { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const formatted = json.map(productFromExcelRow).filter((x) => x.name);

        if (formatted.length === 0) {
          alert("불러올 상품이 없어요.\n엑셀 첫 줄 컬럼명을 확인해줘.");
          return;
        }

        setIsImportingExcel(true);

        const backupOk = await createInventoryBackup(mode === "1" ? "excel_replace_before" : "excel_add_before");
        if (!backupOk) {
          setIsImportingExcel(false);
          return;
        }

        if (mode === "1") {
          const confirmText = window.prompt("기존 재고 전체 삭제 후 교체합니다.\n자동 백업은 완료됐지만 신중하게 진행해야 해요.\n진행하려면 '교체' 라고 입력해줘.");
          if (confirmText !== "교체") {
            setIsImportingExcel(false);
            alert("재고 교체를 취소했어요.");
            return;
          }
          const { error: delErr } = await supabase.from("products").delete().neq("id", 0);
          if (delErr) {
            setIsImportingExcel(false);
            return alert("기존 재고 삭제 실패: " + delErr.message);
          }
        }

        const { error } = await supabase.from("products").insert(formatted);
        if (error) {
          setIsImportingExcel(false);
          return alert("엑셀 업로드 실패: " + error.message);
        }

        await writeAudit("excel_import", `${mode === "1" ? "replace" : "add"} / rows=${formatted.length}`);
        setIsImportingExcel(false);
        alert(`엑셀 ${mode === "1" ? "교체" : "추가"} 완료!\n불러온 상품 수: ${formatted.length}개\n업로드 전 재고는 자동 백업됐어요.`);
        getProducts();
      } catch (err) {
        console.error(err);
        alert("엑셀을 읽는 중 오류가 났어요.\n파일 형식 또는 컬럼명을 확인해줘.");
      } finally {
        setIsImportingExcel(false);
        e.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function downloadInventoryExcel() {
    const data = products.map((p) => ({
      상품명: p.name, 캐릭터1: p.char1, 캐릭터2: p.char2, 카테고리: p.category,
      재고: p.stock, 도매가: p.wholesale, 소비자가: p.retail, 히든: p.hidden ? "Y" : "",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "재고");
    XLSX.writeFile(wb, "재고.xlsx");
  }

  function downloadOrdersExcel() {
    const data = orders.map((o) => ({
      주문ID: o.id, 주문일: o.created_at, 주문자: o.customer,
      재주문: toInt(o.reorder) === 1 ? "Y" : "", 상태: o.status,
      판매가: o.sale_price, 수수료율: o.fee_rate, 수수료: o.fee_amount,
      도매가합: o.wholesale_sum, 소비자가합: o.retail_sum,
      실수령액: o.net_amount, 순이익: o.profit, 취소사유: o.cancel_reason || "", 메모: o.memo || "",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "주문");
    XLSX.writeFile(wb, "주문.xlsx");
  }

  function showChar2Values() {
    const values = Array.from(new Set(products.flatMap((p) => splitMultiValues(p.char2)))).sort();
    alert(values.length ? `캐릭터2 목록 (${values.length}개)\n\n${values.join("\n")}` : "캐릭터2 목록이 없어요.");
  }

  function showCharacterShortage() {
    const map = {};
    products.forEach((p) => {
      splitMultiValues(p.char2).forEach((c) => {
        if (!map[c]) map[c] = { count: 0, stock: 0 };
        map[c].count += 1;
        map[c].stock += toInt(p.stock);
      });
    });
    const rows = Object.entries(map).map(([name, v]) => ({ name, ...v })).sort((a, b) => a.stock - b.stock);
    alert(rows.length ? rows.map((r) => `${r.name} | 상품종류 ${r.count}개 | 총재고 ${r.stock}개`).join("\n") : "계산할 캐릭터가 없어요.");
  }


  function preserveManualProductListScroll(callback) {
    const selector = ".composeConditionResultArea .tableWrap, .manualProductListPanel .tableWrap";
    const beforeEl = document.querySelector(selector);
    const top = beforeEl ? beforeEl.scrollTop : 0;
    const left = beforeEl ? beforeEl.scrollLeft : 0;

    callback();

    requestAnimationFrame(() => {
      const afterEl = document.querySelector(selector);
      if (afterEl) {
        afterEl.scrollTop = top;
        afterEl.scrollLeft = left;
      }
    });
  }

  function addToCompose(product) {
    if (toInt(product.stock) <= 0) return alert("재고가 0개인 상품이에요.");
    preserveManualProductListScroll(() => {
      setComposeItems((prev) => [...prev, product]);
    });
  }

  function clearCompose() {
    setComposeItems([]);
    setCustomer("");
    setMemo("");
    setReorder(false);
  }

  async function checkAndReserveStock(items) {
    const needed = {};
    items.forEach((item) => { needed[item.id] = (needed[item.id] || 0) + 1; });

    const zeroWarnings = [];
    for (const pid of Object.keys(needed)) {
      const { data, error } = await supabase.from("products").select("id,name,stock").eq("id", pid).single();
      if (error || !data) {
        alert(`상품ID ${pid}를 찾지 못했어요.`);
        return false;
      }
      if (toInt(data.stock) < needed[pid]) {
        alert(`${data.name} 재고 부족\n필요 ${needed[pid]}개 / 현재 ${data.stock}개`);
        return false;
      }
      if (toInt(data.stock) - needed[pid] === 0) {
        zeroWarnings.push(`${data.name} | 현재 ${data.stock}개 → 출고 후 0개`);
      }
    }

    if (zeroWarnings.length > 0) {
      const ok = window.confirm(
        "아래 상품은 주문생성/박스출고 처리하면 임시차감되어 재고가 0개가 됩니다.\n\n" +
        zeroWarnings.join("\n") +
        "\n\n그래도 주문생성하고 재고를 임시차감할까요?"
      );
      if (!ok) return false;
    }

    for (const pid of Object.keys(needed)) {
      const { data } = await supabase.from("products").select("stock").eq("id", pid).single();
      await supabase.from("products").update({ stock: Math.max(0, toInt(data?.stock) - needed[pid]) }).eq("id", pid);
    }
    return true;
  }

  async function restoreStockFromItems(items) {
    const needed = {};
    (items || []).forEach((item) => { needed[item.id] = (needed[item.id] || 0) + 1; });

    for (const pid of Object.keys(needed)) {
      const { data: p } = await supabase.from("products").select("stock").eq("id", pid).single();
      if (!p) continue;
      await supabase.from("products").update({ stock: toInt(p.stock) + needed[pid] }).eq("id", pid);
    }
  }

  async function createOrderFromItems(items, orderCustomer, orderMemo, isReorder, orderSalePrice, orderFeeRate) {
    if (isShipping) return alert("이미 출고 처리 중이에요. 잠시만 기다려줘.");
    if (items.length === 0) return alert("상품이 없어요.");
    if (!String(orderCustomer || "").trim()) return alert("주문자명을 입력해줘.");

    setIsShipping(true);
    const finalOk = finalOrderConfirm(items, toInt(orderSalePrice), Number(orderFeeRate || 0), "주문생성 / 재고 임시차감");
    if (!finalOk) {
      setIsShipping(false);
      return;
    }

    const ok = await checkAndReserveStock(items);
    if (!ok) {
      setIsShipping(false);
      return;
    }

    const fin = calcFinance(items, orderSalePrice, orderFeeRate);
    const { data: order, error } = await supabase.from("orders").insert([{
      created_at: nowString(),
      customer: orderCustomer,
      reorder: isReorder ? 1 : 0,
      memo: orderMemo || "",
      sale_price: toInt(orderSalePrice),
      fee_rate: Number(orderFeeRate || 0),
      fee_amount: fin.feeAmount,
      wholesale_sum: fin.wholesaleSum,
      retail_sum: fin.retailSum,
      net_amount: fin.netAmount,
      profit: fin.profit,
      status: "주문접수(재고임시차감)",
      cancel_reason: "",
    }]).select().single();

    if (error) {
      await restoreStockFromItems(items);
      setIsShipping(false);
      return alert("주문 저장 실패로 임시차감 재고를 다시 복구했어요.\n" + error.message);
    }

    const payload = items.map((p) => ({
      order_id: order.id, product_id: p.id, name: p.name, qty: 1,
      wholesale: toInt(p.wholesale), retail: toInt(p.retail),
    }));
    const { error: itemErr } = await supabase.from("order_items").insert(payload);
    if (itemErr) {
      await restoreStockFromItems(items);
      if (order?.id) await supabase.from("orders").delete().eq("id", order.id);
      setIsShipping(false);
      return alert("주문 상품 저장 실패로 주문을 취소하고 임시차감 재고를 다시 복구했어요.\n" + itemErr.message);
    }

    await writeAudit("order_create", `order_id=${order.id} / customer=${orderCustomer} / items=${items.length}`);
    setIsShipping(false);
    alert(`주문 등록 완료! 주문ID: ${order.id}\n재고는 주문접수 상태에서 임시차감됐어요.\n취소하면 재고가 복구되고, 출고확정은 상태만 출고완료로 바뀝니다.`);
    getProducts();
    getOrders();
    getOrderItems();
    setActiveTab("주문관리");
    return true;
  }

  async function createOrderFromCompose() {
    const result = await createOrderFromItems(composeItems, customer, memo, reorder, salePrice, feeRate);
    if (result === true) {
      const clearOk = window.confirm("박스출고가 완료됐어요. 현재 조합 리스트를 비울까요?\n\n취소를 누르면 현재 조합 리스트가 그대로 유지됩니다.");
      if (clearOk) clearCompose();
    }
  }

  async function restoreStockByOrder(orderId) {
    const { data: items, error } = await supabase.from("order_items").select("*").eq("order_id", orderId);
    if (error) {
      alert("주문상품을 불러오지 못했어요.");
      return false;
    }
    for (const item of items || []) {
      const { data: p } = await supabase.from("products").select("stock").eq("id", item.product_id).single();
      if (!p) continue;
      await supabase.from("products").update({ stock: toInt(p.stock) + toInt(item.qty || 1) }).eq("id", item.product_id);
    }
    return true;
  }

  async function shipSelectedOrder() {
    if (!selectedOrderId) return alert("출고확정할 주문을 선택해줘.");
    const order = orders.find((o) => o.id === selectedOrderId);
    if (!order) return alert("주문 정보를 찾을 수 없어요.");
    if (order.status === "출고완료") return alert("이미 출고완료된 주문이에요.");
    if (order.status === "취소") return alert("취소된 주문은 출고확정할 수 없어요.");
    const ok = window.confirm("출고확정은 재고를 추가로 차감하지 않습니다.\n이미 주문생성 때 임시차감된 재고를 확정 처리하는 단계예요.\n출고완료로 변경할까요?");
    if (!ok) return;
    const { error } = await supabase.from("orders").update({ status: "출고완료" }).eq("id", selectedOrderId);
    if (error) return alert("출고확정 실패: " + error.message);
    alert("출고확정 완료! 재고는 추가 차감되지 않았어요.");
    setSelectedOrderId(null);
    getOrders();
  }

  async function cancelSelectedOrder() {
    if (!selectedOrderId) return alert("취소할 주문을 선택해줘.");
    const order = orders.find((o) => o.id === selectedOrderId);
    if (!order) return alert("주문 정보를 찾을 수 없어요.");
    if (order.status === "취소") return alert("이미 취소된 주문이에요.");
    if (order.deleted_at) return alert("이미 취소보관함에 들어간 주문이에요.");

    const reason = window.prompt(
      "취소사유를 입력해줘.\n\n" +
      "아래 중 하나로 입력:\n" +
      "환불 / 반품 / 취소 / 연습 / 기타\n\n" +
      "주의: 취소 시 주문 구성은 현재 주문상태로 보관되며, 취소 후에는 구성을 다시 복구해 출고상태로 되돌릴 수 없습니다.\n" +
      "재고는 주문생성 때 임시차감된 수량만큼 복구됩니다.",
      "취소"
    );
    if (reason === null) return;

    const cleanReason = reason.trim();
    const allowed = ["환불", "반품", "취소", "연습", "기타"];
    if (!allowed.includes(cleanReason)) {
      alert("취소사유는 환불 / 반품 / 취소 / 연습 / 기타 중 하나로 입력해줘.");
      return;
    }

    const warning = [
      "주문취소 최종 확인",
      "",
      `주문ID: ${order.id}`,
      `주문자: ${order.customer || "-"}`,
      `현재상태: ${order.status}`,
      `취소사유: ${cleanReason}`,
      "",
      "안내:",
      "- 주문접수건/출고확정건 모두 취소하면 재고가 복구됩니다.",
      "- 취소된 주문은 취소보관함에 30일 동안 보관됩니다.",
      "- 취소 후에는 이 주문 구성을 다시 출고상태로 복구할 수 없습니다.",
      "- 출고완료 건을 취소하는 경우, 환불/반품 처리를 실제로 했는지 꼭 확인하세요.",
      "",
      "정말 주문취소할까요?"
    ].join("\\n");

    if (!window.confirm(warning)) return;

    await restoreStockByOrder(selectedOrderId);

    if (cleanReason === "연습") {
      // 연습도 바로 삭제하지 않고 보관함으로 이동
      const { error } = await supabase.from("orders").update({
        status: "취소",
        cancel_reason: cleanReason,
        cancel_detail: "연습 주문",
        canceled_at: nowString(),
        deleted_at: nowString(),
      }).eq("id", selectedOrderId);
      if (error) return alert("연습 주문 취소 실패: " + error.message);
      await writeAudit("order_cancel_practice_to_trash", `order_id=${selectedOrderId}`);
      alert("연습 주문을 취소보관함으로 이동했고 재고를 복구했어요.");
    } else {
      const detail = window.prompt("추가 메모가 있으면 적어줘. 없으면 빈칸으로 확인.", "");
      const { error } = await supabase.from("orders").update({
        status: "취소",
        cancel_reason: cleanReason,
        cancel_detail: detail || "",
        canceled_at: nowString(),
        deleted_at: nowString(),
      }).eq("id", selectedOrderId);
      if (error) return alert("주문취소 실패: " + error.message);
      await writeAudit("order_cancel_to_trash", `order_id=${selectedOrderId} / reason=${cleanReason}`);
      alert("주문취소 완료! 재고가 복구됐고 취소보관함에 30일 보관됩니다.");
    }

    setSelectedOrderId(null);
    getProducts();
    getOrders();
    getOrderItems();
  }

  function showSelectedOrderItems() {
    if (!selectedOrderId) return alert("주문을 선택해줘.");
    document.querySelector(".v51OrderDetailPanel")?.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function downloadCustomerOrderExcel() {
    const data = orders.map((o) => ({
      주문자명: o.customer,
      박스수량: 1,
      판매가: o.sale_price,
      상태: o.status,
      메모: o.memo || "",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "고객용주문");
    XLSX.writeFile(wb, "고객용_주문목록.xlsx");
  }

  function baseManualCandidates() {
    const v48StrictManualChars = v48ManualStrictCharsOnly ? v48SelectedChars(manualPrefChar1, manualPrefChar2) : [];
    return products.filter((p) => {
      if (toInt(p.stock) <= 0) return false;
      if (!valueMatchesSelected(p.char1, manualPrefChar1)) return false;
      if (!valueMatchesSelected(p.char2, manualPrefChar2)) return false;
      return true;
    });
  }

  function targetItemCountByStyle() {
    if (manualStyle === "자잘자잘") return 8;
    if (manualStyle === "큼직큼직") return 4;
    if (manualStyle === "믹스") return 6;
    return 6;
  }

  function chooseUniqueByStock(pool, count, usedIds = new Set()) {
    const result = [];
    const localUsed = new Set(usedIds);
    for (const p of pool) {
      if (result.length >= count) break;
      if (localUsed.has(p.id)) continue;
      if (toInt(p.stock) <= 0) continue;
      result.push(p);
      localUsed.add(p.id);
    }
    return result;
  }


  function warnLowPreferredCharacters() {
    const selected = [...manualPrefChar1, ...manualPrefChar2];
    if (selected.length === 0) return true;

    const lowRows = selected.map((char) => {
      const rows = products.filter((p) => {
        const c1 = splitMultiValues(p.char1);
        const c2 = splitMultiValues(p.char2);
        return c1.includes(char) || c2.includes(char);
      });
      const stock = rows.reduce((s, p) => s + toInt(p.stock), 0);
      return { char, stock };
    }).filter((x) => x.stock <= 3);

    if (lowRows.length === 0) return true;

    return window.confirm(
      "선호 캐릭터 중 재고가 적은 캐릭터가 있어요.\n\n" +
      lowRows.map((x) => `${x.char}: 재고 ${x.stock}개`).join("\n") +
      "\n\n그래도 이 캐릭터들을 포함해서 추천안을 만들까요?"
    );
  }

  function overlapRate(aItems, bItems) {
    const a = new Set((aItems || []).map((p) => p.id));
    const b = new Set((bItems || []).map((p) => p.id));
    if (a.size === 0 || b.size === 0) return 0;
    let same = 0;
    a.forEach((id) => { if (b.has(id)) same += 1; });
    return same / Math.min(a.size, b.size);
  }

  function isTooSimilarToExisting(items, existingRecs, limit = 0.65) {
    return existingRecs.some((r) => overlapRate(items, r.items || []) >= limit);
  }

  function getZeroStockWarnings(items) {
    const need = {};
    (items || []).forEach((p) => { need[p.id] = (need[p.id] || 0) + 1; });
    return Object.entries(need).map(([id, qty]) => {
      const p = products.find((x) => String(x.id) === String(id));
      if (!p) return null;
      return toInt(p.stock) - qty === 0 ? `${p.name} | 현재 ${p.stock}개 → 출고 후 0개` : null;
    }).filter(Boolean);
  }

  function finalOrderConfirm(items, sale, fee, label = "출고") {
    const fin = calcFinance(items || [], sale, fee);
    const zeroWarnings = getZeroStockWarnings(items || []);
    const body = [
      `${label} 전 최종 확인`,
      "",
      `상품 수: ${(items || []).length}개`,
      `판매가: ${money(sale)}`,
      `도매가합: ${money(fin.wholesaleSum)}`,
      `소비자가합: ${money(fin.retailSum)}`,
      `수수료: ${money(fin.feeAmount)}`,
      `실수령액: ${money(fin.netAmount)}`,
      `순이익: ${money(realNetProfit)}`,
      `마진율: ${fin.margin.toFixed(1)}%`,
      "",
      zeroWarnings.length ? "[출고 후 재고 0개 상품]\n" + zeroWarnings.join("\n") : "출고 후 재고 0개 상품 없음",
      "",
      "이대로 출고할까요?"
    ].join("\n");
    return window.confirm(body);
  }

  async function showBackupListAndRestore() {
    const { data, error } = await supabase
      .from("inventory_backups")
      .select("id,created_at,reason,data")
      .order("id", { ascending: false })
      .limit(10);

    if (error) return alert("백업 목록 불러오기 실패: " + error.message);
    if (!data || data.length === 0) return alert("저장된 백업이 없어요.");

    const msg = data.map((b) => {
      let count = 0;
      try { count = JSON.parse(b.data || "[]").length; } catch {}
      return `${b.id}: ${String(b.created_at).slice(0,19)} / ${b.reason || ""} / ${count}개`;
    }).join("\n");

    const id = window.prompt("복구할 백업 ID를 입력해줘.\n\n" + msg);
    if (!id) return;

    const picked = data.find((b) => String(b.id) === String(id));
    if (!picked) return alert("해당 백업 ID를 찾지 못했어요.");

    const previewRows = JSON.parse(picked.data || "[]");
    const ok = window.confirm(
      `백업 ID ${picked.id}로 복구할까요?\n` +
      `백업 시각: ${String(picked.created_at).slice(0,19)}\n` +
      `상품 수: ${previewRows.length}개\n\n` +
      "현재 재고는 복구 전 자동 백업됩니다."
    );
    if (!ok) return;

    const beforeOk = await createInventoryBackup("restore_before_selected_backup");
    if (!beforeOk) return;

    const restoreRows = previewRows.map((p) => ({
      name: p.name,
      char1: p.char1,
      char2: p.char2,
      category: p.category,
      stock: toInt(p.stock),
      wholesale: toInt(p.wholesale),
      retail: toInt(p.retail),
      hidden: !!p.hidden,
    }));

    const { error: delErr } = await supabase.from("products").delete().neq("id", 0);
    if (delErr) return alert("현재 재고 삭제 실패: " + delErr.message);

    if (restoreRows.length > 0) {
      const { error: insErr } = await supabase.from("products").insert(restoreRows);
      if (insErr) return alert("백업 복구 실패: " + insErr.message);
    }

    await writeAudit("inventory_restore_selected", `backup_id=${picked.id} / rows=${restoreRows.length}`);
    alert("선택한 백업으로 복구했어요.");
    getProducts();
  }

  function generateManualRecommendations() {
    if (!warnLowPreferredCharacters()) return;
    const poolRaw = baseManualCandidates();
    if (poolRaw.length === 0) return alert("추천할 후보 상품이 없어요.");

    const saleEach = toInt(salePrice || defaultSale);
    const boxCount = Math.max(1, toInt(manualBoxCount));
    const saleTotal = saleEach * boxCount;
    const fee = Number(feeRate || defaultFee || 0);
    const targetMargin = Number(manualTargetMargin || 0);
    const extraRetail = Math.max(0, toInt(manualRetailExtra || 0));

    // v16: 프리미엄/히든 특수 기준 제거.
    // 모든 유형은 기본적으로 목표마진~목표+5% 범위.
    // 본품 소비자가합은 판매가 + 추가소비자가 이상.
    // 소확행만 본품 완성 후 랜덤선물 추가.
    const minMargin = targetMargin;
    const maxMargin = targetMargin + 5;
    const retailTarget = saleTotal + extraRetail;
    const targetItemCount = targetItemCountByStyle() * boxCount;

    const normalPool = poolRaw.filter((p) => !p.hidden || manualType !== "히든박스");
    const hiddenPool = poolRaw.filter((p) => p.hidden);
    const giftPool = poolRaw.filter(isGiftCandidate);

    const recs = [];
    const signatures = new Set();
    let attempts = 0;

    while (recs.length < Math.max(3, Math.min(30, toInt(manualCandidateCount || 12))) && attempts < 1800) {
      attempts += 1;

      let items = [];
      const used = new Set();
      let gift = null;
      let note = `본품 소비자가합 ${money(retailTarget)} 이상 목표`;

      let pool = [...poolRaw].sort((a, b) => {
        const stockScore = (toInt(b.stock) - toInt(a.stock)) * 25;
        const randomScore = (Math.random() - 0.5) * 10000;
        const styleScore = scoreProductForStyle(b, manualStyle) - scoreProductForStyle(a, manualStyle);
        return styleScore + stockScore + randomScore;
      });

      if (manualType === "히든박스") {
        // 특수 마진 기준은 없애되, 히든박스 유형을 골랐으면 히든템 후보가 있으면 하나 정도 섞어볼 수 있게만 함.
        const hidden = hiddenPool
          .filter((p) => !used.has(p.id))
          .sort(() => 0.5 - Math.random())[0];
        if (hidden && Math.random() < 0.7) {
          items.push({ ...hidden, _tag: "히든 후보" });
          used.add(hidden.id);
          note += " / 히든 후보 포함";
        }
      }

      // 다양한 조합을 만들기 위해 시작 구간을 랜덤하게 밀어줌
      const offset = Math.floor(Math.random() * Math.max(1, Math.min(pool.length, 30)));
      pool = [...pool.slice(offset), ...pool.slice(0, offset)];

      for (const p of pool) {
        const bodyRetail = retailSumOf(bodyItemsOf(items));
        if (items.length >= targetItemCount && bodyRetail >= retailTarget) break;
        if (used.has(p.id)) continue;
        items.push({ ...p, _tag: manualType === "프리미엄박스" ? "본품" : "본품" });
        used.add(p.id);
      }

      let guard = 0;
      while (retailSumOf(bodyItemsOf(items)) < retailTarget && guard < 50) {
        const gap = retailTarget - retailSumOf(bodyItemsOf(items));
        const add = poolRaw
          .filter((p) => !used.has(p.id))
          .sort((a, b) => {
            const aScore = Math.abs(toInt(a.retail) - gap) - toInt(a.stock) * 20 + Math.random() * 3000;
            const bScore = Math.abs(toInt(b.retail) - gap) - toInt(b.stock) * 20 + Math.random() * 3000;
            return aScore - bScore;
          })[0];
        if (!add) break;
        items.push({ ...add, _tag: "본품 보정" });
        used.add(add.id);
        guard += 1;
      }

      if (manualType === "소확행") {
        gift = giftPool
          .filter((p) => !used.has(p.id))
          .sort((a, b) => toInt(b.stock) - toInt(a.stock) || Math.random() - 0.5)[0];
        if (gift) {
          items.push({ ...gift, _tag: "랜덤선물" });
          used.add(gift.id);
          note += ` / 랜덤선물 추가: ${gift.name}`;
        } else {
          note += " / 랜덤선물 후보 없음";
        }
      }

      let fin = calcFinance(items, saleTotal, fee);

      // 마진이 너무 높으면 원가가 더 높은 상품으로 교체
      guard = 0;
      while (fin.margin > maxMargin && guard < 60) {
        const candidates = items
          .map((p, idx) => ({ p, idx }))
          .filter(({ p }) => p._tag !== "랜덤선물")
          .sort((a, b) => toInt(a.p.wholesale) - toInt(b.p.wholesale));

        let replaced = false;
        for (const { p: oldItem, idx } of candidates) {
          const replacement = poolRaw
            .filter((p) => !items.some((x, j) => j !== idx && x.id === p.id))
            .filter((p) => toInt(p.wholesale) > toInt(oldItem.wholesale))
            .sort((a, b) => {
              const aFin = calcFinance([...items.slice(0, idx), { ...a, _tag: oldItem._tag || "마진상한보정" }, ...items.slice(idx + 1)], saleTotal, fee);
              const bFin = calcFinance([...items.slice(0, idx), { ...b, _tag: oldItem._tag || "마진상한보정" }, ...items.slice(idx + 1)], saleTotal, fee);
              return Math.abs(aFin.margin - targetMargin) - Math.abs(bFin.margin - targetMargin) + (Math.random() - 0.5) * 2;
            })[0];

          if (replacement) {
            items[idx] = { ...replacement, _tag: oldItem._tag || "마진상한보정" };
            fin = calcFinance(items, saleTotal, fee);
            replaced = true;
            break;
          }
        }
        if (!replaced) break;
        guard += 1;
      }

      // 마진이 너무 낮으면 원가 낮은 유사 소비자가 상품으로 교체
      guard = 0;
      while (fin.margin < minMargin && guard < 60) {
        const candidates = items
          .map((p, idx) => ({ p, idx }))
          .filter(({ p }) => p._tag !== "랜덤선물")
          .sort((a, b) => toInt(b.p.wholesale) - toInt(a.p.wholesale));

        let replaced = false;
        for (const { p: oldItem, idx } of candidates) {
          const replacement = poolRaw
            .filter((p) => !items.some((x, j) => j !== idx && x.id === p.id))
            .filter((p) => toInt(p.wholesale) < toInt(oldItem.wholesale))
            .sort((a, b) => {
              const aFin = calcFinance([...items.slice(0, idx), { ...a, _tag: oldItem._tag || "마진하한보정" }, ...items.slice(idx + 1)], saleTotal, fee);
              const bFin = calcFinance([...items.slice(0, idx), { ...b, _tag: oldItem._tag || "마진하한보정" }, ...items.slice(idx + 1)], saleTotal, fee);
              return Math.abs(aFin.margin - targetMargin) - Math.abs(bFin.margin - targetMargin) + (Math.random() - 0.5) * 2;
            })[0];

          if (replacement) {
            items[idx] = { ...replacement, _tag: oldItem._tag || "마진하한보정" };
            fin = calcFinance(items, saleTotal, fee);
            replaced = true;
            break;
          }
        }
        if (!replaced) break;
        guard += 1;
      }

      const bodyItems = bodyItemsOf(items);
      const bodyRetailSum = retailSumOf(bodyItems);
      const ids = bodyItems.map((p) => p.id).sort((a, b) => a - b).join("-");
      if (signatures.has(ids)) continue;

      const chars = Array.from(new Set(items.flatMap((p) => splitMultiValues(p.char2)))).slice(0, 8).join(", ");

      if (bodyRetailSum >= retailTarget && fin.margin >= minMargin && fin.margin <= maxMargin && !isTooSimilarToExisting(items, recs, 0.65)) {
        signatures.add(ids);
        recs.push({
          name: `추천안${recs.length + 1}`,
          type: manualType,
          boxCount,
          saleTotal,
          feeRate: fee,
          items,
          finance: fin,
          chars,
          note,
          retailGap: retailTarget - bodyRetailSum,
          bodyRetailSum,
          retailTarget,
          marginRangeText: `${minMargin}%~${maxMargin}%`,
          giftName: gift?.name || "",
          diversityText: recs.length === 0 ? "첫 추천안" : `겹침 최대 ${(Math.max(...recs.map((r) => overlapRate(items, r.items || [])), 0) * 100).toFixed(0)}%`,
        });
      }
    }

    if (recs.length === 0) {
      alert(`조건에 맞는 추천안을 만들지 못했어요.\\n마진 허용범위 ${minMargin}%~${maxMargin}% / 본품 소비자가 목표 ${money(retailTarget)} 조건을 만족하는 조합이 부족해요.`);
      setManualRecommendations([]);
      setSelectedManualIndex(null);
      return;
    }

    setManualRecommendations(recs);
    setSelectedManualIndex(0);
  }

  function applyManualRecommendation() {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    setComposeItems(rec.items);
    setSalePrice(String(rec.saleTotal));
    setFeeRate(String(rec.feeRate));
    setActiveTab("수동박스");
  }

  function showCharacterBoxCapacity() {
    const all = products.filter((p) => toInt(p.stock) > 0);
    if (all.length === 0) return alert("계산할 재고가 없어요.");

    const sale = toInt(salePrice || defaultSale);
    const fee = Number(feeRate || defaultFee || 0);
    const targetMargin = Number(manualTargetMargin || 0);
    const minMargin = manualType === "히든박스" ? Math.max(0, targetMargin - Number(manualHiddenDiscount || 0)) : targetMargin;
    const maxMargin = targetMargin + 5;
    const itemCountText = window.prompt("박스 1개당 대략 상품 개수를 입력해줘.\\n예: 6", String(targetItemCountByStyle()));
    if (itemCountText === null) return;
    const itemCount = Math.max(1, toInt(itemCountText));

    const charMap = {};
    all.forEach((p) => {
      const chars = splitMultiValues(p.char2 || p.char1 || "미분류");
      chars.forEach((c) => {
        if (!charMap[c]) charMap[c] = [];
        charMap[c].push(p);
      });
    });

    const rows = Object.entries(charMap).map(([char, rows]) => {
      const totalStock = rows.reduce((s, p) => s + toInt(p.stock), 0);
      const sorted = [...rows].sort((a, b) => {
        const aScore = Math.abs(toInt(a.retail) - sale / itemCount) + Math.abs(toInt(a.wholesale) - (sale * (1 - targetMargin / 100)) / itemCount);
        const bScore = Math.abs(toInt(b.retail) - sale / itemCount) + Math.abs(toInt(b.wholesale) - (sale * (1 - targetMargin / 100)) / itemCount);
        return aScore - bScore;
      });

      let possible = 0;
      const maxByStock = Math.floor(totalStock / itemCount);
      for (let box = 1; box <= maxByStock; box++) {
        let virtual = [];
        const stockLeft = {};
        sorted.forEach((p) => { stockLeft[p.id] = toInt(p.stock); });

        for (let i = 0; i < itemCount; i++) {
          const pick = sorted.find((p) => stockLeft[p.id] > 0);
          if (!pick) break;
          virtual.push(pick);
          stockLeft[pick.id] -= 1;
        }

        const fin = calcFinance(virtual, sale, fee);
        if (virtual.length === itemCount && fin.retailSum >= sale && fin.margin >= minMargin && fin.margin <= maxMargin) {
          possible = Math.max(possible, box);
          sorted.forEach((p) => { stockLeft[p.id] = Math.max(0, stockLeft[p.id] - 1); });
        } else {
          // 단순 재고 기반 대략 계산이므로 조건 실패 시 재고 기반 최소값까지만 표시
          possible = Math.min(possible || maxByStock, maxByStock);
          break;
        }
      }

      const avgRetail = rows.reduce((s, p) => s + toInt(p.retail), 0) / Math.max(1, rows.length);
      const avgWholesale = rows.reduce((s, p) => s + toInt(p.wholesale), 0) / Math.max(1, rows.length);
      return {
        char,
        itemKinds: rows.length,
        totalStock,
        avgRetail: Math.round(avgRetail),
        avgWholesale: Math.round(avgWholesale),
        possible: Math.max(0, possible),
        maxByStock,
      };
    }).sort((a, b) => b.possible - a.possible || b.totalStock - a.totalStock);

    const text = rows.map((r) =>
      `${r.char} | 재고 ${r.totalStock}개 | 상품종류 ${r.itemKinds}종 | 평균소비자가 ${money(r.avgRetail)} | 평균도매가 ${money(r.avgWholesale)} | 대략 가능 ${r.possible}박스 (재고상 최대 ${r.maxByStock}박스)`
    ).join("\\n");

    alert(
      `계산 기준: 판매가 ${money(sale)} / 수수료율 ${fee}% / 마진 ${minMargin}%~${maxMargin}% / 박스당 ${itemCount}개\\n\\n` +
      (text || "계산할 캐릭터가 없어요.")
    );
  }

  function manualGapRecommendations() {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    const gap = rec.saleTotal - rec.finance.retailSum;
    const pool = baseManualCandidates()
      .filter((p) => !rec.items.some((x) => x.id === p.id))
      .sort((a, b) => Math.abs(toInt(a.retail) - gap) - Math.abs(toInt(b.retail) - gap))
      .slice(0, 20);
    alert(pool.length ? pool.map((p) => `${p.name} | 소비자가 ${money(p.retail)} | 도매가 ${money(p.wholesale)}`).join("\n") : "추천할 상품이 없어요.");
  }

  function addManualGapItem() {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    const gap = rec.saleTotal - rec.finance.retailSum;
    const add = baseManualCandidates()
      .filter((p) => !rec.items.some((x) => x.id === p.id))
      .sort((a, b) => Math.abs(toInt(a.retail) - gap) - Math.abs(toInt(b.retail) - gap))[0];
    if (!add) return alert("추가할 상품이 없어요.");
    const next = [...manualRecommendations];
    const items = [...rec.items, add];
    next[selectedManualIndex] = { ...rec, items, finance: calcFinance(items, rec.saleTotal, rec.feeRate), retailGap: rec.saleTotal - items.reduce((s,p)=>s+toInt(p.retail),0), note: rec.note + " / 부족금액 추가" };
    setManualRecommendations(next);
  }

  function addSelectedProductToManualRecommendation() {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    const product = products.find((p) => p.id === selectedProductId);
    if (!product) return alert("조건 상품 리스트에서 추가할 상품을 선택해줘.");
    const items = [...rec.items, product];
    const next = [...manualRecommendations];
    next[selectedManualIndex] = {
      ...rec,
      items,
      finance: calcFinance(items, rec.saleTotal, rec.feeRate),
      retailGap: rec.saleTotal - items.reduce((s, p) => s + toInt(p.retail), 0),
      note: rec.note + " / 수동상품 추가",
    };
    setManualRecommendations(next);
  }

  function removeManualRecommendationItem(index) {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    const items = rec.items.filter((_, i) => i !== index);
    const next = [...manualRecommendations];
    next[selectedManualIndex] = {
      ...rec,
      items,
      finance: calcFinance(items, rec.saleTotal, rec.feeRate),
      retailGap: rec.saleTotal - items.reduce((s, p) => s + toInt(p.retail), 0),
      note: rec.note + " / 상품 삭제 수정",
    };
    setManualRecommendations(next);
  }

  async function createOrderFromManualRecommendation() {
    const rec = manualRecommendations[selectedManualIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    await createOrderFromItems(rec.items, manualCustomer, manualMemo, manualReorder, rec.saleTotal, rec.feeRate);
  }

  function scoopCandidateProductsWithExcluded() {
    const v48StrictScoopChars = v48ScoopStrictCharsOnly ? v48SelectedChars(scoopChar1Selected, scoopChar2Selected) : [];
    const available = [];
    const excluded = [];
    products.forEach((p) => {
      if (toInt(p.stock) <= 0) return;
      if (!valueMatchesSelected(p.char1, scoopChar1Selected)) return;
      if (!valueMatchesSelected(p.char2, scoopChar2Selected)) return;
      if (!inPriceRange(p.retail, scoopPrice)) return;
      if (scoopRetailLimit && toInt(p.retail) >= toInt(scoopRetailLimit)) {
        excluded.push(p);
        return;
      }
      available.push(p);
    });
    return { available, excluded };
  }

  function scoopCandidateProducts() {
    return scoopCandidateProductsWithExcluded().available;
  }

  function buildCategoryStats(pool) {
    const map = {};
    pool.forEach((p) => {
      const cat = p.category || "미분류";
      if (!map[cat]) map[cat] = { category: cat, products: [], stockSum: 0, retailSum: 0, count: 0 };
      map[cat].products.push(p);
      map[cat].stockSum += toInt(p.stock);
      map[cat].retailSum += toInt(p.retail);
      map[cat].count += 1;
    });
    return Object.values(map).map((g) => ({
      ...g,
      avgRetail: Math.round(g.retailSum / Math.max(1, g.count)),
    })).sort((a, b) => b.stockSum - a.stockSum);
  }

  function makeManualCategoryGroup() {
    if (scoopSelectedCategories.length < 2) return alert("묶을 카테고리를 2개 이상 선택해줘.");
    const pool = scoopCandidateProducts();
    const selectedProducts = pool.filter((p) => scoopSelectedCategories.includes(p.category || "미분류"));
    if (selectedProducts.length === 0) return alert("선택 카테고리에 해당하는 상품이 없어요.");

    const groupName = window.prompt("그룹명을 입력해줘.", scoopSelectedCategories.join(", "));
    if (!groupName) return;

    const group = {
      id: Date.now(),
      name: groupName,
      categories: [...scoopSelectedCategories],
      products: selectedProducts,
      stock: selectedProducts.reduce((s, p) => s + toInt(p.stock), 0),
      avgRetail: Math.round(selectedProducts.reduce((s, p) => s + toInt(p.retail), 0) / Math.max(1, selectedProducts.length)),
      partQty: 1,
      partPercent: 0,
      reason: "사용자 수동 묶기",
    };

    setScoopGroups((prev) => {
      const without = prev.filter((g) => !g.categories.some((c) => scoopSelectedCategories.includes(c)));
      const next = [...without, group].map((g, i) => ({ ...g, id: i + 1 }));
      const totalStock = next.reduce((sum, g) => sum + toInt(g.stock), 0);
      return next.map((g) => {
        const percent = totalStock > 0 ? (toInt(g.stock) / totalStock) * 100 : 0;
        return { ...g, partPercent: Number(percent.toFixed(1)), partQty: g.partQty || Math.max(1, Math.round(percent / 20)) };
      });
    });
    setScoopSelectedCategories([]);
  }

  function toggleScoopCategory(category) {
    setScoopSelectedCategories((prev) => prev.includes(category) ? prev.filter((c) => c !== category) : [...prev, category]);
  }

  function makeAutoScoopGroups(stats) {
    if (!stats.length) return [];
    const avgStock = stats.reduce((s, x) => s + x.stockSum, 0) / Math.max(1, stats.length);
    const threshold = Math.max(1, Math.floor(avgStock * 0.5));
    let large = stats.filter((s) => s.stockSum > threshold);
    let small = stats.filter((s) => s.stockSum <= threshold);

    if (large.length === 0) {
      large = stats.slice(0, 1);
      small = stats.slice(1);
    }

    const groups = large.map((s, i) => ({
      id: i + 1,
      name: `그룹 ${String.fromCharCode(65 + i)}`,
      categories: [s.category],
      products: [...s.products],
      stock: s.stockSum,
      avgRetail: s.avgRetail,
      partQty: null,
      reason: "기본 그룹",
    }));

    small.forEach((s) => {
      const idx = groups.reduce((best, g, i) => g.stock < groups[best].stock ? i : best, 0);
      const baseName = groups[idx].categories[0];
      groups[idx].categories.push(s.category);
      groups[idx].products.push(...s.products);
      groups[idx].stock += s.stockSum;
      groups[idx].avgRetail = Math.round(groups[idx].products.reduce((sum, p) => sum + toInt(p.retail), 0) / Math.max(1, groups[idx].products.length));
      groups[idx].reason = `${s.category}: 재고 부족으로 ${baseName} 그룹에 합쳐짐`;
    });

    const totalStock = groups.reduce((sum, g) => sum + toInt(g.stock), 0);
    groups.forEach((g, i) => {
      const percent = totalStock > 0 ? (toInt(g.stock) / totalStock) * 100 : 0;
      g.partPercent = Number(percent.toFixed(1));
      g.partQty = Math.max(1, Math.round(percent / 20));
      g.id = i + 1;
    });
    return groups;
  }

  function buildScoopAnalysisText(stats, groups, excludedCount) {
    if (!stats.length) return "조건에 맞는 상품이 없습니다.";
    const statLines = stats.map((s) => `${s.category}: 재고 ${s.stockSum} / 평균소비자가 ${money(s.avgRetail)} / 상품수 ${s.count}`);
    const groupLines = groups.map((g) => `${g.name}: ${g.categories.join(", ")} / 재고 ${g.stock} / 평균 ${money(g.avgRetail)} / 파츠 ${g.partQty}개(추천 ${g.partPercent || 0}%) / ${g.reason}`);
    return [
      `사용 가능 상품 수 ${stats.reduce((sum, s) => sum + s.count, 0)}개 | 배제된 상품 수 ${excludedCount}개`,
      "",
      "[카테고리 목록]",
      ...statLines,
      "",
      "[자동 그룹 제안]",
      ...groupLines,
    ].join("\n");
  }

  function analyzeScoopCategories() {
    const { available, excluded } = scoopCandidateProductsWithExcluded();
    const stats = buildCategoryStats(available);
    const groups = makeAutoScoopGroups(stats);
    setScoopCategoryStats(stats);
    setScoopExcludedCount(excluded.length);
    setScoopAnalysisText(buildScoopAnalysisText(stats, groups, excluded.length));

    if (!stats.length) {
      setScoopGroups([]);
      setScoopRecommendations([]);
      return alert("조건에 맞는 상품이 없어요.");
    }

    const merged = groups.some((g) => g.categories.length > 1);
    const ok = window.confirm(
      `${merged ? "재고가 부족한 카테고리는 자동으로 합쳐졌어요.\n" : ""}` +
      `카테고리 ${stats.length}개를 ${groups.length}개 그룹으로 제안합니다.\n\n` +
      groups.map((g) => `${g.name}: ${g.categories.join(", ")} / 파츠 ${g.partQty}개`).join("\n") +
      "\n\n이 그룹으로 적용할까요?"
    );

    if (ok) {
      setScoopGroups(groups);
      setScoopRecommendations([]);
    }
  }

  function generateScoopGroups() {
    const count = Math.max(1, toInt(scoopGroupCount));
    const { available, excluded } = scoopCandidateProductsWithExcluded();
    const pool = available;
    if (pool.length === 0) return alert("그룹을 만들 후보 상품이 없어요.");

    const groups = Array.from({ length: count }, (_, i) => ({
      id: i + 1,
      name: `그룹 ${String.fromCharCode(65 + i)}`,
      products: [],
      stock: 0,
      avgRetail: 0,
      partQty: 1,
      partPercent: 0,
      categories: [],
      reason: scoopMode,
    }));

    const sorted = [...pool].sort((a, b) => {
      if (scoopMode === "소비자가 균등") return toInt(b.retail) - toInt(a.retail);
      if (scoopMode === "도매가 균등") return toInt(b.wholesale) - toInt(a.wholesale);
      if (scoopMode === "혼합 균형") return (toInt(b.retail) + toInt(b.wholesale)) - (toInt(a.retail) + toInt(a.wholesale));
      return 0.5 - Math.random();
    });

    sorted.forEach((p, idx) => {
      const target = groups[idx % groups.length];
      target.products.push(p);
      target.stock += toInt(p.stock);
      if (p.category && !target.categories.includes(p.category)) target.categories.push(p.category);
    });

    const totalStock = groups.reduce((sum, g) => sum + toInt(g.stock), 0);
    groups.forEach((g) => {
      g.avgRetail = Math.round(g.products.reduce((s, p) => s + toInt(p.retail), 0) / Math.max(1, g.products.length));
      const percent = totalStock > 0 ? (toInt(g.stock) / totalStock) * 100 : 0;
      g.partPercent = Number(percent.toFixed(1));
      g.partQty = Math.max(1, Math.round(percent / 20));
    });

    const stats = buildCategoryStats(pool);
    setScoopCategoryStats(stats);
    setScoopExcludedCount(excluded.length);
    setScoopAnalysisText(buildScoopAnalysisText(stats, groups, excluded.length));
    setScoopGroups(groups);
    setScoopRecommendations([]);
  }

  function updateScoopPartQty(groupId, qty) {
    setScoopGroups((prev) => prev.map((g) => g.id === groupId ? { ...g, partQty: Math.max(0, toInt(qty)) } : g));
  }

  function renameScoopGroup(groupId) {
    const g = scoopGroups.find((x) => x.id === groupId);
    if (!g) return;
    const name = window.prompt("그룹명 수정", g.name);
    if (!name) return;
    setScoopGroups((prev) => prev.map((x) => x.id === groupId ? { ...x, name } : x));
  }

  async function saveScoopGroups() {
    if (scoopGroups.length === 0) return alert("저장할 그룹이 없어요.");
    const name = window.prompt("저장할 그룹 이름", `스쿱그룹 ${nowString()}`);
    if (!name) return;
    const { error } = await supabase.from("saved_scoop_groups").insert([{ name, data: JSON.stringify(scoopGroups) }]);
    if (error) return alert("그룹 저장 실패: " + error.message);
    alert("그룹 저장 완료");
  }

  async function loadScoopGroups() {
    const { data, error } = await supabase.from("saved_scoop_groups").select("*").order("id", { ascending: false });
    if (error) return alert("그룹 불러오기 실패: " + error.message);
    if (!data?.length) return alert("저장된 그룹이 없어요.");
    const msg = data.map((g) => `${g.id}: ${g.name}`).join("\n");
    const id = window.prompt(`불러올 그룹 ID 입력\n\n${msg}`);
    if (!id) return;
    const row = data.find((g) => String(g.id) === String(id));
    if (!row) return alert("해당 ID를 찾지 못했어요.");
    try {
      setScoopGroups(JSON.parse(row.data || "[]"));
      alert("그룹을 불러왔어요.");
    } catch {
      alert("저장 데이터가 깨졌어요.");
    }
  }


  function selectedScoopPrefChars() {
    return Array.from(new Set([...(scoopChar1Selected || []), ...(scoopChar2Selected || [])])).filter(Boolean);
  }

  function checkScoopPreferredCharStock() {
    const selected = selectedScoopPrefChars();
    if (selected.length === 0) return true;

    const rows = selected.map((char) => {
      const stock = products
        .filter((p) => productCharacters(p).includes(char))
        .reduce((s, p) => s + toInt(p.stock), 0);
      return { char, stock };
    });

    const low = rows.filter((x) => x.stock <= 3);
    if (low.length === 0) return true;

    return window.confirm(
      "선택한 선호 캐릭터 중 재고가 적어서 추천안에 충분히 반영되지 않을 수 있어요.\\n\\n" +
      low.map((x) => `${x.char}: 재고 ${x.stock}개`).join("\\n") +
      "\\n\\n재고가 부족한 경우 다른 캐릭터가 섞일 수 있습니다. 그래도 추천안을 만들까요?"
    );
  }

  function scoopPreferenceNote(items) {
    const selected = selectedScoopPrefChars();
    if (selected.length === 0) return "선호 캐릭터 미선택";
    const included = Array.from(new Set((items || []).flatMap((p) => productCharacters(p))));
    const reflected = selected.filter((c) => included.includes(c));
    const missing = selected.filter((c) => !included.includes(c));
    const others = included.filter((c) => !selected.includes(c));

    const parts = [];
    parts.push(reflected.length ? `선호 반영: ${reflected.join(", ")}` : "선호 반영 없음");
    if (missing.length) parts.push(`부족/미반영: ${missing.join(", ")}`);
    if (others.length) parts.push(`다른 캐릭터 섞임: ${others.slice(0, 6).join(", ")}`);
    return parts.join(" / ");
  }

  function generateScoopRecommendations() {
    if (scoopGroups.length === 0) return alert("먼저 그룹을 만들어줘.");
    if (!checkScoopPreferredCharStock()) return;

    const sale = toInt(salePrice || defaultSale);
    const fee = Number(feeRate || defaultFee || 0);
    const targetMargin = Number(scoopTargetMargin || 0);
    const minMargin = targetMargin;
    const maxMargin = targetMargin + 5;
    const allPool = scoopCandidateProducts();
    const preferredSet = new Set(selectedScoopPrefChars());
    const preferredScoopPool = [...scoopCandidateProducts()].sort((a, b) => {
      const aPref = productCharacters(a).some((c) => preferredSet.has(c)) ? 1 : 0;
      const bPref = productCharacters(b).some((c) => preferredSet.has(c)) ? 1 : 0;
      return bPref - aPref || toInt(b.stock) - toInt(a.stock);
    });
    const recs = [];
    let attempts = 0;

    while (recs.length < 12 && attempts < 600) {
      attempts += 1;
      let items = [];
      const used = new Set();
      let type = "기본";

      scoopGroups.forEach((g) => {
        let pool = [...(g.products || [])].filter((p) => toInt(p.stock) > 0 && !used.has(p.id));
        if (scoopMode === "소비자가 균등") pool.sort((a, b) => toInt(b.retail) - toInt(a.retail));
        else if (scoopMode === "도매가 균등") pool.sort((a, b) => toInt(a.wholesale) - toInt(b.wholesale));
        else if (scoopMode === "혼합 균형") pool.sort((a, b) => (toInt(b.retail) - toInt(b.wholesale)) - (toInt(a.retail) - toInt(a.wholesale)));
        else pool.sort(() => 0.5 - Math.random());

        const qty = Math.max(1, toInt(g.partQty || 1));
        pool.slice(0, qty).forEach((p) => {
          if (!used.has(p.id)) {
            items.push({ ...p, _groupName: g.name, _tag: "기본" });
            used.add(p.id);
          }
        });
      });

      let fin = calcFinance(items, sale, fee);

      let guard = 0;
      while (fin.retailSum < sale && guard < 30) {
        const gap = sale - fin.retailSum;
        const add = findClosestRetailProduct(allPool, gap, used);
        if (!add) break;
        items.push({ ...add, _tag: "소비자가 보정" });
        used.add(add.id);
        fin = calcFinance(items, sale, fee);
        type = "부분 업그레이드";
        guard += 1;
      }

      guard = 0;
      while (fin.margin > maxMargin && guard < 40) {
        const candidates = items
          .map((p, idx) => ({ p, idx }))
          .sort((a, b) => toInt(a.p.wholesale) - toInt(b.p.wholesale));

        let replaced = false;
        for (const { p: oldItem, idx } of candidates) {
          const replacement = allPool
            .filter((p) => !items.some((x, j) => j !== idx && x.id === p.id))
            .filter((p) => toInt(p.wholesale) > toInt(oldItem.wholesale))
            .sort((a, b) => {
              const aFin = calcFinance([...items.slice(0, idx), { ...a, _tag: "마진상한보정" }, ...items.slice(idx + 1)], sale, fee);
              const bFin = calcFinance([...items.slice(0, idx), { ...b, _tag: "마진상한보정" }, ...items.slice(idx + 1)], sale, fee);
              return Math.abs(aFin.margin - targetMargin) - Math.abs(bFin.margin - targetMargin);
            })[0];

          if (replacement) {
            items[idx] = { ...replacement, _tag: "마진상한보정" };
            fin = calcFinance(items, sale, fee);
            type = "부분 업그레이드";
            replaced = true;
            break;
          }
        }
        if (!replaced) break;
        guard += 1;
      }

      guard = 0;
      while (fin.margin < minMargin && guard < 40) {
        const candidates = items
          .map((p, idx) => ({ p, idx }))
          .sort((a, b) => toInt(b.p.wholesale) - toInt(a.p.wholesale));

        let replaced = false;
        for (const { p: oldItem, idx } of candidates) {
          const replacement = allPool
            .filter((p) => !items.some((x, j) => j !== idx && x.id === p.id))
            .filter((p) => toInt(p.wholesale) < toInt(oldItem.wholesale))
            .sort((a, b) => {
              const aFin = calcFinance([...items.slice(0, idx), { ...a, _tag: "마진하한보정" }, ...items.slice(idx + 1)], sale, fee);
              const bFin = calcFinance([...items.slice(0, idx), { ...b, _tag: "마진하한보정" }, ...items.slice(idx + 1)], sale, fee);
              return Math.abs(aFin.margin - targetMargin) - Math.abs(bFin.margin - targetMargin);
            })[0];

          if (replacement) {
            items[idx] = { ...replacement, _tag: "마진하한보정" };
            fin = calcFinance(items, sale, fee);
            type = "전체 업그레이드";
            replaced = true;
            break;
          }
        }
        if (!replaced) break;
        guard += 1;
      }

      const chars = Array.from(new Set(items.flatMap((p) => splitMultiValues(p.char2)))).slice(0, 10).join(", ");

      if (fin.retailSum >= sale && isWithinMargin(fin, targetMargin)) {
        recs.push({ name: `추천안${recs.length + 1}`, type, items, finance: fin, chars });
      }
    }

    let out = recs;
    if (scoopRecType === "기본만") out = recs.filter((r) => r.type === "기본");
    if (scoopRecType === "부분 업그레이드만") out = recs.filter((r) => r.type === "부분 업그레이드");
    if (scoopRecType === "전체 업그레이드만") out = recs.filter((r) => r.type === "전체 업그레이드");

    if (scoopRecSort === "수량 적은 순") out = [...out].sort((a, b) => a.items.length - b.items.length);
    if (scoopRecSort === "수량 많은 순") out = [...out].sort((a, b) => b.items.length - a.items.length);
    if (scoopRecSort === "마진율 높은 순") out = [...out].sort((a, b) => b.finance.margin - a.finance.margin);
    if (scoopRecSort === "소비자가 높은 순") out = [...out].sort((a, b) => b.finance.retailSum - a.finance.retailSum);

    if (out.length === 0) {
      alert(`조건에 맞는 추천안이 없어요.\n소비자가합 ${money(sale)} 이상, 마진율 ${minMargin}%~${maxMargin}% 범위로 만들 수 있는 조합이 부족합니다.`);
      setScoopRecommendations([]);
      setSelectedScoopIndex(null);
      return;
    }

    setScoopRecommendations(out);
    setSelectedScoopIndex(0);
  }

  function sendScoopToCompose() {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    setComposeItems(rec.items);
    setActiveTab("수동박스");
  }

  function replaceScoopItem(index) {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return;
    const old = rec.items[index];
    const pool = scoopCandidateProducts()
      .filter((p) => !rec.items.some((x) => x.id === p.id))
      .sort((a, b) => Math.abs(toInt(a.retail) - toInt(old.retail)) - Math.abs(toInt(b.retail) - toInt(old.retail)));
    const msg = pool.slice(0, 30).map((p) => `${p.id}: ${p.name} | 도매가 ${money(p.wholesale)} | 소비자가 ${money(p.retail)} | 카테고리 ${p.category}`).join("\n");
    const id = window.prompt(`교체할 상품 ID 입력\n\n${msg}`);
    if (!id) return;
    const picked = pool.find((p) => String(p.id) === String(id));
    if (!picked) return alert("상품 ID를 찾지 못했어요.");
    const next = [...scoopRecommendations];
    const items = [...rec.items];
    items[index] = picked;
    next[selectedScoopIndex] = { ...rec, items, finance: calcFinance(items, salePrice, feeRate) };
    setScoopRecommendations(next);
  }

  function removeScoopItem(index) {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return;
    const next = [...scoopRecommendations];
    const items = rec.items.filter((_, i) => i !== index);
    next[selectedScoopIndex] = { ...rec, items, finance: calcFinance(items, salePrice, feeRate) };
    setScoopRecommendations(next);
  }

  function showScoopGapRecommendations() {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    const gap = Math.max(0, toInt(salePrice) - rec.finance.retailSum);
    const recChars = Array.from(new Set((rec.items || []).flatMap((p) => productCharacters(p))));

    let pool = scoopCandidateProducts().filter((p) => !rec.items.some((x) => x.id === p.id));

    if (scoopGapScope === "same") {
      pool = pool.filter((p) => hasSharedCharacter(recChars, p));
    }

    pool = pool
      .sort((a, b) => Math.abs(toInt(a.retail) - gap) - Math.abs(toInt(b.retail) - gap))
      .slice(0, 30);

    const title = scoopGapScope === "same" ? "같은 캐릭터 상품" : "모든 캐릭터 상품";
    alert(pool.length ? `[${title}]\\n부족금액: ${money(gap)}\\n\\n` + pool.map((p) => `${p.id}: ${p.name} | ${p.char1}/${p.char2} | 소비자가 ${money(p.retail)} | 도매가 ${money(p.wholesale)}`).join("\\n") : "추천할 상품이 없어요.");
  }

  async function createOrderFromScoop() {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return alert("추천안을 선택해줘.");
    await createOrderFromItems(rec.items, scoopCustomer, scoopMemo, scoopReorder, salePrice, feeRate);
  }


  function getStockZeroNamesForItems(items) {
    const count = {};
    items.forEach((p) => { count[p.id] = (count[p.id] || 0) + 1; });
    return Object.entries(count)
      .map(([id, qty]) => {
        const p = products.find((x) => String(x.id) === String(id));
        if (!p) return null;
        return toInt(p.stock) - qty === 0 ? p.name : null;
      })
      .filter(Boolean);
  }

  function recommendationCheckText(rec, type = "수동") {
    if (!rec) return "-";
    const sale = toInt(rec.saleTotal || salePrice || defaultSale);
    const bodyRetail = rec.bodyRetailSum || rec.finance?.retailSum || 0;
    const retailTarget = rec.retailTarget || sale;
    const targetMargin = Number(manualTargetMargin || scoopTargetMargin || 0);
    const min = rec.marginRangeText ? Number(String(rec.marginRangeText).split("%")[0]) : targetMargin;
    const max = rec.marginRangeText ? Number(String(rec.marginRangeText).split("~")[1]?.replace("%", "")) || targetMargin + 5 : targetMargin + 5;
    const zeroNames = getStockZeroNamesForItems(rec.items || []);
    return buildRecommendationCheck({
      type: rec.type || type,
      saleTotal: sale,
      retailTarget,
      bodyRetailSum: bodyRetail,
      totalRetailSum: rec.finance?.retailSum || 0,
      margin: rec.finance?.margin || 0,
      minMargin: min,
      maxMargin: max,
      giftName: rec.giftName || "",
      zeroStockNames: zeroNames,
    });
  }

  function FilterBox() {
    return (
      <>
        <div className="filterRow">
          <label>상품명</label>
          <input id="manual-product-search-input" name="manual-product-search" value={searchInput} onChange={(e) => setSearchInput(e.target.value)} placeholder="상품명 검색" autoComplete="off" onKeyDown={(e) => { if (!e.nativeEvent?.isComposing && e.key === "Enter") { e.preventDefault(); runManualProductSearch(); } }} />
        <button type="button" onClick={runManualProductSearch}>검색</button>
        <button type="button" onClick={clearManualProductSearch}>검색초기화</button>
          <MultiCheckFilter label="캐릭터1" options={char1Options} selected={char1Selected} setSelected={setChar1Selected} />
          <MultiCheckFilter label="캐릭터2" options={char2Options} selected={char2Selected} setSelected={setChar2Selected} />
          <label>카테고리</label>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)}>{categoryOptions.map((v) => <option key={v}>{v}</option>)}</select>
          <label>가격대</label>
          <select value={priceFilter} onChange={(e) => setPriceFilter(e.target.value)}>{PRICE_RANGES.map((v) => <option key={v}>{v}</option>)}</select>
          <label>정렬</label>
          <select value={productSort} onChange={(e) => setProductSort(e.target.value)}>
            <option>기본순</option>
            <option>도매가 낮은순</option>
            <option>도매가 높은순</option>
            <option>소비자가 낮은순</option>
            <option>소비자가 높은순</option>
            <option>재고 많은순</option>
            <option>재고 적은순</option>
            <option>상품명순</option>
          </select>
          <button onClick={resetFilters}>초기화</button>
        </div>
        <label className="checkLine"><input checked={excludeLowStock} onChange={(e) => setExcludeLowStock(e.target.checked)} type="checkbox" /> 재고 1개 제외</label>
        <p className="statusLine">조회 결과: {filteredProducts.length.toLocaleString()}종 / 재고 {filteredProducts.reduce((s, p) => s + toInt(p.stock), 0).toLocaleString()}개</p>
      </>
    );
  }

  function isRecentProduct(p) {
    const raw = p.updated_at || p.created_at || p.inserted_at || p.createdAt || p.date;
    if (!raw) return false;
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return false;
    return Date.now() - d.getTime() <= 3 * 24 * 60 * 60 * 1000;
  }

  function ProductTable({ mode }) {
    return (
      <div className="tableWrap">
        <table className="productTable">
          <thead><tr><th>ID</th><th>상품명</th><th>캐릭터1</th><th>캐릭터2</th><th>카테고리</th><th>재고</th><th>도매가</th><th>소비자가</th><th>{mode === "compose" ? "추가" : "삭제"}</th></tr></thead>
          <tbody>
            {filteredProducts.map((p) => (
              <tr key={p.id} onClick={(e) => { if (["INPUT", "BUTTON", "SELECT", "TEXTAREA"].includes(e.target.tagName)) return; preserveManualProductListScroll(() => setSelectedProductId(p.id)); }} className={selectedProductId === p.id ? "selectedRow" : ""} title={p.name}>
                <td>{p.id}</td><td><span className="productNameCell">{p.name}{isRecentProduct(p) && <span className="newBadge">NEW</span>}</span></td><td>{p.char1}</td><td>{p.char2}</td><td>{p.category}</td><td>{p.stock}</td><td>{money(p.wholesale)}</td><td>{money(p.retail)}</td>
                <td>{mode === "compose" ? <button onClick={(e) => { e.stopPropagation(); addToCompose(p); }}>추가</button> : <div className="bulkActionCell"><input type="checkbox" checked={bulkSelectedProductIds.map(String).includes(String(p.id))} onClick={(e) => { e.stopPropagation(); v59ToggleBulkProduct(p.id); }} onChange={() => {}} /><button onClick={(e) => { e.stopPropagation(); startEditProductById(p.id); }}>수정</button><button className="deleteBtn" onClick={(e) => { e.stopPropagation(); deleteProduct(p.id); }}>삭제</button></div>}</td>
              </tr>
            ))}
            {filteredProducts.length === 0 && <tr><td colSpan="11" className="empty">등록된 상품이 없어요.</td></tr>}
          </tbody>
        </table>
      </div>
    );
  }


  const v48CurrentInventoryCost = useMemo(() => {
    return products.reduce((sum, p) => sum + productWholesaleValue(p) * toInt(p.stock), 0);
  }, [products]);

  const v48MaterialsTotal = useMemo(() => {
    return materials.reduce((sum, m) => sum + toInt(m.amount || m.price || m.cost || 0), 0);
  }, [materials]);

  const v48TotalInvestment = useMemo(() => {
    // 전체 투자원금: 가능한 경우 initial_stock/original_stock를 우선 사용하고,
    // 없으면 현재 stock 기준으로 계산합니다. 매출/순이익으로 차감하지 않습니다.
    const stockCost = products.reduce((sum, p) => {
      const baseQty = toInt(p.initial_stock || p.initialStock || p.original_stock || p.stock);
      return sum + toInt(p.wholesale) * baseQty;
    }, 0);
    return stockCost + v48MaterialsTotal;
  }, [products, v48MaterialsTotal]);


  const v50SoldItemsWholesaleTotal = useMemo(() => {
    return orderItems.reduce((sum, x) => sum + toInt(x.wholesale || x.wholesale_price || x.cost || 0) * toInt(x.qty || 1), 0);
  }, [orderItems]);

  const v50SoldItemsRetailTotal = useMemo(() => {
    return orderItems.reduce((sum, x) => sum + toInt(x.retail || x.retail_price || x.consumer_price || 0) * toInt(x.qty || 1), 0);
  }, [orderItems]);

  const v50TotalInvestment = v48CurrentInventoryCost + v50SoldItemsWholesaleTotal + v48MaterialsTotal;


  const materialTotalForProfit = useMemo(() => {
    return materials.reduce((sum, m) => sum + toInt(m.amount || m.price || m.cost || 0), 0);
  }, [materials]);

  const shippedOrdersForProfit = useMemo(() => {
    return orders.filter((o) => String(o.status || "").includes("출고"));
  }, [orders]);

  const shippedWholesaleTotal = useMemo(() => {
    const shippedIds = new Set(shippedOrdersForProfit.map((o) => String(o.id)));
    return orderItems
      .filter((x) => shippedIds.has(String(x.order_id)))
      .reduce((sum, x) => sum + toInt(x.wholesale || x.wholesale_price || x.cost || 0) * toInt(x.qty || 1), 0);
  }, [orderItems, shippedOrdersForProfit]);

  const realNetProfit = useMemo(() => {
    const received = shippedOrdersForProfit.reduce((sum, o) => sum + toInt(o.net_amount || o.netAmount || o.received || o.real_amount || o.sale_price || o.price || 0), 0);
    return received - shippedWholesaleTotal - materialTotalForProfit;
  }, [shippedOrdersForProfit, shippedWholesaleTotal, materialTotalForProfit]);

  const pendingShippingCount = useMemo(() => {
    return orders.filter((o) => {
      const s = String(o.status || "");
      return s.includes("주문접수") || s.includes("재고임시차감");
    }).length;
  }, [orders]);

  useEffect(() => {
    setTodoItems((prev) => prev.filter((t) => !(t.done && t.doneAt && Date.now() - t.doneAt > 12 * 60 * 60 * 1000)));
  }, [orders.length, activeTab]);

  useEffect(() => {
    setTodoItems((prev) => {
      const autoId = "auto-pending-shipping";
      const text = `주문건 ${pendingShippingCount}개 우체국 출고하기`;
      const exists = prev.find((t) => t.id === autoId);
      if (pendingShippingCount > 0) {
        if (exists) return prev.map((t) => t.id === autoId ? { ...t, text, done: false, doneAt: null, auto: true } : t);
        return [{ id: autoId, text, done: false, doneAt: null, auto: true }, ...prev];
      }
      if (exists && !exists.done) return prev.map((t) => t.id === autoId ? { ...t, done: true, doneAt: Date.now() } : t);
      return prev;
    });
  }, [pendingShippingCount]);

  function addTodoItem(text) {
    const clean = String(text || "").trim();
    if (!clean) return;
    setTodoItems((prev) => [{ id: `todo-${Date.now()}`, text: clean, done: false, doneAt: null, auto: false }, ...prev]);
    setNewTodoText("");
    setQuickTodo("");
  }

  function toggleTodoItem(id) {
    setTodoItems((prev) => prev.map((t) => t.id === id ? { ...t, done: !t.done, doneAt: !t.done ? Date.now() : null } : t));
  }

  function deleteTodoItem(id) {
    setTodoItems((prev) => prev.filter((t) => t.id !== id));
  }

  function orderDateString(o) {
    return String(o.created_at || o.order_date || o.date || "").slice(0, 10);
  }

  const financeYears = useMemo(() => {
    const years = new Set();
    orders.forEach((o) => {
      const y = orderDateString(o).slice(0, 4);
      if (y) years.add(y);
    });
    years.add(String(new Date().getFullYear()));
    return Array.from(years).sort().reverse();
  }, [orders]);

  function financeRows() {
    const monthly = {};
    orders.forEach((o) => {
      const date = orderDateString(o);
      const year = date.slice(0, 4) || financeYear;
      const month = date.slice(5, 7) || "01";
      if (String(year) !== String(financeYear)) return;
      if (financeMonth !== "전체" && month !== financeMonth) return;
      const key = `${year}-${month}`;
      if (!monthly[key]) monthly[key] = { period: key, orderCount: 0, sales: 0, netSales: 0, productCost: 0, materialCost: 0, realProfit: 0 };
      const statusText = String(o.status || "");
      if (!statusText.includes("출고")) return;

      monthly[key].orderCount += 1;
      monthly[key].sales += toInt(o.sale_price || o.price || o.total_price || o.sales || 0);
      monthly[key].netSales += toInt(o.net_amount || o.netAmount || o.received || o.real_amount || o.sale_price || o.price || 0);
      const items = orderItems.filter((x) => String(x.order_id) === String(o.id));
      monthly[key].productCost += items.reduce((sum, x) => sum + toInt(x.wholesale || x.wholesale_price || x.cost || 0) * toInt(x.qty || 1), 0);
    });

    materials.forEach((m) => {
      const date = String(m.created_at || m.date || nowString()).slice(0, 10);
      const year = date.slice(0, 4) || financeYear;
      const month = date.slice(5, 7) || "01";
      if (String(year) !== String(financeYear)) return;
      if (financeMonth !== "전체" && month !== financeMonth) return;
      const key = `${year}-${month}`;
      if (!monthly[key]) monthly[key] = { period: key, orderCount: 0, sales: 0, netSales: 0, productCost: 0, materialCost: 0, realProfit: 0 };
      monthly[key].materialCost += toInt(m.amount || m.price || m.cost || 0);
    });

    return Object.values(monthly).sort((a, b) => a.period.localeCompare(b.period)).map((r) => ({
      ...r,
      realProfit: r.netSales - r.productCost - r.materialCost,
    }));
  }

  function downloadFinanceExcel() {
    const rows = financeRows();
    if (rows.length === 0) return alert("다운로드할 매입매출 데이터가 없어요.");
    const data = rows.map((r) => ({
      기간: r.period,
      주문수: r.orderCount,
      총매출: r.sales,
      실수령액: r.netSales,
      상품매입가: r.productCost,
      재료비: r.materialCost,
      진짜순이익: r.realProfit,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "매입매출현황");
    XLSX.writeFile(wb, `매입매출현황_${financeYear}_${financeMonth}.xlsx`);
  }

  function DashboardPage() {
    return (
      <>
        <div className="v49InvestmentSummary">
          <div><b>총 투자원금</b><span>{money(v50TotalInvestment)}</span></div>
          <div><b>현재 남은 재고 매입가</b><span>{money(v48CurrentInventoryCost)}</span></div>
          <p>총 투자원금은 현재 남은 재고 매입가 + 출고/주문상품 도매가합 + 재료비합입니다. 매출/순이익으로 차감하지 않습니다.</p>
        </div>
        <section className="cards">
          <div className="card"><span>상품종류</span><strong>{products.length.toLocaleString()}</strong></div>
          <div className="card"><span>재고수량</span><strong>{totalStock.toLocaleString()}</strong></div>
          <div className="card"><span>총매입가격</span><strong>{money(totalWholesale)}</strong></div>
          <div className="card"><span>출고완료 주문수</span><strong>{orders.filter((o) => o.status !== "취소").length.toLocaleString()}</strong></div>
          <div className="card"><span>총매출</span><strong>{money(totalSales)}</strong></div>
          <div className="card"><span>실수령액</span><strong>{money(totalNet)}</strong></div>
          <div className="card"><span>순이익</span><strong>{money(totalProfit)}</strong></div>
          <div className="card"><span>재료비</span><strong>{money(totalMaterials)}</strong></div>
        </section>
        
        <section className="panel dashboardTodoPanel">
          <h2>오늘 일정표</h2>
          <p className="statusLine">완료 체크한 할 일은 12시간 뒤 자동 삭제됩니다. 주문접수/재고임시차감 건은 자동으로 출고 할 일에 표시됩니다.</p>
          <div className="filterRow">
            <input
              placeholder="할 일 직접 입력"
              value={newTodoText}
              onChange={(e) => setNewTodoText(e.target.value)}
              onKeyDown={(e) => { if (!e.nativeEvent?.isComposing && e.key === "Enter") addTodoItem(newTodoText); }}
            />
            <button type="button" onClick={() => addTodoItem(newTodoText)}>할 일 추가</button>
            <select value={quickTodo} onChange={(e) => { setQuickTodo(e.target.value); if (e.target.value) addTodoItem(e.target.value); }}>
              <option value="">자주 하는 업무 추가</option>
              <option value="매입정리">매입정리</option>
              <option value="매출정리">매출정리</option>
              <option value="창고정리">창고정리</option>
              <option value="고객응대">고객응대</option>
              <option value="영상업로드">영상업로드</option>
              <option value="사입주문">사입주문</option>
              <option value="주문포장">주문포장</option>
            </select>
          </div>
          <div className="todoList">
            {todoItems.length === 0 && <div className="empty">오늘 할 일이 없어요.</div>}
            {todoItems.map((t) => (
              <div key={t.id} className={`todoItem ${t.done ? "done" : ""}`}>
                <label><input type="checkbox" checked={t.done} onChange={() => toggleTodoItem(t.id)} /><span>{t.text}</span></label>
                <button type="button" onClick={() => deleteTodoItem(t.id)}>삭제</button>
              </div>
            ))}
          </div>
        </section>

<section className="panel">
          <h2>재료비 관리</h2>
          <div className="filterRow">
            <label>재료비명</label><input value={materialName} onChange={(e) => setMaterialName(e.target.value)} />
            <label>금액</label><input value={materialAmount} onChange={(e) => setMaterialAmount(e.target.value)} type="number" />
            <button onClick={addMaterial}>저장</button>
            <button className="deleteBtn" onClick={deleteMaterial}>삭제</button>
          </div>
          <div className="tableWrap smallTable">
            <table><thead><tr><th>ID</th><th>재료비명</th><th>금액</th></tr></thead><tbody>
              {materials.map((m) => <tr key={m.id} onClick={() => setSelectedMaterialId(m.id)} className={selectedMaterialId === m.id ? "selectedRow" : ""}><td>{m.id}</td><td>{m.name}</td><td>{money(m.amount)}</td></tr>)}
              {materials.length === 0 && <tr><td colSpan="3" className="empty">등록된 재료비가 없어요.</td></tr>}
            </tbody></table>
          </div>
        </section>
      </>
    );
  }


  function startEditSelectedProduct() {
    const p = products.find((x) => String(x.id) === String(selectedProductId));
    if (!p) return alert("수정할 상품을 선택해줘.");
    setEditProductForm({
      id: p.id,
      name: p.name || "",
      char1: p.char1 || "",
      char2: p.char2 || "",
      category: p.category || "",
      stock: String(toInt(p.stock)),
      wholesale: String(toInt(p.wholesale)),
      retail: String(toInt(p.retail)),
    });
  }

  function cancelEditProduct() {
    setEditProductForm(null);
  }

  async function saveEditedProduct() {
    if (!editProductForm?.id) return alert("수정할 상품이 없어요.");

    const payload = {
      name: editProductForm.name || "",
      char1: editProductForm.char1 || "",
      char2: editProductForm.char2 || "",
      category: editProductForm.category || "",
      stock: toInt(editProductForm.stock),
      wholesale: toInt(editProductForm.wholesale),
      retail: toInt(editProductForm.retail),
    };

    const { error } = await supabase.from("products").update(payload).eq("id", editProductForm.id);
    if (error) return alert("재고 상품 수정 실패: " + error.message);

    const linked = orderItems.filter((x) => String(x.product_id) === String(editProductForm.id));
    if (linked.length > 0) {
      const ok = window.confirm(
        `이 상품이 기존 주문/출고 상품목록 ${linked.length}건에 포함되어 있어요.\n\n` +
        "해당 주문/출고건의 상품명, 캐릭터1, 캐릭터2, 카테고리, 도매가, 소비자가에도 수정사항을 적용할까요?"
      );

      if (ok) {
        const { error: itemError } = await supabase
          .from("order_items")
          .update({
            name: payload.name,
            product_name: payload.name,
            char1: payload.char1,
            char2: payload.char2,
            category: payload.category,
            wholesale: payload.wholesale,
            retail: payload.retail,
            wholesale_price: payload.wholesale,
            retail_price: payload.retail,
          })
          .eq("product_id", editProductForm.id);
        if (itemError) return alert("주문/출고 상품목록 반영 실패: " + itemError.message);
      }
    }

    alert("상품 수정 완료!");
    setEditProductForm(null);
    getProducts();
    getOrderItems();
  }


  function startEditProductById(productId) {
    const p = products.find((x) => String(x.id) === String(productId));
    if (!p) return alert("수정할 상품을 선택해줘.");
    setSelectedProductId(p.id);
    setEditProductForm({
      id: p.id,
      name: p.name || "",
      char1: p.char1 || "",
      char2: p.char2 || "",
      category: p.category || "",
      stock: String(toInt(p.stock)),
      wholesale: String(toInt(p.wholesale)),
      retail: String(toInt(p.retail)),
    });
  }

  function startEditSelectedProduct() {
    if (!selectedProductId) return alert("수정할 상품을 선택해줘.");
    startEditProductById(selectedProductId);
  }

  function cancelEditProduct() {
    setEditProductForm(null);
  }

  async function saveEditedProduct() {
    if (!editProductForm?.id) return alert("수정할 상품이 없어요.");

    const payload = {
      name: editProductForm.name || "",
      char1: editProductForm.char1 || "",
      char2: editProductForm.char2 || "",
      category: editProductForm.category || "",
      stock: toInt(editProductForm.stock),
      wholesale: toInt(editProductForm.wholesale),
      retail: toInt(editProductForm.retail),
    };

    const { error } = await supabase.from("products").update(payload).eq("id", editProductForm.id);
    if (error) return alert("재고 상품 수정 실패: " + error.message);

    const linked = orderItems.filter((x) => String(x.product_id) === String(editProductForm.id));
    if (linked.length > 0) {
      const ok = window.confirm(
        `이 상품이 기존 주문/출고 상품목록 ${linked.length}건에 포함되어 있어요.\n\n` +
        "해당 주문/출고건의 상품명, 캐릭터1, 캐릭터2, 카테고리, 도매가, 소비자가에도 수정사항을 적용할까요?\n\n" +
        "확인 = 기존 주문/출고건에도 반영\n취소 = 재고관리 상품만 수정"
      );

      if (ok) {
        const { error: itemError } = await supabase
          .from("order_items")
          .update({
            name: payload.name,
            product_name: payload.name,
            char1: payload.char1,
            char2: payload.char2,
            category: payload.category,
            wholesale: payload.wholesale,
            retail: payload.retail,
            wholesale_price: payload.wholesale,
            retail_price: payload.retail,
          })
          .eq
  function v59ToggleBulkProduct(productId) {
    setBulkSelectedProductIds((prev) => {
      const id = String(productId);
      return prev.map(String).includes(id) ? prev.filter((x) => String(x) !== id) : [...prev, productId];
    });
  }

  function v59ClearBulkProducts() {
    setBulkSelectedProductIds([]);
    setBulkEditForm(null);
  }

  function v59StartBulkEditProducts() {
    if (bulkSelectedProductIds.length === 0) return alert("일괄 수정할 상품을 체크해줘.");
    setBulkEditForm({ name: "", char1: "", char2: "", category: "", stock: "", wholesale: "", retail: "", hidden: "" });
  }

  async function v59SaveBulkEditedProducts() {
    if (!bulkEditForm || bulkSelectedProductIds.length === 0) return alert("일괄 수정할 상품이 없어요.");
    const payload = {};
    if (bulkEditForm.name.trim()) payload.name = bulkEditForm.name.trim();
    if (bulkEditForm.char1.trim()) payload.char1 = bulkEditForm.char1.trim();
    if (bulkEditForm.char2.trim()) payload.char2 = bulkEditForm.char2.trim();
    if (bulkEditForm.category.trim()) payload.category = bulkEditForm.category.trim();
    if (bulkEditForm.stock !== "") payload.stock = toInt(bulkEditForm.stock);
    if (bulkEditForm.wholesale !== "") {
      payload.wholesale = toInt(bulkEditForm.wholesale);
      payload.wholesale_price = toInt(bulkEditForm.wholesale);
    }
    if (bulkEditForm.retail !== "") {
      payload.retail = toInt(bulkEditForm.retail);
      payload.retail_price = toInt(bulkEditForm.retail);
      payload.consumer_price = toInt(bulkEditForm.retail);
    }
    if (Object.keys(payload).length === 0) return alert("변경할 내용을 하나 이상 입력해줘.");
    if (!window.confirm(`${bulkSelectedProductIds.length}개 상품을 일괄 수정할까요?\n빈칸은 수정하지 않습니다.`)) return;

    const { error } = await supabase.from("products").update(payload).in("id", bulkSelectedProductIds);
    if (error) return alert("상품 일괄 수정 실패: " + error.message);

    const linked = orderItems.filter((x) => bulkSelectedProductIds.map(String).includes(String(x.product_id)));
    if (linked.length > 0 && window.confirm(`선택한 상품이 기존 주문/출고 상품목록 ${linked.length}건에 포함되어 있어요.\n주문/출고건에도 이번 일괄 수정사항을 반영할까요?`)) {
      for (const productId of bulkSelectedProductIds) {
        const latest = { ...products.find((p) => String(p.id) === String(productId)), ...payload };
        const itemPayload = {};
        if (payload.name !== undefined) { itemPayload.name = latest.name; itemPayload.product_name = latest.name; }
        if (payload.char1 !== undefined) itemPayload.char1 = latest.char1;
        if (payload.char2 !== undefined) itemPayload.char2 = latest.char2;
        if (payload.category !== undefined) itemPayload.category = latest.category;
        if (payload.wholesale !== undefined || payload.wholesale_price !== undefined) {
          itemPayload.wholesale = productWholesaleValue(latest);
          itemPayload.wholesale_price = productWholesaleValue(latest);
        }
        if (payload.retail !== undefined || payload.retail_price !== undefined || payload.consumer_price !== undefined) {
          itemPayload.retail = productRetailValue(latest);
          itemPayload.retail_price = productRetailValue(latest);
          itemPayload.consumer_price = productRetailValue(latest);
        }
        if (Object.keys(itemPayload).length > 0) {
          const { error: itemError } = await supabase.from("order_items").update(itemPayload).eq("product_id", productId);
          if (itemError) return alert("주문/출고 상품목록 일괄 반영 실패: " + itemError.message);
        }
      }
    }
    alert("상품 일괄 수정 완료!");
    v59ClearBulkProducts();
    getProducts();
    getOrderItems();
  }

("product_id", editProductForm.id);

        if (itemError) return alert("주문/출고 상품목록 반영 실패: " + itemError.message);
      }
    }

    alert("상품 수정 완료!");
    setEditProductForm(null);
    getProducts();
    getOrderItems();
  }


  function runInventorySearchOnEnter(e) {
    if (!e.nativeEvent?.isComposing && e.key === "Enter") {
      e.preventDefault();
      getProducts();
    }
  }


  function v59ToggleBulkProduct(productId) {
    setBulkSelectedProductIds((prev) => {
      const id = String(productId);
      return prev.map(String).includes(id)
        ? prev.filter((x) => String(x) !== id)
        : [...prev, productId];
    });
  }

  function v59ClearBulkProducts() {
    setBulkSelectedProductIds([]);
    setBulkEditForm(null);
  }

  function v59StartBulkEditProducts() {
    if (!bulkSelectedProductIds || bulkSelectedProductIds.length === 0) {
      return alert("일괄 수정할 상품을 체크해줘.");
    }
    setBulkEditForm({
      name: "",
      char1: "",
      char2: "",
      category: "",
      stock: "",
      wholesale: "",
      retail: "",
    });
  }

  async function v59SaveBulkEditedProducts() {
    if (!bulkEditForm || bulkSelectedProductIds.length === 0) {
      return alert("일괄 수정할 상품이 없어요.");
    }

    const payload = {};
    if (bulkEditForm.name.trim()) payload.name = bulkEditForm.name.trim();
    if (bulkEditForm.char1.trim()) payload.char1 = bulkEditForm.char1.trim();
    if (bulkEditForm.char2.trim()) payload.char2 = bulkEditForm.char2.trim();
    if (bulkEditForm.category.trim()) payload.category = bulkEditForm.category.trim();
    if (bulkEditForm.stock !== "") payload.stock = toInt(bulkEditForm.stock);
    if (bulkEditForm.wholesale !== "") {
      payload.wholesale = toInt(bulkEditForm.wholesale);
      payload.wholesale_price = toInt(bulkEditForm.wholesale);
    }
    if (bulkEditForm.retail !== "") {
      payload.retail = toInt(bulkEditForm.retail);
      payload.retail_price = toInt(bulkEditForm.retail);
      payload.consumer_price = toInt(bulkEditForm.retail);
    }

    if (Object.keys(payload).length === 0) return alert("변경할 내용을 하나 이상 입력해줘.");
    if (!window.confirm(`${bulkSelectedProductIds.length}개 상품을 일괄 수정할까요?\n\n빈칸은 수정하지 않습니다.`)) return;

    const { error } = await supabase.from("products").update(payload).in("id", bulkSelectedProductIds);
    if (error) return alert("상품 일괄 수정 실패: " + error.message);

    const linked = orderItems.filter((x) => bulkSelectedProductIds.map(String).includes(String(x.product_id)));
    if (linked.length > 0) {
      const apply = window.confirm(
        `선택한 상품이 기존 주문/출고 상품목록 ${linked.length}건에 포함되어 있어요.\n\n` +
        "주문/출고건에도 이번 일괄 수정사항을 반영할까요?"
      );

      if (apply) {
        for (const productId of bulkSelectedProductIds) {
          const latest = { ...products.find((p) => String(p.id) === String(productId)), ...payload };
          const itemPayload = {};
          if (payload.name !== undefined) {
            itemPayload.name = latest.name;
            itemPayload.product_name = latest.name;
          }
          if (payload.char1 !== undefined) itemPayload.char1 = latest.char1;
          if (payload.char2 !== undefined) itemPayload.char2 = latest.char2;
          if (payload.category !== undefined) itemPayload.category = latest.category;
          if (payload.wholesale !== undefined || payload.wholesale_price !== undefined) {
            itemPayload.wholesale = productWholesaleValue(latest);
            itemPayload.wholesale_price = productWholesaleValue(latest);
          }
          if (payload.retail !== undefined || payload.retail_price !== undefined || payload.consumer_price !== undefined) {
            itemPayload.retail = productRetailValue(latest);
            itemPayload.retail_price = productRetailValue(latest);
            itemPayload.consumer_price = productRetailValue(latest);
          }
          if (Object.keys(itemPayload).length > 0) {
            const { error: itemError } = await supabase.from("order_items").update(itemPayload).eq("product_id", productId);
            if (itemError) return alert("주문/출고 상품목록 일괄 반영 실패: " + itemError.message);
          }
        }
      }
    }

    alert("상품 일괄 수정 완료!");
    v59ClearBulkProducts();
    getProducts();
    getOrderItems();
  }


  function FinanceReportPage() {
    const rows = financeRows();
    const totals = rows.reduce((acc, r) => {
      acc.orderCount += r.orderCount;
      acc.sales += r.sales;
      acc.netSales += r.netSales;
      acc.productCost += r.productCost;
      acc.materialCost += r.materialCost;
      acc.realProfit += r.realProfit;
      return acc;
    }, { orderCount: 0, sales: 0, netSales: 0, productCost: 0, materialCost: 0, realProfit: 0 });

    return (
      <>
        <section className="panel financeReportPage">
          <h2>매입매출현황</h2>
          <p className="statusLine">월별 매출/실수령액/상품매입가/재료비/진짜 순이익을 정리합니다. 세금신고 전에는 증빙자료와 홈택스 자료를 함께 확인해주세요.</p>
          <div className="filterRow">
            <label>연도</label>
            <select value={financeYear} onChange={(e) => setFinanceYear(e.target.value)}>
              {financeYears.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
            <label>월</label>
            <select value={financeMonth} onChange={(e) => setFinanceMonth(e.target.value)}>
              <option value="전체">전체</option>
              {Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, "0")).map((m) => <option key={m} value={m}>{m}월</option>)}
            </select>
            <button type="button" onClick={downloadFinanceExcel}>엑셀 다운로드</button>
          </div>
          <div className="summaryGrid financeSummary">
            <div><b>주문수</b><span>{totals.orderCount.toLocaleString()}건</span></div>
            <div><b>총매출</b><span>{money(totals.sales)}</span></div>
            <div><b>실수령액</b><span>{money(totals.netSales)}</span></div>
            <div><b>상품매입가</b><span>{money(totals.productCost)}</span></div>
            <div><b>재료비</b><span>{money(totals.materialCost)}</span></div>
            <div><b>진짜 순이익</b><span>{money(totals.realProfit)}</span></div>
          </div>
        </section>
        <section className="panel financeTablePanel">
          <h3>월별 현황표</h3>
          <div className="tableWrap">
            <table>
              <thead><tr><th>기간</th><th>주문수</th><th>총매출</th><th>실수령액</th><th>상품매입가</th><th>재료비</th><th>진짜 순이익</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.period}>
                    <td>{r.period}</td><td>{r.orderCount}</td><td>{money(r.sales)}</td><td>{money(r.netSales)}</td><td>{money(r.productCost)}</td><td>{money(r.materialCost)}</td><td>{money(r.realProfit)}</td>
                  </tr>
                ))}
                {rows.length === 0 && <tr><td colSpan="7" className="empty">해당 기간 데이터가 없어요.</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
      </>
    );
  }

  function extractJsonArrayFromText(text) {
    const raw = String(text || "");
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
      if (Array.isArray(parsed.items)) return parsed.items;
      if (Array.isArray(parsed.products)) return parsed.products;
    } catch (_) {}

    const block = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (block) {
      try {
        const parsed = JSON.parse(block[1]);
        if (Array.isArray(parsed)) return parsed;
        if (Array.isArray(parsed.items)) return parsed.items;
        if (Array.isArray(parsed.products)) return parsed.products;
      } catch (_) {}
    }

    const first = raw.indexOf("[");
    const last = raw.lastIndexOf("]");
    if (first !== -1 && last !== -1 && last > first) {
      try {
        const parsed = JSON.parse(raw.slice(first, last + 1));
        return Array.isArray(parsed) ? parsed : [];
      } catch (_) {}
    }
    return [];
  }

  function normalizeAiImportRow(row) {
    return {
      name: String(row.name || row.product_name || row.상품명 || row.품명 || row.item || "").trim(),
      char1: String(row.char1 || row.character1 || row.캐릭터1 || row.브랜드 || row.대표캐릭터 || "").trim(),
      char2: String(row.char2 || row.character2 || row.캐릭터2 || row.캐릭터 || row.상세캐릭터 || "").trim(),
      category: String(row.category || row.카테고리 || row.분류 || "").trim(),
      stock: toInt(row.stock || row.qty || row.quantity || row.수량 || row.재고 || 1),
      wholesale: toInt(row.wholesale || row.cost || row.도매가 || row.매입가 || row.원가 || 0),
      retail: toInt(row.retail || row.consumer_price || row.소비자가 || row.판매가 || row.정가 || 0),
      hidden: false,
    };
  }

  function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = String(reader.result || "");
        resolve(result.includes(",") ? result.split(",")[1] : result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }


  async function analyzeInventoryFile(fileOrEvent, options = {}) {
    const filesFromEvent = Array.from(fileOrEvent?.target?.files || []);
    if (filesFromEvent.length > 1) return analyzeInventoryFiles(filesFromEvent);
    const file = filesFromEvent[0] || fileOrEvent;
    if (fileOrEvent?.target) fileOrEvent.target.value = "";
    if (!file) return;

    const append = options.append !== false;
    setAiImportFileName((prev) => prev ? `${prev}, ${file.name || "붙여넣은 파일"}` : (file.name || "붙여넣은 이미지/파일"));
    setAiImportLoading(true);
    if (!append) {
      setAiImportRawText("");
      setAiImportRows([]);
    }

    try {
      const base64 = await readFileAsBase64(file);
      const prompt = `거래명세서, 영수증, 주문내역서 또는 상품 입고 이미지에서 재고 등록에 필요한 상품을 추출해줘.
반드시 JSON 배열만 반환해줘. 설명문 금지.
각 항목 필드: name, char1, char2, category, stock, wholesale, retail.
모르는 값은 빈 문자열 또는 0으로 둬.
상품명은 실제 판매/재고 관리에 쓸 수 있게 정리하고, 수량/도매가/소비자가가 보이면 숫자로 넣어줘.`;

      const res = await fetch("/api/gemini", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: prompt,
          imageBase64: base64,
          mimeType: file.type || "image/png",
          task: "inventory_image_parse",
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.message || data?.error?.message || data?.error || "AI 이미지 분석 실패");
      const text = data.text || "";
      setAiImportRawText((prev) => [prev, `\n\n--- ${file.name || "파일"} ---\n${text}`].filter(Boolean).join(""));
      const rows = extractJsonArrayFromText(text).map(normalizeAiImportRow).filter((x) => x.name);
      setAiImportRows((prev) => append ? [...prev, ...rows] : rows);
      if (rows.length === 0) alert(`${file.name || "파일"} 분석은 완료됐지만 상품 JSON을 찾지 못했어요. 원문 결과를 확인해줘.`);
    } catch (err) {
      alert("이미지 분석 실패: " + (err?.message || err));
    } finally {
      setAiImportLoading(false);
    }
  }

  async function analyzeInventoryFiles(files) {
    const valid = Array.from(files || []).filter((file) => file && (file.type?.startsWith("image/") || file.type === "application/pdf" || /\.pdf$/i.test(file.name || "")));
    if (valid.length === 0) return alert("분석할 이미지/PDF 파일이 없어요.");
    setAiImportFileName("");
    setAiImportRawText("");
    setAiImportRows([]);
    for (const file of valid) {
      await analyzeInventoryFile(file, { append: true });
    }
  }

  async function analyzeInventoryImageFile(e) {
    const files = Array.from(e?.target?.files || []);
    if (files.length > 1) return analyzeInventoryFiles(files);
    return analyzeInventoryFile(e, { append: false });
  }

  async function handleAiImportPaste(e) {
    const clipboard = e.clipboardData;
    if (!clipboard) return;

    const files = Array.from(clipboard.files || []).filter((file) => file.type?.startsWith("image/") || file.type === "application/pdf");
    if (files.length > 0) {
      e.preventDefault();
      await analyzeInventoryFiles(files);
      return;
    }

    const itemFiles = [];
    const items = Array.from(clipboard.items || []);
    for (const item of items) {
      if (item.kind === "file") {
        const file = item.getAsFile();
        if (file && (file.type?.startsWith("image/") || file.type === "application/pdf")) itemFiles.push(file);
      }
    }
    if (itemFiles.length > 0) {
      e.preventDefault();
      await analyzeInventoryFiles(itemFiles);
      return;
    }

    const text = clipboard.getData("text/plain");
    if (text && text.trim()) {
      e.preventDefault();
      setAiImportPasteText((prev) => prev ? `${prev}\n\n${text}` : text);
      await analyzeInventoryText(text, { append: true, label: "붙여넣은 텍스트" });
    }
  }

  async function analyzeInventoryText(text, options = {}) {
    const append = options.append !== false;
    setAiImportFileName((prev) => prev ? `${prev}, ${options.label || "붙여넣은 텍스트"}` : (options.label || "붙여넣은 텍스트"));
    setAiImportLoading(true);
    if (!append) {
      setAiImportRawText("");
      setAiImportRows([]);
    }

    try {
      const prompt = `아래 붙여넣은 거래명세서/영수증/주문내역 텍스트에서 재고 등록에 필요한 상품을 추출해줘.
여러 건이 섞여 있어도 전부 추출해줘.
반드시 JSON 배열만 반환해줘. 설명문 금지.
각 항목 필드: name, char1, char2, category, stock, wholesale, retail.
모르는 값은 빈 문자열 또는 0으로 둬.
상품명은 실제 판매/재고 관리에 쓸 수 있게 정리하고, 수량/도매가/소비자가가 보이면 숫자로 넣어줘.

붙여넣은 내용:
${text}`;

      const res = await fetch("/api/gemini", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: prompt,
          task: "inventory_text_parse",
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.message || data?.error?.message || data?.error || "AI 텍스트 분석 실패");
      const resultText = data.text || "";
      setAiImportRawText((prev) => [prev, `\n\n--- ${options.label || "붙여넣은 텍스트"} ---\n${resultText}`].filter(Boolean).join(""));
      const rows = extractJsonArrayFromText(resultText).map(normalizeAiImportRow).filter((x) => x.name);
      setAiImportRows((prev) => append ? [...prev, ...rows] : rows);
      if (rows.length === 0) alert("분석은 완료됐지만 상품 JSON을 찾지 못했어요. 원문 결과를 확인해줘.");
    } catch (err) {
      alert("텍스트 분석 실패: " + (err?.message || err));
    } finally {
      setAiImportLoading(false);
    }
  }


  function updateAiImportRow(index, key, value) {
    setAiImportRows((prev) => prev.map((row, i) => i === index ? { ...row, [key]: value } : row));
  }

  function removeAiImportRow(index) {
    setAiImportRows((prev) => prev.filter((_, i) => i !== index));
  }

  async function applyAiInventoryRows() {
    const rows = aiImportRows.map(normalizeAiImportRow).filter((x) => x.name);
    if (rows.length === 0) return alert("반영할 상품이 없어요.");
    const ok = window.confirm(`${rows.length}개 상품을 재고에 ${aiImportMode === "add" ? "추가" : "업데이트/추가"}할까요?`);
    if (!ok) return;

    const backupOk = await createInventoryBackup("ai_image_import_before");
    if (!backupOk) return;

    if (aiImportMode === "add") {
      const { error } = await supabase.from("products").insert(rows);
      if (error) return alert("AI 입고 추가 실패: " + error.message);
    } else {
      for (const row of rows) {
        const existing = products.find((p) => String(p.name || "").trim() === row.name.trim());
        if (existing) {
          const { error } = await supabase.from("products").update({
            char1: row.char1 || existing.char1,
            char2: row.char2 || existing.char2,
            category: row.category || existing.category,
            stock: toInt(existing.stock) + toInt(row.stock),
            wholesale: row.wholesale || existing.wholesale,
            retail: row.retail || existing.retail,
          }).eq("id", existing.id);
          if (error) return alert("AI 입고 업데이트 실패: " + error.message);
        } else {
          const { error } = await supabase.from("products").insert([row]);
          if (error) return alert("AI 입고 추가 실패: " + error.message);
        }
      }
    }
    await writeAudit("ai_inventory_import", `rows=${rows.length} / file=${aiImportFileName}`);
    alert(`AI 입고 반영 완료!\n${rows.length}개 상품이 처리됐어요. 최근 3일 내 생성/수정된 상품은 NEW 배지로 표시됩니다.`);
    setAiImportRows([]);
    setAiImportRawText("");
    getProducts();
  }

  function AiImportPage() {
    return (
      <section className="panel aiImportPage" onPaste={handleAiImportPaste} tabIndex={0}>
        <div className="pageTitleRow">
          <div>
            <h2>AI 입고 분석</h2>
            <p className="statusLine">거래명세서, 영수증, 주문내역서, 상품 사진을 업로드하거나 Ctrl+V로 붙여넣으면 Gemini가 상품명/수량/캐릭터/가격을 표로 정리합니다. 반영 전 직접 수정할 수 있어요.</p>
          </div>
          <label className="uploadBtn aiUploadBtn">이미지/PDF 여러개 업로드<input type="file" accept="image/*,.pdf" multiple onChange={analyzeInventoryImageFile} /></label>
        </div>

        <div className="aiPasteGuide">
          이 영역을 한 번 클릭한 뒤 <b>Ctrl+V</b>를 누르면 캡처 이미지, 여러 파일, 여러 텍스트를 계속 누적 분석할 수 있어요.
        </div>

        <div className="aiBulkPasteBox">
          <label>여러 거래명세서/주문목록 복붙 분석</label>
          <textarea
            value={aiImportPasteText}
            onChange={(e) => setAiImportPasteText(e.target.value)}
            placeholder={"여기에 거래명세서, 사입내역, 주문목록을 여러 건 붙여넣어도 됩니다.\n분석 결과는 아래 표에 누적되고, 반영 전 직접 수정할 수 있어요."}
          />
          <div className="buttonRow">
            <button type="button" onClick={() => analyzeInventoryText(aiImportPasteText, { append: true, label: "복붙 입력" })} disabled={!aiImportPasteText.trim() || aiImportLoading}>복붙 내용 AI 분석</button>
            <button type="button" onClick={() => { setAiImportPasteText(""); setAiImportRows([]); setAiImportRawText(""); setAiImportFileName(""); }}>분석내용 초기화</button>
          </div>
        </div>

        <div className="aiImportToolbar">
          <span>파일: {aiImportFileName || "-"}</span>
          <label>반영방식</label>
          <select value={aiImportMode} onChange={(e) => setAiImportMode(e.target.value)}>
            <option value="add">그대로 새 상품 추가</option>
            <option value="merge">같은 상품명은 재고수량 더하기</option>
          </select>
          <button onClick={applyAiInventoryRows} disabled={aiImportRows.length === 0 || aiImportLoading}>분석 결과 재고 반영</button>
        </div>

        {aiImportLoading && <div className="loadingBox">AI가 이미지에서 상품 정보를 읽는 중이에요...</div>}

        <div className="tableWrap aiImportTableWrap">
          <table>
            <thead><tr><th>상품명</th><th>캐릭터1</th><th>캐릭터2</th><th>카테고리</th><th>수량</th><th>도매가</th><th>소비자가</th><th>삭제</th></tr></thead>
            <tbody>
              {aiImportRows.map((row, i) => (
                <tr key={i}>
                  <td><input value={row.name} onChange={(e) => updateAiImportRow(i, "name", e.target.value)} /></td>
                  <td><input value={row.char1} onChange={(e) => updateAiImportRow(i, "char1", e.target.value)} /></td>
                  <td><input value={row.char2} onChange={(e) => updateAiImportRow(i, "char2", e.target.value)} /></td>
                  <td><input value={row.category} onChange={(e) => updateAiImportRow(i, "category", e.target.value)} /></td>
                  <td><input value={row.stock} onChange={(e) => updateAiImportRow(i, "stock", e.target.value)} /></td>
                  <td><input value={row.wholesale} onChange={(e) => updateAiImportRow(i, "wholesale", e.target.value)} /></td>
                  <td><input value={row.retail} onChange={(e) => updateAiImportRow(i, "retail", e.target.value)} /></td>
                  <td><button className="deleteBtn" onClick={() => removeAiImportRow(i)}>삭제</button></td>
                </tr>
              ))}
              {aiImportRows.length === 0 && <tr><td colSpan="8" className="empty">이미지/PDF를 업로드하면 분석 결과가 여기에 표시됩니다.</td></tr>}
            </tbody>
          </table>
        </div>

        <details className="aiRawBox">
          <summary>AI 원문 결과 보기</summary>
          <pre>{aiImportRawText || "아직 분석 결과가 없어요."}</pre>
        </details>
      </section>
    );
  }

  function InventoryPage() {
    return (
      <>
        <section className="panel inventoryPageFixed stickyControlPanel">
          <div className="filterRow">
            <input value={productForm.name} onChange={(e) => setProductForm({ ...productForm, name: e.target.value })} placeholder="상품명" />
            <input value={productForm.char1} onChange={(e) => setProductForm({ ...productForm, char1: e.target.value })} placeholder="캐릭터1" />
            <input value={productForm.char2} onChange={(e) => setProductForm({ ...productForm, char2: e.target.value })} placeholder="캐릭터2" />
            <input value={productForm.category} onChange={(e) => setProductForm({ ...productForm, category: e.target.value })} placeholder="카테고리" />
            <input value={productForm.stock} onChange={(e) => setProductForm({ ...productForm, stock: e.target.value })} placeholder="재고" />
            <input value={productForm.wholesale} onChange={(e) => setProductForm({ ...productForm, wholesale: e.target.value })} placeholder="도매가" />
            <input value={productForm.retail} onChange={(e) => setProductForm({ ...productForm, retail: e.target.value })} placeholder="소비자가" />
            <button onClick={addProduct}>상품 저장</button>
          </div>

          <div className="buttonRow">
            <label className="uploadBtn">엑셀 불러오기<input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} /></label>
            <button onClick={downloadInventoryExcel}>현재 재고 엑셀</button>
            <button onClick={downloadCurrentInventoryBackupFile}>재고 백업 파일</button>
            <button onClick={() => createInventoryBackup("manual_backup").then((ok) => ok && alert("재고 백업 완료!"))}>DB 백업 저장</button>
            <button onClick={restoreLatestInventoryBackup}>최근 백업 복구</button>
            <button onClick={showBackupListAndRestore}>백업 목록 선택복구</button>
            <button onClick={showChar2Values}>캐릭터2 목록 확인</button>
            <button onClick={showCharacterShortage}>부족 캐릭터 보기</button>
            <button type="button" onClick={startEditSelectedProduct}>상품 수정</button>
            <button className="deleteBtn" onClick={() => deleteProduct(selectedProductId)}>상품 삭제</button>
          </div>
        </section>
        <section className="panel inventoryListPanel">
          {FilterBox()}

        
        {editProductForm && (
          <section className="panel v53EditProductPanel">
            <h3>선택 상품 수정</h3>
            <div className="filterRow">
              <label>상품명</label><input value={editProductForm.name} onChange={(e) => setEditProductForm({ ...editProductForm, name: e.target.value })} />
              <label>캐릭터1</label><input value={editProductForm.char1} onChange={(e) => setEditProductForm({ ...editProductForm, char1: e.target.value })} />
              <label>캐릭터2</label><input value={editProductForm.char2} onChange={(e) => setEditProductForm({ ...editProductForm, char2: e.target.value })} />
              <label>카테고리</label><input value={editProductForm.category} onChange={(e) => setEditProductForm({ ...editProductForm, category: e.target.value })} />
            </div>
            <div className="filterRow">
              <label>재고수량</label><input value={editProductForm.stock} onChange={(e) => setEditProductForm({ ...editProductForm, stock: e.target.value })} />
              <label>도매가</label><input value={editProductForm.wholesale} onChange={(e) => setEditProductForm({ ...editProductForm, wholesale: e.target.value })} />
              <label>소비자가</label><input value={editProductForm.retail} onChange={(e) => setEditProductForm({ ...editProductForm, retail: e.target.value })} />
            </div>
            <div className="buttonRow">
              <button type="button" onClick={saveEditedProduct}>수정완료</button>
              <button type="button" onClick={cancelEditProduct}>수정취소</button>
            </div>
            <p className="statusLine">수정완료 시 기존 주문/출고건에 포함된 상품이면 반영 여부를 물어봅니다.</p>
          </section>
        )}

        
        <section className="panel v54BulkEditPanel inventorySelectedPanel">
          <div className="buttonRow">
            <span className="statusLine">체크된 상품: {bulkSelectedProductIds.length}개</span>
            <button type="button" onClick={v59StartBulkEditProducts}>체크상품 일괄수정</button>
            <button type="button" onClick={v59ClearBulkProducts}>체크해제</button>
          </div>
          {bulkEditForm && (
            <>
              <div className="filterRow">
                <label>상품명</label><input placeholder="빈칸=유지" value={bulkEditForm.name} onChange={(e) => setBulkEditForm({ ...bulkEditForm, name: e.target.value })} />
                <label>캐릭터1</label><input placeholder="빈칸=유지" value={bulkEditForm.char1} onChange={(e) => setBulkEditForm({ ...bulkEditForm, char1: e.target.value })} />
                <label>캐릭터2</label><input placeholder="빈칸=유지" value={bulkEditForm.char2} onChange={(e) => setBulkEditForm({ ...bulkEditForm, char2: e.target.value })} />
                <label>카테고리</label><input placeholder="빈칸=유지" value={bulkEditForm.category} onChange={(e) => setBulkEditForm({ ...bulkEditForm, category: e.target.value })} />
              </div>
              <div className="filterRow">
                <label>재고수량</label><input placeholder="빈칸=유지" value={bulkEditForm.stock} onChange={(e) => setBulkEditForm({ ...bulkEditForm, stock: e.target.value })} />
                <label>도매가</label><input placeholder="빈칸=유지" value={bulkEditForm.wholesale} onChange={(e) => setBulkEditForm({ ...bulkEditForm, wholesale: e.target.value })} />
                <label>소비자가</label><input placeholder="빈칸=유지" value={bulkEditForm.retail} onChange={(e) => setBulkEditForm({ ...bulkEditForm, retail: e.target.value })} />
              </div>
              <div className="buttonRow">
                <button type="button" onClick={v59SaveBulkEditedProducts}>일괄 수정완료</button>
                <button type="button" onClick={() => setBulkEditForm(null)}>닫기</button>
              </div>
            </>
          )}
        </section>

        <div className="inventoryResultArea"><ProductTable mode="inventory" /></div>
        </section>
      </>
    );
  }


  function saveManualAiPreset() {
    const name = manualPresetName.trim() || `수동박스 조건 ${manualSavedPresets.length + 1}`;
    const preset = {
      id: Date.now(),
      name,
      manualType,
      manualBoxCount,
      manualTargetMargin,
      manualRetailExtra,
      manualStyle,
      manualPrefChar1,
      manualPrefChar2,
      manualCharStrategy,
      manualCandidateCount,
      salePrice,
      feeRate,
      request: manualAiRequest,
    };
    const next = [preset, ...manualSavedPresets].slice(0, 50);
    setManualSavedPresets(next);
    localStorage.setItem("manual_ai_presets", JSON.stringify(next));
    setManualPresetName("");
    alert("수동박스 AI 조건을 저장했어요.");
  }

  function loadManualAiPreset(id) {
    const preset = manualSavedPresets.find((x) => String(x.id) === String(id));
    if (!preset) return;
    setManualType(preset.manualType || manualType);
    setManualBoxCount(preset.manualBoxCount || manualBoxCount);
    setManualTargetMargin(preset.manualTargetMargin || manualTargetMargin);
    setManualRetailExtra(preset.manualRetailExtra || manualRetailExtra);
    setManualStyle(preset.manualStyle || manualStyle);
    setManualPrefChar1(preset.manualPrefChar1 || []);
    setManualPrefChar2(preset.manualPrefChar2 || []);
    setManualCharStrategy(preset.manualCharStrategy || manualCharStrategy);
    setManualCandidateCount(preset.manualCandidateCount || manualCandidateCount);
    setSalePrice(preset.salePrice || salePrice);
    setFeeRate(preset.feeRate || feeRate);
    setManualAiRequest(preset.request || "");
  }

  function deleteManualAiPreset(id) {
    const next = manualSavedPresets.filter((x) => String(x.id) !== String(id));
    setManualSavedPresets(next);
    localStorage.setItem("manual_ai_presets", JSON.stringify(next));
  }

  function saveScoopAiPreset() {
    const name = scoopPresetName.trim() || `삭제된 조건 ${scoopSavedPresets.length + 1}`;
    const preset = {
      id: Date.now(),
      name,
      scoopGroupCount,
      scoopMode,
      scoopPrice,
      scoopRetailLimit,
      scoopChar1Selected,
      scoopChar2Selected,
      scoopTargetMargin,
      scoopRecType,
      scoopRecSort,
      salePrice,
      feeRate,
      request: scoopAiRequest,
      groups: scoopGroups.map((g) => ({ ...g, products: [] })),
    };
    const next = [preset, ...scoopSavedPresets].slice(0, 50);
    setScoopSavedPresets(next);
    localStorage.setItem("scoop_ai_presets", JSON.stringify(next));
    setScoopPresetName("");
    alert("삭제된 조건을 저장했어요.");
  }

  function loadScoopAiPreset(id) {
    const preset = scoopSavedPresets.find((x) => String(x.id) === String(id));
    if (!preset) return;
    setScoopGroupCount(preset.scoopGroupCount || scoopGroupCount);
    setScoopMode(preset.scoopMode || scoopMode);
    setScoopPrice(preset.scoopPrice || scoopPrice);
    setScoopRetailLimit(preset.scoopRetailLimit || "");
    setScoopChar1Selected(preset.scoopChar1Selected || []);
    setScoopChar2Selected(preset.scoopChar2Selected || []);
    setScoopTargetMargin(preset.scoopTargetMargin || scoopTargetMargin);
    setScoopRecType(preset.scoopRecType || scoopRecType);
    setScoopRecSort(preset.scoopRecSort || scoopRecSort);
    setSalePrice(preset.salePrice || salePrice);
    setFeeRate(preset.feeRate || feeRate);
    setScoopAiRequest(preset.request || "");
    if (preset.groups?.length) {
      setScoopGroups(preset.groups.map((g, idx) => ({ ...g, id: g.id || idx + 1, products: scoopCandidateProducts().filter((p) => (g.categories || []).includes(p.category)) })));
    }
  }

  function deleteScoopAiPreset(id) {
    const next = scoopSavedPresets.filter((x) => String(x.id) !== String(id));
    setScoopSavedPresets(next);
    localStorage.setItem("scoop_ai_presets", JSON.stringify(next));
  }

  function previousCustomerItems(customerName) {
    const key = String(customerName || "").trim();
    if (!key) return [];
    const matchedOrders = orders.filter((o) => String(o.customer || "").trim() === key);
    const ids = new Set(matchedOrders.map((o) => o.id));
    return orderItems
      .filter((it) => ids.has(it.order_id))
      .map((it) => it.name || products.find((p) => String(p.id) === String(it.product_id))?.name || "")
      .filter(Boolean);
  }

  async function askAiForManualRecommendation() {
    setManualAiLoading(true);
    try {
      generateManualRecommendations();

      const context = {
        mode: "manual_box_recommendation",
        salePrice,
        feeRate,
        manualType,
        manualBoxCount,
        manualTargetMargin,
        manualRetailExtra,
        manualStyle,
        manualCharStrategy,
        preferredCharacters: [...manualPrefChar1, ...manualPrefChar2],
        customer: manualCustomer || customer,
        previousCustomerItems: previousCustomerItems(manualCustomer || customer),
        userRequest: manualAiRequest,
        revisionRequest: manualAiRevision,
        previousAiRecommendation: manualAiMemo,
        availableProducts: compactProductsForGemini(),
      };

      const res = await fetch("/api/gemini", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message:
            "수동박스 추천안을 딱 1개만 만들어줘. 마진율, 소비자가합, 도매가합, 재주문 고객이면 이전 상품 중복 제외를 고려해줘. " +
            "선택된 캐릭터가 있으면 그 캐릭터 재고 안에서 우선 구성하고, 선택된 캐릭터가 없으면 전체 재고에서 무작위/균형 추천으로 판단해줘. " +
            "수정요청이 있으면 이전 추천안을 기준으로 수정해줘. 실패하거나 조건을 못 맞추면 왜 실패했는지 이유와 부족한 조건을 먼저 말해줘. " +
            "상품명은 실제 재고명 그대로 쓰고, 수동박스로 옮기기 쉽게 상품명을 줄마다 정리해줘.",
          context,
        }),
      });

      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const reason = data?.message || data?.error?.message || (typeof data?.error === "string" ? data.error : JSON.stringify(data?.error || data));
        throw new Error(reason || "AI 추천 실패");
      }
      setManualAiMemo(data.text || "AI 답변을 가져오지 못했어요.");
    } catch (err) {
      setManualAiMemo("AI 추천 오류: " + String(err?.message || err));
    } finally {
      setManualAiLoading(false);
    }
  }

  async function askAiForScoopRecommendation() {
    setScoopAiLoading(true);
    try {
      if (scoopGroups.length === 0) {
        analyzeScoopCategories();
        generateScoopGroups();
      }
      setTimeout(() => {
        try { generateScoopRecommendations(); } catch {}
      }, 0);

      const context = {
        mode: "random_scoop_recommendation",
        salePrice,
        feeRate,
        scoopGroupCount,
        scoopMode,
        scoopPrice,
        scoopRetailLimit,
        scoopTargetMargin,
        selectedCharacters: [...scoopChar1Selected, ...scoopChar2Selected],
        groups: scoopGroups.map((g) => ({
          name: g.name,
          categories: g.categories,
          partQty: g.partQty,
          stock: g.stock,
          avgRetail: g.avgRetail,
        })),
        customer: scoopCustomer,
        previousCustomerItems: previousCustomerItems(scoopCustomer),
        userRequest: scoopAiRequest,
        revisionRequest: scoopAiRevision,
        previousAiRecommendation: scoopAiMemo,
        availableProducts: compactProductsForGemini(),
      };

      const res = await fetch("/api/gemini", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message:
            "삭제된 추천안을 딱 1개만 만들어줘. 반드시 현재 그룹명, 그룹별 카테고리, 그룹별 파츠 개수를 먼저 지켜줘. " +
            "선택된 캐릭터가 있으면 그 재고 안에서 우선 구성하고, 선택된 캐릭터가 없으면 전체 재고에서 무작위/균형 추천으로 판단해줘. " +
            "수정요청이 있으면 이전 추천안을 기준으로 수정해줘. 부족하면 같은 가격대 업그레이드나 소형 여러 개 추가로 맞춰줘. " +
            "실패하거나 조건을 못 맞추면 왜 실패했는지 이유와 부족한 조건을 먼저 말해줘. 실제 출고 가능한 추천안 1개와 조건 요약을 써줘.",
          context,
        }),
      });

      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const reason = data?.message || data?.error?.message || (typeof data?.error === "string" ? data.error : JSON.stringify(data?.error || data));
        throw new Error(reason || "AI 추천 실패");
      }
      setScoopAiMemo(data.text || "AI 답변을 가져오지 못했어요.");
    } catch (err) {
      setScoopAiMemo("AI 추천 오류: " + String(err?.message || err));
    } finally {
      setScoopAiLoading(false);
    }
  }


  function parseJsonMaybe(value, fallback) {
    if (Array.isArray(value)) return value;
    if (value && typeof value === "object") return value;
    if (typeof value === "string" && value.trim()) {
      try { return JSON.parse(value); } catch {}
    }
    return fallback;
  }

  function liveSessionFromDb(r) {
    return {
      id: r.id,
      createdAt: r.created_at || r.createdAt || "",
      title: r.title || "",
      date: r.date || "",
      keepDays: String(r.keep_days ?? r.keepDays ?? "7"),
      shippingFee: String(r.shipping_fee ?? r.shippingFee ?? "4000"),
      cardFeeRate: String(r.card_fee_rate ?? r.cardFeeRate ?? "3"),
      bankName: r.bank_name ?? r.bankName ?? "",
      accountNumber: r.account_number ?? r.accountNumber ?? "",
      accountHolder: r.account_holder ?? r.accountHolder ?? "여깁니다유",
      notice: r.notice || "",
      products: parseJsonMaybe(r.products, []),
    };
  }

  function liveSessionToDb(row) {
    return {
      id: String(row.id),
      created_at: row.createdAt || row.created_at || nowString(),
      title: row.title || "",
      date: row.date || "",
      keep_days: String(row.keepDays || "7"),
      shipping_fee: String(row.shippingFee || "4000"),
      card_fee_rate: String(row.cardFeeRate || "3"),
      bank_name: row.bankName || "",
      account_number: row.accountNumber || "",
      account_holder: row.accountHolder || "여깁니다유",
      notice: row.notice || "",
      products: row.products || [],
    };
  }

  function liveMemberFromDb(r) {
    return {
      id: r.id,
      updatedAt: r.updated_at || r.updatedAt || "",
      name: r.name || "",
      phone: r.phone || "",
      postalCode: r.postal_code || r.postalCode || "",
      baseAddress: r.base_address || r.baseAddress || "",
      detailAddress: r.detail_address || r.detailAddress || "",
      address: r.address || [r.base_address || r.baseAddress || "", r.detail_address || r.detailAddress || ""].filter(Boolean).join(" "),
      points: String(r.points ?? "0"),
      usedPoints: toInt(r.used_points ?? r.usedPoints),
      memo: r.memo || "",
      keepStart: r.keep_start || r.keepStart || "",
      keepDays: String(r.keep_days ?? r.keepDays ?? "7"),
    };
  }

  function liveMemberToDb(row) {
    return {
      id: String(row.id),
      updated_at: row.updatedAt || nowString(),
      name: row.name || "",
      phone: row.phone || "",
      postal_code: row.postalCode || "",
      base_address: row.baseAddress || "",
      detail_address: row.detailAddress || "",
      address: row.address || [row.baseAddress || "", row.detailAddress || ""].filter(Boolean).join(" "),
      points: String(row.points ?? "0"),
      used_points: toInt(row.usedPoints),
      memo: row.memo || "",
      keep_start: row.keepStart || "",
      keep_days: String(row.keepDays || "7"),
    };
  }

  function liveOrderFromDb(r) {
    return {
      id: r.id,
      sessionId: r.session_id || r.sessionId || "",
      liveTitle: r.live_title || r.liveTitle || "",
      liveDate: r.live_date || r.liveDate || "",
      createdAt: r.created_at || r.createdAt || "",
      updatedAt: r.updated_at || r.updatedAt || "",
      buyer: r.buyer || "",
      phone: r.phone || "",
      postalCode: r.postal_code || r.postalCode || "",
      baseAddress: r.base_address || r.baseAddress || "",
      detailAddress: r.detail_address || r.detailAddress || "",
      address: r.address || [r.base_address || r.baseAddress || "", r.detail_address || r.detailAddress || ""].filter(Boolean).join(" "),
      paymentMethod: r.payment_method || r.paymentMethod || "계좌이체",
      status: r.status || "미입금",
      trackingNo: r.tracking_no || r.trackingNo || "",
      memo: r.memo || "",
      shippingApply: r.shipping_apply ?? r.shippingApply ?? true,
      cardApply: r.card_apply ?? r.cardApply ?? false,
      items: parseJsonMaybe(r.items, []),
      subtotal: toInt(r.subtotal),
      paySubtotal: toInt(r.pay_subtotal ?? r.paySubtotal),
      shipping: toInt(r.shipping),
      cardFee: toInt(r.card_fee ?? r.cardFee),
      total: toInt(r.total),
      locked: !!r.locked,
      canceledAt: r.canceled_at || r.canceledAt || "",
      cancelReason: r.cancel_reason || r.cancelReason || "",
      boxWeight: String(r.box_weight ?? r.boxWeight ?? "2"),
      boxVolume: String(r.box_volume ?? r.boxVolume ?? "60"),
      household: r.household || "생활용품",
      deliveryMessage: r.delivery_message || r.deliveryMessage || "",
      memberKey: r.member_key || r.memberKey || "",
      bundleId: r.bundle_id || r.bundleId || "",
      deducted: !!r.deducted,
      paidAt: r.paid_at || r.paidAt || "",
      usedPoints: toInt(r.used_points ?? r.usedPoints),
    };
  }

  function liveOrderToDb(row) {
    return {
      id: String(row.id),
      session_id: String(row.sessionId || ""),
      live_title: row.liveTitle || "",
      live_date: row.liveDate || "",
      created_at: row.createdAt || nowString(),
      updated_at: row.updatedAt || "",
      buyer: row.buyer || "",
      phone: row.phone || "",
      postal_code: row.postalCode || "",
      base_address: row.baseAddress || "",
      detail_address: row.detailAddress || "",
      address: row.address || [row.baseAddress || "", row.detailAddress || ""].filter(Boolean).join(" "),
      payment_method: row.paymentMethod || "계좌이체",
      status: row.status || "미입금",
      tracking_no: row.trackingNo || "",
      memo: row.memo || "",
      shipping_apply: !!row.shippingApply,
      card_apply: !!row.cardApply,
      items: row.items || [],
      subtotal: toInt(row.subtotal),
      pay_subtotal: toInt(row.paySubtotal),
      shipping: toInt(row.shipping),
      card_fee: toInt(row.cardFee),
      total: toInt(row.total),
      locked: !!row.locked,
      canceled_at: row.canceledAt || "",
      cancel_reason: row.cancelReason || "",
      box_weight: String(row.boxWeight || "2"),
      box_volume: String(row.boxVolume || "60"),
      household: row.household || "생활용품",
      delivery_message: row.deliveryMessage || "",
      member_key: row.memberKey || makeMemberKey(row.buyer, row.phone),
      bundle_id: row.bundleId || "",
      deducted: !!row.deducted,
      paid_at: row.paidAt || "",
      used_points: toInt(row.usedPoints),
    };
  }

  function explainLiveTableMissing(error) {
    const msg = String(error?.message || error || "");
    if (msg.includes("does not exist") || msg.includes("schema cache")) {
      console.warn("라방 테이블이 아직 Supabase에 없어요. supabase_setup.sql을 실행해주세요.");
      return true;
    }
    return false;
  }

  async function getLiveSessions() {
    setLiveDataLoading(true);
    const { data, error } = await supabase.from("live_sessions").select("*").order("created_at", { ascending: false });
    setLiveDataLoading(false);
    if (error) { if (!explainLiveTableMissing(error)) console.log(error); return; }
    setLiveSessions((data || []).map(liveSessionFromDb));
  }

  async function getLiveMembers() {
    const { data, error } = await supabase.from("live_members").select("*").order("updated_at", { ascending: false });
    if (error) { if (!explainLiveTableMissing(error)) console.log(error); return; }
    setLiveMembers((data || []).map(liveMemberFromDb));
  }

  async function getLiveOrders() {
    const { data, error } = await supabase.from("live_orders").select("*").order("created_at", { ascending: false });
    if (error) { if (!explainLiveTableMissing(error)) console.log(error); return; }
    setLiveOrders((data || []).map(liveOrderFromDb));
  }

  async function saveLiveSessionDb(row) {
    const { error } = await supabase.from("live_sessions").upsert(liveSessionToDb(row));
    if (error) throw error;
  }

  async function saveLiveMemberDb(row) {
    const { error } = await supabase.from("live_members").upsert(liveMemberToDb(row));
    if (error) throw error;
  }

  async function saveLiveOrderDb(row) {
    const { error } = await supabase.from("live_orders").upsert(liveOrderToDb(row));
    if (error) throw error;
  }

  function preserveLiveScroll(callback) {
    const selectors = [".liveProductSourceTable", ".liveSelectedTable", ".liveOrdersTable"];
    const positions = selectors.map((sel) => {
      const el = document.querySelector(sel);
      return { sel, top: el?.scrollTop || 0, left: el?.scrollLeft || 0 };
    });
    callback();
    requestAnimationFrame(() => {
      positions.forEach(({ sel, top, left }) => {
        const el = document.querySelector(sel);
        if (el) { el.scrollTop = top; el.scrollLeft = left; }
      });
    });
  }


  const selectedLiveSession = useMemo(() => liveSessions.find((s) => String(s.id) === String(selectedLiveSessionId)) || liveSessions[0] || null, [liveSessions, selectedLiveSessionId]);

  const selectedLiveProducts = selectedLiveSession?.products || [];

  const liveFilteredProducts = useMemo(() => {
    const kw = liveProductSearch.trim().toLowerCase();
    return products.filter((p) => {
      if (toInt(p.stock) <= 0) return false;
      if (!kw) return true;
      return String(p.name || "").toLowerCase().includes(kw) || String(p.char1 || "").toLowerCase().includes(kw) || String(p.char2 || "").toLowerCase().includes(kw) || String(p.category || "").toLowerCase().includes(kw);
    }).slice(0, 400);
  }, [products, liveProductSearch]);

  const liveFilteredOrders = useMemo(() => {
    const kw = liveOrderSearch.trim().toLowerCase();
    return liveOrders
      .filter((o) => !selectedLiveSession || String(o.sessionId) === String(selectedLiveSession.id))
      .filter((o) => !liveDueOnly || isLiveKeepDueSoon(o))
      .filter((o) => !kw || String(o.buyer || "").toLowerCase().includes(kw) || String(o.phone || "").includes(kw) || String(o.trackingNo || "").toLowerCase().includes(kw) || String(o.memo || "").toLowerCase().includes(kw))
      .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
  }, [liveOrders, selectedLiveSession, liveStatusFilter, livePaymentFilter, liveOrderSearch, liveDueOnly]);

  const liveFilteredMembers = useMemo(() => {
    const kw = liveMemberSearch.trim().toLowerCase();
    return liveMembers.filter((m) => !kw || String(m.name || "").toLowerCase().includes(kw) || String(m.phone || "").includes(kw) || String(m.memo || "").toLowerCase().includes(kw));
  }, [liveMembers, liveMemberSearch]);

  const liveMemberLookupResults = useMemo(() => {
    const kw = liveMemberLookupSearch.trim().toLowerCase();
    if (!kw) return [];
    return liveMembers
      .filter((m) => String(m.name || "").toLowerCase().includes(kw) || phoneLast4(m.phone).includes(kw) || String(m.phone || "").includes(kw))
      .slice(0, 8);
  }, [liveMembers, liveMemberLookupSearch]);

  function makeLiveId(prefix = "live") {
    return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2, 7)}`;
  }


  function onlyDigits(v) {
    return String(v || "").replace(/[^0-9]/g, "");
  }

  function phoneLast4(v) {
    const d = onlyDigits(v);
    return d.slice(-4);
  }

  function makeMemberKey(name, phone) {
    const n = String(name || "").trim().replace(/\s+/g, "");
    const last = phoneLast4(phone);
    return n && last ? `${n}-${last}` : "";
  }

  function orderAddressOf(row) {
    return row.address || [row.baseAddress || "", row.detailAddress || ""].filter(Boolean).join(" ");
  }

  function liveOrderKeepDday(order) {
    if (String(order?.status || "") !== "정산후킵") return "-";
    const session = liveSessions.find((s) => String(s.id) === String(order.sessionId)) || selectedLiveSession;
    const startValue = order.liveDate || session?.date || order.createdAt?.slice(0, 10);
    if (!startValue) return "-";
    const start = new Date(startValue);
    if (Number.isNaN(start.getTime())) return "-";
    const end = new Date(start);
    end.setDate(end.getDate() + toInt(session?.keepDays || 0));
    const diff = Math.ceil((end.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
    if (diff > 0) return `D-${diff}`;
    if (diff === 0) return "D-DAY";
    return `출고필요 D+${Math.abs(diff)}`;
  }

  function isLiveKeepDueSoon(order) {
    const text = liveOrderKeepDday(order);
    if (text.includes("출고필요") || text === "D-DAY") return true;
    const m = text.match(/^D-(\d+)$/);
    return !!m && Number(m[1]) <= 2;
  }

  function sameLiveMemberOrders(form) {
    const key = makeMemberKey(form.buyer, form.phone);
    if (!key) return [];
    return liveOrders.filter((o) => !o.canceledAt && String(o.status || "") !== "출고완료" && makeMemberKey(o.buyer, o.phone) === key);
  }

  function getBundleOrders(order) {
    if (!order) return [];
    if (order.bundleId) return liveOrders.filter((o) => !o.canceledAt && String(o.bundleId) === String(order.bundleId));
    return [order];
  }

  async function createLiveSession() {
    const date = liveNewSession.date || new Date().toISOString().slice(0, 10);
    const title = liveNewSession.title.trim() || `${date} 라이브`;
    const row = { id: makeLiveId("session"), createdAt: nowString(), ...liveNewSession, title, date, products: [] };
    try {
      await saveLiveSessionDb(row);
      setLiveSessions((prev) => [row, ...prev.filter((s) => String(s.id) !== String(row.id))]);
      setSelectedLiveSessionId(row.id);
      alert("라방이 생성됐어요.");
    } catch (error) {
      alert("라방 생성 실패: " + error.message);
    }
  }

  async function updateLiveSession(patch) {
    if (!selectedLiveSession) return;
    const next = { ...selectedLiveSession, ...patch };
    try {
      await saveLiveSessionDb(next);
      setLiveSessions((prev) => prev.map((s) => String(s.id) === String(selectedLiveSession.id) ? next : s));
    } catch (error) {
      alert("라방 설정 저장 실패: " + error.message);
    }
  }

  async function addProductToLive(product) {
    if (!selectedLiveSession) return alert("먼저 라방을 생성하거나 선택해줘.");
    if (toInt(product.stock) <= 0) return alert("본재고가 부족해요.");
    const session = selectedLiveSession;
    try {
      const current = products.find((p) => String(p.id) === String(product.id)) || product;
      if (toInt(current.stock) <= 0) return alert("본재고가 부족해요.");
      const nextStock = toInt(current.stock) - 1;
      const { error: stockError } = await supabase.from("products").update({ stock: nextStock }).eq("id", product.id);
      if (stockError) throw stockError;

      let nextProducts = [];
      const exists = (session.products || []).find((x) => String(x.productId) === String(product.id));
      if (exists) {
        nextProducts = (session.products || []).map((x) => String(x.productId) === String(product.id)
          ? { ...x, liveQty: String(toInt(x.liveQty) + 1), remainingQty: String(toInt(x.remainingQty) + 1) }
          : x);
      } else {
        const item = {
          id: makeLiveId("liveitem"), productId: product.id, name: product.name, char1: product.char1, char2: product.char2,
          category: product.category, wholesale: toInt(product.wholesale), retail: toInt(product.retail), livePrice: toInt(product.retail), discountRate: "0",
          liveQty: "1", remainingQty: "1", memo: ""
        };
        nextProducts = [item, ...(session.products || [])];
      }
      const nextSession = { ...session, products: nextProducts };
      await saveLiveSessionDb(nextSession);
      preserveLiveScroll(() => setLiveSessions((prev) => prev.map((s) => String(s.id) === String(session.id) ? nextSession : s)));
      await getProducts();
    } catch (error) {
      alert("라방 상품 추가 실패: " + error.message);
      getProducts();
    }
  }

  async function updateLiveItem(itemId, patch) {
    if (!selectedLiveSession) return;
    const nextSession = {
      ...selectedLiveSession,
      products: (selectedLiveSession.products || []).map((it) => String(it.id) === String(itemId) ? { ...it, ...patch } : it)
    };
    try {
      await saveLiveSessionDb(nextSession);
      preserveLiveScroll(() => setLiveSessions((prev) => prev.map((s) => String(s.id) === String(selectedLiveSession.id) ? nextSession : s)));
    } catch (error) {
      alert("라방 상품 수정 실패: " + error.message);
    }
  }

  async function changeLiveQty(item, value) {
    if (!selectedLiveSession) return;
    const currentLiveQty = toInt(item.liveQty);
    const currentRemaining = toInt(item.remainingQty);
    const reservedOrSold = Math.max(0, currentLiveQty - currentRemaining);
    const requested = Math.max(reservedOrSold, toInt(value));
    if (toInt(value) < reservedOrSold) alert(`이미 주문에 잡힌 수량이 ${reservedOrSold}개라 배정수량을 낮출 수 없어요.`);
    const delta = requested - currentLiveQty;
    const product = products.find((p) => String(p.id) === String(item.productId));
    if (delta > 0 && toInt(product?.stock) < delta) return alert(`본재고가 부족해요. 추가 가능: ${toInt(product?.stock)}개`);
    try {
      if (delta !== 0) {
        const nextStock = toInt(product?.stock) - delta;
        const { error } = await supabase.from("products").update({ stock: nextStock }).eq("id", item.productId);
        if (error) throw error;
      }
      await updateLiveItem(item.id, { liveQty: String(requested), remainingQty: String(Math.max(0, currentRemaining + delta)) });
      await getProducts();
    } catch (error) {
      alert("라방 배정수량 수정 실패: " + error.message);
      getProducts();
    }
  }

  function changeLivePrice(item, value) {
    const price = toInt(value);
    const retail = toInt(item.retail);
    const discountRate = retail > 0 && price > 0 ? Math.max(0, Math.round((1 - price / retail) * 1000) / 10) : 0;
    updateLiveItem(item.id, { livePrice: value, discountRate: String(discountRate) });
  }

  function changeLiveDiscount(item, value) {
    const rate = Number(value || 0);
    const retail = toInt(item.retail);
    const price = Math.round(retail * (1 - rate / 100));
    updateLiveItem(item.id, { discountRate: value, livePrice: price });
  }

  async function removeLiveItem(itemId) {
    if (!selectedLiveSession) return;
    const target = (selectedLiveSession.products || []).find((it) => String(it.id) === String(itemId));
    if (!target) return;
    if (toInt(target.liveQty) !== toInt(target.remainingQty)) return alert("이미 주문에 잡힌 수량이 있는 상품은 삭제할 수 없어요. 주문 취소/복구 후 삭제해줘.");
    const ok = window.confirm(`${target.name} 라방 배정 ${target.liveQty}개를 본재고로 돌리고 삭제할까요?`);
    if (!ok) return;
    try {
      const current = products.find((p) => String(p.id) === String(target.productId));
      const { error: stockError } = await supabase.from("products").update({ stock: toInt(current?.stock) + toInt(target.liveQty) }).eq("id", target.productId);
      if (stockError) throw stockError;
      const nextSession = { ...selectedLiveSession, products: (selectedLiveSession.products || []).filter((it) => String(it.id) !== String(itemId)) };
      await saveLiveSessionDb(nextSession);
      setLiveSessions((prev) => prev.map((s) => String(s.id) === String(selectedLiveSession.id) ? nextSession : s));
      await getProducts();
    } catch (error) {
      alert("라방 상품 삭제 실패: " + error.message);
      getProducts();
    }
  }


  async function deleteLiveSessionWithRestore() {
    if (!selectedLiveSession) return alert("삭제할 라방을 선택해줘.");
    const sessionOrders = liveOrders.filter((o) => String(o.sessionId) === String(selectedLiveSession.id) && !o.canceledAt);
    if (sessionOrders.some((o) => String(o.status) === "출고완료")) {
      return alert("출고완료 주문이 있는 라방은 삭제할 수 없어요. 먼저 출고완료 주문을 확인해줘.");
    }
    const ok = window.confirm(`${selectedLiveSession.title} 라방을 삭제하고, 라방에 배정된 모든 수량을 본재고로 되돌릴까요?\n\n미입금/입금확인/킵/송장입력 주문은 취소 처리됩니다.`);
    if (!ok) return;
    try {
      for (const li of selectedLiveSession.products || []) {
        const current = products.find((p) => String(p.id) === String(li.productId));
        const restoreQty = toInt(li.liveQty);
        const { error } = await supabase.from("products").update({ stock: toInt(current?.stock) + restoreQty }).eq("id", li.productId);
        if (error) throw error;
      }
      for (const o of sessionOrders) {
        if (toInt(o.usedPoints) > 0) await adjustMemberPointsByOrder(o, -toInt(o.usedPoints));
        await saveLiveOrderDb({ ...o, status: "취소", canceledAt: nowString(), cancelReason: "라방 삭제로 취소", updatedAt: nowString(), usedPoints: 0 });
      }
      const { error: delError } = await supabase.from("live_sessions").delete().eq("id", selectedLiveSession.id);
      if (delError) throw delError;
      await writeAudit("live_session_delete_restore", `${selectedLiveSession.title} / orders=${sessionOrders.length}`);
      setLiveSessions((prev) => prev.filter((s) => String(s.id) !== String(selectedLiveSession.id)));
      setLiveOrders((prev) => prev.map((o) => String(o.sessionId) === String(selectedLiveSession.id) && !o.canceledAt ? { ...o, status: "취소", canceledAt: nowString(), cancelReason: "라방 삭제로 취소" } : o));
      setSelectedLiveSessionId(null);
      await getProducts();
      alert("라방을 삭제하고 본재고로 복구했어요.");
    } catch (error) {
      alert("라방 삭제/복구 실패: " + String(error.message || error));
      getProducts();
    }
  }

  async function saveLiveMember() {
    const name = liveMemberForm.name.trim();
    if (!name) return alert("회원 이름을 입력해줘.");
    const key = makeMemberKey(name, liveMemberForm.phone);
    const existing = liveMembers.find((m) => (key && makeMemberKey(m.name, m.phone) === key) || String(m.id) === String(liveMemberForm.id || ""));
    const row = { id: existing?.id || makeLiveId("member"), updatedAt: nowString(), ...liveMemberForm, name, address: [liveMemberForm.baseAddress || "", liveMemberForm.detailAddress || ""].filter(Boolean).join(" ") };
    try {
      await saveLiveMemberDb(row);
      setLiveMembers((prev) => existing ? prev.map((m) => String(m.id) === String(existing.id) ? { ...m, ...row } : m) : [row, ...prev]);
      setLiveOrderForm((prev) => ({ ...prev, buyer: row.name, phone: row.phone || "", postalCode: row.postalCode || "", baseAddress: row.baseAddress || "", detailAddress: row.detailAddress || "", address: row.address || "", points: String(row.points || "0"), usedPoints: 0 }));
      setLiveMemberForm({ name: "", phone: "", postalCode: "", baseAddress: "", detailAddress: "", address: "", points: "0", memo: "" });
    } catch (error) {
      alert("회원 저장 실패: " + error.message);
    }
  }

  function loadMemberToOrder(member) {
    setSelectedLiveMemberId(member.id || "");
    setLiveMemberLookupSearch(`${member.name || ""} ${phoneLast4(member.phone)}`.trim());
    setLiveOrderForm((prev) => ({ ...prev, buyer: member.name || "", phone: member.phone || "", postalCode: member.postalCode || "", baseAddress: member.baseAddress || "", detailAddress: member.detailAddress || "", address: member.address || "", points: String(member.points || "0"), usedPoints: 0, memo: member.memo || prev.memo }));
    setLiveMemberForm({ name: member.name || "", phone: member.phone || "", postalCode: member.postalCode || "", baseAddress: member.baseAddress || "", detailAddress: member.detailAddress || "", address: member.address || "", points: String(member.points || 0), memo: member.memo || "" });
  }

  async function deleteLiveMember(member) {
    if (!member?.id) return;
    if (!window.confirm(`${member.name} 회원 정보를 삭제할까요?`)) return;
    const { error } = await supabase.from("live_members").delete().eq("id", member.id);
    if (error) return alert("회원 삭제 실패: " + error.message);
    setLiveMembers((prev) => prev.filter((m) => String(m.id) !== String(member.id)));
  }

  function keepDday(member) {
    return "-";
  }

  function addLiveItemToCart(item) {
    if (toInt(item.remainingQty) <= 0) return alert("라방 남은 수량이 없어요.");
    setLiveCart((prev) => {
      const exists = prev.find((x) => String(x.liveItemId) === String(item.id));
      if (exists) return prev.map((x) => String(x.liveItemId) === String(item.id) ? { ...x, qty: Math.min(toInt(item.remainingQty), toInt(x.qty) + 1) } : x);
      return [{ liveItemId: item.id, productId: item.productId, name: item.name, char1: item.char1, char2: item.char2, wholesale: toInt(item.wholesale), qty: 1, price: toInt(item.livePrice), prepaid: "N" }, ...prev];
    });
  }

  function updateLiveCartItem(idx, patch) {
    setLiveCart((prev) => prev.map((it, i) => {
      if (i !== idx) return it;
      const next = { ...it, ...patch };
      if (Object.prototype.hasOwnProperty.call(patch, "qty")) {
        const liveItem = selectedLiveProducts.find((li) => String(li.id) === String(it.liveItemId));
        const editingOrder = editingLiveOrderId ? liveOrders.find((o) => String(o.id) === String(editingLiveOrderId)) : null;
        const oldQty = (editingOrder?.items || []).filter((oldIt) => String(oldIt.liveItemId) === String(it.liveItemId)).reduce((sum, oldIt) => sum + toInt(oldIt.qty), 0);
        const maxQty = Math.max(1, toInt(liveItem?.remainingQty || 0) + oldQty);
        next.qty = String(Math.min(maxQty, Math.max(1, toInt(patch.qty))));
      }
      if (Object.prototype.hasOwnProperty.call(patch, "price")) next.price = String(Math.max(0, toInt(patch.price)));
      return next;
    }));
  }

  function liveCartSummary() {
    const subtotal = liveCart.reduce((sum, it) => sum + toInt(it.price) * toInt(it.qty), 0);
    const paySubtotal = liveCart.reduce((sum, it) => String(it.prepaid).toUpperCase() === "Y" ? sum : sum + toInt(it.price) * toInt(it.qty), 0);
    const shipping = liveOrderForm.shippingApply && subtotal > 0 ? toInt(selectedLiveSession?.shippingFee || 0) : 0;
    const cardFee = liveOrderForm.paymentMethod === "카드결제" && paySubtotal > 0 ? Math.round((paySubtotal + shipping) * Number(selectedLiveSession?.cardFeeRate || 0) / 100) : 0;
    const usedPoints = Math.min(toInt(liveOrderForm.usedPoints), paySubtotal + shipping + cardFee);
    return { subtotal, paySubtotal, shipping, cardFee, usedPoints, total: Math.max(0, paySubtotal + shipping + cardFee - usedPoints) };
  }

  function memberFormFromOrderForm(extra = {}, pointDelta = 0) {
    const existingByKey = liveMembers.find((m) => makeMemberKey(m.name, m.phone) === makeMemberKey(liveOrderForm.buyer, liveOrderForm.phone));
    const id = selectedLiveMemberId || existingByKey?.id || makeMemberKey(liveOrderForm.buyer, liveOrderForm.phone) || makeLiveId("member");
    const existing = liveMembers.find((m) => String(m.id) === String(id)) || existingByKey || {};
    const delta = toInt(pointDelta);
    return {
      id,
      updatedAt: nowString(),
      name: liveOrderForm.buyer || "",
      phone: liveOrderForm.phone || "",
      postalCode: liveOrderForm.postalCode || "",
      baseAddress: liveOrderForm.baseAddress || "",
      detailAddress: liveOrderForm.detailAddress || "",
      address: orderAddressOf(liveOrderForm),
      points: String(Math.max(0, toInt(liveOrderForm.points) - delta)),
      usedPoints: Math.max(0, toInt(existing.usedPoints) + delta),
      memo: liveOrderForm.memo || existing.memo || "",
      ...extra,
    };
  }

  async function saveMemberFromOrderForm(showAlert = true, pointDelta = 0) {
    if (!liveOrderForm.buyer.trim()) return alert("고객명을 입력해줘.");
    if (!liveOrderForm.phone.trim()) return alert("전화번호를 입력해줘.");
    const row = memberFormFromOrderForm({}, pointDelta);
    await saveLiveMemberDb(row);
    setSelectedLiveMemberId(row.id);
    setLiveMembers((prev) => [row, ...prev.filter((m) => String(m.id) !== String(row.id))]);
    setLiveOrderForm((prev) => ({ ...prev, points: String(row.points || "0") }));
    if (showAlert) alert("회원 정보가 저장됐어요.");
    return row;
  }

  function useAllMemberPoints() {
    const pts = toInt(liveOrderForm.points);
    if (pts <= 0) return alert("사용할 포인트가 없어요.");
    setLiveOrderForm((prev) => ({ ...prev, usedPoints: pts }));
  }

  function findLiveMemberForOrderLike(row) {
    const key = makeMemberKey(row?.buyer || row?.name, row?.phone);
    return liveMembers.find((m) => {
      const mk = makeMemberKey(m.name, m.phone);
      return (key && mk === key) || (row?.memberKey && mk === row.memberKey) || (row?.memberKey && String(m.id) === String(row.memberKey));
    });
  }

  async function adjustMemberPointsByOrder(row, deltaUsedPoints) {
    const delta = toInt(deltaUsedPoints);
    if (delta === 0) return;
    const member = findLiveMemberForOrderLike(row);
    if (!member) return;
    const next = {
      ...member,
      updatedAt: nowString(),
      points: String(Math.max(0, toInt(member.points) - delta)),
      usedPoints: Math.max(0, toInt(member.usedPoints) + delta),
    };
    await saveLiveMemberDb(next);
    setLiveMembers((prev) => prev.map((m) => String(m.id) === String(member.id) ? next : m));
  }

  function resetLiveOrderFormAfterSave() {
    setLiveCart([]);
    setLiveOrderForm((prev) => ({ ...prev, buyer: "", phone: "", postalCode: "", baseAddress: "", detailAddress: "", address: "", memo: "", trackingNo: "", status: "미입금", deliveryMessage: "", points: "0", usedPoints: 0 }));
    setSelectedLiveMemberId("");
    setLiveMemberLookupSearch("");
    setEditingLiveOrderId("");
  }

  function beginEditLiveOrder(order) {
    if (!order) return;
    if (order.locked) return alert("구매확정 잠금된 주문이에요. 잠금해제 후 수정해줘.");
    if (order.canceledAt || order.status === "취소") return alert("취소된 주문은 수정할 수 없어요.");
    const member = findLiveMemberForOrderLike(order);
    setEditingLiveOrderId(order.id);
    setSelectedLiveMemberId(member?.id || "");
    setLiveMemberLookupSearch(member ? `${member.name || ""} ${phoneLast4(member.phone)}`.trim() : "");
    setLiveOrderForm({
      buyer: order.buyer || "",
      phone: order.phone || "",
      postalCode: order.postalCode || "",
      baseAddress: order.baseAddress || "",
      detailAddress: order.detailAddress || "",
      address: orderAddressOf(order),
      paymentMethod: order.paymentMethod || "계좌이체",
      status: order.status || "미입금",
      trackingNo: order.trackingNo || "",
      memo: order.memo || "",
      shippingApply: order.shippingApply !== false,
      cardApply: !!order.cardApply,
      boxWeight: order.boxWeight || "2",
      boxVolume: order.boxVolume || "60",
      household: order.household || "생활용품",
      deliveryMessage: order.deliveryMessage || "",
      points: String(toInt(member?.points) + toInt(order.usedPoints || 0)),
      usedPoints: toInt(order.usedPoints || 0),
    });
    setLiveCart((order.items || []).map((it) => ({
      liveItemId: it.liveItemId,
      productId: it.productId,
      name: it.name,
      char1: it.char1,
      char2: it.char2,
      wholesale: toInt(it.wholesale),
      qty: toInt(it.qty),
      price: toInt(it.price),
      prepaid: String(it.prepaid || "N").toUpperCase() === "Y" ? "Y" : "N",
    })));
    alert("주문을 수정 모드로 불러왔어요. 품목/주소/금액 수정 후 저장을 눌러줘.");
  }

  function cancelLiveOrderEdit() {
    resetLiveOrderFormAfterSave();
  }

  async function openDaumPostcode(target = "order") {
    const loadScript = () => new Promise((resolve, reject) => {
      if (window.daum?.Postcode) return resolve();
      const old = document.getElementById("daum-postcode-script");
      if (old) {
        old.addEventListener("load", resolve, { once: true });
        old.addEventListener("error", reject, { once: true });
        return;
      }
      const script = document.createElement("script");
      script.id = "daum-postcode-script";
      script.src = "https://t1.daumcdn.net/mapjsapi/bundle/postcode/prod/postcode.v2.js";
      script.onload = resolve;
      script.onerror = reject;
      document.body.appendChild(script);
    });
    try {
      await loadScript();
      new window.daum.Postcode({
        oncomplete: function(data) {
          const postalCode = data.zonecode || "";
          const baseAddress = data.roadAddress || data.jibunAddress || "";
          if (target === "member") {
            setLiveMemberForm((prev) => ({ ...prev, postalCode, baseAddress, address: [baseAddress, prev.detailAddress || ""].filter(Boolean).join(" ") }));
          } else {
            setLiveOrderForm((prev) => ({ ...prev, postalCode, baseAddress, address: [baseAddress, prev.detailAddress || ""].filter(Boolean).join(" ") }));
          }
        }
      }).open();
    } catch (error) {
      alert("우편번호 검색창을 불러오지 못했어요. 우편번호/주소를 직접 입력해줘.");
    }
  }

  async function saveLiveOrderAndDeduct() {
    if (!selectedLiveSession) return alert("라방을 선택해줘.");
    if (!liveOrderForm.buyer.trim()) return alert("구매자명을 입력해줘.");
    if (!liveOrderForm.phone.trim()) return alert("전화번호를 입력해줘. 같은 회원 판단에 필요해.");
    if (liveCart.length === 0) return alert("주문 품목을 추가해줘.");

    const oldOrder = editingLiveOrderId ? liveOrders.find((o) => String(o.id) === String(editingLiveOrderId)) : null;
    const session = liveSessions.find((s) => String(s.id) === String(oldOrder?.sessionId || selectedLiveSession.id)) || selectedLiveSession;
    const productsForCalc = session.products || [];

    const oldQtyByLiveItem = {};
    (oldOrder?.items || []).forEach((it) => { oldQtyByLiveItem[it.liveItemId] = (oldQtyByLiveItem[it.liveItemId] || 0) + toInt(it.qty); });
    const newQtyByLiveItem = {};
    liveCart.forEach((it) => { newQtyByLiveItem[it.liveItemId] = (newQtyByLiveItem[it.liveItemId] || 0) + toInt(it.qty); });

    for (const item of liveCart) {
      const liveItem = productsForCalc.find((li) => String(li.id) === String(item.liveItemId));
      if (!liveItem) return alert(`라방 상품을 찾을 수 없어요: ${item.name}`);
      if (toInt(item.qty) <= 0) return alert(`수량을 확인해줘: ${item.name}`);
      const availableForEdit = toInt(liveItem.remainingQty) + toInt(oldQtyByLiveItem[item.liveItemId] || 0);
      if (toInt(newQtyByLiveItem[item.liveItemId] || 0) > availableForEdit) return alert(`라방 남은 수량을 초과했어요: ${item.name}`);
    }

    const summary = liveCartSummary();
    const ok = window.confirm(`${oldOrder ? "주문을 수정" : "주문을 미입금 상태로 저장"}할까요?\n\n구매자: ${liveOrderForm.buyer}\n품목: ${liveCart.length}종\n배송비: ${money(summary.shipping)}\n최종 결제금액: ${money(summary.total)}`);
    if (!ok) return;

    try {
      const nextSession = {
        ...session,
        products: (productsForCalc || []).map((li) => {
          const oldReserved = toInt(oldQtyByLiveItem[li.id] || 0);
          const nextReserved = toInt(newQtyByLiveItem[li.id] || 0);
          const nextRemaining = Math.max(0, toInt(li.remainingQty) + oldReserved - nextReserved);
          return oldReserved || nextReserved ? { ...li, remainingQty: String(nextRemaining) } : li;
        })
      };
      await saveLiveSessionDb(nextSession);

      const memberKey = makeMemberKey(liveOrderForm.buyer, liveOrderForm.phone);
      const order = {
        id: oldOrder?.id || makeLiveId("liveorder"), sessionId: session.id, liveTitle: session.title, liveDate: session.date,
        createdAt: oldOrder?.createdAt || nowString(), updatedAt: nowString(), locked: oldOrder?.locked || false, canceledAt: "", cancelReason: "", deducted: oldOrder?.deducted || false, paidAt: oldOrder?.paidAt || "", memberKey,
        ...liveOrderForm, status: oldOrder?.status || "미입금", address: orderAddressOf(liveOrderForm), items: liveCart.map((it) => ({ ...it, qty: toInt(it.qty), price: toInt(it.price) })), ...summary,
      };
      await saveLiveOrderDb(order);

      setLiveSessions((prev) => prev.map((s) => String(s.id) === String(session.id) ? nextSession : s));
      setLiveOrders((prev) => oldOrder ? prev.map((o) => String(o.id) === String(order.id) ? order : o) : [order, ...prev]);
      if (oldOrder) {
        await adjustMemberPointsByOrder(order, toInt(summary.usedPoints) - toInt(oldOrder.usedPoints));
      } else {
        await saveMemberFromOrderForm(false, summary.usedPoints);
      }
      await writeAudit(oldOrder ? "live_order_updated" : "live_order_saved", `${session.title} / ${liveOrderForm.buyer} / ${summary.total}`);
      resetLiveOrderFormAfterSave();
      alert(oldOrder ? "라방 주문을 수정했어요." : "라방 주문이 미입금 상태로 저장됐어요. 입금 확인 후 상태를 입금확인으로 바꾸면 매출에 반영돼요.");
    } catch (error) {
      alert(String(error.message || error));
    }
  }

  async function updateLiveOrder(orderId, patch) {
    const current = liveOrders.find((o) => String(o.id) === String(orderId));
    if (!current) return;
    if (current.locked && !Object.prototype.hasOwnProperty.call(patch, "locked")) return alert("구매확정 잠금된 주문이에요. 수정하려면 잠금해제 해줘.");
    let next = { ...current, ...patch, updatedAt: nowString() };
    if (Object.prototype.hasOwnProperty.call(patch, "status") && patch.status === "입금확인" && current.status !== "입금확인") {
      next = { ...next, deducted: true, paidAt: nowString() };
    }
    if (Object.prototype.hasOwnProperty.call(patch, "status") && patch.status === "정산후킵") {
      next = { ...next, trackingNo: "" };
    }
    try {
      await saveLiveOrderDb(next);
      setLiveOrders((prev) => prev.map((o) => String(o.id) === String(orderId) ? next : o));
    } catch (error) {
      alert("주문 수정 실패: " + error.message);
    }
  }

  async function cancelLiveOrderWithRestore(order) {
    if (!order || order.canceledAt) return;
    if (order.locked) return alert("구매확정 잠금된 주문이에요. 잠금해제 후 취소해줘.");
    const reason = window.prompt("취소 사유를 적어줘", order.cancelReason || "고객 취소");
    if (reason === null) return;
    const ok = window.confirm(`주문을 취소하고 라방 남은수량을 복구할까요?\n\n구매자: ${order.buyer}\n금액: ${money(order.total)}`);
    if (!ok) return;
    try {
      const session = liveSessions.find((s) => String(s.id) === String(order.sessionId));
      if (session) {
        const nextSession = {
          ...session,
          products: (session.products || []).map((li) => {
            const restored = (order.items || []).filter((it) => String(it.liveItemId) === String(li.id)).reduce((sum, it) => sum + toInt(it.qty), 0);
            return restored ? { ...li, remainingQty: String(Math.min(toInt(li.liveQty), toInt(li.remainingQty) + restored)) } : li;
          })
        };
        await saveLiveSessionDb(nextSession);
        setLiveSessions((prev) => prev.map((s) => String(s.id) === String(session.id) ? nextSession : s));
      }
      if (toInt(order.usedPoints) > 0) await adjustMemberPointsByOrder(order, -toInt(order.usedPoints));
      await updateLiveOrder(order.id, { status: "취소", canceledAt: nowString(), cancelReason: reason, usedPoints: 0 });
      await writeAudit("live_order_cancel_restore", `${order.buyer} / ${order.id} / ${reason}`);
    } catch (error) {
      alert(String(error.message || error));
    }
  }


  async function deleteLiveOrderWithRestore(order) {
    if (!order) return;
    if (order.locked) return alert("구매확정 잠금된 주문이에요. 잠금해제 후 삭제해줘.");
    const ok = window.confirm(`주문을 완전히 삭제할까요?\n\n취소 전 주문이면 라방 남은수량을 먼저 복구한 뒤 삭제됩니다.\n구매자: ${order.buyer}`);
    if (!ok) return;
    try {
      if (!order.canceledAt) {
        const session = liveSessions.find((s) => String(s.id) === String(order.sessionId));
        if (session) {
          const nextSession = {
            ...session,
            products: (session.products || []).map((li) => {
              const restored = (order.items || []).filter((it) => String(it.liveItemId) === String(li.id)).reduce((sum, it) => sum + toInt(it.qty), 0);
              return restored ? { ...li, remainingQty: String(Math.min(toInt(li.liveQty), toInt(li.remainingQty) + restored)) } : li;
            })
          };
          await saveLiveSessionDb(nextSession);
          setLiveSessions((prev) => prev.map((s) => String(s.id) === String(session.id) ? nextSession : s));
        }
      }
      if (!order.canceledAt && toInt(order.usedPoints) > 0) await adjustMemberPointsByOrder(order, -toInt(order.usedPoints));
      const { error } = await supabase.from("live_orders").delete().eq("id", order.id);
      if (error) throw error;
      setLiveOrders((prev) => prev.filter((o) => String(o.id) !== String(order.id)));
      await writeAudit("live_order_delete_restore", `${order.buyer} / ${order.id}`);
      alert("주문을 삭제했어요.");
    } catch (error) {
      alert("주문 삭제 실패: " + String(error.message || error));
    }
  }

  async function bulkMarkLiveOrdersPaid() {
    const targets = liveFilteredOrders.filter((o) => !o.locked && !o.canceledAt && ["미입금"].includes(String(o.status || "")));
    if (!targets.length) return alert("입금확인으로 바꿀 주문이 없어요.");
    if (!window.confirm(`현재 필터의 ${targets.length}건을 입금확인으로 변경할까요?`)) return;
    for (const o of targets) await updateLiveOrder(o.id, { status: "입금확인" });
  }

  function liveSalesSummary() {
    const rows = liveOrders.filter((o) => selectedLiveSession && String(o.sessionId) === String(selectedLiveSession.id) && !o.canceledAt);
    const paidRows = rows.filter((o) => ["입금확인", "정산후킵", "송장입력", "출고완료"].includes(String(o.status || "")));
    const soldItems = paidRows.flatMap((o) => o.items || []);
    const cost = soldItems.reduce((s, it) => s + toInt(it.wholesale) * toInt(it.qty), 0);
    const paySubtotal = paidRows.reduce((s, o) => s + toInt(o.paySubtotal), 0);
    const allocated = (selectedLiveSession?.products || []).reduce((s, it) => s + toInt(it.liveQty), 0);
    const remaining = (selectedLiveSession?.products || []).reduce((s, it) => s + toInt(it.remainingQty), 0);
    return {
      orderCount: rows.length,
      buyerCount: new Set(rows.map((o) => makeMemberKey(o.buyer, o.phone) || o.buyer)).size,
      paidCount: paidRows.length,
      total: paidRows.reduce((s, o) => s + toInt(o.total), 0),
      paySubtotal,
      profit: paySubtotal - cost,
      unpaid: rows.filter((o) => ["미입금"].includes(String(o.status || ""))).reduce((s, o) => s + toInt(o.total), 0),
      cardFee: paidRows.reduce((s, o) => s + toInt(o.cardFee), 0),
      shipping: paidRows.reduce((s, o) => s + toInt(o.shipping), 0),
      allocated,
      remaining,
      soldQty: soldItems.reduce((s, it) => s + toInt(it.qty), 0),
      keepDue: rows.filter(isLiveKeepDueSoon).length,
    };
  }

  async function bundleLiveOrdersFor(order) {
    const key = makeMemberKey(order.buyer, order.phone);
    if (!key) return alert("고객명과 전화번호 뒷 4자리가 있어야 같은 회원으로 묶을 수 있어요.");
    const targets = liveOrders.filter((o) => !o.canceledAt && String(o.status || "") !== "출고완료" && makeMemberKey(o.buyer, o.phone) === key);
    if (targets.length < 2) return alert("합칠 수 있는 미출고 주문이 2건 이상 없어요.");
    const bundleId = order.bundleId || makeLiveId("bundle");
    if (!window.confirm(`${order.buyer}님의 미출고 주문 ${targets.length}건을 합배송으로 묶을까요?`)) return;
    for (const o of targets) await updateLiveOrder(o.id, { bundleId });
    alert("주문을 합배송 묶음으로 연결했어요.");
  }

  function downloadLiveShippingExcel() {
    const targets = liveOrders.filter((o) => selectedLiveSession && String(o.sessionId) === String(selectedLiveSession.id) && !o.canceledAt && ["입금확인", "송장입력"].includes(String(o.status || "")));
    if (!targets.length) return alert("택배접수로 내보낼 입금확인/송장입력 주문이 없어요.");
    const seen = new Set();
    const rows = [];
    targets.forEach((o) => {
      const key = o.bundleId || o.id;
      if (seen.has(key)) return;
      seen.add(key);
      const bundle = o.bundleId ? liveOrders.filter((x) => String(x.bundleId) === String(o.bundleId) && !x.canceledAt) : [o];
      const first = bundle[0] || o;
      const memo = bundle.map((bo) => `${bo.liveDate || "날짜없음"} ${bo.liveTitle || ""} / ${bo.items?.length || 0}종`).join("\n");
      rows.push([
        first.buyer || "",
        first.postalCode || "",
        first.baseAddress || first.address || "",
        first.detailAddress || "",
        first.phone || "",
        first.boxWeight || "2",
        first.boxVolume || "60",
        "1",
        first.household || "생활용품",
        [first.deliveryMessage || "", first.memo || "", memo].filter(Boolean).join("\n"),
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 42 }, { wch: 24 }, { wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 45 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "택배접수");
    XLSX.writeFile(wb, `라방_택배접수_${selectedLiveSession?.date || ""}.xlsx`);
  }

  function htmlSafe(value) {
    return String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[ch]));
  }

  function openLiveInvoicePdf(order) {
    const session = liveSessions.find((s) => String(s.id) === String(order.sessionId)) || selectedLiveSession || {};
    const items = order.items || [];
    const page1 = items.slice(0, 15);
    const page2 = items.slice(15, 30);
    const itemRows = (list, offset = 0) => list.map((it, idx) => `
      <tr><td>${offset + idx + 1}</td><td>${htmlSafe(it.name || "")}</td><td>${toInt(it.qty)}</td><td>${money(toInt(it.price) * toInt(it.qty))}</td><td>${String(it.prepaid).toUpperCase() === "Y" ? "Y" : "N"}</td><td>${String(it.prepaid).toUpperCase() === "Y" ? "0원" : money(toInt(it.price) * toInt(it.qty))}</td></tr>
    `).join("");
    const emptyRows = (count) => Array.from({ length: Math.max(0, count) }).map(() => `<tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>`).join("");
    const summary = `<div class="summary"><div class="summaryTitle">최종 결제금액</div><div><span>상품합계</span><b>${money(order.paySubtotal)}</b></div><div><span>배송비</span><b>${money(order.shipping)}</b></div><div><span>카드수수료 (${order.paymentMethod === "카드결제" ? session.cardFeeRate || 0 : 0}%)</span><b>${money(order.cardFee)}</b></div><div><span>포인트 차감</span><b>-${money(order.usedPoints || 0)}</b></div><div class="total"><span>총 결제금액</span><b>${money(order.total)}</b></div></div>`;
    const page = (list, offset, withInfo, last) => `
      <section class="page">
        <div class="watermark">여깁니다유</div>
        ${withInfo ? `<div class="invoiceNo">정산번호 #${htmlSafe(order.id || "")}</div><h1>여깁니다유 라이브 정산서</h1>
        <table class="info"><tr><th>라방날짜</th><td>${htmlSafe(order.liveDate || "")}</td><th>구매자명</th><td>${htmlSafe(order.buyer || "")}</td><th>연락처</th><td>${htmlSafe(order.phone || "")}</td><th>킵</th><td>${order.status === "정산후킵" ? "Y" : "N"}</td></tr><tr><th>결제방법</th><td>${htmlSafe(order.paymentMethod || "")}</td><th>입금기한</th><td>${htmlSafe((session.notice || "").split("\n")[0] || "")}</td><th>배송비</th><td>${money(order.shipping)}</td><th>별명/아이디</th><td></td></tr><tr><th>주소</th><td colspan="7">${htmlSafe(orderAddressOf(order))}</td></tr><tr><th>입금계좌</th><td colspan="7">${htmlSafe([session.bankName, session.accountNumber, session.accountHolder].filter(Boolean).join(" / "))}</td></tr></table>` : `<div class="pageSub"><b>라방날짜</b> ${htmlSafe(order.liveDate || "")} &nbsp; <b>정산번호</b> ${htmlSafe(order.id || "")} &nbsp; <b>구매자명</b> ${htmlSafe(order.buyer || "")}</div>`}
        <table class="items"><thead><tr><th>No</th><th>상품명</th><th>수량</th><th>금액</th><th>선결제 유무</th><th>실결제금액</th></tr></thead><tbody>${itemRows(list, offset)}${emptyRows((withInfo ? 15 : 15) - list.length)}</tbody></table>
        ${withInfo ? `<div class="notice">${htmlSafe(session.notice || "입금 확인 순서대로 포장 후 출고됩니다.")}</div>` : ""}
        ${last ? summary : ""}
        <div class="footerNotice">본 정산서는 여깁니다유 라이브 구매 확인용으로만 제공되며, 무단 공유 및 외부 유출을 금합니다.</div>
      </section>`;
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>여깁니다유 정산서</title><style>
      @page{size:A4;margin:0}*{box-sizing:border-box}body{margin:0;background:#eee;font-family:Arial,'맑은 고딕',sans-serif;color:#4a3b00}.no-print{position:fixed;right:14px;top:14px;z-index:99;background:#ffd84d;border:1px solid #d0aa00;padding:8px 12px;font-weight:800}.page{position:relative;width:210mm;min-height:297mm;margin:0 auto 10px;background:#fff9e6;padding:12mm;page-break-after:always;overflow:hidden}.page:last-child{page-break-after:auto}.watermark{position:absolute;left:50%;top:49%;transform:translate(-50%,-50%);font-size:56px;font-weight:900;opacity:.035;white-space:nowrap;pointer-events:none}.invoiceNo{text-align:right;font-weight:800}h1{text-align:center;font-size:25px;margin:8px 0 18px}.info,.items{width:100%;border-collapse:collapse;background:white}.info th{background:#fff2b3;width:11%}.info th,.info td,.items th,.items td{border:1px solid #d6c15c;padding:6px 7px;font-size:12px}.items{margin-top:12px}.items th{background:#ffd84d;text-align:center}.items td{text-align:center;height:24px}.items td:nth-child(2){text-align:left}.notice{white-space:pre-wrap;margin-top:12px;border:1px solid #d6c15c;background:#fffdf3;padding:10px;line-height:1.5}.pageSub{margin:8px 0 10px;border:1px solid #d6c15c;background:#fff2b3;padding:8px}.summary{width:360px;margin:18px auto 8px;border:2px solid #d0aa00;background:#fff}.summaryTitle{text-align:center!important;justify-content:center!important;background:#fff2b3;font-weight:900}.summary div{display:flex;justify-content:space-between;border-bottom:1px solid #eadb91;padding:8px 12px}.summary div:last-child{border-bottom:none}.summary .total{background:#ffd84d;font-weight:900;font-size:18px}.footerNotice{position:absolute;left:12mm;right:12mm;bottom:9mm;text-align:center;font-size:11px;color:#806600}@media print{body{background:white}.no-print{display:none}.page{margin:0}}
    </style></head><body><button class="no-print" onclick="window.print()">PDF 저장/인쇄</button>${page(page1, 0, true, items.length <= 15)}${items.length > 15 ? page(page2, 15, false, true) : ""}</body></html>`;
    const w = window.open("", "_blank");
    if (!w) return alert("팝업이 차단됐어요. 팝업 허용 후 다시 눌러줘.");
    w.document.write(html);
    w.document.close();
  }

  async function downloadLiveInvoiceExcel(order) {
    const session = liveSessions.find((s) => String(s.id) === String(order.sessionId)) || selectedLiveSession || {};
    try {
      const res = await fetch("/invoice_template.xlsx");
      if (!res.ok) throw new Error("정산서 템플릿 파일을 불러오지 못했어요.");
      const buffer = await res.arrayBuffer();
      const wb = await XlsxPopulate.fromDataAsync(buffer);
      const sheet = wb.sheet("고객정산서") || wb.sheet(0);
      sheet.cell("A1").value(`정산번호  #${order.id || ""}`);
      sheet.cell("B7").value(order.liveDate || "");
      sheet.cell("D7").value(order.buyer || "");
      sheet.cell("F7").value(order.phone || "");
      sheet.cell("H7").value(order.status === "정산후킵" ? "Y" : "N");
      sheet.cell("B8").value(order.paymentMethod || "");
      sheet.cell("D8").value((session.notice || "").split("\n")[0] || "");
      sheet.cell("F8").value(toInt(order.shipping));
      sheet.cell("H8").value("");
      sheet.cell("B9").value(orderAddressOf(order));
      sheet.cell("B10").value([session.bankName, session.accountNumber, session.accountHolder].filter(Boolean).join(" / "));
      sheet.cell("B29").value(session.notice || "입금 확인 순서대로 포장 후 출고됩니다.");
      sheet.cell("B36").value(order.liveDate || "");
      sheet.cell("D36").value(order.id || "");
      sheet.cell("F36").value(order.buyer || "");
      const rowNums = [...Array.from({length:15}, (_,i)=>13+i), ...Array.from({length:15}, (_,i)=>39+i)];
      rowNums.forEach((r, i) => {
        const it = (order.items || [])[i];
        sheet.cell(`A${r}`).value(it ? i + 1 : "");
        sheet.cell(`B${r}`).value(it ? it.name || "" : "");
        sheet.cell(`D${r}`).value(it ? toInt(it.qty) : "");
        sheet.cell(`E${r}`).value(it ? toInt(it.price) * toInt(it.qty) : "");
        sheet.cell(`F${r}`).value(it ? (String(it.prepaid).toUpperCase() === "Y" ? "Y" : "N") : "");
        sheet.cell(`G${r}`).formula(`IF(COUNTA(B${r}:F${r})=0,"",IF(UPPER(TRIM(F${r}))="Y",0,E${r}))`);
      });
      sheet.cell("E56").value("최종 결제금액");
      sheet.cell("E57").value("상품합계");
      sheet.cell("F57").value(toInt(order.paySubtotal));
      sheet.cell("E58").value("배송비");
      sheet.cell("F58").value(toInt(order.shipping));
      sheet.cell("E59").value(`카드수수료 (${order.paymentMethod === "카드결제" ? session.cardFeeRate || 0 : 0}%)`);
      sheet.cell("F59").value(toInt(order.cardFee));
      sheet.cell("E60").value("포인트 차감");
      sheet.cell("F60").value(-toInt(order.usedPoints || 0));
      try { sheet.cell("E61").style(sheet.cell("E60").style()); sheet.cell("F61").style(sheet.cell("F60").style()); } catch {}
      sheet.cell("E61").value("총 결제금액");
      sheet.cell("F61").value(toInt(order.total));
      const blob = await wb.outputAsync({ type: "blob" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `여깁니다유_라방정산서_${order.buyer || "고객"}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      alert("정산서 엑셀 생성 실패: " + String(error.message || error));
    }
  }

  function LiveOrderPage() {
    const summary = liveCartSummary();
    const sales = liveSalesSummary();
    const statusOptions = ["미입금", "입금확인", "정산후킵", "송장입력", "출고완료", "취소"];
    const matchingOrders = sameLiveMemberOrders(liveOrderForm).filter((o) => String(o.sessionId) !== String(selectedLiveSession?.id) || String(o.status) === "정산후킵");
    return (
      <section className="livePage">
        <div className="panel liveTopPanel">
          <h2>라방주문</h2>
          <div className="filterRow">
            <label>라방명</label><input value={liveNewSession.title} onChange={(e) => setLiveNewSession({ ...liveNewSession, title: e.target.value })} placeholder="예: 6/20 치이카와 라방" />
            <label>라방날짜</label><input type="date" value={liveNewSession.date} onChange={(e) => setLiveNewSession({ ...liveNewSession, date: e.target.value })} />
            <label>기본 킵기간</label><input className="tinyInput" value={liveNewSession.keepDays} onChange={(e) => setLiveNewSession({ ...liveNewSession, keepDays: e.target.value })} />일
            <label>배송비</label><input value={liveNewSession.shippingFee} onChange={(e) => setLiveNewSession({ ...liveNewSession, shippingFee: e.target.value })} />
            <label>카드%</label><input value={liveNewSession.cardFeeRate} onChange={(e) => setLiveNewSession({ ...liveNewSession, cardFeeRate: e.target.value })} />
            <button onClick={createLiveSession}>새 라방 생성</button>
            <label>라방선택</label><select value={selectedLiveSession?.id || ""} onChange={(e) => setSelectedLiveSessionId(e.target.value)}>{liveSessions.map((s) => <option key={s.id} value={s.id}>{s.date} {s.title}</option>)}</select>
          </div>
          {selectedLiveSession && <>
            <div className="filterRow">
              <label>안내사항</label><textarea className="liveNoticeInput" value={selectedLiveSession.notice || ""} onChange={(e) => updateLiveSession({ notice: e.target.value })} />
              <label>은행</label><input value={selectedLiveSession.bankName || ""} onChange={(e) => updateLiveSession({ bankName: e.target.value })} />
              <label>계좌</label><input value={selectedLiveSession.accountNumber || ""} onChange={(e) => updateLiveSession({ accountNumber: e.target.value })} />
              <label>예금주</label><input value={selectedLiveSession.accountHolder || ""} onChange={(e) => updateLiveSession({ accountHolder: e.target.value })} />
            </div>
            <div className="liveSummaryCards">
              <div><span>주문자수</span><b>{sales.buyerCount.toLocaleString()}명</b></div>
              <div><span>입금확인 주문</span><b>{sales.paidCount.toLocaleString()}건</b></div>
              <div><span>확정매출</span><b>{money(sales.total)}</b></div>
              <div><span>순수익</span><b>{money(sales.profit)}</b></div>
              <div><span>판매수량</span><b>{sales.soldQty.toLocaleString()}개</b></div>
              <div><span>라방재고</span><b>{sales.remaining}/{sales.allocated}</b></div>
              <div><span>미입금</span><b>{money(sales.unpaid)}</b></div>
              <div><span>킵 D-2/출고필요</span><b>{sales.keepDue.toLocaleString()}건</b></div>
            </div>
            <div className="buttonRow"><button onClick={downloadLiveShippingExcel}>입금확인 주문 택배접수 엑셀</button><button className="deleteBtn" onClick={deleteLiveSessionWithRestore}>라방 삭제/재고복구</button></div>
          </>}
        </div>

        <div className="liveGrid liveWorkflowGrid">
          <div className="panel liveProductPanel">
            <h2>1. 라방 상품 등록</h2>
            <p className="statusLine">라방추가 시 본재고에서 라방재고로 수량이 이동돼요. 주문 저장은 라방재고를 예약하고, 입금확인부터 매출에 반영돼요.</p>
            <div className="filterRow"><label>상품검색</label><input value={liveProductSearch} onChange={(e) => setLiveProductSearch(e.target.value)} placeholder="상품명/캐릭터/카테고리" /></div>
            <div className="tableWrap liveProductSourceTable"><table><thead><tr><th>상품명</th><th>본재고</th><th>소비자가</th><th>추가</th></tr></thead><tbody>
              {liveFilteredProducts.map((p) => <tr key={p.id}><td title={p.name}>{p.name}</td><td>{p.stock}</td><td>{money(p.retail)}</td><td><button onClick={() => addProductToLive(p)}>라방추가</button></td></tr>)}
              {liveFilteredProducts.length === 0 && <tr><td colSpan="4" className="empty">상품이 없어요.</td></tr>}
            </tbody></table></div>
            <h3>라방용 상품 목록</h3>
            <div className="tableWrap liveSelectedTable"><table><thead><tr><th>상품명</th><th>캐릭터</th><th>배정</th><th>남음</th><th>정가</th><th>할인율</th><th>라방가</th><th>담기</th><th>삭제</th></tr></thead><tbody>
              {selectedLiveProducts.map((it) => <tr key={it.id}><td title={it.name}>{it.name}</td><td>{[it.char1, it.char2].filter(Boolean).join("/")}</td><td><input className="tinyInput" value={it.liveQty} onChange={(e) => changeLiveQty(it, e.target.value)} /></td><td>{it.remainingQty}</td><td>{money(it.retail)}</td><td><input className="tinyInput" value={it.discountRate} onChange={(e) => changeLiveDiscount(it, e.target.value)} />%</td><td><input value={it.livePrice} onChange={(e) => changeLivePrice(it, e.target.value)} /></td><td><button onClick={() => addLiveItemToCart(it)}>담기</button></td><td><button className="deleteBtn" onClick={() => removeLiveItem(it.id)}>삭제</button></td></tr>)}
              {selectedLiveProducts.length === 0 && <tr><td colSpan="9" className="empty">라방에 올릴 상품을 추가해줘.</td></tr>}
            </tbody></table></div>
          </div>

          <div className="panel liveOrderPanel">
            <h2>2. 구매자 주문 생성 {editingLiveOrderId ? "(수정중)" : ""}</h2>
            <div className="filterRow liveMemberLookupRow">
              <label>회원검색</label><input className="wideInput" value={liveMemberLookupSearch} onChange={(e) => setLiveMemberLookupSearch(e.target.value)} placeholder="이름/전화번호/뒷4자리 검색" />
              {liveMemberLookupSearch && liveMemberLookupResults.length > 0 && <select value="" onChange={(e) => { const m = liveMembers.find((x) => String(x.id) === e.target.value); if (m) loadMemberToOrder(m); }}>
                <option value="">검색 결과 선택</option>{liveMemberLookupResults.map((m) => <option key={m.id} value={m.id}>{m.name} / {phoneLast4(m.phone)} / {toInt(m.points).toLocaleString()}P</option>)}
              </select>}
            </div>
            <div className="filterRow">
              <label>고객명</label><input value={liveOrderForm.buyer} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, buyer: e.target.value })} />
              <label>전화번호</label><input value={liveOrderForm.phone} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, phone: e.target.value })} />
              <label>포인트</label><input value={liveOrderForm.points} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, points: e.target.value, usedPoints: 0 })} />
              <button type="button" onClick={useAllMemberPoints}>포인트 전액사용</button>
              <span className="statusLine">사용: {money(liveOrderForm.usedPoints || 0)}</span>
              <button type="button" onClick={() => saveMemberFromOrderForm(true)}>회원저장</button>
              <label>결제</label><select value={liveOrderForm.paymentMethod} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, paymentMethod: e.target.value })}><option>계좌이체</option><option>카드결제</option></select>
            </div>
            {matchingOrders.length > 0 && <p className="statusLine dangerText">⚠ 같은 회원의 미출고/킵 주문 {matchingOrders.length}건이 있어요. 주문관리에서 합배송 묶기를 눌러 합칠 수 있어요.</p>}
            <div className="filterRow"><label>우편번호</label><input value={liveOrderForm.postalCode} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, postalCode: e.target.value })} /><button type="button" onClick={() => openDaumPostcode("order")}>우편번호 검색</button><label>기본주소</label><input className="wideInput" value={liveOrderForm.baseAddress} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, baseAddress: e.target.value, address: [e.target.value, liveOrderForm.detailAddress].filter(Boolean).join(" ") })} /></div>
            <div className="filterRow"><label>상세주소</label><input className="wideInput" value={liveOrderForm.detailAddress} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, detailAddress: e.target.value, address: [liveOrderForm.baseAddress, e.target.value].filter(Boolean).join(" ") })} /><label className="checkLine"><input type="checkbox" checked={liveOrderForm.shippingApply} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, shippingApply: e.target.checked })} />배송비 적용</label></div>
            <div className="filterRow"><label>박스무게</label><select value={liveOrderForm.boxWeight} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, boxWeight: e.target.value })}><option>2</option><option>5</option></select><label>박스부피</label><select value={liveOrderForm.boxVolume} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, boxVolume: e.target.value })}><option>60</option><option>80</option><option>100</option></select><label>내용품</label><input value={liveOrderForm.household} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, household: e.target.value })} /><label>배송메모</label><input className="wideInput" value={liveOrderForm.deliveryMessage} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, deliveryMessage: e.target.value })} /></div>
            <div className="tableWrap liveCartTable"><table><thead><tr><th>상품명</th><th>수량</th><th>금액</th><th>선결제</th><th>실결제</th><th>삭제</th></tr></thead><tbody>
              {liveCart.map((it, idx) => <tr key={`${it.liveItemId}-${idx}`}><td title={it.name}>{it.name}</td><td><input className="tinyInput" value={it.qty} onChange={(e) => updateLiveCartItem(idx, { qty: e.target.value })} /></td><td><input value={it.price} onChange={(e) => updateLiveCartItem(idx, { price: e.target.value })} /></td><td><select value={it.prepaid} onChange={(e) => updateLiveCartItem(idx, { prepaid: e.target.value })}><option>N</option><option>Y</option></select></td><td>{String(it.prepaid).toUpperCase() === "Y" ? "0원" : money(toInt(it.price) * toInt(it.qty))}</td><td><button className="deleteBtn" onClick={() => setLiveCart(liveCart.filter((_, i) => i !== idx))}>삭제</button></td></tr>)}
              {liveCart.length === 0 && <tr><td colSpan="6" className="empty">라방 상품에서 담기를 눌러줘.</td></tr>}
            </tbody></table></div>
            <p className="statusLine">상품합계 {money(summary.paySubtotal)} | 배송비 {money(summary.shipping)} | 카드수수료 {money(summary.cardFee)} ({liveOrderForm.paymentMethod === "카드결제" ? selectedLiveSession?.cardFeeRate || 0 : 0}%) | 포인트차감 -{money(summary.usedPoints)} | 최종 {money(summary.total)}</p>
            <div className="filterRow"><label>주문메모</label><input className="wideInput" value={liveOrderForm.memo} onChange={(e) => setLiveOrderForm({ ...liveOrderForm, memo: e.target.value })} /><button onClick={saveLiveOrderAndDeduct}>{editingLiveOrderId ? "주문수정 저장" : "미입금 주문저장"}</button>{editingLiveOrderId && <button type="button" onClick={cancelLiveOrderEdit}>수정취소</button>}</div>
          </div>

          <div className="panel liveManagePanel">
            <h2>3. 주문관리</h2>
            <div className="filterRow">
              <label>주문검색</label><input value={liveOrderSearch} onChange={(e) => setLiveOrderSearch(e.target.value)} placeholder="구매자/전화/송장/메모" />
              <label className="checkLine"><input type="checkbox" checked={liveDueOnly} onChange={(e) => setLiveDueOnly(e.target.checked)} />출고필요만 보기</label>
              <span className="statusLine">상태 변경은 각 주문 행의 드롭다운에서 저장돼요.</span>
            </div>
            <div className="tableWrap liveOrdersTable"><table><thead><tr><th>구매자</th><th>라방일</th><th>금액</th><th>포인트</th><th>상태</th><th>킵</th><th>송장</th><th>묶음</th><th>정산서</th><th>수정</th><th>취소</th><th>삭제</th></tr></thead><tbody>
              {liveFilteredOrders.map((o) => <tr key={o.id} className={o.canceledAt ? "dangerRow" : isLiveKeepDueSoon(o) ? "dangerRow" : o.locked ? "lockedRow" : ""}><td>{o.buyer}<br/><small>{phoneLast4(o.phone)}</small></td><td>{o.liveDate}</td><td>{money(o.total)}</td><td>{toInt(o.usedPoints) > 0 ? `-${money(o.usedPoints)}` : "-"}</td><td><select disabled={o.locked} value={o.status} onChange={(e) => updateLiveOrder(o.id, { status: e.target.value })}>{statusOptions.map((s) => <option key={s}>{s}</option>)}</select></td><td>{liveOrderKeepDday(o)}</td><td><input disabled={o.locked || o.status === "정산후킵"} value={o.trackingNo || ""} onChange={(e) => updateLiveOrder(o.id, { trackingNo: e.target.value, status: e.target.value ? "송장입력" : o.status })} /></td><td><button onClick={() => bundleLiveOrdersFor(o)}>{o.bundleId ? "묶임" : "합치기"}</button></td><td><button onClick={() => downloadLiveInvoiceExcel(o)}>엑셀</button><button onClick={() => openLiveInvoicePdf(o)}>PDF</button><button type="button" onClick={() => updateLiveOrder(o.id, { locked: !o.locked })}>{o.locked ? "해제" : "잠금"}</button></td><td><button type="button" disabled={!!o.canceledAt || o.locked} onClick={() => beginEditLiveOrder(o)}>수정</button></td><td><button className="deleteBtn" disabled={!!o.canceledAt || o.locked} onClick={() => cancelLiveOrderWithRestore(o)}>{o.canceledAt ? "취소됨" : "취소"}</button></td><td><button className="deleteBtn" disabled={o.locked} onClick={() => deleteLiveOrderWithRestore(o)}>삭제</button></td></tr>)}
              {liveFilteredOrders.length === 0 && <tr><td colSpan="12" className="empty">주문 기록이 없어요.</td></tr>}
            </tbody></table></div>
            <h3>통합 회원관리</h3>
            <div className="filterRow"><label>검색</label><input value={liveMemberSearch} onChange={(e) => setLiveMemberSearch(e.target.value)} /><label>이름</label><input value={liveMemberForm.name} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, name: e.target.value })} /><label>연락처</label><input value={liveMemberForm.phone} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, phone: e.target.value })} /><label>포인트</label><input value={liveMemberForm.points} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, points: e.target.value })} /><button onClick={saveLiveMember}>회원 저장</button></div>
            <div className="filterRow"><label>우편번호</label><input value={liveMemberForm.postalCode} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, postalCode: e.target.value })} /><button type="button" onClick={() => openDaumPostcode("member")}>우편번호 검색</button><label>기본주소</label><input className="wideInput" value={liveMemberForm.baseAddress} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, baseAddress: e.target.value, address: [e.target.value, liveMemberForm.detailAddress].filter(Boolean).join(" ") })} /><label>상세주소</label><input className="wideInput" value={liveMemberForm.detailAddress} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, detailAddress: e.target.value, address: [liveMemberForm.baseAddress, e.target.value].filter(Boolean).join(" ") })} /></div>
            <div className="filterRow"><label>회원메모</label><input className="wideInput" value={liveMemberForm.memo} onChange={(e) => setLiveMemberForm({ ...liveMemberForm, memo: e.target.value })} /></div>
            <div className="tableWrap liveMembersTable"><table><thead><tr><th>이름</th><th>연락처</th><th>주소</th><th>포인트</th><th>사용누적</th><th>메모</th><th>관리</th></tr></thead><tbody>
              {liveFilteredMembers.map((m) => <tr key={m.id}><td>{m.name}</td><td>{m.phone}</td><td title={m.address}>{m.postalCode} {m.baseAddress} {m.detailAddress}</td><td>{toInt(m.points).toLocaleString()}</td><td>{toInt(m.usedPoints).toLocaleString()}</td><td title={m.memo}>{m.memo}</td><td><button onClick={() => setLiveMemberForm(m)}>수정</button><button onClick={() => loadMemberToOrder(m)}>주문에 불러오기</button><button className="deleteBtn" onClick={() => deleteLiveMember(m)}>삭제</button></td></tr>)}
              {liveFilteredMembers.length === 0 && <tr><td colSpan="7" className="empty">저장된 회원이 없어요.</td></tr>}
            </tbody></table></div>
          </div>
        </div>
      </section>
    );
  }

  function ComposePage() {
    return (
      <>
        <section className="panel composePageFixed">{FilterBox()}</section>
        <section className="splitLayout manualAiOnlyLayout">
          <div className="panel manualProductListPanel">
            <h2>조건 상품 리스트</h2>
            <div className="composeConditionResultArea"><ProductTable mode="compose" /></div>
          </div>
          <div className="panel manualWorkPanel">
            <h2>현재 조합 리스트</h2>
            <div className="tableWrap composeNow">
              <table><thead><tr><th>상품ID</th><th>상품명</th><th>재고</th><th>도매가</th><th>소비자가</th><th>체크/수정</th></tr></thead><tbody>
                {composeItems.map((p, i) => <tr key={`${p.id}-${i}`}><td>{p.id}</td><td title={p.name}>{p.name}</td><td>{p.stock}</td><td>{money(p.wholesale)}</td><td>{money(p.retail)}</td><td><button className="deleteBtn" onClick={() => setComposeItems(composeItems.filter((_, idx) => idx !== i))}>삭제</button></td></tr>)}
                {composeItems.length === 0 && <tr><td colSpan="6" className="empty">아직 조합한 상품이 없어요.</td></tr>}
              </tbody></table>
            </div>

            <div className="filterRow calcRow">
              <label>판매가</label><input value={salePrice} onChange={(e) => setSalePrice(e.target.value)} />
              <label>수수료율</label><input value={feeRate} onChange={(e) => setFeeRate(e.target.value)} />
              <label>목표마진율</label><input value={manualTargetMargin} onChange={(e) => setManualTargetMargin(e.target.value)} />
              <label>고객명</label><input value={customer} onChange={(e) => setCustomer(e.target.value)} />
              <label className="checkLine"><input checked={reorder} onChange={(e) => setReorder(e.target.checked)} type="checkbox" /> 재주문</label>
              <label>메모</label><input value={memo} onChange={(e) => setMemo(e.target.value)} />
              <button onClick={clearCompose}>주문초기화</button>
              <button onClick={createOrderFromCompose}>박스출고</button>
            </div>

            <p className="statusLine">도매가합 {money(finance.wholesaleSum)} | 소비자가합 {money(finance.retailSum)} | 수수료 {money(finance.feeAmount)} | 실수령액 {money(finance.netAmount)} | 순이익 {money(finance.profit)} | 마진율 {finance.margin.toFixed(1)}%</p>

            <div className="subPanel manualAiPanel">
              <h2>AI 수동박스 추천</h2>
              <p className="statusLine">위 통합 입력값(박스수·판매가·수수료율·목표마진율·고객명·메모)과 선택 캐릭터를 기준으로 AI가 추천합니다.</p>

              <div className="filterRow aiSimpleControls manualAiTopControls">
                <label>구성느낌</label><select value={manualStyle} onChange={(e) => setManualStyle(e.target.value)}><option>선택안함</option><option>자잘자잘</option><option>믹스</option><option>큼직큼직</option></select>
                <label>캐릭터비중</label><select value={manualCharStrategy} onChange={(e) => setManualCharStrategy(e.target.value)}><option>골고루</option><option>선택 캐릭터 위주</option><option>완전 랜덤</option><option>재고 많은 캐릭터 우선</option></select>
              </div>

              <div className="characterPickPanel">
                <h3>캐릭터 선택</h3>
                <p className="statusLine">버튼으로 눌러두면 AI가 선택 캐릭터 안에서 우선 구성해요. 선택하지 않으면 전체 재고에서 무작위/균형으로 추천합니다.</p>
                <MultiCheckFilter label="캐릭터1" options={char1Options} selected={manualPrefChar1} setSelected={setManualPrefChar1} />
                <MultiCheckFilter label="캐릭터2" options={char2Options} selected={manualPrefChar2} setSelected={setManualPrefChar2} />
              </div>

              <div className="aiPresetGrid">
                <div className="aiRequestBox">
                  <label>고객 요청사항 / 내가 원하는 조건</label>
                  <textarea
                    value={manualAiRequest}
                    onChange={(e) => setManualAiRequest(e.target.value)}
                    placeholder={"예: 시나모롤 위주, 문구류 적게, 실용적인 소품 위주\n예: 재주문 고객이면 지난번 상품은 빼고 5만원 프리미엄 느낌으로"}
                  />
                </div>
                <div className="aiRequestBox">
                  <label>AI 추천안 수정 요청</label>
                  <textarea
                    value={manualAiRevision}
                    onChange={(e) => setManualAiRevision(e.target.value)}
                    placeholder={"예: 1번 후보에서 키링 빼고 파우치로 바꿔줘\n예: 소비자가를 조금 더 높이고 마진율 20% 이상으로 다시 맞춰줘"}
                  />
                </div>
                <div className="aiPresetBox">
                  <label>저장된 조건</label>
                  <select onChange={(e) => e.target.value && loadManualAiPreset(e.target.value)} value="">
                    <option value="">조건 선택</option>
                    {manualSavedPresets.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                  </select>
                  <input value={manualPresetName} onChange={(e) => setManualPresetName(e.target.value)} placeholder="조건 이름 예: 키티 실용템 5만원" />
                  <div className="buttonRow compactButtons">
                    <button type="button" onClick={saveManualAiPreset}>현재 조건 저장</button>
                    <button type="button" className="deleteBtn" onClick={() => {
                      const id = window.prompt("삭제할 조건 번호를 선택창에서 확인 후 입력하거나, 조건을 다시 저장해 새로 사용해줘.");
                      if (id) deleteManualAiPreset(id);
                    }}>조건 삭제</button>
                  </div>
                  <button type="button" className="primaryAiBtn fullBtn" onClick={askAiForManualRecommendation} disabled={manualAiLoading}>{manualAiLoading ? "AI 추천 중..." : "AI가 추천안 짜기 / 수정하기"}</button>
                </div>
              </div>

              {manualAiMemo && <textarea className="aiMemoBox manualAiMemoBox" value={manualAiMemo} onChange={(e) => setManualAiMemo(e.target.value)} />}
              <div className="buttonRow">
                <button type="button" onClick={() => exportGeminiAnswerToManualBox(manualAiMemo)}>AI 답변 수동박스로 내보내기</button>
                <button type="button" onClick={() => setManualAiRevision("")}>수정요청 비우기</button>
              </div>
            </div>
          </div>
        </section>
      </>
    );
  }


  async function deleteSavedScoopGroupByPicker() {
    const { data, error } = await supabase
      .from("saved_scoop_groups")
      .select("*")
      .order("id", { ascending: false });

    if (error) return alert("저장 그룹 목록 불러오기 실패: " + error.message);
    if (!data || data.length === 0) return alert("삭제할 저장 그룹이 없어요.");

    const listText = data.map((g, idx) => {
      const title = g.name || g.title || g.group_name || `저장그룹 ${g.id}`;
      const date = String(g.created_at || "").slice(0, 10);
      return `${idx + 1}. ID ${g.id} | ${title}${date ? " | " + date : ""}`;
    }).join("\n");

    const pick = window.prompt(
      "삭제할 저장 그룹 번호를 입력해줘.\n\n" +
      listText +
      "\n\n예: 1"
    );

    if (!pick) return;

    const idx = Number(pick) - 1;
    if (Number.isNaN(idx) || idx < 0 || idx >= data.length) {
      alert("번호가 올바르지 않아요.");
      return;
    }

    const target = data[idx];
    const title = target.name || target.title || target.group_name || `저장그룹 ${target.id}`;

    const ok = window.confirm(
      `저장 그룹을 삭제할까요?\n\n` +
      `ID: ${target.id}\n` +
      `이름: ${title}\n\n` +
      "삭제 후 되돌릴 수 없어요."
    );
    if (!ok) return;

    const { error: delError } = await supabase
      .from("saved_scoop_groups")
      .delete()
      .eq("id", target.id);

    if (delError) return alert("저장 그룹 삭제 실패: " + delError.message);

    alert("저장 그룹 삭제 완료!");
  }

  async function permanentlyDeleteOrder(orderId) {
    const order = orders.find((o) => o.id === orderId);
    if (!order) return alert("주문을 찾을 수 없어요.");
    if (!order.deleted_at) return alert("취소보관함 주문만 영구삭제할 수 있어요.");

    const ok = window.confirm(
      `주문ID ${orderId}를 영구삭제할까요?\n\n` +
      "영구삭제하면 주문 기록과 주문상품 기록이 완전히 삭제됩니다.\n" +
      "재고는 이미 취소 시 복구되었으므로 여기서는 재고 변화가 없습니다."
    );
    if (!ok) return;

    await supabase.from("order_items").delete().eq("order_id", orderId);
    const { error } = await supabase.from("orders").delete().eq("id", orderId);
    if (error) return alert("영구삭제 실패: " + error.message);

    await writeAudit("order_permanent_delete", `order_id=${orderId}`);
    alert("영구삭제 완료!");
    getOrders();
    getOrderItems();
  }

  function OrderTable({ title, rows }) {
    return (
      <div className="orderBox">
        <h3>{title}</h3>
        <div className="tableWrap">
          <table><thead><tr><th>주문ID</th><th>주문일</th><th>주문자</th><th>재주문</th><th>상태</th><th>판매가</th><th>실수령액</th><th>순이익</th><th>취소사유</th></tr></thead><tbody>
            {rows.map((o) => <tr key={o.id} onClick={(e) => { if (["INPUT","TEXTAREA","SELECT","BUTTON"].includes(e.target.tagName)) return; setSelectedOrderId(o.id); }} className={selectedOrderId === o.id ? "selectedRow" : ""}><td>{o.id}</td><td>{String(o.created_at || "").replace("T", " ").slice(0, 19)}</td><td>{o.customer}</td><td>{toInt(o.reorder) === 1 ? "Y" : ""}</td><td>{o.status}</td><td>{money(o.sale_price)}</td><td>{money(o.net_amount)}</td><td>{money(o.profit)}</td><td>{o.cancel_reason || ""}</td></tr>)}
            {rows.length === 0 && <tr><td colSpan="9" className="empty">표시할 주문이 없어요.</td></tr>}
          </tbody></table>
        </div>
      </div>
    );
  }


  const selectedOrderItemsRows = useMemo(() => {
    if (!selectedOrderId) return [];
    return orderItems.filter((x) => String(x.order_id) === String(selectedOrderId));
  }, [orderItems, selectedOrderId]);

  const selectedOrderItemsWholesaleTotal = useMemo(() => {
    return selectedOrderItemsRows.reduce((sum, x) => sum + toInt(x.wholesale || x.wholesale_price || x.cost || 0) * toInt(x.qty || 1), 0);
  }, [selectedOrderItemsRows]);

  const selectedOrderItemsRetailTotal = useMemo(() => {
    return selectedOrderItemsRows.reduce((sum, x) => sum + toInt(x.retail || x.retail_price || x.consumer_price || 0) * toInt(x.qty || 1), 0);
  }, [selectedOrderItemsRows]);

  function openSelectedOrderItemsPanel() {
    if (!selectedOrderId) return alert("주문을 선택해줘.");
    setSelectedOrderItemsOpen(true);
  }


  const v50SelectedOrderItemsRows = useMemo(() => {
    if (!selectedOrderId) return [];
    return orderItems.filter((x) => String(x.order_id) === String(selectedOrderId));
  }, [orderItems, selectedOrderId]);

  const v50SelectedOrderWholesaleTotal = useMemo(() => {
    return v50SelectedOrderItemsRows.reduce((sum, x) => sum + toInt(x.wholesale || x.wholesale_price || x.cost || 0) * toInt(x.qty || 1), 0);
  }, [v50SelectedOrderItemsRows]);

  const v50SelectedOrderRetailTotal = useMemo(() => {
    return v50SelectedOrderItemsRows.reduce((sum, x) => sum + toInt(x.retail || x.retail_price || x.consumer_price || 0) * toInt(x.qty || 1), 0);
  }, [v50SelectedOrderItemsRows]);


  const v85SelectedOrder = useMemo(() => orders.find((o) => String(o.id) === String(selectedOrderId)), [orders, selectedOrderId]);
  const v85SelectedCustomerInfo = v85SelectedOrder ? `${v85SelectedOrder.customer || v85SelectedOrder.customer_name || "-"} / ${v85SelectedOrder.status || "-"} / ${v85SelectedOrder.memo || ""}` : "주문을 클릭하면 고객정보와 상품목록이 여기에 표시됩니다.";
  function OrdersPage() {
    return (
      <div className="ordersPageFit">
        <section className="panel orderTopPanel stickyControlPanel">
          <div className="filterRow">
            <label>주문자명</label><input value={orderSearchCustomer} onChange={(e) => setOrderSearchCustomer(e.target.value)} />
            <label>주문일</label><input value={orderSearchDate} onChange={(e) => setOrderSearchDate(e.target.value)} placeholder="YYYY-MM-DD" />
            <label className="checkLine"><input checked={orderReorderOnly} onChange={(e) => setOrderReorderOnly(e.target.checked)} type="checkbox" /> 재구매자만</label>
            <button onClick={getOrders}>검색</button>
            <button onClick={() => { setOrderSearchCustomer(""); setOrderSearchDate(""); setOrderReorderOnly(false); setSelectedOrderId(null); }}>초기화</button>
            <button onClick={shipSelectedOrder}>출고확정</button>
            <button className="deleteBtn" onClick={cancelSelectedOrder}>주문취소</button>
            <button onClick={() => selectedOrderId ? copyOrderToManualComposition(selectedOrderId) : alert("복사할 주문을 선택해줘.")}>구성복사 수동박스</button>
            <button onClick={downloadOrdersExcel}>주문 엑셀</button>
            <button onClick={downloadCustomerOrderExcel}>고객용 엑셀</button>
          </div>
          <p className="statusLine">선택된 주문ID: {selectedOrderId || "-"}</p>
        </section>
        <section className="ordersGrid fixedOrdersGrid"><OrderTable title="주문접수 / 재고임시차감" rows={pendingOrders} /><OrderTable title="출고확정 / 발송완료" rows={shippedOrders} /></section>
        <section className="panel orderDetailPanel v51OrderDetailPanel">
          <div className="v51OrderDetailHeader">
            <h2>선택 주문 상품목록</h2><p className="statusLine">{v85SelectedCustomerInfo}</p>
            <div className="v51OrderTotals">
              <b>주문ID:</b> {selectedOrderId || "-"}　
              <b>총 도매가합:</b> {money(selectedOrderItemsWholesaleTotal || v50SelectedOrderWholesaleTotal || 0)}　
              <b>총 소비자가합:</b> {money(selectedOrderItemsRetailTotal || v50SelectedOrderRetailTotal || 0)}
            </div>
          </div>
          <div className="tableWrap v51OrderItemsTableWrap">
            <table>
              <thead>
                <tr>
                  <th>번호</th>
                  <th>상품ID</th>
                  <th>상품명</th>
                  <th>캐릭터1</th>
                  <th>캐릭터2</th>
                  <th>카테고리</th>
                  <th>수량</th>
                  <th>도매가</th>
                  <th>소비자가</th>
                </tr>
              </thead>
              <tbody>
                {(selectedOrderItemsRows.length ? selectedOrderItemsRows : v50SelectedOrderItemsRows).map((x, i) => (
                  <tr key={x.id || i}>
                    <td>{i + 1}</td>
                    <td>{x.product_id || x.id}</td>
                    <td>{x.product_name || x.name || x.item_name || "-"}</td>
                    <td>{x.char1 || products.find((p) => String(p.id) === String(x.product_id))?.char1 || "-"}</td>
                    <td>{x.char2 || products.find((p) => String(p.id) === String(x.product_id))?.char2 || "-"}</td>
                    <td>{x.category || products.find((p) => String(p.id) === String(x.product_id))?.category || "-"}</td>
                    <td>{x.qty || 1}</td>
                    <td>{money(x.wholesale || x.wholesale_price || x.cost || 0)}</td>
                    <td>{money(x.retail || x.retail_price || x.consumer_price || 0)}</td>
                  </tr>
                ))}
                {!selectedOrderId && <tr><td colSpan="9" className="empty">주문접수 또는 출고완료 표에서 주문을 클릭하면 상품목록이 여기에 표시됩니다.</td></tr>}
                {selectedOrderId && (selectedOrderItemsRows.length ? selectedOrderItemsRows : v50SelectedOrderItemsRows).length === 0 && <tr><td colSpan="9" className="empty">이 주문의 상품목록이 없어요.</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    );
  }


  async function copyOrderToManualComposition(orderId) {
    const oldOrder = orders.find((o) => o.id === orderId);
    if (!oldOrder) return alert("주문을 찾을 수 없어요.");

    const rows = orderItems.filter((x) => x.order_id === orderId);
    if (rows.length === 0) return alert("복사할 주문상품이 없어요.");

    const buildCopiedItems = (stockRestoreMap = {}) => {
      const items = [];
      const missingRows = [];
      const shortageRows = [];

      for (const x of rows) {
        const p = products.find((prod) => String(prod.id) === String(x.product_id));
        if (!p) {
          missingRows.push(`${x.name || "상품명 없음"} / 상품ID ${x.product_id}`);
          continue;
        }

        const qty = toInt(x.qty || 1);
        const available = toInt(p.stock) + toInt(stockRestoreMap[String(p.id)] || 0);
        if (available < qty) {
          shortageRows.push(`${p.name} | 필요 ${qty}개 / 현재 ${p.stock}개`);
        }

        const copiedProduct = { ...p, stock: available };
        for (let i = 0; i < qty; i++) items.push(copiedProduct);
      }

      return { items, missingRows, shortageRows };
    };

    const restoreMap = {};
    rows.forEach((x) => {
      const key = String(x.product_id);
      restoreMap[key] = (restoreMap[key] || 0) + toInt(x.qty || 1);
    });

    let { items: copiedItems, missingRows, shortageRows } = buildCopiedItems();

    if (missingRows.length > 0) {
      alert(
        "현재 재고 목록에서 찾을 수 없는 상품이 있어요.\n" +
        "삭제된 상품은 수동박스 조합으로 복사할 수 없어요.\n\n" +
        missingRows.join("\n")
      );
      return;
    }

    let orderCanceledForCopy = false;
    if (shortageRows.length > 0) {
      if (oldOrder.status === "취소" || oldOrder.deleted_at) {
        alert(
          "복사하려는 구성 중 현재 재고가 부족한 상품이 있어요.\n" +
          "이미 취소된 주문이라 재고를 추가 복구할 수 없어서 수동박스로 복사하지 않았습니다.\n\n" +
          shortageRows.join("\n")
        );
        return;
      }

      const afterRestoreCheck = buildCopiedItems(restoreMap);
      if (afterRestoreCheck.shortageRows.length > 0) {
        alert(
          "이 주문을 취소해 재고를 복구해도 부족한 상품이 있어요.\n" +
          "수동박스로 복사하지 않았습니다.\n\n" +
          afterRestoreCheck.shortageRows.join("\n")
        );
        return;
      }

      const okRestore = window.confirm(
        "복사하려는 구성 중 현재 재고가 부족한 상품이 있어요.\n" +
        "이 주문에서 마지막 재고가 이미 임시차감된 상태일 수 있습니다.\n\n" +
        shortageRows.join("\n") +
        "\n\n해당 주문건을 취소하고 재고를 복구한 뒤 수동박스로 복사하시겠어요?\n" +
        "확인하면 기존 주문은 취소보관함으로 이동하고, 같은 구성은 수동박스 현재 조합으로 옮겨집니다."
      );
      if (!okRestore) return;

      const restored = await restoreStockByOrder(orderId);
      if (!restored) return;

      const { error } = await supabase.from("orders").update({
        status: "취소",
        cancel_reason: "수동박스 재복사",
        cancel_detail: "재고 0개 상품 포함으로 주문 취소 후 수동박스 복사",
        canceled_at: nowString(),
        deleted_at: nowString(),
      }).eq("id", orderId);
      if (error) return alert("주문 취소 처리 실패: " + error.message);

      await writeAudit("order_cancel_for_manual_copy", `order_id=${orderId}`);
      orderCanceledForCopy = true;
      copiedItems = afterRestoreCheck.items;
    }

    const ok = window.confirm(
      `주문ID ${orderId}의 구성을 수동박스 현재 조합 리스트로 복사할까요?\n\n` +
      `상품 수: ${copiedItems.length}개\n` +
      `주문자명: ${oldOrder.customer || ""}\n` +
      `판매가: ${money(oldOrder.sale_price)}\n\n` +
      (orderCanceledForCopy ? "기존 주문은 이미 취소 처리했고 재고를 복구했습니다.\n" : "") +
      "복사 후 수동박스 화면에서 상품 목록을 확인하고 박스출고를 눌러야 새 주문이 생성됩니다."
    );
    if (!ok) {
      if (orderCanceledForCopy) alert("기존 주문은 이미 취소/재고복구 처리됐어요. 필요하면 주문관리 취소보관함에서 확인해줘.");
      return;
    }

    setComposeItems(copiedItems);
    setCustomer(oldOrder.customer || "");
    setReorder(toInt(oldOrder.reorder) === 1);
    setMemo(`주문ID ${orderId} 구성 재복사`);
    setSalePrice(String(toInt(oldOrder.sale_price || salePrice || defaultSale)));
    setFeeRate(String(oldOrder.fee_rate ?? feeRate ?? defaultFee));

    setActiveTab("수동박스");
    if (orderCanceledForCopy) {
      await Promise.all([getProducts(), getOrders(), getOrderItems()]);
      setSelectedOrderId(null);
    }
    alert("구성 복사 완료!\n수동박스의 현재 조합 리스트에서 확인한 뒤 박스출고를 눌러주세요.");
  }


  function parsePastedTsv(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      const next = text[i + 1];

      if (ch === '"') {
        if (inQuotes && next === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === "\t" && !inQuotes) {
        row.push(cell);
        cell = "";
      } else if ((ch === "\n" || ch === "\r") && !inQuotes) {
        if (ch === "\r" && next === "\n") i++;
        row.push(cell);
        if (row.some((x) => String(x).trim() !== "")) rows.push(row);
        row = [];
        cell = "";
      } else {
        cell += ch;
      }
    }

    row.push(cell);
    if (row.some((x) => String(x).trim() !== "")) rows.push(row);
    return rows;
  }

  function normalizePhone(v) {
    return String(v || "").trim();
  }

  function normalizeZip(v) {
    return String(v || "").replace(/[^0-9]/g, "").trim();
  }

  function convertShippingPaste() {
    const parsed = parsePastedTsv(shippingPasteText);
    if (parsed.length === 0) return alert("붙여넣은 주문 데이터가 없어요.");

    const converted = parsed.map((cols, index) => {
      const clean = cols.map((x) => String(x ?? "").trim());
      return {
        id: Date.now() + index,
        selected: false,
        receiverName: clean[0] || "",
        zipcode: normalizeZip(clean[9] || ""),
        baseAddress: clean[6] || "",
        detailAddress: clean[7] || "",
        receiverPhone: normalizePhone(clean[5] || clean[8] || ""),
        boxWeight: "2",
        boxVolume: "60",
        boxCount: "1",
        content: "생활용품",
        deliveryMessage: clean.slice(10).join("\n").trim(),
      };
    }).filter((x) => x.receiverName || x.zipcode || x.baseAddress || x.receiverPhone);

    if (converted.length === 0) return alert("변환할 수 있는 주문 데이터가 없어요. 복사한 데이터 순서를 확인해줘.");

    setShippingRows(converted);
    alert(`${converted.length}건을 택배접수 양식으로 변환했어요.`);
  }

  function updateShippingRow(id, key, value) {
    setShippingRows((prev) => prev.map((row) => row.id === id ? { ...row, [key]: value } : row));
  }

  function toggleShippingRow(id) {
    setShippingRows((prev) => prev.map((row) => row.id === id ? { ...row, selected: !row.selected } : row));
  }

  function deleteSelectedShippingRows() {
    const count = shippingRows.filter((x) => x.selected).length;
    if (count === 0) return alert("삭제할 행을 체크해줘.");
    if (!window.confirm(`${count}건을 삭제할까요?`)) return;
    setShippingRows((prev) => prev.filter((x) => !x.selected));
  }

  function clearShippingRows() {
    if (shippingRows.length === 0) return;
    if (!window.confirm("택배접수 목록을 모두 비울까요?")) return;
    setShippingRows([]);
  }

  function downloadShippingExcel() {
    if (shippingRows.length === 0) return alert("다운로드할 택배접수 목록이 없어요.");

    const rows = shippingRows.map((row) => [
      row.receiverName,
      row.zipcode,
      row.baseAddress,
      row.detailAddress,
      row.receiverPhone,
      row.boxWeight,
      row.boxVolume,
      row.boxCount,
      "생활용품",
      row.deliveryMessage,
    ]);

    const ws = XLSX.utils.aoa_to_sheet(rows);

    ws["!cols"] = [
      { wch: 12 }, { wch: 10 }, { wch: 42 }, { wch: 24 }, { wch: 16 },
      { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 45 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "택배접수");
    XLSX.writeFile(wb, "택배접수.xlsx");

    const clearOk = window.confirm("엑셀 다운로드가 완료됐어요.\n\n택배접수 목록도 삭제할까요?");
    if (clearOk) {
      setShippingRows([]);
      setShippingPasteText("");
    }
  }

  function ShippingRegisterPage() {
    return (
      <>
        <section className="panel shippingRegisterPage">
          <h2>택배접수</h2>
          <p className="statusLine">
            네이버 주문 데이터를 그대로 복사해서 붙여넣으면 우체국 접수 양식으로 자동 정리됩니다.
            내용품은 항상 생활용품으로 들어가고, 배송메시지는 아래 표에서 직접 수정할 수 있어요.
          </p>

          <textarea
            className="shippingPasteBox"
            value={shippingPasteText}
            onChange={(e) => setShippingPasteText(e.target.value)}
            placeholder={"수취인명\\t상품명\\t옵션상품\\t좋아하는캐릭터/영상촬영\\t수량\\t수취인연락처\\t기본주소\\t상세주소\\t구매자연락처\\t우편번호\\t배송메시지\\n여기에 주문 데이터를 그대로 붙여넣어줘."}
          />

          <div className="buttonRow">
            <button type="button" onClick={convertShippingPaste}>자동 변환</button>
            <button type="button" onClick={downloadShippingExcel}>엑셀 다운로드</button>
            <button type="button" onClick={deleteSelectedShippingRows}>선택 삭제</button>
            <button type="button" className="deleteBtn" onClick={clearShippingRows}>전체 삭제</button>
          </div>

          <p className="statusLine">변환된 택배접수 건수: {shippingRows.length.toLocaleString()}건</p>
        </section>

        <section className="panel shippingTablePanel">
          <h3>택배접수 목록</h3>
          <div className="tableWrap shippingTableWrap">
            <table>
              <thead>
                <tr>
                  <th>선택</th>
                  <th>수취인명</th>
                  <th>우편번호</th>
                  <th>기본주소</th>
                  <th>상세주소</th>
                  <th>수취인연락처</th>
                  <th>박스무게</th>
                  <th>박스부피</th>
                  <th>박스수량</th>
                  <th>내용품</th>
                  <th>배송메시지</th>
                </tr>
              </thead>
              <tbody>
                {shippingRows.map((row) => (
                  <tr key={row.id}>
                    <td><input type="checkbox" checked={row.selected} onChange={() => toggleShippingRow(row.id)} /></td>
                    <td><input value={row.receiverName} onChange={(e) => updateShippingRow(row.id, "receiverName", e.target.value)} /></td>
                    <td><input value={row.zipcode} onChange={(e) => updateShippingRow(row.id, "zipcode", e.target.value)} /></td>
                    <td><input className="shippingAddressInput" value={row.baseAddress} onChange={(e) => updateShippingRow(row.id, "baseAddress", e.target.value)} /></td>
                    <td><input className="shippingDetailInput" value={row.detailAddress} onChange={(e) => updateShippingRow(row.id, "detailAddress", e.target.value)} /></td>
                    <td><input value={row.receiverPhone} onChange={(e) => updateShippingRow(row.id, "receiverPhone", e.target.value)} /></td>
                    <td>
                      <select value={row.boxWeight} onChange={(e) => updateShippingRow(row.id, "boxWeight", e.target.value)}>
                        <option value="2">2</option>
                        <option value="5">5</option>
                        <option value="10">10</option>
                      </select>
                    </td>
                    <td>
                      <select value={row.boxVolume} onChange={(e) => updateShippingRow(row.id, "boxVolume", e.target.value)}>
                        <option value="60">60</option>
                        <option value="80">80</option>
                        <option value="100">100</option>
                      </select>
                    </td>
                    <td><input className="boxCountInput" value={row.boxCount} onChange={(e) => updateShippingRow(row.id, "boxCount", e.target.value)} /></td>
                    <td>생활용품</td>
                    <td>
                      <textarea
                        className="shippingMessageInput"
                        value={row.deliveryMessage}
                        onChange={(e) => updateShippingRow(row.id, "deliveryMessage", e.target.value)}
                      />
                    </td>
                  </tr>
                ))}
                {shippingRows.length === 0 && (
                  <tr><td colSpan="11" className="empty">붙여넣기 후 자동 변환을 누르면 목록이 표시됩니다.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </>
    );
  }

  function TrashPage() {
    return (
      <section className="panel trashPage">
        <h2>취소보관함</h2>
        <p className="statusLine">취소된 주문은 30일 보관용으로 표시됩니다. 재고는 취소 시 이미 복구됩니다.</p>
        <div className="tableWrap trashTable">
          <table>
            <thead>
              <tr>
                <th>주문ID</th><th>주문자</th><th>취소사유</th><th>메모</th><th>취소일</th><th>보관 남은일</th><th>판매가</th><th>순이익</th><th>수동박스복사</th><th>영구삭제</th>
              </tr>
            </thead>
            <tbody>
              {trashOrders.map((o) => (
                <tr key={o.id}>
                  <td>{o.id}</td>
                  <td>{o.customer}</td>
                  <td>{o.cancel_reason || "-"}</td>
                  <td>{o.cancel_detail || ""}</td>
                  <td>{o.deleted_at || o.canceled_at || "-"}</td>
                  <td>{daysLeftForTrash(o)}일</td>
                  <td>{money(o.sale_price)}</td>
                  <td>{money(o.profit)}</td>
                  <td><button onClick={() => copyOrderToManualComposition(o.id)}>수동박스로 복사</button></td>
                  <td><button className="deleteBtn" onClick={() => permanentlyDeleteOrder(o.id)}>영구삭제</button></td>
                </tr>
              ))}
              {trashOrders.length === 0 && <tr><td colSpan="10" className="empty">취소보관함이 비어 있어요.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    );
  }


  function SettingsPage() {
    return (
      <section className="manualPage">
        <h2>사용 설명서</h2>

        <div className="manualGrid">
          <div className="manualCard">
            <h3>1. AI입고분석</h3>
            <p>거래명세서, 영수증, 주문내역서, 이미지/PDF를 올리면 AI가 상품명, 수량, 캐릭터, 카테고리, 도매가, 소비자가를 분석합니다.</p>
            <p>분석 결과는 바로 반영하지 않고, 표에서 직접 수정한 뒤 재고에 반영할 수 있습니다.</p>
            <p>최근 3일 안에 추가/수정된 상품은 재고관리에서 NEW 표시로 구분됩니다.</p>
          </div>

          <div className="manualCard">
            <h3>2. 재고관리</h3>
            <p>상품을 직접 추가하거나 엑셀을 불러와 재고를 등록합니다.</p>
            <p>상품명, 캐릭터1/2, 카테고리, 재고, 도매가, 소비자가 기준으로 검색·정렬할 수 있습니다.</p>
            <p>엑셀 불러오기 전 재고는 자동 백업되며, 재고관리에서 최근 백업 복구와 백업 목록 선택복구가 가능합니다.</p>
          </div>

          <div className="manualCard">
            <h3>3. 수동박스</h3>
            <p>왼쪽 조건 상품 리스트에서 상품을 추가하면 오른쪽 현재 조합 리스트에 담깁니다.</p>
            <p>판매가, 수수료율, 주문자명, 재주문 여부, 메모를 입력하고 박스출고를 누르면 주문이 등록되고 재고가 임시차감됩니다.</p>
            <p>AI 수동박스 추천에서는 판매가, 목표마진율, 추가 소비자가, 구성느낌, 캐릭터 선택값이 자동으로 AI에게 전달됩니다.</p>
            <p>캐릭터를 선택하지 않으면 전체 재고에서 무작위/균형 추천으로 판단합니다.</p>
            <p>AI 추천안이 마음에 안 들면 수정 요청칸에 “이 상품 빼고 다른 걸로”, “마진율 맞춰줘”처럼 적고 다시 추천하면 됩니다.</p>
          </div>

          <div className="manualCard">
            <h3>5. 주문관리</h3>
            <p>주문접수/재고임시차감과 출고확정/발송완료를 나눠서 봅니다.</p>
            <p>주문을 클릭하면 아래 선택 주문 상품목록에서 구성 상품을 확인할 수 있습니다.</p>
            <p>주문취소 시 재고가 복구되고, 출고확정은 추가 차감 없이 상태만 발송완료로 바뀝니다.</p>
            <p>이전 주문 구성을 수동박스로 복사해 재주문 조합을 다시 만들 수 있습니다.</p>
          </div>

          <div className="manualCard">
            <h3>6. AI 운영 비서</h3>
            <p>오른쪽 아래 Gemini 비서는 고객 요청사항 분석, 상품명 수정, 카테고리 정리, 랜덤박스 구성 상담을 도와줍니다.</p>
            <p>AI 답변에서 실제 재고 상품명이 줄마다 정리되어 있으면 수동박스로 내보낼 수 있습니다.</p>
            <p>AI가 실패하면 환경변수 GEMINI_API_KEY, Supabase 연결, 재고 데이터 수량을 먼저 확인하세요.</p>
          </div>

          <div className="manualCard">
            <h3>7. 실시간 동기화</h3>
            <p>같은 Supabase 프로젝트를 사용하는 PC/휴대폰에서는 재고, 주문, 재료비 변경이 실시간으로 반영됩니다.</p>
            <p>반영이 늦거나 화면이 이상하면 새로고침 후 다시 확인하세요.</p>
          </div>
        </div>
      </section>
    );
  }


  async function handleLoginSubmit(e) {
    e.preventDefault();
    setLoginError("");

    if (!loginPassword.trim()) {
      setLoginError("비밀번호를 입력해줘.");
      return;
    }

    const { error } = await supabase.auth.signInWithPassword({
      email: ADMIN_EMAIL,
      password: loginPassword,
    });

    if (error) {
      setLoginPassword("");
      setLoginError("비밀번호가 맞지 않아요.");
      return;
    }
  }

  async function handleLogout() {
    await supabase.auth.signOut();
    setAuthUser(null);
    setProducts([]);
    setOrders([]);
    setOrderItems([]);
    setMaterials([]);
  }

  function AuthScreen() {
    return (
      <div className="authPage">
        <form className="authBox" onSubmit={handleLoginSubmit}>
          <h1>랜덤박스 운영 프로그램</h1>
          <p>관리자 비밀번호를 입력해야 재고/주문 데이터를 볼 수 있어요.</p>

          <input type="text" name="fake-user" autoComplete="username" style={{ display: "none" }} />
          <label>관리자 비밀번호</label>
          <input
            type="password"
            name="randombox-admin-pass"
            value={loginPassword}
            onChange={(e) => setLoginPassword(e.target.value)}
            placeholder="비밀번호"
            autoFocus
            autoComplete="new-password"
            spellCheck="false"
          />

          {loginError && <div className="authError">{loginError}</div>}

          <button type="submit">들어가기</button>


        </form>
      </div>
    );
  }


  function runManualProductSearch() {
    setSearch(searchInput);
  }

  function clearManualProductSearch() {
    setSearchInput("");
    setSearch("");
  }


  function sendSelectedScoopToManual() {
    const rec = scoopRecommendations[selectedScoopIndex];
    if (!rec) return alert("수동박스로 보낼 추천안을 선택해줘.");
    setComposeItems(rec.items || []);
    setActiveTab("수동박스");
    alert("삭제된 추천안을 수동박스 현재 조합 리스트로 옮겼어요. 수동박스에서 수정 후 박스출고해주세요.");
  }


  function compactProductsForGemini(limit = 160) {
    return products.slice(0, limit).map((p) => ({
      id: p.id,
      name: p.name,
      char1: p.char1,
      char2: p.char2,
      category: p.category,
      stock: toInt(p.stock),
      wholesale: toInt(p.wholesale || p.wholesale_price || p.cost || 0),
      retail: toInt(p.retail || p.retail_price || p.consumer_price || 0),
    }));
  }

  function localFindProducts(keyword, limit = 20) {
    const k = String(keyword || "").trim().toLowerCase();
    if (!k) return [];
    return products
      .filter((p) =>
        String(p.name || "").toLowerCase().includes(k) ||
        String(p.char1 || "").toLowerCase().includes(k) ||
        String(p.char2 || "").toLowerCase().includes(k) ||
        String(p.category || "").toLowerCase().includes(k)
      )
      .slice(0, limit);
  }

  function parseAssistantProductNames(text) {
    const names = [];
    String(text || "").split(/\n+/).forEach((line) => {
      const clean = line.replace(/^[\s\-\*\d\.\)\(]+/, "").replace(/[`"']/g, "").trim();
      if (clean && clean.length >= 2) names.push(clean);
    });
    return names.slice(0, 30);
  }

  function exportGeminiAnswerToManualBox(messageText) {
    const names = parseAssistantProductNames(messageText);
    const picked = [];
    names.forEach((name) => {
      const found = products.find((p) => String(p.name || "").includes(name) || name.includes(String(p.name || "")));
      if (found && !picked.some((x) => String(x.id) === String(found.id))) picked.push(found);
    });

    if (picked.length === 0) {
      alert("답변에서 현재 재고 상품명을 찾지 못했어요. Gemini가 실제 상품명을 줄마다 하나씩 쓰게 다시 요청해보세요.");
      return;
    }

    preserveManualProductListScroll(() => {
      setComposeItems((prev) => [...prev, ...picked]);
    });
    setActiveTab("수동박스");
    alert(`${picked.length}개 상품을 수동박스 조합 리스트로 보냈어요.`);
  }

  function guessGeminiLocalAction(userText) {
    const text = String(userText || "").trim();

    const renameMatch = text.match(/(.+?)(?:상품명|이름)?\s*(?:을|를)?\s*(.+?)(?:으로|로)\s*(?:바꿔|변경|수정)/);
    if (renameMatch) {
      const target = renameMatch[1].replace(/그|상품|이름|상품명/g, "").trim();
      const newName = renameMatch[2].trim();
      const product = localFindProducts(target, 1)[0];
      if (product && newName) {
        return { type: "rename_product", label: `상품명 변경: ${product.name} → ${newName}`, productId: product.id, payload: { name: newName } };
      }
    }

    const categoryMatch = text.match(/(.+?)\s*(?:카테고리|분류)(?:를|을)?\s*(.+?)(?:으로|로)\s*(?:바꿔|변경|수정|분류)/);
    if (categoryMatch) {
      const target = categoryMatch[1].replace(/그|상품/g, "").trim();
      const newCategory = categoryMatch[2].trim();
      const product = localFindProducts(target, 1)[0];
      if (product && newCategory) {
        return { type: "update_product", label: `카테고리 변경: ${product.name} → ${newCategory}`, productId: product.id, payload: { category: newCategory } };
      }
    }

    const stockMatch = text.match(/(.+?)\s*(?:재고|수량)(?:를|을)?\s*(\d+)\s*(?:개)?(?:로)?\s*(?:바꿔|변경|수정|해줘)/);
    if (stockMatch) {
      const target = stockMatch[1].replace(/그|상품/g, "").trim();
      const stock = Number(stockMatch[2]);
      const product = localFindProducts(target, 1)[0];
      if (product && Number.isFinite(stock)) {
        return { type: "update_product", label: `재고수량 변경: ${product.name} → ${stock}개`, productId: product.id, payload: { stock } };
      }
    }

    return null;
  }

  async function applyGeminiAction() {
    if (!geminiActionDraft) return;
    const { productId, payload, label } = geminiActionDraft;
    const ok = window.confirm(`${label}\n\n이 수정사항을 실제 재고 데이터에 반영할까요?`);
    if (!ok) return;

    const { error } = await supabase.from("products").update(payload).eq("id", productId);
    if (error) {
      alert("수정 실패: " + error.message);
      return;
    }

    setGeminiMessages((prev) => [...prev, { role: "assistant", text: `반영 완료: ${label}` }]);
    setGeminiActionDraft(null);
    getProducts();
  }

  async function sendGeminiMessage() {
    const userText = geminiInput.trim();
    if (!userText || geminiLoading) return;

    setGeminiInput("");
    setGeminiMessages((prev) => [...prev, { role: "user", text: userText }]);

    const localAction = guessGeminiLocalAction(userText);
    if (localAction) {
      setGeminiActionDraft(localAction);
      setGeminiMessages((prev) => [
        ...prev,
        { role: "assistant", text: `제가 이해한 수정사항은 아래와 같아요.\n\n${localAction.label}\n\n아래 “수정사항 반영” 버튼을 누르면 실제 재고에 반영됩니다.` }
      ]);
      return;
    }

    setGeminiLoading(true);
    try {
      const context = {
        activeTab,
        salePrice,
        feeRate,
        customer,
        memo,
        reorder,
        products: compactProductsForGemini(),
        currentManualBox: composeItems.map((p) => ({
          id: p.id,
          name: p.name,
          char1: p.char1,
          char2: p.char2,
          category: p.category,
          stock: p.stock,
          wholesale: p.wholesale,
          retail: p.retail,
        })),
        pendingOrders: orders.filter((o) => !String(o.status || "").includes("출고")).slice(0, 30),
      };

      const res = await fetch("/api/gemini", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: userText,
          context,
        }),
      });

      const data = await res.json().catch(() => ({}));

      if (!res.ok) {
        const msg = data?.message || data?.error?.message || data?.error || "Gemini 호출에 실패했어요.";
        throw new Error(String(msg));
      }

      const answer = data?.text || "답변을 가져오지 못했어요.";
      setGeminiMessages((prev) => [...prev, { role: "assistant", text: answer }]);
    } catch (err) {
      const raw = String(err?.message || err);
      const friendly = raw.includes("API key not valid") || raw.includes("API_KEY_INVALID")
        ? "Gemini API 키가 유효하지 않다고 나와요.\n\nVercel 환경변수에서 GEMINI_API_KEY 값을 새 Gemini API 키로 다시 넣어주세요.\n주의: VITE_GEMINI_API_KEY가 아니라 GEMINI_API_KEY 입니다."
        : `Gemini 호출 오류:\n${raw}`;

      setGeminiMessages((prev) => [...prev, { role: "assistant", text: friendly }]);
    } finally {
      setGeminiLoading(false);
    }
  }

  function renderGeminiAssistantWidget() {
    return (
      <div className={`geminiWidget ${geminiOpen ? "open" : ""}`}>
        {!geminiOpen && (
          <button type="button" className="geminiFloatingButton geminiIconButton" aria-label="Gemini 비서 열기" title="Gemini 비서" onClick={() => setGeminiOpen(true)}>💬</button>
        )}

        {geminiOpen && (
          <div className="geminiPanel">
            <div className="geminiHeader">
              <b>Gemini 운영 비서</b>
              <button type="button" onClick={() => setGeminiOpen(false)}>닫기</button>
            </div>

            <div className="geminiMessages">
              {geminiMessages.map((m, idx) => (
                <div key={idx} className={`geminiMessage ${m.role}`}>
                  <pre>{m.text}</pre>
                  {m.role === "assistant" && idx === geminiMessages.length - 1 && (
                    <button type="button" onClick={() => exportGeminiAnswerToManualBox(m.text)}>수동박스로 내보내기</button>
                  )}
                </div>
              ))}
              {geminiLoading && <div className="geminiMessage assistant"><pre>생각 중...</pre></div>}
            </div>

            {geminiActionDraft && (
              <div className="geminiActionBox">
                <b>수정 대기중</b>
                <p>{geminiActionDraft.label}</p>
                <button type="button" onClick={applyGeminiAction}>수정사항 반영</button>
                <button type="button" onClick={() => setGeminiActionDraft(null)}>취소</button>
              </div>
            )}

            <div className="geminiInputRow">
              <textarea
                value={geminiInput}
                onChange={(e) => setGeminiInput(e.target.value)}
                onKeyDown={(e) => {
                  if (!e.nativeEvent?.isComposing && e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    sendGeminiMessage();
                  }
                }}
                placeholder={"예: 고객 요청사항 분석해줘\n예: 키티 위주 5만원 랜박 조합해줘\n예: 시나모롤 파우치 이름 바꿔줘"}
              />
              <button type="button" onClick={sendGeminiMessage}>전송</button>
            </div>
          </div>
        )}
      </div>
    );
  }

  function renderPage() {
    try {
      if (activeTab === "대시보드") return DashboardPage();
      if (activeTab === "정산매입매출") return FinanceReportPage();
      if (activeTab === "재고관리") return InventoryPage();
      if (activeTab === "수동박스") return ComposePage();
      if (activeTab === "AI사입입고분석") return AiImportPage();
      if (activeTab === "주문관리") return OrdersPage();
      if (activeTab === "택배접수") return ShippingRegisterPage();
      if (activeTab === "라방주문") return LiveOrderPage();
      if (activeTab === "취소보관함") return TrashPage();
      if (activeTab === "설정") return SettingsPage();
      return DashboardPage();
    } catch (err) {
      console.error(err);
      return (
        <section className="panel">
          <h2>페이지 표시 오류</h2>
          <p className="statusLine">페이지를 표시하는 중 오류가 발생했어요. 새로고침 후 다시 시도해줘.</p>
          <pre className="errorBox">{String(err?.message || err)}</pre>
        </section>
      );
    }
  }

  if (authLoading) {
    return <div className="authPage"><div className="authBox"><h1>확인 중...</h1></div></div>;
  }

  if (!authUser) {
    return AuthScreen();
  }

  return (
    <div className="app">
      <header className="header">
        <h1>랜덤박스 운영 프로그램</h1>
        <div className="loginInfo">
          <span>관리자 로그인 중</span>
          <button onClick={handleLogout}>로그아웃</button>
        </div>
      </header>
      <nav className="tabs">{TABS.map((tab) => <button key={tab} className={activeTab === tab ? "tab activeTab" : "tab"} onClick={() => { if (activeTab !== tab) setActiveTab(tab); }}>{tab}</button>)}</nav>
      {renderPage()}
      {renderGeminiAssistantWidget()}
    </div>
  );
}
