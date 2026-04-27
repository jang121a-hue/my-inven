import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Calculator,
  Coins,
  Download,
  Package,
  Plus,
  RefreshCw,
  Save,
  Search,
  Trash2,
  Edit3,
  RotateCcw,
  ChevronDown,
  ChevronUp,
  Archive,
  Tag,
  CheckCircle2,
  Image as ImageIcon,
} from "lucide-react";
import { supabase } from "./lib/supabase.js";

function uid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formatNumber(value) {
  return new Intl.NumberFormat("ko-KR", {
    maximumFractionDigits: 0,
  }).format(Math.round(toNumber(value)));
}

function formatPercent(value) {
  return `${toNumber(value).toFixed(2)}%`;
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function makeEmptyProduct(index = 1) {
  return {
    id: uid(),
    name: `상품 ${index}`,
    sku: "",
    imageUrl: "",
    qty: "1",
    buyUsd: "0",
    buyCny: "0",
    buyKrw: "0",       // ← 실제 원화 결제액 추가
    salePrice: "0",
    salesStatus: "미판매",
    restockStatus: "보통",
  };
}

function makeSampleProducts() {
  return [
    {
      id: uid(),
      name: "포스터 A",
      sku: "POSTER-A-50X70",
      imageUrl: "",
      qty: "2",
      buyUsd: "0",
      buyCny: "13",
      buyKrw: "0",
      salePrice: "29900",
      salesStatus: "판매중",
      restockStatus: "보통",
    },
    {
      id: uid(),
      name: "포스터 B",
      sku: "POSTER-B-60X80",
      imageUrl: "",
      qty: "2",
      buyUsd: "0",
      buyCny: "18",
      buyKrw: "0",
      salePrice: "34900",
      salesStatus: "미판매",
      restockStatus: "재주문 필요",
    },
  ];
}

function calculateRows(products, shared) {
  const usdRate = toNumber(shared.usdRate);
  const cnyRate = toNumber(shared.cnyRate);
  const overseasShipping = toNumber(shared.overseasShipping);
  const domesticShipping = toNumber(shared.domesticShipping);
  const packagingCost = toNumber(shared.packagingCost);
  const adCostFixed = toNumber(shared.adCostFixed);
  const smartStoreFeeRate = toNumber(shared.smartStoreFeeRate) / 100;
  const adCostRate = toNumber(shared.adCostRate) / 100;
  const simpleVatRate = toNumber(shared.simpleVatRate) / 100;

  const baseRows = products.map((item) => {
    const qty = Math.max(1, toNumber(item.qty));
    const buyUsd = toNumber(item.buyUsd);
    const buyCny = toNumber(item.buyCny);
    const buyKrw = toNumber(item.buyKrw);
    const salePrice = toNumber(item.salePrice);

    // 실제 원화 결제액이 입력된 경우 우선 사용, 없으면 USD/CNY 환산
    const calculatedBuyCostKrw = buyUsd * usdRate + buyCny * cnyRate;
    const unitBuyCostKrw = buyKrw > 0 ? buyKrw : calculatedBuyCostKrw;
    const totalBuyCostKrw = unitBuyCostKrw * qty;

    // 환산 차이 (실제 - 계산)
    const krwDiff = buyKrw > 0 ? buyKrw - calculatedBuyCostKrw : 0;

    return {
      ...item, qty, buyUsd, buyCny, buyKrw, salePrice,
      unitBuyCostKrw, totalBuyCostKrw, calculatedBuyCostKrw, krwDiff,
    };
  });

  const totalQty = baseRows.reduce((sum, row) => sum + row.qty, 0);
  const totalBuyCost = baseRows.reduce((sum, row) => sum + row.totalBuyCostKrw, 0);
  const domesticPerUnit = totalQty > 0 ? domesticShipping / totalQty : 0;
  const packagingPerUnit = totalQty > 0 ? packagingCost / totalQty : 0;
  const fixedAdPerUnit = totalQty > 0 ? adCostFixed / totalQty : 0;

  const rows = baseRows.map((row) => {
    const buyRatio = totalBuyCost > 0 ? row.totalBuyCostKrw / totalBuyCost : 0;
    const overseasAllocatedTotal = overseasShipping * buyRatio;
    const overseasAllocatedPerUnit = row.qty > 0 ? overseasAllocatedTotal / row.qty : 0;
    const smartStoreFeePerUnit = row.salePrice * smartStoreFeeRate;
    const adCostPerUnit = row.salePrice * adCostRate;
    const simpleVatPerUnit = row.salePrice * simpleVatRate * 0.1;

    const totalCostPerUnit =
      row.unitBuyCostKrw + overseasAllocatedPerUnit + domesticPerUnit +
      packagingPerUnit + fixedAdPerUnit + smartStoreFeePerUnit +
      adCostPerUnit + simpleVatPerUnit;

    const totalCostAll = totalCostPerUnit * row.qty;
    const revenueAll = row.salePrice * row.qty;
    const profitPerUnit = row.salePrice - totalCostPerUnit;
    const profitAll = revenueAll - totalCostAll;
    const marginRate = row.salePrice > 0 ? (profitPerUnit / row.salePrice) * 100 : 0;
    const costRate = row.salePrice > 0 ? (totalCostPerUnit / row.salePrice) * 100 : 0;
    const target30 = totalCostPerUnit / 0.7;
    const target40 = totalCostPerUnit / 0.6;

    return {
      ...row, buyRatio, overseasAllocatedTotal, overseasAllocatedPerUnit,
      domesticPerUnit, packagingPerUnit, fixedAdPerUnit,
      smartStoreFeePerUnit, adCostPerUnit, simpleVatPerUnit,
      totalCostPerUnit, totalCostAll, revenueAll,
      profitPerUnit, profitAll, marginRate, costRate, target30, target40,
    };
  });

  const summary = {
    totalQty,
    totalBuyCost,
    totalRevenue: rows.reduce((sum, row) => sum + row.revenueAll, 0),
    totalCost: rows.reduce((sum, row) => sum + row.totalCostAll, 0),
    totalProfit: rows.reduce((sum, row) => sum + row.profitAll, 0),
  };
  summary.costRate = summary.totalRevenue > 0 ? (summary.totalCost / summary.totalRevenue) * 100 : 0;
  summary.marginRate = summary.totalRevenue > 0 ? (summary.totalProfit / summary.totalRevenue) * 100 : 0;

  return { rows, summary };
}

function exportToCsv(records) {
  const headers = [
    "묶음명","매입일","저장일","상품명","SKU","판매상태","재주문상태",
    "수량","판매가","개당매입원가(실제)","개당환산원가","원화차이","개당총원가",
    "총매출","총원가","개당순이익","총순이익","원가율","마진율","이미지URL",
  ];
  const lines = [headers.join(",")];
  records.forEach((record) => {
    record.items.forEach((item) => {
      lines.push(
        [
          record.bundleName, record.purchaseDate, record.savedAt || "",
          item.name, item.sku || "", item.salesStatus || "미판매",
          item.restockStatus || "보통", item.qty, item.salePrice,
          Math.round(item.unitBuyCostKrw),
          Math.round(item.calculatedBuyCostKrw || 0),
          Math.round(item.krwDiff || 0),
          Math.round(item.totalCostPerUnit),
          Math.round(item.revenueAll), Math.round(item.totalCostAll),
          Math.round(item.profitPerUnit), Math.round(item.profitAll),
          item.costRate.toFixed(2), item.marginRate.toFixed(2), item.imageUrl || "",
        ]
          .map((v) => `"${String(v).replace(/"/g, '""')}"`)
          .join(",")
      );
    });
  });
  downloadFile(`원가계산기록_${new Date().toISOString().slice(0, 10)}.csv`, lines.join("\n"), "text/csv;charset=utf-8;");
}

function exportToXlsx(records) {
  const workbook = XLSX.utils.book_new();
  const summaryRows = records.map((record) => ({
    묶음명: record.bundleName, 매입일: record.purchaseDate, 저장일: record.savedAt,
    상품종수: record.items.length, 총수량: record.summary.totalQty,
    전체매출: Math.round(record.summary.totalRevenue),
    전체원가: Math.round(record.summary.totalCost),
    전체순이익: Math.round(record.summary.totalProfit),
    전체원가율: Number(record.summary.costRate.toFixed(2)),
    전체마진율: Number(record.summary.marginRate.toFixed(2)),
  }));
  const detailRows = records.flatMap((record) =>
    record.items.map((item) => ({
      묶음명: record.bundleName, 매입일: record.purchaseDate, 저장일: record.savedAt,
      상품명: item.name, SKU: item.sku || "",
      판매상태: item.salesStatus || "미판매",
      재주문상태: item.restockStatus || "보통",
      수량: item.qty, 판매가: Math.round(item.salePrice),
      개당매입원가실제: Math.round(item.unitBuyCostKrw),
      개당환산원가: Math.round(item.calculatedBuyCostKrw || 0),
      원화차이: Math.round(item.krwDiff || 0),
      개당총원가: Math.round(item.totalCostPerUnit),
      총매출: Math.round(item.revenueAll), 총원가: Math.round(item.totalCostAll),
      개당순이익: Math.round(item.profitPerUnit), 총순이익: Math.round(item.profitAll),
      원가율: Number(item.costRate.toFixed(2)),
      마진율: Number(item.marginRate.toFixed(2)),
      이미지URL: item.imageUrl || "",
    }))
  );
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), "요약");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailRows), "상품상세");
  XLSX.writeFile(workbook, `원가계산기록_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function mapDbBundleToUi(bundle) {
  const items = (bundle.cost_bundle_items || []).map((item) => ({
    id: item.id, name: item.name ?? "", sku: item.sku ?? "",
    imageUrl: item.image_url ?? "", qty: item.qty ?? 1,
    buyUsd: item.buy_usd ?? 0, buyCny: item.buy_cny ?? 0,
    buyKrw: item.buy_krw ?? 0,
    salePrice: item.sale_price ?? 0,
    salesStatus: item.sales_status ?? "미판매",
    restockStatus: item.restock_status ?? "보통",
    unitBuyCostKrw: item.unit_buy_cost_krw ?? 0,
    calculatedBuyCostKrw: item.calculated_buy_cost_krw ?? 0,
    krwDiff: item.krw_diff ?? 0,
    totalBuyCostKrw: item.total_buy_cost_krw ?? 0,
    buyRatio: item.buy_ratio ?? 0,
    overseasAllocatedTotal: item.overseas_allocated_total ?? 0,
    overseasAllocatedPerUnit: item.overseas_allocated_per_unit ?? 0,
    domesticPerUnit: item.domestic_per_unit ?? 0,
    packagingPerUnit: item.packaging_per_unit ?? 0,
    fixedAdPerUnit: item.fixed_ad_per_unit ?? 0,
    smartStoreFeePerUnit: item.smartstore_fee_per_unit ?? 0,
    adCostPerUnit: item.ad_cost_per_unit ?? 0,
    simpleVatPerUnit: item.simple_vat_per_unit ?? 0,
    totalCostPerUnit: item.total_cost_per_unit ?? 0,
    totalCostAll: item.total_cost_all ?? 0,
    revenueAll: item.revenue_all ?? 0,
    profitPerUnit: item.profit_per_unit ?? 0,
    profitAll: item.profit_all ?? 0,
    costRate: item.cost_rate ?? 0, marginRate: item.margin_rate ?? 0,
    target30: item.target30 ?? 0, target40: item.target40 ?? 0,
  }));

  return {
    id: bundle.id, bundleName: bundle.bundle_name ?? "",
    purchaseDate: bundle.purchase_date ?? "",
    sharedInput: {
      usdRate: String(bundle.usd_rate ?? 0), cnyRate: String(bundle.cny_rate ?? 0),
      overseasShipping: String(bundle.overseas_shipping ?? 0),
      domesticShipping: String(bundle.domestic_shipping ?? 0),
      packagingCost: String(bundle.packaging_cost ?? 0),
      adCostFixed: String(bundle.ad_cost_fixed ?? 0),
      smartStoreFeeRate: String(bundle.smartstore_fee_rate ?? 0),
      adCostRate: String(bundle.ad_cost_rate ?? 0),
      simpleVatRate: String(bundle.simple_vat_rate ?? 0),
    },
    items,
    sourceProducts: items.map((item) => ({
      id: item.id || uid(), name: item.name, sku: item.sku,
      imageUrl: item.imageUrl, qty: String(item.qty ?? 1),
      buyUsd: String(item.buyUsd ?? 0), buyCny: String(item.buyCny ?? 0),
      buyKrw: String(item.buyKrw ?? 0),
      salePrice: String(item.salePrice ?? 0),
      salesStatus: item.salesStatus ?? "미판매",
      restockStatus: item.restockStatus ?? "보통",
    })),
    summary: {
      totalQty: bundle.total_qty ?? 0, totalBuyCost: bundle.total_buy_cost ?? 0,
      totalRevenue: bundle.total_revenue ?? 0, totalCost: bundle.total_cost ?? 0,
      totalProfit: bundle.total_profit ?? 0,
      costRate: bundle.cost_rate ?? 0, marginRate: bundle.margin_rate ?? 0,
    },
    savedAt: bundle.created_at ? new Date(bundle.created_at).toLocaleString("ko-KR") : "",
  };
}

function InputField({ label, value, onChange, suffix, placeholder, type = "text", highlight = false }) {
  return (
    <label className="block">
      <div className="mb-1.5 text-sm font-medium text-slate-700">{label}</div>
      <div className={`flex items-center rounded-2xl border px-4 py-2.5 shadow-sm transition focus-within:border-slate-400 ${highlight ? "border-amber-300 bg-amber-50" : "border-slate-200 bg-white"}`}>
        <input
          type={type}
          inputMode={type === "number" ? "decimal" : undefined}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          className="w-full bg-transparent text-base outline-none"
        />
        {suffix && <span className="ml-2 whitespace-nowrap text-sm text-slate-500">{suffix}</span>}
      </div>
    </label>
  );
}

function SelectField({ label, value, onChange, options }) {
  return (
    <label className="block">
      <div className="mb-1.5 text-sm font-medium text-slate-700">{label}</div>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-base outline-none shadow-sm"
      >
        {options.map((option) => {
          const val = typeof option === "string" ? option : option.value;
          const text = typeof option === "string" ? option : option.label;
          return <option key={val} value={val}>{text}</option>;
        })}
      </select>
    </label>
  );
}

function SummaryCard({ title, value, sub, icon: Icon }) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-4 shadow-sm">
      <div className="mb-2 flex items-center justify-between">
        <div className="text-sm font-medium text-slate-500">{title}</div>
        <Icon className="h-4 w-4 text-slate-400" />
      </div>
      <div className="text-xl font-bold tracking-tight text-slate-900">{value}</div>
      {sub && <div className="mt-1 text-xs text-slate-500">{sub}</div>}
    </div>
  );
}

export default function App() {
  const today = new Date().toISOString().slice(0, 10);

  const [bundleName, setBundleName] = useState("4월 포스터 합배송 1차");
  const [purchaseDate, setPurchaseDate] = useState(today);
  const [usdRate, setUsdRate] = useState("1380");
  const [cnyRate, setCnyRate] = useState("190");
  const [overseasShipping, setOverseasShipping] = useState("22000");
  const [domesticShipping, setDomesticShipping] = useState("4000");
  const [packagingCost, setPackagingCost] = useState("6000");
  const [adCostFixed, setAdCostFixed] = useState("3000");
  const [smartStoreFeeRate, setSmartStoreFeeRate] = useState("2");
  const [adCostRate, setAdCostRate] = useState("3");
  const [simpleVatRate, setSimpleVatRate] = useState("15");

  const [products, setProducts] = useState(makeSampleProducts());
  const [savedRecords, setSavedRecords] = useState([]);
  const [editingRecordId, setEditingRecordId] = useState(null);
  const [searchKeyword, setSearchKeyword] = useState("");
  const [isFetchingRates, setIsFetchingRates] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isLoadingRecords, setIsLoadingRecords] = useState(false);
  const [rateMessage, setRateMessage] = useState("");
  const [dbMessage, setDbMessage] = useState("");
  const [expandedIds, setExpandedIds] = useState({});
  const [expandedResults, setExpandedResults] = useState({});

  const sharedInput = useMemo(() => ({
    usdRate, cnyRate, overseasShipping, domesticShipping, packagingCost,
    adCostFixed, smartStoreFeeRate, adCostRate, simpleVatRate,
  }), [usdRate, cnyRate, overseasShipping, domesticShipping, packagingCost, adCostFixed, smartStoreFeeRate, adCostRate, simpleVatRate]);

  const result = useMemo(() => calculateRows(products, sharedInput), [products, sharedInput]);

  const filteredRecords = useMemo(() => {
    const keyword = searchKeyword.trim().toLowerCase();
    if (!keyword) return savedRecords;
    return savedRecords.filter((record) => {
      const target = [
        record.bundleName, record.purchaseDate,
        ...(record.items || []).flatMap((item) => [item.name, item.sku, item.imageUrl]),
      ].join(" ").toLowerCase();
      return target.includes(keyword);
    });
  }, [savedRecords, searchKeyword]);

  const loadRecords = async () => {
    try {
      setIsLoadingRecords(true);
      setDbMessage("");
      const { data, error } = await supabase
        .from("cost_bundles")
        .select(`*, cost_bundle_items (*)`)
        .order("created_at", { ascending: false });
      if (error) throw error;
      setSavedRecords((data || []).map(mapDbBundleToUi));
    } catch (error) {
      setDbMessage(error.message || "DB 조회 중 오류가 발생했습니다.");
    } finally {
      setIsLoadingRecords(false);
    }
  };

  useEffect(() => { loadRecords(); }, []);

  const fetchRates = async () => {
    try {
      setIsFetchingRates(true);
      setRateMessage("");
      const [usdRes, cnyRes] = await Promise.all([
        fetch("https://api.frankfurter.dev/v2/rate/USD/KRW"),
        fetch("https://api.frankfurter.dev/v2/rate/CNY/KRW"),
      ]);
      if (!usdRes.ok || !cnyRes.ok) throw new Error("환율 정보를 불러오지 못했습니다.");
      const usdData = await usdRes.json();
      const cnyData = await cnyRes.json();
      if (!usdData?.rate || !cnyData?.rate) throw new Error("환율 데이터 형식이 올바르지 않습니다.");
      setUsdRate(String(Math.round(usdData.rate)));
      setCnyRate(String(Math.round(cnyData.rate)));
      setRateMessage(`환율 반영 완료 (${new Date().toLocaleString("ko-KR")})`);
    } catch (error) {
      setRateMessage(error.message || "환율 반영 실패");
    } finally {
      setIsFetchingRates(false);
    }
  };

  const resetForm = () => {
    setBundleName("4월 포스터 합배송 1차");
    setPurchaseDate(today);
    setUsdRate("1380"); setCnyRate("190");
    setOverseasShipping("22000"); setDomesticShipping("4000");
    setPackagingCost("6000"); setAdCostFixed("3000");
    setSmartStoreFeeRate("2"); setAdCostRate("3"); setSimpleVatRate("15");
    setProducts(makeSampleProducts());
    setEditingRecordId(null);
    setDbMessage("");
  };

  const updateProduct = (id, key, value) => {
    setProducts((prev) => prev.map((item) => item.id === id ? { ...item, [key]: value } : item));
  };

  const addProduct = () => {
    setProducts((prev) => [...prev, makeEmptyProduct(prev.length + 1)]);
  };

  const removeProduct = (id) => {
    setProducts((prev) => prev.length <= 1 ? prev : prev.filter((item) => item.id !== id));
  };

  const buildBundlePayload = () => ({
    bundle_name: bundleName.trim() || "이름 없는 묶음",
    purchase_date: purchaseDate || null,
    usd_rate: toNumber(usdRate), cny_rate: toNumber(cnyRate),
    overseas_shipping: toNumber(overseasShipping), domestic_shipping: toNumber(domesticShipping),
    packaging_cost: toNumber(packagingCost), ad_cost_fixed: toNumber(adCostFixed),
    smartstore_fee_rate: toNumber(smartStoreFeeRate), ad_cost_rate: toNumber(adCostRate),
    simple_vat_rate: toNumber(simpleVatRate),
    total_qty: Math.round(toNumber(result.summary.totalQty)),
    total_buy_cost: toNumber(result.summary.totalBuyCost),
    total_revenue: toNumber(result.summary.totalRevenue),
    total_cost: toNumber(result.summary.totalCost),
    total_profit: toNumber(result.summary.totalProfit),
    cost_rate: toNumber(result.summary.costRate),
    margin_rate: toNumber(result.summary.marginRate),
  });

  const buildItemPayloads = (bundleId) =>
    result.rows.map((row) => ({
      bundle_id: bundleId, name: row.name || "", sku: row.sku || "",
      image_url: row.imageUrl || "", qty: Math.round(toNumber(row.qty)),
      buy_usd: toNumber(row.buyUsd), buy_cny: toNumber(row.buyCny),
      buy_krw: toNumber(row.buyKrw),
      sale_price: toNumber(row.salePrice),
      sales_status: row.salesStatus || "미판매",
      restock_status: row.restockStatus || "보통",
      unit_buy_cost_krw: toNumber(row.unitBuyCostKrw),
      calculated_buy_cost_krw: toNumber(row.calculatedBuyCostKrw),
      krw_diff: toNumber(row.krwDiff),
      total_buy_cost_krw: toNumber(row.totalBuyCostKrw),
      buy_ratio: toNumber(row.buyRatio),
      overseas_allocated_total: toNumber(row.overseasAllocatedTotal),
      overseas_allocated_per_unit: toNumber(row.overseasAllocatedPerUnit),
      domestic_per_unit: toNumber(row.domesticPerUnit),
      packaging_per_unit: toNumber(row.packagingPerUnit),
      fixed_ad_per_unit: toNumber(row.fixedAdPerUnit),
      smartstore_fee_per_unit: toNumber(row.smartStoreFeePerUnit),
      ad_cost_per_unit: toNumber(row.adCostPerUnit),
      simple_vat_per_unit: toNumber(row.simpleVatPerUnit),
      total_cost_per_unit: toNumber(row.totalCostPerUnit),
      total_cost_all: toNumber(row.totalCostAll),
      revenue_all: toNumber(row.revenueAll),
      profit_per_unit: toNumber(row.profitPerUnit),
      profit_all: toNumber(row.profitAll),
      cost_rate: toNumber(row.costRate), margin_rate: toNumber(row.marginRate),
      target30: toNumber(row.target30), target40: toNumber(row.target40),
    }));

  const saveRecord = async () => {
    try {
      setIsSaving(true);
      setDbMessage("");
      const bundlePayload = buildBundlePayload();
      let bundleId = editingRecordId;

      if (editingRecordId) {
        const { error: updateError } = await supabase.from("cost_bundles").update(bundlePayload).eq("id", editingRecordId);
        if (updateError) throw updateError;
        const { error: deleteItemsError } = await supabase.from("cost_bundle_items").delete().eq("bundle_id", editingRecordId);
        if (deleteItemsError) throw deleteItemsError;
      } else {
        const { data: insertedBundle, error: insertBundleError } = await supabase.from("cost_bundles").insert(bundlePayload).select("id").single();
        if (insertBundleError) throw insertBundleError;
        bundleId = insertedBundle.id;
      }

      const itemPayloads = buildItemPayloads(bundleId);
      if (itemPayloads.length > 0) {
        const { error: insertItemsError } = await supabase.from("cost_bundle_items").insert(itemPayloads);
        if (insertItemsError) throw insertItemsError;
      }

      setEditingRecordId(bundleId);
      setDbMessage(editingRecordId ? "수정 완료" : "저장 완료");
      await loadRecords();
      alert(editingRecordId ? "수정되었습니다." : "저장되었습니다.");
    } catch (error) {
      setDbMessage(error.message || "DB 저장 중 오류가 발생했습니다.");
      alert(error.message || "DB 저장 중 오류가 발생했습니다.");
    } finally {
      setIsSaving(false);
    }
  };

  const editRecord = (id) => {
    const record = savedRecords.find((item) => item.id === id);
    if (!record) return;
    setEditingRecordId(record.id);
    setBundleName(record.bundleName || "");
    setPurchaseDate(record.purchaseDate || today);
    setUsdRate(record.sharedInput.usdRate || "0");
    setCnyRate(record.sharedInput.cnyRate || "0");
    setOverseasShipping(record.sharedInput.overseasShipping || "0");
    setDomesticShipping(record.sharedInput.domesticShipping || "0");
    setPackagingCost(record.sharedInput.packagingCost || "0");
    setAdCostFixed(record.sharedInput.adCostFixed || "0");
    setSmartStoreFeeRate(record.sharedInput.smartStoreFeeRate || "0");
    setAdCostRate(record.sharedInput.adCostRate || "0");
    setSimpleVatRate(record.sharedInput.simpleVatRate || "15");
    setProducts((record.sourceProducts || []).map((item) => ({ ...item, id: item.id || uid() })));
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const deleteRecord = async (id) => {
    const ok = window.confirm("정말 삭제하시겠습니까?");
    if (!ok) return;
    try {
      setDbMessage("");
      const { error } = await supabase.from("cost_bundles").delete().eq("id", id);
      if (error) throw error;
      if (editingRecordId === id) resetForm();
      setDbMessage("삭제 완료");
      await loadRecords();
    } catch (error) {
      setDbMessage(error.message || "DB 삭제 중 오류가 발생했습니다.");
      alert(error.message || "DB 삭제 중 오류가 발생했습니다.");
    }
  };

  const toggleExpand = (id) => setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  const toggleResult = (id) => setExpandedResults((prev) => ({ ...prev, [id]: !prev[id] }));

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="mx-auto max-w-7xl p-4 sm:p-6">

        {/* 헤더 */}
        <div className="mb-6 rounded-[28px] bg-gradient-to-r from-slate-900 to-slate-700 p-5 text-white shadow-xl">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <div className="mb-1 flex items-center gap-2 text-xs text-slate-300">
                <Archive className="h-3.5 w-3.5" /> 원가 계산 앱
              </div>
              <h1 className="text-2xl font-bold tracking-tight">합배송 원가 계산 앱</h1>
              <p className="mt-1 text-xs text-slate-300">
                환율, 배송비, 포장비, 광고비, 수수료를 반영한 상품별 원가율과 마진율 계산
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button onClick={fetchRates} className="inline-flex items-center gap-2 rounded-2xl bg-emerald-500 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-emerald-400">
                <RefreshCw className={`h-4 w-4 ${isFetchingRates ? "animate-spin" : ""}`} />
                {isFetchingRates ? "불러오는 중" : "오늘 환율 반영"}
              </button>
              <button onClick={() => exportToCsv(filteredRecords)} className="inline-flex items-center gap-2 rounded-2xl bg-white/10 px-4 py-2.5 text-sm font-medium text-white backdrop-blur transition hover:bg-white/20">
                <Download className="h-4 w-4" /> CSV
              </button>
              <button onClick={() => exportToXlsx(filteredRecords)} className="inline-flex items-center gap-2 rounded-2xl bg-white/10 px-4 py-2.5 text-sm font-medium text-white backdrop-blur transition hover:bg-white/20">
                <Download className="h-4 w-4" /> XLSX
              </button>
              <button onClick={resetForm} className="inline-flex items-center gap-2 rounded-2xl bg-white/10 px-4 py-2.5 text-sm font-medium text-white backdrop-blur transition hover:bg-white/20">
                <RotateCcw className="h-4 w-4" /> 새 작업
              </button>
            </div>
          </div>
        </div>

        {rateMessage && (
          <div className="mb-4 rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-700">{rateMessage}</div>
        )}
        {dbMessage && (
          <div className="mb-4 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-600 shadow-sm">{dbMessage}</div>
        )}

        {/* 요약 카드 */}
        <div className="mb-6 grid grid-cols-3 gap-3">
          <SummaryCard title="전체 매출" value={`${formatNumber(result.summary.totalRevenue)}원`} sub={`총 ${formatNumber(result.summary.totalQty)}개`} icon={Coins} />
          <SummaryCard title="전체 원가" value={`${formatNumber(result.summary.totalCost)}원`} sub={`원가율 ${formatPercent(result.summary.costRate)}`} icon={Calculator} />
          <SummaryCard title="전체 순이익" value={`${formatNumber(result.summary.totalProfit)}원`} sub={`마진율 ${formatPercent(result.summary.marginRate)}`} icon={Package} />
        </div>

        <div className="space-y-4">

          {/* 기본 정보 */}
          <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-bold">기본 정보</h2>
            <div className="grid gap-3 sm:grid-cols-2">
              <InputField label="묶음명" value={bundleName} onChange={setBundleName} placeholder="예: 4월 포스터 합배송 1차" />
              <InputField label="매입일" value={purchaseDate} onChange={setPurchaseDate} type="date" />
            </div>
          </div>

          {/* 공통 비용 */}
          <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-bold">공통 비용 설정</h2>
            <div className="grid gap-3 grid-cols-2 sm:grid-cols-3">
              <InputField label="달러 환율" value={usdRate} onChange={setUsdRate} suffix="원" />
              <InputField label="위안 환율" value={cnyRate} onChange={setCnyRate} suffix="원" />
              <InputField label="국제배송비" value={overseasShipping} onChange={setOverseasShipping} suffix="원" />
              <InputField label="국내배송비" value={domesticShipping} onChange={setDomesticShipping} suffix="원" />
              <InputField label="포장비" value={packagingCost} onChange={setPackagingCost} suffix="원" />
              <InputField label="정액 광고비" value={adCostFixed} onChange={setAdCostFixed} suffix="원" />
              <InputField label="스마트스토어 수수료율" value={smartStoreFeeRate} onChange={setSmartStoreFeeRate} suffix="%" />
              <InputField label="광고비율" value={adCostRate} onChange={setAdCostRate} suffix="%" />
              <InputField label="간이과세 부가가치율" value={simpleVatRate} onChange={setSimpleVatRate} suffix="%" />
            </div>
            <div className="mt-3 rounded-2xl bg-slate-50 px-4 py-2.5 text-xs text-slate-500">
              국제배송비는 매입금액 비율로 배분, 국내배송비·포장비·정액광고비는 전체 수량 기준 균등 배분
            </div>
          </div>

          {/* 상품 목록 */}
          <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
            <div className="mb-4 flex items-center justify-between">
              <h2 className="text-lg font-bold">상품 목록</h2>
              <button onClick={addProduct} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
                <Plus className="h-4 w-4" /> 상품 추가
              </button>
            </div>

            <div className="space-y-4">
              {products.map((product, index) => {
                const hasKrw = toNumber(product.buyKrw) > 0;
                const calcKrw = toNumber(product.buyUsd) * toNumber(usdRate) + toNumber(product.buyCny) * toNumber(cnyRate);
                const diff = hasKrw ? toNumber(product.buyKrw) - calcKrw : 0;

                return (
                  <div key={product.id} className="rounded-2xl border border-slate-200 p-4">
                    <div className="mb-3 flex items-center justify-between">
                      <div className="text-sm font-semibold text-slate-700">상품 {index + 1}</div>
                      <button onClick={() => removeProduct(product.id)} className="inline-flex items-center gap-1 text-sm text-rose-500 transition hover:text-rose-600">
                        <Trash2 className="h-4 w-4" /> 삭제
                      </button>
                    </div>

                    <div className="grid gap-3 grid-cols-2 sm:grid-cols-4">
                      <InputField label="상품명" value={product.name} onChange={(v) => updateProduct(product.id, "name", v)} placeholder="상품명" />
                      <InputField label="SKU" value={product.sku} onChange={(v) => updateProduct(product.id, "sku", v)} placeholder="SKU" />
                      <InputField label="이미지 URL" value={product.imageUrl} onChange={(v) => updateProduct(product.id, "imageUrl", v)} placeholder="https://..." />
                      <InputField label="수량" value={product.qty} onChange={(v) => updateProduct(product.id, "qty", v)} suffix="개" />
                      <InputField label="개당 매입가 (USD)" value={product.buyUsd} onChange={(v) => updateProduct(product.id, "buyUsd", v)} suffix="$" />
                      <InputField label="개당 매입가 (CNY)" value={product.buyCny} onChange={(v) => updateProduct(product.id, "buyCny", v)} suffix="¥" />
                      <InputField label="판매가" value={product.salePrice} onChange={(v) => updateProduct(product.id, "salePrice", v)} suffix="원" />
                      <SelectField label="판매 상태" value={product.salesStatus} onChange={(v) => updateProduct(product.id, "salesStatus", v)} options={["미판매", "판매중", "판매완료"]} />
                    </div>

                    {/* 실제 원화 결제액 섹션 */}
                    <div className="mt-3 rounded-2xl border border-amber-200 bg-amber-50 p-3">
                      <div className="mb-2 text-xs font-semibold text-amber-800">
                        💳 실제 원화 결제액 (선택)
                      </div>
                      <div className="grid gap-3 grid-cols-1 sm:grid-cols-3 items-end">
                        <InputField
                          label="개당 실제 결제액 (원화)"
                          value={product.buyKrw}
                          onChange={(v) => updateProduct(product.id, "buyKrw", v)}
                          suffix="원"
                          highlight={true}
                          placeholder="직접 입력 (없으면 0)"
                        />
                        <div className="rounded-2xl bg-white border border-amber-200 px-4 py-2.5 text-sm">
                          <div className="text-xs text-slate-500 mb-1">환산 원가 (USD+CNY)</div>
                          <div className="font-semibold text-slate-700">{formatNumber(calcKrw)}원</div>
                        </div>
                        <div className={`rounded-2xl border px-4 py-2.5 text-sm ${diff > 0 ? "border-rose-200 bg-rose-50" : diff < 0 ? "border-emerald-200 bg-emerald-50" : "border-slate-200 bg-white"}`}>
                          <div className="text-xs text-slate-500 mb-1">실제 vs 환산 차이</div>
                          <div className={`font-semibold ${diff > 0 ? "text-rose-600" : diff < 0 ? "text-emerald-600" : "text-slate-400"}`}>
                            {hasKrw ? `${diff > 0 ? "+" : ""}${formatNumber(diff)}원` : "미입력"}
                          </div>
                        </div>
                      </div>
                      <div className="mt-2 text-xs text-amber-700">
                        {hasKrw
                          ? "✅ 실제 결제액으로 원가 계산 중 (수수료·환율 차이 반영됨)"
                          : "💡 알리페이·위챗페이 실제 결제액 입력 시 수수료·환율 차이까지 반영돼요"}
                      </div>
                    </div>

                    <div className="mt-3 grid gap-3 grid-cols-1">
                      <SelectField label="재주문 상태" value={product.restockStatus} onChange={(v) => updateProduct(product.id, "restockStatus", v)} options={["보통", "재주문 필요", "재주문 완료"]} />
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="mt-4">
              <button
                onClick={saveRecord}
                disabled={isSaving}
                className="w-full inline-flex items-center justify-center gap-2 rounded-2xl bg-slate-900 px-5 py-3 text-sm font-medium text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
              >
                <Save className="h-4 w-4" />
                {isSaving ? "저장 중..." : editingRecordId ? "수정 저장" : "현재 묶음 저장"}
              </button>
            </div>
          </div>

          {/* 계산 결과 */}
          <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-bold">현재 계산 결과</h2>
            <div className="space-y-3">
              {result.rows.map((row) => {
                const isExpanded = !!expandedResults[row.id];
                return (
                  <div key={row.id} className="rounded-2xl border border-slate-200">
                    <button
                      onClick={() => toggleResult(row.id)}
                      className="flex w-full items-center justify-between gap-3 p-4 text-left hover:bg-slate-50 rounded-2xl"
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        {row.imageUrl ? (
                          <img src={row.imageUrl} alt={row.name} className="h-10 w-10 rounded-xl object-cover flex-shrink-0" />
                        ) : (
                          <div className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-400">
                            <ImageIcon className="h-4 w-4" />
                          </div>
                        )}
                        <div className="min-w-0">
                          <div className="font-semibold text-slate-900 truncate">{row.name}</div>
                          <div className="flex flex-wrap gap-2 text-xs text-slate-500 mt-0.5">
                            <span className="flex items-center gap-1"><Tag className="h-3 w-3" />{row.sku || "SKU 없음"}</span>
                            <span className="flex items-center gap-1"><CheckCircle2 className="h-3 w-3" />{row.salesStatus || "미판매"}</span>
                            {row.buyKrw > 0 && (
                              <span className="rounded-full bg-amber-100 px-2 py-0.5 text-amber-700">실결제 반영</span>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center gap-3 flex-shrink-0">
                        <div className={`rounded-2xl px-3 py-2 text-right ${row.profitPerUnit >= 0 ? "bg-emerald-50 text-emerald-600" : "bg-rose-50 text-rose-600"}`}>
                          <div className="text-base font-bold">{formatNumber(row.profitPerUnit)}원</div>
                          <div className="text-xs text-slate-500">개당 순이익</div>
                        </div>
                        {isExpanded ? <ChevronUp className="h-4 w-4 text-slate-400" /> : <ChevronDown className="h-4 w-4 text-slate-400" />}
                      </div>
                    </button>

                    {isExpanded && (
                      <div className="border-t border-slate-100 p-4">
                        <div className="grid gap-3 sm:grid-cols-2">
                          <div className="rounded-2xl bg-slate-50 p-4 text-sm">
                            <div className="mb-2 font-semibold text-slate-700">개당 비용 요약</div>
                            <div className="flex justify-between py-1">
                              <span className="text-slate-600">환산 매입원가</span>
                              <span className="text-slate-500">{formatNumber(row.calculatedBuyCostKrw)}원</span>
                            </div>
                            {row.buyKrw > 0 && (
                              <>
                                <div className="flex justify-between py-1">
                                  <span className="text-amber-700 font-medium">실제 결제액</span>
                                  <strong className="text-amber-700">{formatNumber(row.buyKrw)}원</strong>
                                </div>
                                <div className="flex justify-between py-1">
                                  <span className="text-slate-600">수수료·환율 차이</span>
                                  <span className={row.krwDiff > 0 ? "text-rose-600" : "text-emerald-600"}>
                                    {row.krwDiff > 0 ? "+" : ""}{formatNumber(row.krwDiff)}원
                                  </span>
                                </div>
                              </>
                            )}
                            <div className="flex justify-between py-1 border-t border-slate-200 mt-1 pt-2">
                              <span className="font-medium">개당 매입원가 (적용)</span>
                              <strong>{formatNumber(row.unitBuyCostKrw)}원</strong>
                            </div>
                            <div className="flex justify-between py-1">
                              <span>개당 총원가</span>
                              <strong>{formatNumber(row.totalCostPerUnit)}원</strong>
                            </div>
                            <div className="flex justify-between py-1">
                              <span>원가율</span>
                              <strong>{formatPercent(row.costRate)}</strong>
                            </div>
                            <div className="flex justify-between py-1">
                              <span>마진율</span>
                              <strong>{formatPercent(row.marginRate)}</strong>
                            </div>
                          </div>
                          <div className="rounded-2xl bg-slate-50 p-4 text-sm">
                            <div className="mb-2 font-semibold text-slate-700">목표 판매가</div>
                            <div className="flex justify-between py-1"><span>30% 마진 목표</span><strong>{formatNumber(row.target30)}원</strong></div>
                            <div className="flex justify-between py-1"><span>40% 마진 목표</span><strong>{formatNumber(row.target40)}원</strong></div>
                            <div className="flex justify-between py-1"><span>총 순이익</span><strong>{formatNumber(row.profitAll)}원</strong></div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* 저장 기록 */}
          <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <h2 className="text-lg font-bold">저장 기록</h2>
              <div className="flex flex-wrap gap-2">
                <button onClick={loadRecords} disabled={isLoadingRecords} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50 disabled:opacity-60">
                  <RefreshCw className={`h-4 w-4 ${isLoadingRecords ? "animate-spin" : ""}`} /> 새로고침
                </button>
                <button onClick={() => exportToCsv(filteredRecords)} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
                  <Download className="h-4 w-4" /> CSV
                </button>
                <button onClick={() => exportToXlsx(filteredRecords)} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
                  <Download className="h-4 w-4" /> XLSX
                </button>
              </div>
            </div>

            <div className="relative mb-4">
              <Search className="pointer-events-none absolute left-4 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={searchKeyword}
                onChange={(e) => setSearchKeyword(e.target.value)}
                placeholder="묶음명, 날짜, 상품명, SKU 검색"
                className="w-full rounded-2xl border border-slate-200 bg-white py-3 pl-11 pr-4 text-sm outline-none shadow-sm"
              />
            </div>

            <div className="space-y-3">
              {isLoadingRecords ? (
                <div className="rounded-2xl border border-dashed border-slate-200 py-10 text-center text-sm text-slate-400">불러오는 중...</div>
              ) : filteredRecords.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-200 py-10 text-center text-sm text-slate-400">저장된 기록이 없습니다.</div>
              ) : (
                filteredRecords.map((record) => {
                  const expanded = !!expandedIds[record.id];
                  return (
                    <div key={record.id} className="rounded-2xl border border-slate-200">
                      <div className="flex flex-col gap-3 p-4 sm:flex-row sm:items-center sm:justify-between">
                        <div className="min-w-0">
                          <div className="font-bold text-slate-900 truncate">{record.bundleName}</div>
                          <div className="mt-1 text-xs text-slate-500">
                            매입일 {record.purchaseDate} · 저장일 {record.savedAt}
                          </div>
                          <div className="mt-2 flex flex-wrap gap-3 text-sm">
                            <span className="text-slate-600">수량 <strong>{formatNumber(record.summary.totalQty)}개</strong></span>
                            <span className="text-slate-600">매출 <strong>{formatNumber(record.summary.totalRevenue)}원</strong></span>
                            <span className={record.summary.totalProfit >= 0 ? "text-emerald-600" : "text-rose-600"}>
                              순이익 <strong>{formatNumber(record.summary.totalProfit)}원</strong>
                            </span>
                          </div>
                        </div>
                        <div className="flex flex-wrap gap-2">
                          <button onClick={() => toggleExpand(record.id)} className="inline-flex items-center gap-1 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
                            {expanded ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
                            {expanded ? "접기" : "상세"}
                          </button>
                          <button onClick={() => editRecord(record.id)} className="inline-flex items-center gap-1 rounded-2xl border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
                            <Edit3 className="h-4 w-4" /> 수정
                          </button>
                          <button onClick={() => deleteRecord(record.id)} className="inline-flex items-center gap-1 rounded-2xl border border-rose-200 px-3 py-2 text-sm font-medium text-rose-600 transition hover:bg-rose-50">
                            <Trash2 className="h-4 w-4" /> 삭제
                          </button>
                        </div>
                      </div>

                      {expanded && (
                        <div className="border-t border-slate-100 p-4">
                          <div className="overflow-hidden rounded-2xl border border-slate-200">
                            <div className="overflow-x-auto">
                              <table className="min-w-full text-sm">
                                <thead className="bg-slate-50 text-slate-600">
                                  <tr>
                                    <th className="px-4 py-3 text-left">상품명</th>
                                    <th className="px-4 py-3 text-right">수량</th>
                                    <th className="px-4 py-3 text-right">판매가</th>
                                    <th className="px-4 py-3 text-right">매입원가</th>
                                    <th className="px-4 py-3 text-right">개당원가</th>
                                    <th className="px-4 py-3 text-right">원가율</th>
                                    <th className="px-4 py-3 text-right">마진율</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {record.items.map((item) => (
                                    <tr key={item.id} className="border-t border-slate-100">
                                      <td className="px-4 py-3">
                                        <div>{item.name}</div>
                                        {item.buyKrw > 0 && (
                                          <div className="text-xs text-amber-600 mt-0.5">실결제 {formatNumber(item.buyKrw)}원</div>
                                        )}
                                      </td>
                                      <td className="px-4 py-3 text-right">{formatNumber(item.qty)}</td>
                                      <td className="px-4 py-3 text-right">{formatNumber(item.salePrice)}원</td>
                                      <td className="px-4 py-3 text-right">{formatNumber(item.unitBuyCostKrw)}원</td>
                                      <td className="px-4 py-3 text-right">{formatNumber(item.totalCostPerUnit)}원</td>
                                      <td className="px-4 py-3 text-right">{formatPercent(item.costRate)}</td>
                                      <td className="px-4 py-3 text-right font-medium text-emerald-600">{formatPercent(item.marginRate)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
