const sampleMonths = [
  {
    date: "2026-01-15",
    salesFactor: 0.88,
    marginShift: 0.012,
    boosts: { Electronics: 0.96, Apparel: 0.9, Home: 1.02, Beauty: 0.97, Grocery: 1.04 },
  },
  {
    date: "2026-02-15",
    salesFactor: 0.91,
    marginShift: 0.014,
    boosts: { Electronics: 0.98, Apparel: 0.93, Home: 1.01, Beauty: 0.98, Grocery: 1.03 },
  },
  {
    date: "2026-03-15",
    salesFactor: 0.97,
    marginShift: 0.011,
    boosts: { Electronics: 1.02, Apparel: 1.01, Home: 1.03, Beauty: 1.02, Grocery: 1.01 },
  },
  {
    date: "2026-04-15",
    salesFactor: 1.01,
    marginShift: 0.006,
    boosts: { Electronics: 0.99, Apparel: 1.05, Home: 1.01, Beauty: 1.03, Grocery: 1.02 },
  },
  {
    date: "2026-05-15",
    salesFactor: 1.05,
    marginShift: 0.008,
    boosts: { Electronics: 1.01, Apparel: 1.08, Home: 1.04, Beauty: 1.01, Grocery: 1.02 },
  },
  {
    date: "2026-06-15",
    salesFactor: 1.09,
    marginShift: 0.01,
    boosts: { Electronics: 1.03, Apparel: 1.02, Home: 1.06, Beauty: 1.04, Grocery: 1.03 },
  },
  {
    date: "2026-07-15",
    salesFactor: 1.1,
    marginShift: 0.009,
    boosts: { Electronics: 1.05, Apparel: 0.99, Home: 1.05, Beauty: 1.03, Grocery: 1.02 },
  },
  {
    date: "2026-08-15",
    salesFactor: 1.06,
    marginShift: -0.002,
    boosts: { Electronics: 1.07, Apparel: 0.95, Home: 1.02, Beauty: 1.01, Grocery: 1.01 },
  },
  {
    date: "2026-09-15",
    salesFactor: 1.03,
    marginShift: 0.001,
    boosts: { Electronics: 1.01, Apparel: 0.97, Home: 1.01, Beauty: 1.02, Grocery: 1.02 },
  },
  {
    date: "2026-10-15",
    salesFactor: 1.12,
    marginShift: 0.01,
    boosts: { Electronics: 1.08, Apparel: 1.03, Home: 1.04, Beauty: 1.08, Grocery: 1.01 },
  },
  {
    date: "2026-11-15",
    salesFactor: 1.24,
    marginShift: 0.018,
    boosts: { Electronics: 1.24, Apparel: 1.04, Home: 1.08, Beauty: 1.1, Grocery: 0.99 },
  },
  {
    date: "2026-12-15",
    salesFactor: 1.31,
    marginShift: 0.02,
    boosts: { Electronics: 1.28, Apparel: 1.06, Home: 1.09, Beauty: 1.16, Grocery: 1.01 },
  },
];

const sampleProducts = [
  { product: "4K TV", category: "Electronics", baseSales: 190000, baseMargin: 0.18, marginBias: 0.014 },
  { product: "Wireless Earbuds", category: "Electronics", baseSales: 118000, baseMargin: 0.16, marginBias: 0.006 },
  { product: "Running Shoes", category: "Apparel", baseSales: 98000, baseMargin: 0.24, marginBias: 0.009 },
  { product: "Urban Jacket", category: "Apparel", baseSales: 86000, baseMargin: 0.2, marginBias: -0.01 },
  { product: "Coffee Maker", category: "Home", baseSales: 91000, baseMargin: 0.22, marginBias: 0.008 },
  { product: "Air Purifier", category: "Home", baseSales: 124000, baseMargin: 0.19, marginBias: -0.004 },
  { product: "Skin Serum", category: "Beauty", baseSales: 76000, baseMargin: 0.31, marginBias: 0.018 },
  { product: "Fragrance Set", category: "Beauty", baseSales: 69000, baseMargin: 0.28, marginBias: 0.012 },
  { product: "Organic Snacks", category: "Grocery", baseSales: 118000, baseMargin: 0.11, marginBias: 0.005 },
  { product: "Protein Pack", category: "Grocery", baseSales: 102000, baseMargin: 0.13, marginBias: 0.007 },
];

const currencyCompact = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  notation: "compact",
  maximumFractionDigits: 1,
});

const currencyWhole = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0,
});

const numberWhole = new Intl.NumberFormat("en-US");

const percentFormat = new Intl.NumberFormat("en-US", {
  style: "percent",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});

const initialDataset = createDataset(buildSampleRecords());

const state = {
  startDate: "",
  endDate: "",
  category: "All",
  product: "All",
  dataset: initialDataset,
  filteredRecords: [],
  filteredRange: null,
  sourceLabel: "Built-in sample data",
  sourceKind: "sample",
  uploadTone: "neutral",
  uploadMessage: "Using built-in sample dataset generated from the required columns.",
};

const chartInstances = {};
let chartFallbackRendered = false;
let renderFrameHandle = 0;

const elements = {
  csvInput: document.getElementById("csvInput"),
  resetDataButton: document.getElementById("resetDataButton"),
  startDateFilter: document.getElementById("startDateFilter"),
  endDateFilter: document.getElementById("endDateFilter"),
  categoryFilter: document.getElementById("categoryFilter"),
  productFilter: document.getElementById("productFilter"),
  exportCsvButton: document.getElementById("exportCsvButton"),
  dataSourceBadge: document.getElementById("dataSourceBadge"),
  uploadStatus: document.getElementById("uploadStatus"),
  heroPeriodLabel: document.getElementById("heroPeriodLabel"),
  storyTitle: document.getElementById("storyTitle"),
  storyBody: document.getElementById("storyBody"),
  heroDateRange: document.getElementById("heroDateRange"),
  heroCategory: document.getElementById("heroCategory"),
  heroProduct: document.getElementById("heroProduct"),
  kpiSales: document.getElementById("kpiSales"),
  kpiSalesDelta: document.getElementById("kpiSalesDelta"),
  kpiProfit: document.getElementById("kpiProfit"),
  kpiProfitDelta: document.getElementById("kpiProfitDelta"),
  kpiAvgSale: document.getElementById("kpiAvgSale"),
  kpiAvgSaleDelta: document.getElementById("kpiAvgSaleDelta"),
  kpiTopCategory: document.getElementById("kpiTopCategory"),
  kpiTopCategoryDelta: document.getElementById("kpiTopCategoryDelta"),
  kpiTopCategoryFoot: document.getElementById("kpiTopCategoryFoot"),
  signalList: document.getElementById("signalList"),
  mixTitle: document.getElementById("mixTitle"),
  mixBreakdown: document.getElementById("mixBreakdown"),
  comparisonTitle: document.getElementById("comparisonTitle"),
  comparisonSummary: document.getElementById("comparisonSummary"),
  productTableBody: document.getElementById("productTableBody"),
  monthSummary: document.getElementById("monthSummary"),
  watchList: document.getElementById("watchList"),
  performanceChart: document.getElementById("performanceChart"),
  mixChart: document.getElementById("mixChart"),
  comparisonChart: document.getElementById("comparisonChart"),
};

elements.startDateFilter.addEventListener("change", (event) => {
  state.startDate = event.target.value || state.dataset.minDateInput;

  if (state.endDate && state.startDate > state.endDate) {
    state.endDate = state.startDate;
    elements.endDateFilter.value = state.endDate;
  }

  syncProductFilter();
  scheduleRender();
});

elements.categoryFilter.addEventListener("change", (event) => {
  state.category = event.target.value;
  syncProductFilter();
  scheduleRender();
});

elements.endDateFilter.addEventListener("change", (event) => {
  state.endDate = event.target.value || state.dataset.maxDateInput;

  if (state.startDate && state.endDate < state.startDate) {
    state.startDate = state.endDate;
    elements.startDateFilter.value = state.startDate;
  }

  syncProductFilter();
  scheduleRender();
});

elements.productFilter.addEventListener("change", (event) => {
  state.product = event.target.value;
  scheduleRender();
});

elements.exportCsvButton.addEventListener("click", () => {
  exportFilteredCsv();
});

elements.resetDataButton.addEventListener("click", () => {
  state.dataset = createDataset(buildSampleRecords());
  state.sourceLabel = "Built-in sample data";
  state.sourceKind = "sample";
  state.uploadTone = "neutral";
  state.uploadMessage = "Using built-in sample dataset generated from the required columns.";
  if (elements.csvInput) {
    elements.csvInput.value = "";
  }
  resetFiltersToDataset();
  scheduleRender();
});

elements.csvInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const parsed = parseCsvData(text);
    state.dataset = createDataset(parsed.records);
    state.sourceLabel = file.name;
    state.sourceKind = "upload";
    state.uploadTone = "success";
    state.uploadMessage =
      parsed.skippedRows > 0
        ? `Loaded ${numberWhole.format(parsed.records.length)} valid rows from ${file.name}. Skipped ${numberWhole.format(parsed.skippedRows)} invalid rows.`
        : `Loaded ${numberWhole.format(parsed.records.length)} valid rows from ${file.name}.`;
    resetFiltersToDataset();
    scheduleRender();
  } catch (error) {
    state.uploadTone = "error";
    state.uploadMessage = error.message;
    updateUploadState();
  } finally {
    event.target.value = "";
  }
});

if (window.Chart) {
  Chart.defaults.font.family = '"Plus Jakarta Sans", sans-serif';
  Chart.defaults.color = "rgba(21, 52, 59, 0.72)";
}

resetFiltersToDataset();
scheduleRender();

// Batch rapid UI changes into one paint so larger uploads still feel responsive.
function scheduleRender() {
  if (renderFrameHandle) {
    return;
  }

  const schedule =
    typeof window.requestAnimationFrame === "function"
      ? window.requestAnimationFrame.bind(window)
      : (callback) => window.setTimeout(callback, 16);

  renderFrameHandle = schedule(() => {
    renderFrameHandle = 0;
    renderDashboard();
  });
}

function buildSampleRecords() {
  return sampleMonths.flatMap((month, monthIndex) =>
    sampleProducts.map((item, productIndex) => {
      const wave =
        1 +
        Math.sin((monthIndex + 1) * (productIndex + 2)) * 0.04 +
        Math.cos((monthIndex + 2) * (productIndex + 1)) * 0.02;
      const sales = Math.round(
        item.baseSales * month.salesFactor * month.boosts[item.category] * wave
      );
      const marginRate = clamp(
        item.baseMargin +
          item.marginBias +
          month.marginShift +
          getMarginEvent(item.product, monthIndex),
        0.03,
        0.42
      );

      return {
        date: month.date,
        product: item.product,
        category: item.category,
        sales,
        profit: Math.round(sales * marginRate),
      };
    })
  );
}

function getMarginEvent(product, monthIndex) {
  if (product === "Urban Jacket" && monthIndex >= 3 && monthIndex <= 4) {
    return -0.085;
  }

  if (product === "Wireless Earbuds" && monthIndex === 7) {
    return -0.055;
  }

  if (product === "4K TV" && monthIndex >= 10) {
    return -0.018;
  }

  if (product === "Fragrance Set" && monthIndex === 11) {
    return 0.024;
  }

  return 0;
}

function createDataset(records) {
  const normalizedRecords = records
    .map((record) => {
      const date = record.date instanceof Date ? new Date(record.date.getTime()) : new Date(record.date);
      return {
        date,
        product: String(record.product).trim(),
        category: String(record.category).trim(),
        sales: Number(record.sales),
        profit: Number(record.profit),
        monthKey: formatMonthKey(date),
        monthLabel: formatMonthLabel(date),
      };
    })
    .filter((record) => !Number.isNaN(record.date.getTime()) && Number.isFinite(record.sales) && Number.isFinite(record.profit))
    .sort((left, right) => left.date - right.date);

  const monthMap = new Map();
  const categorySet = new Set();
  const productSet = new Set();

  normalizedRecords.forEach((record) => {
    if (!monthMap.has(record.monthKey)) {
      monthMap.set(record.monthKey, {
        key: record.monthKey,
        label: record.monthLabel,
      });
    }

    categorySet.add(record.category);
    productSet.add(record.product);
  });

  const minDate = normalizedRecords.length ? startOfDay(normalizedRecords[0].date) : null;
  const maxDate = normalizedRecords.length ? startOfDay(normalizedRecords[normalizedRecords.length - 1].date) : null;

  return {
    records: normalizedRecords,
    months: Array.from(monthMap.values()),
    categories: Array.from(categorySet).sort((left, right) => left.localeCompare(right)),
    products: Array.from(productSet).sort((left, right) => left.localeCompare(right)),
    minDate,
    maxDate,
    minDateInput: minDate ? formatDateInput(minDate) : "",
    maxDateInput: maxDate ? formatDateInput(maxDate) : "",
  };
}

function formatMonthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function formatMonthLabel(date) {
  return date.toLocaleDateString("en-US", {
    month: "short",
    year: "numeric",
  });
}

function resetFiltersToDataset() {
  state.startDate = state.dataset.minDateInput;
  state.endDate = state.dataset.maxDateInput;
  state.category = "All";
  state.product = "All";
  syncDateFilters();
  syncCategoryFilter();
  syncProductFilter();
}

function syncDateFilters() {
  elements.startDateFilter.min = state.dataset.minDateInput;
  elements.startDateFilter.max = state.dataset.maxDateInput;
  elements.endDateFilter.min = state.dataset.minDateInput;
  elements.endDateFilter.max = state.dataset.maxDateInput;
  elements.startDateFilter.value = state.startDate;
  elements.endDateFilter.value = state.endDate;
}

function syncCategoryFilter() {
  const availableCategories = ["All"].concat(state.dataset.categories);
  const currentValue = availableCategories.includes(state.category) ? state.category : "All";
  elements.categoryFilter.innerHTML = availableCategories
    .map((category) => {
      const label = category === "All" ? "All Categories" : category;
      return `<option value="${escapeHtml(category)}">${escapeHtml(label)}</option>`;
    })
    .join("");
  elements.categoryFilter.value = currentValue;
  state.category = currentValue;
}

function syncProductFilter() {
  const productSet = new Set();
  const rangeStart = startOfDay(parseInputDate(state.startDate) || state.dataset.minDate);
  const rangeEnd = endOfDay(parseInputDate(state.endDate) || state.dataset.maxDate);

  state.dataset.records.forEach((record) => {
    if (
      record.date >= rangeStart &&
      record.date <= rangeEnd &&
      (state.category === "All" || record.category === state.category)
    ) {
      productSet.add(record.product);
    }
  });

  const availableProducts = ["All"].concat(
    Array.from(productSet).sort((left, right) => left.localeCompare(right))
  );
  const currentValue = availableProducts.includes(state.product) ? state.product : "All";

  elements.productFilter.innerHTML = availableProducts
    .map((product) => {
      const label = product === "All" ? "All Products" : product;
      return `<option value="${escapeHtml(product)}">${escapeHtml(label)}</option>`;
    })
    .join("");

  elements.productFilter.value = currentValue;
  state.product = currentValue;
}

function renderDashboard() {
  const range = getSelectedRange();
  const previousRange = getPreviousRange(range.start, range.end);
  const currentMonths = buildMonthsInRange(range.start, range.end);
  const currentRecords = filterRecords(
    state.dataset.records,
    range.start,
    range.end,
    state.category,
    state.product
  );
  const previousRecords = filterRecords(
    state.dataset.records,
    previousRange.start,
    previousRange.end,
    state.category,
    state.product
  );
  const currentMetrics = aggregateRecords(currentRecords);
  const previousMetrics = aggregateRecords(previousRecords);
  const monthlySeries = buildMonthlySeries(currentMonths, currentRecords);
  const categorySeries = buildGroupSeries(currentRecords, "category");
  const productSeries = buildGroupSeries(currentRecords, "product");
  const categorySalesSeries = categorySeries;
  const profitDistributionSeries = buildProfitDistributionSeries(categorySeries, productSeries);
  const topCategory = categorySeries[0] || null;
  const topProduct = productSeries[0] || null;
  const weakestCategory = categorySeries.length
    ? categorySeries.slice().sort((left, right) => left.marginRate - right.marginRate)[0]
    : null;
  const weakestProduct = productSeries.length
    ? productSeries.slice().sort((left, right) => left.profit - right.profit)[0]
    : null;
  const negativeProducts = productSeries.filter((item) => item.profit < 0);
  const bestMonth = monthlySeries.length
    ? monthlySeries.slice().sort((left, right) => right.sales - left.sales)[0]
    : null;

  state.filteredRecords = currentRecords;
  state.filteredRange = range;

  updateUploadState();
  updateContextTitles();
  updateExportButton(currentRecords.length);
  renderHero({
    range,
    currentMonths,
    currentMetrics,
    previousMetrics,
    topCategory,
    topProduct,
  });
  renderKpis(currentMetrics, previousMetrics, topCategory);
  renderSignals({
    currentMetrics,
    previousMetrics,
    topCategory,
    topProduct,
    weakestCategory,
    weakestProduct,
    monthlySeries,
  });
  renderMixBreakdown(categorySalesSeries);
  renderComparisonSummary(profitDistributionSeries);
  renderProductTable(productSeries, currentMetrics.sales);
  renderMonthSummary(monthlySeries);
  renderWatchList({
    currentMetrics,
    previousMetrics,
    topCategory,
    topProduct,
    weakestCategory,
    weakestProduct,
    negativeProducts,
    bestMonth,
  });
  renderCharts({
    monthlySeries,
    categorySalesSeries,
    profitDistributionSeries,
  });
}

function getSelectedRange() {
  return {
    start: startOfDay(parseInputDate(state.startDate) || state.dataset.minDate),
    end: endOfDay(parseInputDate(state.endDate) || state.dataset.maxDate),
  };
}

function getPreviousRange(start, end) {
  const startTime = start.getTime();
  const endTime = end.getTime();
  const rangeLength = endTime - startTime;
  const previousEnd = endOfDay(new Date(startTime - 24 * 60 * 60 * 1000));
  const previousStart = startOfDay(new Date(previousEnd.getTime() - rangeLength));

  return {
    start: previousStart,
    end: previousEnd,
  };
}

function buildMonthsInRange(start, end) {
  const months = [];
  const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  const lastMonth = new Date(end.getFullYear(), end.getMonth(), 1);

  while (cursor <= lastMonth) {
    months.push({
      key: formatMonthKey(cursor),
      label: formatMonthLabel(cursor),
    });
    cursor.setMonth(cursor.getMonth() + 1);
  }

  return months;
}

function filterRecords(records, startDate, endDate, category, product) {
  return records.filter((record) => {
    return (
      record.date >= startDate &&
      record.date <= endDate &&
      (category === "All" || record.category === category) &&
      (product === "All" || record.product === product)
    );
  });
}

function aggregateRecords(records) {
  const productSet = new Set();
  const categorySet = new Set();
  let sales = 0;
  let profit = 0;

  records.forEach((record) => {
    sales += record.sales;
    profit += record.profit;
    productSet.add(record.product);
    categorySet.add(record.category);
  });

  return {
    sales,
    profit,
    rows: records.length,
    marginRate: sales ? profit / sales : 0,
    avgSale: records.length ? sales / records.length : 0,
    activeProducts: productSet.size,
    activeCategories: categorySet.size,
  };
}

function buildMonthlySeries(months, records) {
  const monthMap = new Map();

  months.forEach((month) => {
    monthMap.set(month.key, {
      key: month.key,
      label: month.label,
      sales: 0,
      profit: 0,
      rows: 0,
    });
  });

  records.forEach((record) => {
    const bucket = monthMap.get(record.monthKey);

    if (!bucket) {
      return;
    }

    bucket.sales += record.sales;
    bucket.profit += record.profit;
    bucket.rows += 1;
  });

  return Array.from(monthMap.values()).map((bucket) => {
    return {
      key: bucket.key,
      label: bucket.label,
      sales: bucket.sales,
      profit: bucket.profit,
      rows: bucket.rows,
      marginRate: bucket.sales ? bucket.profit / bucket.sales : 0,
    };
  });
}

function buildGroupSeries(records, groupKey) {
  const groupMap = new Map();
  let totalSales = 0;

  records.forEach((record) => {
    const groupValue = record[groupKey];

    if (!groupMap.has(groupValue)) {
      groupMap.set(groupValue, {
        name: groupValue,
        category: record.category,
        sales: 0,
        profit: 0,
        rows: 0,
      });
    }

    const bucket = groupMap.get(groupValue);
    bucket.sales += record.sales;
    bucket.profit += record.profit;
    bucket.rows += 1;
    totalSales += record.sales;
  });

  return Array.from(groupMap.values())
    .map((bucket) => {
      return {
        name: bucket.name,
        category: bucket.category,
        sales: bucket.sales,
        profit: bucket.profit,
        rows: bucket.rows,
        marginRate: bucket.sales ? bucket.profit / bucket.sales : 0,
        mix: totalSales ? bucket.sales / totalSales : 0,
      };
    })
    .sort((left, right) => right.sales - left.sales);
}

function buildProfitDistributionSeries(categorySeries, productSeries) {
  const baseSeries = state.category === "All" ? categorySeries : productSeries;
  const positiveSeries = baseSeries.filter((item) => item.profit > 0);

  if (positiveSeries.length <= 6) {
    return positiveSeries;
  }

  const topSeries = positiveSeries.slice(0, 5);
  const remainingSeries = positiveSeries.slice(5);
  const otherProfit = remainingSeries.reduce((sum, item) => sum + item.profit, 0);
  const otherSales = remainingSeries.reduce((sum, item) => sum + item.sales, 0);
  const totalProfit = positiveSeries.reduce((sum, item) => sum + item.profit, 0);

  topSeries.push({
    name: "Other",
    category: state.category,
    sales: otherSales,
    profit: otherProfit,
    rows: remainingSeries.reduce((sum, item) => sum + item.rows, 0),
    marginRate: otherSales ? otherProfit / otherSales : 0,
    mix: totalProfit ? otherProfit / totalProfit : 0,
  });

  return topSeries;
}

function renderHero(context) {
  const salesDelta = calculateRelativeDelta(context.currentMetrics.sales, context.previousMetrics.sales);
  const profitDelta = calculateRelativeDelta(context.currentMetrics.profit, context.previousMetrics.profit);
  const hasRows = context.currentMetrics.rows > 0;
  const topCategoryLabel =
    state.category === "All"
      ? context.topCategory
        ? context.topCategory.name
        : "No data"
      : state.category;
  const topProductLabel = context.topProduct ? context.topProduct.name : "No data";

  elements.dataSourceBadge.textContent =
    state.sourceKind === "upload"
      ? `CSV: ${truncateLabel(state.sourceLabel, 26)}`
      : state.sourceLabel;
  elements.heroPeriodLabel.textContent = getWindowLabel(context.currentMonths.length);
  elements.heroDateRange.textContent = getDateRangeLabelFromBounds(context.range.start, context.range.end);
  elements.heroCategory.textContent = topCategoryLabel;
  elements.heroProduct.textContent = topProductLabel;

  if (!hasRows) {
    elements.storyTitle.textContent = "No data matches the current filters.";
    elements.storyBody.textContent =
      "Try widening the date range or resetting the category and product filters so the dashboard has visible rows to analyze.";
    return;
  }

  if (salesDelta !== null && salesDelta > 0.03 && profitDelta !== null && profitDelta > 0.03) {
    elements.storyTitle.textContent = "Sales and profit are both accelerating.";
  } else if (salesDelta !== null && salesDelta > 0 && profitDelta !== null && profitDelta < 0) {
    elements.storyTitle.textContent = "Sales are growing, but profit quality is slipping.";
  } else if (salesDelta !== null && salesDelta < 0 && profitDelta !== null && profitDelta >= 0) {
    elements.storyTitle.textContent = "Revenue is softer, but profitability is holding.";
  } else if (salesDelta !== null && salesDelta < 0 && profitDelta !== null && profitDelta < 0) {
    elements.storyTitle.textContent = "Sales and profit are both under pressure.";
  } else {
    elements.storyTitle.textContent = "The business is holding a steady pace.";
  }

  const categoryText = context.topCategory
    ? `${context.topCategory.name} contributes ${percentFormat.format(context.topCategory.mix)} of visible sales`
    : "No category signal is available";
  const productText = context.topProduct
    ? `${context.topProduct.name} is the leading product with ${currencyCompact.format(context.topProduct.sales)} in sales`
    : "No product signal is available";

  elements.storyBody.textContent =
    `${categoryText}, ${productText}, and the current view is running at ${percentFormat.format(context.currentMetrics.marginRate)} margin across ${numberWhole.format(context.currentMetrics.rows)} rows.`;
}

function renderKpis(currentMetrics, previousMetrics, topCategory) {
  elements.kpiSales.textContent = currencyCompact.format(currentMetrics.sales);
  setRelativeDelta(elements.kpiSalesDelta, calculateRelativeDelta(currentMetrics.sales, previousMetrics.sales));

  elements.kpiProfit.textContent = currencyCompact.format(currentMetrics.profit);
  setRelativeDelta(elements.kpiProfitDelta, calculateRelativeDelta(currentMetrics.profit, previousMetrics.profit));

  elements.kpiAvgSale.textContent = currencyWhole.format(currentMetrics.avgSale);
  setRelativeDelta(elements.kpiAvgSaleDelta, calculateRelativeDelta(currentMetrics.avgSale, previousMetrics.avgSale));

  if (!topCategory) {
    elements.kpiTopCategory.textContent = "No data";
    setNeutralDelta(elements.kpiTopCategoryDelta, "Adjust filters");
    elements.kpiTopCategoryFoot.textContent = "Highest sales contribution";
    return;
  }

  elements.kpiTopCategory.textContent = topCategory.name;

  if (state.category === "All") {
    setNeutralDelta(elements.kpiTopCategoryDelta, `${formatUnsignedPercent(topCategory.mix)} of sales`);
    elements.kpiTopCategoryFoot.textContent = `${currencyCompact.format(topCategory.profit)} profit contribution`;
    return;
  }

  setNeutralDelta(elements.kpiTopCategoryDelta, "Selected category");
  elements.kpiTopCategoryFoot.textContent = `${currencyCompact.format(topCategory.sales)} visible sales`;
}

function renderSignals(context) {
  if (context.currentMetrics.rows === 0) {
    elements.signalList.innerHTML = `
      <li class="signal-list__item">
        <strong>No data in view</strong>
        <span>The selected filters do not return any rows. Broaden the date range or reset the category and product filters.</span>
      </li>
    `;
    return;
  }

  const salesDelta = calculateRelativeDelta(context.currentMetrics.sales, context.previousMetrics.sales);
  const profitDelta = calculateRelativeDelta(context.currentMetrics.profit, context.previousMetrics.profit);
  const items = [];
  const monthTrendInsight = buildMonthTrendInsight(context.monthlySeries);

  if (monthTrendInsight) {
    items.push({
      title: "Sales Trend",
      text: monthTrendInsight,
    });
  }

  if (state.category === "All" && context.topCategory) {
    items.push({
      title: "Top Category",
      text: `${context.topCategory.name} is the top-performing category, contributing ${formatUnsignedPercent(context.topCategory.mix)} of sales and ${currencyCompact.format(context.topCategory.profit)} in profit.`,
    });
  } else if (context.topProduct) {
    items.push({
      title: "Top Product",
      text: `${context.topProduct.name} is the top-performing product${state.category !== "All" ? ` in ${state.category}` : ""}, generating ${currencyCompact.format(context.topProduct.sales)} in sales and ${currencyCompact.format(context.topProduct.profit)} in profit.`,
    });
  }

  if (context.weakestProduct && context.weakestProduct.profit < 0) {
    items.push({
      title: "Profit Alert",
      text: `${context.weakestProduct.name} is currently loss-making at ${currencyCompact.format(context.weakestProduct.profit)} profit, so it needs immediate attention.`,
    });
  } else if (context.weakestCategory) {
    items.push({
      title: "Margin Insight",
      text: `${context.weakestCategory.name} has the weakest margin at ${percentFormat.format(context.weakestCategory.marginRate)}, while the overall view is at ${percentFormat.format(context.currentMetrics.marginRate)}.`,
    });
  } else {
    items.push({
      title: "Profit Insight",
      text:
        profitDelta === null
          ? `${currencyCompact.format(context.currentMetrics.profit)} in profit is currently visible at a ${percentFormat.format(context.currentMetrics.marginRate)} margin.`
          : `Profit is ${describeDelta(profitDelta)} versus the previous comparison window, with margin at ${percentFormat.format(context.currentMetrics.marginRate)}.`,
    });
  }

  elements.signalList.innerHTML = items
    .slice(0, 3)
    .map((item) => {
      return `
        <li class="signal-list__item">
          <strong>${escapeHtml(item.title)}</strong>
          <span>${escapeHtml(item.text)}</span>
        </li>
      `;
    })
    .join("");
}

function buildMonthTrendInsight(monthlySeries) {
  if (!monthlySeries || monthlySeries.length < 2) {
    return null;
  }

  let bestChange = null;

  for (let index = 1; index < monthlySeries.length; index += 1) {
    const currentMonth = monthlySeries[index];
    const previousMonth = monthlySeries[index - 1];

    if (previousMonth.sales <= 0) {
      continue;
    }

    const delta = currentMonth.sales / previousMonth.sales - 1;

    if (!bestChange || Math.abs(delta) > Math.abs(bestChange.delta)) {
      bestChange = {
        delta,
        currentLabel: currentMonth.label,
        previousLabel: previousMonth.label,
      };
    }
  }

  if (!bestChange) {
    return null;
  }

  if (bestChange.delta >= 0) {
    return `Sales increased by ${formatUnsignedPercent(bestChange.delta)} in ${bestChange.currentLabel} compared with ${bestChange.previousLabel}.`;
  }

  return `Sales decreased by ${formatUnsignedPercent(Math.abs(bestChange.delta))} in ${bestChange.currentLabel} compared with ${bestChange.previousLabel}.`;
}

function renderMixBreakdown(entries) {
  if (!entries.length) {
    elements.mixBreakdown.innerHTML = `<p class="stack-empty">No category sales data is available for the current filters.</p>`;
    return;
  }

  elements.mixBreakdown.innerHTML = entries
    .map((entry) => {
      return `
        <div class="stack-item">
          <div>
            <div class="stack-item__label">${escapeHtml(entry.name)}</div>
            <div class="stack-item__sub">${currencyCompact.format(entry.profit)} profit · ${percentFormat.format(entry.mix)} of sales</div>
          </div>
          <div class="stack-item__value">${currencyCompact.format(entry.sales)}</div>
        </div>
      `;
    })
    .join("");
}

function renderComparisonSummary(entries) {
  if (!entries.length) {
    elements.comparisonSummary.innerHTML = `<p class="stack-empty">No positive profit distribution is available for the current filters.</p>`;
    return;
  }

  const totalProfit = entries.reduce((sum, entry) => sum + entry.profit, 0);
  elements.comparisonSummary.innerHTML = entries
    .map((entry) => {
      const share = totalProfit ? entry.profit / totalProfit : 0;
      return `
        <div class="stack-item">
          <div>
            <div class="stack-item__label">${escapeHtml(entry.name)}</div>
            <div class="stack-item__sub">${percentFormat.format(share)} of visible profit</div>
          </div>
          <div class="stack-item__value">${currencyCompact.format(entry.profit)}</div>
        </div>
      `;
    })
    .join("");
}

function renderProductTable(productSeries, totalSales) {
  if (!productSeries.length) {
    elements.productTableBody.innerHTML = `
      <tr>
        <td class="performance-table__empty" colspan="6">No products are visible for the current filters.</td>
      </tr>
    `;
    return;
  }

  elements.productTableBody.innerHTML = productSeries
    .slice(0, 10)
    .map((product) => {
      const profitTone = product.profit > 0 ? "positive" : product.profit < 0 ? "negative" : "neutral";
      const mix = totalSales ? product.sales / totalSales : 0;
      return `
        <tr>
          <td><strong>${escapeHtml(product.name)}</strong></td>
          <td>${escapeHtml(product.category)}</td>
          <td>${currencyCompact.format(product.sales)}</td>
          <td><span class="table-pill is-${profitTone}">${currencyCompact.format(product.profit)}</span></td>
          <td>${percentFormat.format(product.marginRate)}</td>
          <td>${percentFormat.format(mix)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderMonthSummary(monthlySeries) {
  if (!monthlySeries.length) {
    elements.monthSummary.innerHTML = `<p class="stack-empty">No monthly trend is available for the current filters.</p>`;
    return;
  }

  elements.monthSummary.innerHTML = monthlySeries
    .slice()
    .reverse()
    .map((month) => {
      return `
        <div class="stack-item">
          <div>
            <div class="stack-item__label">${escapeHtml(month.label)}</div>
            <div class="stack-item__sub">${currencyCompact.format(month.profit)} profit · ${percentFormat.format(month.marginRate)} margin</div>
          </div>
          <div class="stack-item__value">${currencyCompact.format(month.sales)}</div>
        </div>
      `;
    })
    .join("");
}

function renderWatchList(context) {
  if (context.currentMetrics.rows === 0) {
    elements.watchList.innerHTML = `
      <li class="watch-list__item">
        <strong>No watchlist items</strong>
        <span>Once rows are visible, this section will highlight concentration, low-profit products, and the strongest month.</span>
      </li>
    `;
    return;
  }

  const salesDelta = calculateRelativeDelta(context.currentMetrics.sales, context.previousMetrics.sales);
  const leadMix =
    state.category === "All" && context.topCategory
      ? context.topCategory.mix
      : context.topProduct
      ? context.topProduct.mix
      : 0;

  const items = [
    {
      tone: salesDelta === null ? "neutral" : salesDelta >= 0 ? "good" : "bad",
      title: "Sales momentum",
      text:
        salesDelta === null
          ? `${currencyCompact.format(context.currentMetrics.sales)} is the current visible sales total, but there is no prior window available for comparison.`
          : `${currencyCompact.format(context.currentMetrics.sales)} is ${describeDelta(salesDelta)} compared with the previous window.`,
    },
    {
      tone: context.negativeProducts.length ? "bad" : "good",
      title: "Loss-making products",
      text: context.negativeProducts.length
        ? `${numberWhole.format(context.negativeProducts.length)} visible products are below zero profit. Start with ${context.negativeProducts[0].name}.`
        : "All visible products are currently profitable in the selected scope.",
    },
    {
      tone: leadMix > 0.45 ? "bad" : "neutral",
      title: "Concentration",
      text:
        state.category === "All" && context.topCategory
          ? `${context.topCategory.name} accounts for ${percentFormat.format(context.topCategory.mix)} of sales, so keep an eye on concentration risk.`
          : context.topProduct
          ? `${context.topProduct.name} accounts for ${percentFormat.format(context.topProduct.mix)} of the selected category's sales.`
          : "No concentration signal is available.",
    },
    {
      tone: "good",
      title: "Best visible month",
      text: context.bestMonth
        ? `${context.bestMonth.label} is the strongest month in view with ${currencyCompact.format(context.bestMonth.sales)} in sales and ${currencyCompact.format(context.bestMonth.profit)} in profit.`
        : "No month signal is available.",
    },
  ];

  if (context.weakestCategory) {
    items[3] = {
      tone: context.weakestCategory.marginRate < context.currentMetrics.marginRate * 0.7 ? "bad" : "neutral",
      title: "Margin gap",
      text: `${context.weakestCategory.name} has the thinnest category margin at ${percentFormat.format(context.weakestCategory.marginRate)}.`,
    };
  }

  elements.watchList.innerHTML = items
    .map((item) => {
      const toneClass = item.tone === "neutral" ? "" : `is-${item.tone}`;
      return `
        <li class="watch-list__item ${toneClass}">
          <strong>${escapeHtml(item.title)}</strong>
          <span>${escapeHtml(item.text)}</span>
        </li>
      `;
    })
    .join("");
}

function renderCharts(context) {
  if (!window.Chart) {
    renderChartFallbacks();
    return;
  }

  upsertChart("performance", elements.performanceChart, {
    type: "line",
    data: {
      labels: context.monthlySeries.map((item) => item.label),
      datasets: [
        {
          label: "Sales",
          data: context.monthlySeries.map((item) => item.sales),
          borderColor: "#2c6c6d",
          backgroundColor: "rgba(44, 108, 109, 0.16)",
          borderWidth: 3,
          tension: 0.32,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: true,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          labels: {
            usePointStyle: true,
            boxWidth: 10,
            boxHeight: 10,
            padding: 18,
          },
        },
        tooltip: {
          callbacks: {
            label(contextValue) {
              return `${contextValue.dataset.label}: ${currencyWhole.format(contextValue.raw)}`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: {
            callback(value) {
              return currencyCompact.format(value);
            },
          },
          grid: {
            color: "rgba(21, 52, 59, 0.08)",
          },
        },
      },
    },
  });

  upsertChart("mix", elements.mixChart, {
    type: "bar",
    data: {
      labels: context.categorySalesSeries.length ? context.categorySalesSeries.map((item) => item.name) : ["No data"],
      datasets: [
        {
          label: "Sales",
          data: context.categorySalesSeries.length ? context.categorySalesSeries.map((item) => item.sales) : [0],
          backgroundColor: ["#2c6c6d", "#de6b48", "#d8a644", "#7b9ca8", "#4f7f72", "#c98558"],
          borderRadius: 12,
          maxBarThickness: 42,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          callbacks: {
            label(contextValue) {
              if (!context.categorySalesSeries.length) {
                return "No data";
              }

              const total = context.categorySalesSeries.reduce((sum, item) => sum + item.sales, 0);
              const share = total ? contextValue.raw / total : 0;
              return `${contextValue.label}: ${currencyWhole.format(contextValue.raw)} (${percentFormat.format(share)})`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: {
            callback(value) {
              return currencyCompact.format(value);
            },
          },
          grid: {
            color: "rgba(21, 52, 59, 0.08)",
          },
        },
      },
    },
  });

  upsertChart("comparison", elements.comparisonChart, {
    type: "pie",
    data: {
      labels: context.profitDistributionSeries.length ? context.profitDistributionSeries.map((item) => item.name) : ["No data"],
      datasets: [
        {
          data: context.profitDistributionSeries.length ? context.profitDistributionSeries.map((item) => item.profit) : [1],
          backgroundColor: ["#2c6c6d", "#de6b48", "#d8a644", "#7b9ca8", "#4f7f72", "#c98558"],
          borderColor: "rgba(255, 250, 244, 0.9)",
          borderWidth: 2,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label(contextValue) {
              if (!context.profitDistributionSeries.length) {
                return "No data";
              }

              const total = context.profitDistributionSeries.reduce((sum, item) => sum + item.profit, 0);
              const share = total ? contextValue.raw / total : 0;
              return `${contextValue.label}: ${currencyWhole.format(contextValue.raw)} (${percentFormat.format(share)})`;
            },
          },
        },
      },
    },
  });
}

function updateContextTitles() {
  elements.mixTitle.textContent = "Category-wise sales";
  elements.comparisonTitle.textContent =
    state.category === "All" ? "Profit distribution" : `Profit distribution in ${state.category}`;
}

function updateUploadState() {
  elements.uploadStatus.textContent = state.uploadMessage;
  elements.uploadStatus.className = "upload-status";

  if (state.uploadTone === "success") {
    elements.uploadStatus.classList.add("is-success");
  } else if (state.uploadTone === "error") {
    elements.uploadStatus.classList.add("is-error");
  }
}

function updateExportButton(recordCount) {
  elements.exportCsvButton.disabled = recordCount === 0;
  elements.exportCsvButton.textContent =
    recordCount === 0 ? "No filtered rows to export" : "Download filtered CSV";
}

function exportFilteredCsv() {
  if (!state.filteredRecords.length || !state.filteredRange) {
    return;
  }

  const headers = ["Date", "Product", "Category", "Sales", "Profit"];
  const rows = state.filteredRecords.map((record) => {
    return [
      formatDateInput(record.date),
      record.product,
      record.category,
      formatCsvNumber(record.sales),
      formatCsvNumber(record.profit),
    ];
  });

  const csv = [headers, ...rows]
    .map((row) => row.map(escapeCsvCell).join(","))
    .join("\r\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = buildExportFilename();
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function buildExportFilename() {
  const startLabel = formatDateInput(state.filteredRange.start);
  const endLabel = formatDateInput(state.filteredRange.end);
  const categoryLabel = slugifyForFilename(state.category === "All" ? "all-categories" : state.category);
  const productLabel = slugifyForFilename(state.product === "All" ? "all-products" : state.product);
  return `filtered-sales-${startLabel}_to_${endLabel}-${categoryLabel}-${productLabel}.csv`;
}

function renderChartFallbacks() {
  if (chartFallbackRendered) {
    return;
  }

  [elements.performanceChart, elements.mixChart, elements.comparisonChart].forEach((canvas) => {
    const fallback = document.createElement("div");
    fallback.className = "chart-fallback";
    fallback.textContent = "Charts need access to the Chart.js CDN. The upload flow, KPI cards, and tables still work without it.";
    canvas.replaceWith(fallback);
  });

  chartFallbackRendered = true;
}

function upsertChart(key, canvas, config) {
  const existingChart = chartInstances[key];

  if (existingChart) {
    // Reuse chart instances instead of recreating them on every filter change.
    existingChart.config.type = config.type;
    existingChart.data = config.data;
    existingChart.options = config.options;
    existingChart.update("none");
    return existingChart;
  }

  chartInstances[key] = new Chart(canvas, config);
  return chartInstances[key];
}

function calculateRelativeDelta(currentValue, previousValue) {
  if (previousValue === 0) {
    return currentValue === 0 ? 0 : null;
  }

  return currentValue / previousValue - 1;
}

function calculatePointDelta(currentValue, previousValue, previousRows) {
  if (!previousRows) {
    return null;
  }

  return currentValue - previousValue;
}

function setRelativeDelta(element, delta) {
  if (delta === null) {
    setNeutralDelta(element, "No prior window");
    return;
  }

  const tone = Math.abs(delta) < 0.005 ? "neutral" : delta > 0 ? "positive" : "negative";
  element.className = `metric-card__delta is-${tone}`;
  element.textContent = `${signedPercent(delta)} vs prev`;
}

function setPointDelta(element, delta) {
  if (delta === null) {
    setNeutralDelta(element, "No prior window");
    return;
  }

  const tone = Math.abs(delta) < 0.002 ? "neutral" : delta > 0 ? "positive" : "negative";
  element.className = `metric-card__delta is-${tone}`;
  element.textContent = `${signedPoints(delta)} vs prev`;
}

function setNeutralDelta(element, text) {
  element.className = "metric-card__delta is-neutral";
  element.textContent = text;
}

function describeDelta(delta) {
  if (delta === null) {
    return "without a prior window available";
  }

  if (Math.abs(delta) < 0.005) {
    return "essentially flat";
  }

  return delta > 0 ? `up ${formatUnsignedPercent(delta)}` : `down ${formatUnsignedPercent(Math.abs(delta))}`;
}

function describePointDelta(delta) {
  if (delta === null) {
    return "with no prior window available";
  }

  if (Math.abs(delta) < 0.002) {
    return "with margin essentially flat versus the previous window";
  }

  return delta > 0
    ? `with margin up ${formatUnsignedPoints(delta)}`
    : `with margin down ${formatUnsignedPoints(Math.abs(delta))}`;
}

function signedPercent(value) {
  const prefix = value >= 0 ? "+" : "";
  return `${prefix}${(value * 100).toFixed(1)}%`;
}

function formatUnsignedPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function signedPoints(value) {
  const prefix = value >= 0 ? "+" : "";
  return `${prefix}${(value * 100).toFixed(1)} pts`;
}

function formatUnsignedPoints(value) {
  return `${(value * 100).toFixed(1)} pts`;
}

function getWindowLabel(monthCount) {
  if (!monthCount) {
    return "No data";
  }

  if (monthCount === 1) {
    return "1 month";
  }

  if (monthCount >= 12) {
    return "Full year";
  }

  return `Last ${monthCount} months`;
}

function getDateRangeLabelFromBounds(start, end) {
  if (!start || !end) {
    return "No data";
  }

  return `${start.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  })} - ${end.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  })}`;
}

function truncateLabel(value, limit) {
  if (value.length <= limit) {
    return value;
  }

  return `${value.slice(0, Math.max(0, limit - 1))}…`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatDateInput(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatCsvNumber(value) {
  return Number.isFinite(value) ? String(value) : "";
}

function escapeCsvCell(value) {
  const text = String(value === undefined || value === null ? "" : value);

  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function slugifyForFilename(value) {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "all";
}

function parseInputDate(value) {
  if (!value) {
    return null;
  }

  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseCsvData(text) {
  const rows = parseCsvRows(text).filter((row) => row.some((cell) => cell.trim() !== ""));

  if (!rows.length) {
    throw new Error("The CSV is empty. Add rows with Date, Product, Category, Sales, and Profit.");
  }

  const headers = rows[0].map(normalizeHeader);
  const requiredColumns = ["date", "product", "category", "sales", "profit"];
  const columnIndexes = {};

  requiredColumns.forEach((column) => {
    const index = headers.indexOf(column);
    if (index === -1) {
      throw new Error(`Missing required column: ${column.charAt(0).toUpperCase()}${column.slice(1)}.`);
    }
    columnIndexes[column] = index;
  });

  const records = [];
  let skippedRows = 0;

  rows.slice(1).forEach((row) => {
    const rawDate = getCell(row, columnIndexes.date).trim();
    const rawProduct = getCell(row, columnIndexes.product).trim();
    const rawCategory = getCell(row, columnIndexes.category).trim();
    const rawSales = getCell(row, columnIndexes.sales).trim();
    const rawProfit = getCell(row, columnIndexes.profit).trim();

    if (!rawDate && !rawProduct && !rawCategory && !rawSales && !rawProfit) {
      return;
    }

    const parsedDate = parseDateValue(rawDate);
    const parsedSales = parseNumberValue(rawSales);
    const parsedProfit = parseNumberValue(rawProfit);

    if (!parsedDate || !rawProduct || !rawCategory || Number.isNaN(parsedSales) || Number.isNaN(parsedProfit)) {
      skippedRows += 1;
      return;
    }

    records.push({
      date: parsedDate,
      product: rawProduct,
      category: rawCategory,
      sales: parsedSales,
      profit: parsedProfit,
    });
  });

  if (!records.length) {
    throw new Error("No valid rows were found. Check that Date, Product, Category, Sales, and Profit are populated correctly.");
  }

  return {
    records,
    skippedRows,
  };
}

function parseCsvRows(text) {
  const cleaned = text.replace(/^\uFEFF/, "");
  const rows = [];
  let currentRow = [];
  let currentValue = "";
  let inQuotes = false;

  for (let index = 0; index < cleaned.length; index += 1) {
    const character = cleaned[index];
    const nextCharacter = cleaned[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        currentValue += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === "," && !inQuotes) {
      currentRow.push(currentValue);
      currentValue = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }

      currentRow.push(currentValue);
      rows.push(currentRow);
      currentRow = [];
      currentValue = "";
      continue;
    }

    currentValue += character;
  }

  if (currentValue.length > 0 || currentRow.length > 0) {
    currentRow.push(currentValue);
    rows.push(currentRow);
  }

  return rows;
}

function normalizeHeader(value) {
  return value.trim().toLowerCase().replace(/\s+/g, "");
}

function getCell(row, index) {
  return row[index] === undefined ? "" : row[index];
}

function parseDateValue(value) {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseNumberValue(value) {
  const trimmed = value.trim();

  if (!trimmed) {
    return Number.NaN;
  }

  const isBracketNegative = trimmed.startsWith("(") && trimmed.endsWith(")");
  const sanitized = trimmed.replace(/[,$%\s()]/g, "");
  const numericValue = Number(sanitized);

  if (Number.isNaN(numericValue)) {
    return Number.NaN;
  }

  return isBracketNegative ? -Math.abs(numericValue) : numericValue;
}
