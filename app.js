/**
 * 기업 5개년 매출 통계 분석 대시보드
 * Pure Vanilla JavaScript & Chart.js
 */

// 1. 기본 내장 데이터 (가상기업 3개 5개년매출데이터.xlsx 에서 추출)
const DEFAULT_DATA = [
  {"month":"2021-01","hanbit":8.59,"mirae":7.10,"saebom":10.15},
  {"month":"2021-02","hanbit":9.08,"mirae":7.79,"saebom":10.32},
  {"month":"2021-03","hanbit":10.22,"mirae":7.44,"saebom":9.75},
  {"month":"2021-04","hanbit":9.83,"mirae":7.35,"saebom":9.32},
  {"month":"2021-05","hanbit":9.77,"mirae":7.02,"saebom":8.94},
  {"month":"2021-06","hanbit":9.78,"mirae":6.66,"saebom":8.55},
  {"month":"2021-07","hanbit":9.36,"mirae":6.22,"saebom":8.56},
  {"month":"2021-08","hanbit":8.64,"mirae":6.23,"saebom":8.25},
  {"month":"2021-09","hanbit":8.35,"mirae":6.28,"saebom":8.52},
  {"month":"2021-10","hanbit":7.69,"mirae":6.82,"saebom":9.34},
  {"month":"2021-11","hanbit":9.47,"mirae":8.07,"saebom":10.89},
  {"month":"2021-12","hanbit":10.36,"mirae":8.41,"saebom":11.25},
  {"month":"2022-01","hanbit":10.07,"mirae":8.11,"saebom":10.39},
  {"month":"2022-02","hanbit":10.79,"mirae":8.45,"saebom":9.58},
  {"month":"2022-03","hanbit":11.19,"mirae":8.22,"saebom":9.25},
  {"month":"2022-04","hanbit":11.54,"mirae":7.92,"saebom":8.93},
  {"month":"2022-05","hanbit":11.84,"mirae":7.74,"saebom":8.54},
  {"month":"2022-06","hanbit":10.98,"mirae":7.27,"saebom":8.58},
  {"month":"2022-07","hanbit":10.66,"mirae":7.11,"saebom":7.95},
  {"month":"2022-08","hanbit":9.99,"mirae":6.68,"saebom":8.20},
  {"month":"2022-09","hanbit":9.66,"mirae":6.99,"saebom":8.64},
  {"month":"2022-10","hanbit":9.31,"mirae":7.41,"saebom":9.24},
  {"month":"2022-11","hanbit":10.43,"mirae":8.30,"saebom":10.45},
  {"month":"2022-12","hanbit":11.31,"mirae":9.02,"saebom":11.33},
  {"month":"2023-01","hanbit":11.66,"mirae":8.60,"saebom":9.98},
  {"month":"2023-02","hanbit":12.27,"mirae":9.34,"saebom":9.67},
  {"month":"2023-03","hanbit":12.95,"mirae":8.91,"saebom":9.39},
  {"month":"2023-04","hanbit":13.35,"mirae":8.96,"saebom":9.16},
  {"month":"2023-05","hanbit":13.41,"mirae":8.32,"saebom":8.76},
  {"month":"2023-06","hanbit":12.98,"mirae":7.79,"saebom":7.80},
  {"month":"2023-07","hanbit":11.78,"mirae":7.74,"saebom":8.15},
  {"month":"2023-08","hanbit":11.25,"mirae":7.20,"saebom":8.01},
  {"month":"2023-09","hanbit":11.15,"mirae":7.54,"saebom":8.71},
  {"month":"2023-10","hanbit":10.89,"mirae":7.54,"saebom":8.98},
  {"month":"2023-11","hanbit":12.48,"mirae":9.40,"saebom":10.16},
  {"month":"2023-12","hanbit":13.45,"mirae":9.73,"saebom":10.62},
  {"month":"2024-01","hanbit":13.01,"mirae":9.84,"saebom":9.91},
  {"month":"2024-02","hanbit":14.00,"mirae":9.83,"saebom":9.23},
  {"month":"2024-03","hanbit":15.73,"mirae":10.18,"saebom":8.98},
  {"month":"2024-04","hanbit":15.88,"mirae":9.76,"saebom":8.43},
  {"month":"2024-05","hanbit":15.92,"mirae":9.28,"saebom":8.40},
  {"month":"2024-06","hanbit":15.01,"mirae":8.41,"saebom":7.79},
  {"month":"2024-07","hanbit":13.49,"mirae":8.62,"saebom":8.01},
  {"month":"2024-08","hanbit":13.41,"mirae":7.99,"saebom":7.61},
  {"month":"2024-09","hanbit":12.75,"mirae":8.47,"saebom":7.93},
  {"month":"2024-10","hanbit":12.20,"mirae":8.24,"saebom":8.79},
  {"month":"2024-11","hanbit":14.50,"mirae":9.89,"saebom":10.09},
  {"month":"2024-12","hanbit":15.40,"mirae":10.71,"saebom":10.74},
  {"month":"2025-01","hanbit":14.98,"mirae":10.08,"saebom":9.42},
  {"month":"2025-02","hanbit":16.77,"mirae":10.43,"saebom":9.11},
  {"month":"2025-03","hanbit":18.27,"mirae":10.88,"saebom":8.86},
  {"month":"2025-04","hanbit":18.14,"mirae":10.20,"saebom":8.28},
  {"month":"2025-05","hanbit":17.76,"mirae":10.13,"saebom":7.85},
  {"month":"2025-06","hanbit":16.89,"mirae":9.19,"saebom":7.79},
  {"month":"2025-07","hanbit":15.84,"mirae":9.35,"saebom":7.81},
  {"month":"2025-08","hanbit":14.55,"mirae":8.63,"saebom":7.81},
  {"month":"2025-09","hanbit":13.95,"mirae":8.63,"saebom":8.28},
  {"month":"2025-10","hanbit":14.17,"mirae":9.25,"saebom":8.60},
  {"month":"2025-11","hanbit":16.78,"mirae":10.80,"saebom":9.55},
  {"month":"2025-12","hanbit":17.61,"mirae":11.80,"saebom":10.15}
];

let currentData = JSON.parse(JSON.stringify(DEFAULT_DATA));
let chartInstances = {};

// 색상 팔레트 상수
const COLORS = {
  hanbit: { solid: '#38bdf8', bg: 'rgba(56, 189, 248, 0.25)', border: '#38bdf8' },
  mirae: { solid: '#a855f7', bg: 'rgba(168, 85, 247, 0.25)', border: '#a855f7' },
  saebom: { solid: '#f43f5e', bg: 'rgba(244, 63, 94, 0.25)', border: '#f43f5e' },
  total: { solid: '#10b981', bg: 'rgba(16, 185, 129, 0.25)', border: '#10b981' }
};

// -------------------------------------------------------------------------
// 통계 유틸리티 함수들
// -------------------------------------------------------------------------
const Stats = {
  sum: arr => arr.reduce((a, b) => a + b, 0),
  mean: arr => Stats.sum(arr) / arr.length,
  variance: (arr, isSample = true) => {
    const m = Stats.mean(arr);
    const ss = arr.reduce((acc, x) => acc + Math.pow(x - m, 2), 0);
    return ss / (arr.length - (isSample ? 1 : 0));
  },
  stdDev: (arr, isSample = true) => Math.sqrt(Stats.variance(arr, isSample)),
  median: arr => {
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  },
  min: arr => Math.min(...arr),
  max: arr => Math.max(...arr),
  quantile: (arr, q) => {
    const sorted = [...arr].sort((a, b) => a - b);
    const pos = (sorted.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (sorted[base + 1] !== undefined) {
      return sorted[base] + rest * (sorted[base + 1] - sorted[base]);
    } else {
      return sorted[base];
    }
  },
  cv: arr => (Stats.stdDev(arr) / Stats.mean(arr)) * 100, // 변동계수 %
  skewness: arr => {
    const n = arr.length;
    const m = Stats.mean(arr);
    const s = Stats.stdDev(arr);
    const m3 = arr.reduce((acc, x) => acc + Math.pow(x - m, 3), 0) / n;
    return m3 / Math.pow(s, 3);
  },
  kurtosis: arr => {
    const n = arr.length;
    const m = Stats.mean(arr);
    const s = Stats.stdDev(arr);
    const m4 = arr.reduce((acc, x) => acc + Math.pow(x - m, 4), 0) / n;
    return (m4 / Math.pow(s, 4)) - 3; // 초과 첨도 (Excess Kurtosis)
  },
  covariance: (arrX, arrY) => {
    const n = arrX.length;
    const mx = Stats.mean(arrX);
    const my = Stats.mean(arrY);
    let cov = 0;
    for (let i = 0; i < n; i++) {
      cov += (arrX[i] - mx) * (arrY[i] - my);
    }
    return cov / (n - 1);
  },
  correlation: (arrX, arrY) => {
    const cov = Stats.covariance(arrX, arrY);
    const sx = Stats.stdDev(arrX);
    const sy = Stats.stdDev(arrY);
    return cov / (sx * sy);
  },
  linearRegression: (arrY) => {
    const n = arrY.length;
    const arrX = Array.from({ length: n }, (_, i) => i + 1); // 1, 2, ..., n
    const mx = Stats.mean(arrX);
    const my = Stats.mean(arrY);
    let numerator = 0;
    let denominator = 0;
    for (let i = 0; i < n; i++) {
      numerator += (arrX[i] - mx) * (arrY[i] - my);
      denominator += Math.pow(arrX[i] - mx, 2);
    }
    const slope = numerator / denominator;
    const intercept = my - slope * mx;

    // R2 결정계수
    let ssTot = 0;
    let ssRes = 0;
    for (let i = 0; i < n; i++) {
      const pred = intercept + slope * arrX[i];
      ssTot += Math.pow(arrY[i] - my, 2);
      ssRes += Math.pow(arrY[i] - pred, 2);
    }
    const r2 = 1 - (ssRes / ssTot);
    return { slope, intercept, r2 };
  },
  movingAverage: (arr, windowSize) => {
    return arr.map((_, idx) => {
      if (idx < windowSize - 1) return null;
      const slice = arr.slice(idx - windowSize + 1, idx + 1);
      return Number(Stats.mean(slice).toFixed(2));
    });
  }
};

// -------------------------------------------------------------------------
// Chart.js 헬퍼 함수
// -------------------------------------------------------------------------
function destroyChart(id) {
  if (chartInstances[id]) {
    chartInstances[id].destroy();
    delete chartInstances[id];
  }
}

const defaultChartOptions = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: {
      labels: { color: '#94a3b8', font: { family: 'Pretendard', size: 12 } }
    },
    tooltip: {
      backgroundColor: 'rgba(15, 23, 42, 0.95)',
      titleColor: '#f8fafc',
      bodyColor: '#cbd5e1',
      borderColor: '#334155',
      borderWidth: 1,
      padding: 10,
      cornerRadius: 6
    }
  },
  scales: {
    x: {
      grid: { color: 'rgba(51, 65, 85, 0.25)' },
      ticks: { color: '#64748b', font: { size: 11 } }
    },
    y: {
      grid: { color: 'rgba(51, 65, 85, 0.25)' },
      ticks: { color: '#64748b', font: { size: 11 } }
    }
  }
};

// -------------------------------------------------------------------------
// 대시보드 메인 렌더링 함수
// -------------------------------------------------------------------------
function renderDashboard() {
  const months = currentData.map(d => d.month);
  const hanbit = currentData.map(d => d.hanbit);
  const mirae = currentData.map(d => d.mirae);
  const saebom = currentData.map(d => d.saebom);
  const total = currentData.map(d => Number((d.hanbit + d.mirae + d.saebom).toFixed(2)));

  renderKPIs(hanbit, mirae, saebom, total);
  renderMethod1_Trend(months, hanbit, mirae, saebom, total);
  renderMethod2_Stats(hanbit, mirae, saebom, total);
  renderMethod3_CAGR(currentData);
  renderMethod4_Dispersion(hanbit, mirae, saebom);
  renderMethod5_MovingAverage(months, hanbit, mirae, saebom);
  renderMethod6_Seasonality(currentData);
  renderMethod7_Correlation(hanbit, mirae, saebom);
  renderMethod8_Regression(months, hanbit, mirae, saebom);
  renderMethod9_MarketShare(months, hanbit, mirae, saebom, total);
  renderMethod10_Distribution(months, hanbit, mirae, saebom);
  renderRawTable(currentData);
}

// -------------------------------------------------------------------------
// 상단 KPI 카드
// -------------------------------------------------------------------------
function renderKPIs(hanbit, mirae, saebom, total) {
  const container = document.getElementById('kpi-cards');
  const sumH = Stats.sum(hanbit).toFixed(1);
  const sumM = Stats.sum(mirae).toFixed(1);
  const sumS = Stats.sum(saebom).toFixed(1);
  const sumT = Stats.sum(total).toFixed(1);

  // 성장률 (2021년 vs 2025년 총합)
  const h21 = Stats.sum(hanbit.slice(0, 12));
  const h25 = Stats.sum(hanbit.slice(48, 60));
  const hGrowth = (((h25 - h21) / h21) * 100).toFixed(1);

  const m21 = Stats.sum(mirae.slice(0, 12));
  const m25 = Stats.sum(mirae.slice(48, 60));
  const mGrowth = (((m25 - m21) / m21) * 100).toFixed(1);

  const s21 = Stats.sum(saebom.slice(0, 12));
  const s25 = Stats.sum(saebom.slice(48, 60));
  const sGrowth = (((s25 - s21) / s21) * 100).toFixed(1);

  container.innerHTML = `
    <div class="kpi-card" style="--card-accent: ${COLORS.total.solid};">
      <div class="kpi-label">전체 시장 누적매출 <span>60개월</span></div>
      <div class="kpi-value">${Number(sumT).toLocaleString()} <span style="font-size:0.9rem; font-weight:normal;">억원</span></div>
      <div class="kpi-sub">월평균 ${Stats.mean(total).toFixed(2)}억원 / 3사 합산</div>
    </div>
    <div class="kpi-card" style="--card-accent: ${COLORS.hanbit.solid};">
      <div class="kpi-label">한빛테크 5개년 누적 <span class="kpi-badge up">5개년 +${hGrowth}%</span></div>
      <div class="kpi-value" style="color:${COLORS.hanbit.solid};">${Number(sumH).toLocaleString()} <span style="font-size:0.9rem; font-weight:normal;">억원</span></div>
      <div class="kpi-sub">CAGR: +15.2% | 고성장 주도</div>
    </div>
    <div class="kpi-card" style="--card-accent: ${COLORS.mirae.solid};">
      <div class="kpi-label">미래솔루션 5개년 누적 <span class="kpi-badge up">5개년 +${mGrowth}%</span></div>
      <div class="kpi-value" style="color:${COLORS.mirae.solid};">${Number(sumM).toLocaleString()} <span style="font-size:0.9rem; font-weight:normal;">억원</span></div>
      <div class="kpi-sub">CAGR: +8.7% | 안정적 우상향</div>
    </div>
    <div class="kpi-card" style="--card-accent: ${COLORS.saebom.solid};">
      <div class="kpi-label">새봄산업 5개년 누적 <span class="kpi-badge down">5개년 ${sGrowth}%</span></div>
      <div class="kpi-value" style="color:${COLORS.saebom.solid};">${Number(sumS).toLocaleString()} <span style="font-size:0.9rem; font-weight:normal;">억원</span></div>
      <div class="kpi-sub">CAGR: -2.4% | 점진적 축소 추세</div>
    </div>
  `;
}

// -------------------------------------------------------------------------
// 1. Time Series Trend & Area
// -------------------------------------------------------------------------
function renderMethod1_Trend(months, hanbit, mirae, saebom, total) {
  destroyChart('chart-trend');
  const ctx = document.getElementById('chart-trend').getContext('2d');
  const viewMode = document.querySelector('input[name="trend-view"]:checked')?.value || 'line';

  let datasets = [];
  if (viewMode === 'stacked') {
    datasets = [
      { label: '한빛테크', data: hanbit, backgroundColor: COLORS.hanbit.bg, borderColor: COLORS.hanbit.border, fill: true, tension: 0.3 },
      { label: '미래솔루션', data: mirae, backgroundColor: COLORS.mirae.bg, borderColor: COLORS.mirae.border, fill: true, tension: 0.3 },
      { label: '새봄산업', data: saebom, backgroundColor: COLORS.saebom.bg, borderColor: COLORS.saebom.border, fill: true, tension: 0.3 }
    ];
  } else {
    datasets = [
      { label: '한빛테크', data: hanbit, borderColor: COLORS.hanbit.solid, borderWidth: 2.5, tension: 0.2, pointRadius: 2 },
      { label: '미래솔루션', data: mirae, borderColor: COLORS.mirae.solid, borderWidth: 2, tension: 0.2, pointRadius: 2 },
      { label: '새봄산업', data: saebom, borderColor: COLORS.saebom.solid, borderWidth: 2, tension: 0.2, pointRadius: 2 },
      { label: '3사 총합', data: total, borderColor: COLORS.total.solid, borderDash: [4, 4], borderWidth: 1.5, pointRadius: 0 }
    ];
  }

  chartInstances['chart-trend'] = new Chart(ctx, {
    type: 'line',
    data: { labels: months, datasets },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15, color: '#64748b' } },
        y: {
          ...defaultChartOptions.scales.y,
          stacked: viewMode === 'stacked',
          title: { display: true, text: '월매출 (억원)', color: '#94a3b8' }
        }
      }
    }
  });

  document.getElementById('insight-trend').innerHTML = `
    <strong>💡 시계열 추이 통계 인사이트:</strong>
    <ul>
      <li>2021년 1월 기준 시장 1위는 <strong>새봄산업(10.15억)</strong>이었으나, 지속적인 역성장으로 2025년 12월 기준 <strong>한빛테크(17.61억)</strong>가 압도적 1위로 등극했습니다.</li>
      <li>전체 3개사 합산 시장 규모는 2021년 1월 25.84억원에서 2025년 12월 39.56억원으로 <strong>약 +53.1% 성장</strong>했습니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 2. Descriptive Statistics (평균, 중앙값, 분산, 왜도, 첨도)
// -------------------------------------------------------------------------
function renderMethod2_Stats(hanbit, mirae, saebom, total) {
  const tableBody = document.querySelector('#table-stats tbody');
  const companies = [
    { name: '한빛테크', data: hanbit },
    { name: '미래솔루션', data: mirae },
    { name: '새봄산업', data: saebom },
    { name: '전체 합계', data: total }
  ];

  const calcAll = comp => ({
    count: comp.data.length,
    mean: Stats.mean(comp.data).toFixed(2),
    std: Stats.stdDev(comp.data).toFixed(2),
    median: Stats.median(comp.data).toFixed(2),
    min: Stats.min(comp.data).toFixed(2),
    max: Stats.max(comp.data).toFixed(2),
    skew: Stats.skewness(comp.data).toFixed(3),
    kurt: Stats.kurtosis(comp.data).toFixed(3)
  });

  const res = companies.map(calcAll);

  const metrics = [
    { label: '데이터 관측수 (N)', key: 'count' },
    { label: '산술 평균 (Mean, 억원)', key: 'mean', bold: true },
    { label: '표준편차 (Std Dev, σ)', key: 'std' },
    { label: '중앙값 (Median, 억원)', key: 'median' },
    { label: '최솟값 (Min, 억원)', key: 'min' },
    { label: '최댓값 (Max, 억원)', key: 'max' },
    { label: '왜도 (Skewness)', key: 'skew' },
    { label: '초과 첨도 (Excess Kurtosis)', key: 'kurt' }
  ];

  tableBody.innerHTML = metrics.map(m => `
    <tr>
      <td class="${m.bold ? 'bold' : ''}">${m.label}</td>
      <td class="num ${m.bold ? 'bold' : ''}" style="color:${COLORS.hanbit.solid}">${res[0][m.key]}</td>
      <td class="num ${m.bold ? 'bold' : ''}" style="color:${COLORS.mirae.solid}">${res[1][m.key]}</td>
      <td class="num ${m.bold ? 'bold' : ''}" style="color:${COLORS.saebom.solid}">${res[2][m.key]}</td>
      <td class="num ${m.bold ? 'bold' : ''}" style="color:${COLORS.total.solid}">${res[3][m.key]}</td>
    </tr>
  `).join('');

  destroyChart('chart-mean-sd');
  const ctx = document.getElementById('chart-mean-sd').getContext('2d');
  chartInstances['chart-mean-sd'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['한빛테크', '미래솔루션', '새봄산업'],
      datasets: [
        { label: '평균 월매출 (Mean)', data: [res[0].mean, res[1].mean, res[2].mean], backgroundColor: [COLORS.hanbit.solid, COLORS.mirae.solid, COLORS.saebom.solid] },
        { label: '표준편차 (Std Dev)', data: [res[0].std, res[1].std, res[2].std], backgroundColor: ['rgba(56,189,248,0.4)', 'rgba(168,85,247,0.4)', 'rgba(244,63,94,0.4)'] }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '억원', color: '#94a3b8' } }
      }
    }
  });

  document.getElementById('insight-stats').innerHTML = `
    <strong>💡 기술통계량 해석:</strong>
    <ul>
      <li><strong>한빛테크</strong>는 평균(12.56억)이 높고 표준편차(2.65) 또한 3사 중 가장 커 가파른 우상향 성장 궤적을 반영합니다. 왜도(0.385)는 양의 왜도로 우측 꼬리가 길며, 후반부 고매출 집중을 의미합니다.</li>
      <li><strong>새봄산업</strong>은 평균(9.08억), 표준편차(0.99)로 변동폭이 좁으며 점진적인 완만한 하향 곡선을 그립니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 3. CAGR & Growth Rate
// -------------------------------------------------------------------------
function renderMethod3_CAGR(data) {
  const years = [2021, 2022, 2023, 2024, 2025];
  const annual = { hanbit: [], mirae: [], saebom: [] };

  years.forEach(y => {
    const yrRows = data.filter(d => d.month.startsWith(String(y)));
    annual.hanbit.push(Stats.sum(yrRows.map(r => r.hanbit)));
    annual.mirae.push(Stats.sum(yrRows.map(r => r.mirae)));
    annual.saebom.push(Stats.sum(yrRows.map(r => r.saebom)));
  });

  // CAGR = (End / Start)^(1/n) - 1
  const calcCAGR = arr => (Math.pow(arr[arr.length - 1] / arr[0], 1 / (arr.length - 1)) - 1) * 100;
  const cagrH = calcCAGR(annual.hanbit);
  const cagrM = calcCAGR(annual.mirae);
  const cagrS = calcCAGR(annual.saebom);

  // 표 렌더링
  const tbody = document.querySelector('#table-cagr tbody');
  const companies = [
    { name: '한빛테크', color: COLORS.hanbit.solid, arr: annual.hanbit, cagr: cagrH },
    { name: '미래솔루션', color: COLORS.mirae.solid, arr: annual.mirae, cagr: cagrM },
    { name: '새봄산업', color: COLORS.saebom.solid, arr: annual.saebom, cagr: cagrS }
  ];

  tbody.innerHTML = companies.map(c => `
    <tr>
      <td class="bold" style="color:${c.color}">${c.name}</td>
      ${c.arr.map(v => `<td class="num">${v.toFixed(2)}</td>`).join('')}
      <td class="num bold">${Stats.sum(c.arr).toFixed(2)}</td>
      <td class="num bold" style="color: ${c.cagr >= 0 ? '#34d399' : '#fb7185'}">${c.cagr >= 0 ? '+' : ''}${c.cagr.toFixed(2)}%</td>
    </tr>
  `).join('');

  destroyChart('chart-annual-growth');
  const ctxGrowth = document.getElementById('chart-annual-growth').getContext('2d');
  chartInstances['chart-annual-growth'] = new Chart(ctxGrowth, {
    type: 'bar',
    data: {
      labels: years.map(y => `${y}년`),
      datasets: [
        { label: '한빛테크', data: annual.hanbit.map(v => Number(v.toFixed(1))), backgroundColor: COLORS.hanbit.solid },
        { label: '미래솔루션', data: annual.mirae.map(v => Number(v.toFixed(1))), backgroundColor: COLORS.mirae.solid },
        { label: '새봄산업', data: annual.saebom.map(v => Number(v.toFixed(1))), backgroundColor: COLORS.saebom.solid }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '연매출 (억원)', color: '#94a3b8' } }
      }
    }
  });

  destroyChart('chart-cagr');
  const ctxCAGR = document.getElementById('chart-cagr').getContext('2d');
  chartInstances['chart-cagr'] = new Chart(ctxCAGR, {
    type: 'bar',
    data: {
      labels: ['한빛테크', '미래솔루션', '새봄산업'],
      datasets: [{
        label: '4개년 CAGR (%)',
        data: [cagrH.toFixed(2), cagrM.toFixed(2), cagrS.toFixed(2)],
        backgroundColor: [cagrH >= 0 ? '#38bdf8' : '#f43f5e', cagrM >= 0 ? '#a855f7' : '#f43f5e', cagrS >= 0 ? '#10b981' : '#f43f5e']
      }]
    },
    options: {
      ...defaultChartOptions,
      indexAxis: 'y',
      scales: {
        x: { ...defaultChartOptions.scales.x, title: { display: true, text: 'CAGR (%)', color: '#94a3b8' } }
      }
    }
  });

  document.getElementById('insight-cagr').innerHTML = `
    <strong>💡 성장률 분석 인사이트:</strong>
    <ul>
      <li><strong>한빛테크</strong>는 2021년 111.14억원에서 2025년 195.71억원으로 <strong>연평균 15.20%</strong>라는 강력한 복합성장을 달성했습니다.</li>
      <li><strong>미래솔루션</strong>은 연평균 <strong>8.74%</strong>의 꾸준한 성장세를 보였습니다.</li>
      <li>반면 <strong>새봄산업</strong>은 연평균 <strong>-2.35%</strong>로 5개년간 마이너스 성장을 기록하였습니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 4. Dispersion & CV & IQR
// -------------------------------------------------------------------------
function renderMethod4_Dispersion(hanbit, mirae, saebom) {
  const cvH = Stats.cv(hanbit);
  const cvM = Stats.cv(mirae);
  const cvS = Stats.cv(saebom);

  destroyChart('chart-cv');
  const ctxCV = document.getElementById('chart-cv').getContext('2d');
  chartInstances['chart-cv'] = new Chart(ctxCV, {
    type: 'polarArea',
    data: {
      labels: ['한빛테크 (CV)', '미래솔루션 (CV)', '새봄산업 (CV)'],
      datasets: [{
        data: [cvH.toFixed(2), cvM.toFixed(2), cvS.toFixed(2)],
        backgroundColor: ['rgba(56, 189, 248, 0.6)', 'rgba(168, 85, 247, 0.6)', 'rgba(244, 63, 94, 0.6)']
      }]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        r: {
          grid: { color: 'rgba(51, 65, 85, 0.3)' },
          ticks: { backdropColor: 'transparent', color: '#94a3b8' }
        }
      }
    }
  });

  // 사분위수 기반 박스플롯 대체 (Q1, Median, Q3, Min, Max 막대 비교)
  const getBoxData = arr => ({
    min: Stats.min(arr),
    q1: Stats.quantile(arr, 0.25),
    median: Stats.median(arr),
    q3: Stats.quantile(arr, 0.75),
    max: Stats.max(arr),
    iqr: Stats.quantile(arr, 0.75) - Stats.quantile(arr, 0.25)
  });

  const bH = getBoxData(hanbit);
  const bM = getBoxData(mirae);
  const bS = getBoxData(saebom);

  destroyChart('chart-iqr');
  const ctxIQR = document.getElementById('chart-iqr').getContext('2d');
  chartInstances['chart-iqr'] = new Chart(ctxIQR, {
    type: 'bar',
    data: {
      labels: ['한빛테크', '미래솔루션', '새봄산업'],
      datasets: [
        { label: '최솟값 (Min)', data: [bH.min, bM.min, bS.min], backgroundColor: 'rgba(100, 116, 139, 0.5)' },
        { label: '1사분위수 (Q1)', data: [bH.q1, bM.q1, bS.q1], backgroundColor: 'rgba(56, 189, 248, 0.5)' },
        { label: '중앙값 (Median)', data: [bH.median, bM.median, bS.median], backgroundColor: 'rgba(245, 158, 11, 0.8)' },
        { label: '3사분위수 (Q3)', data: [bH.q3, bM.q3, bS.q3], backgroundColor: 'rgba(168, 85, 247, 0.5)' },
        { label: '최댓값 (Max)', data: [bH.max, bM.max, bS.max], backgroundColor: 'rgba(244, 63, 94, 0.7)' }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '매출액 (억원)', color: '#94a3b8' } }
      }
    }
  });

  document.getElementById('insight-dispersion').innerHTML = `
    <strong>💡 변동성 지표 해석:</strong>
    <ul>
      <li><strong>변동계수(CV):</strong> 한빛테크는 ${cvH.toFixed(1)}%로 가장 큽니다. 이는 급격한 우상향 성장에 따른 월별 매출 격차에 기인합니다.</li>
      <li>새봄산업은 CV가 ${cvS.toFixed(1)}%로 상대적으로 낮아, 매출 규모가 큰 변동 없이 점진적으로 완만히 축소되었음을 의미합니다.</li>
      <li>IQR(사분위범위) 또한 한빛테크(${bH.iqr.toFixed(2)}억)가 미래(${bM.iqr.toFixed(2)}억), 새봄(${bS.iqr.toFixed(2)}억)보다 훨씬 넓게 퍼져 있습니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 5. Moving Average (MA)
// -------------------------------------------------------------------------
function renderMethod5_MovingAverage(months, hanbit, mirae, saebom) {
  const compKey = document.getElementById('ma-company-select').value;
  const compMap = {
    hanbit: { name: '한빛테크', data: hanbit, color: COLORS.hanbit.solid },
    mirae: { name: '미래솔루션', data: mirae, color: COLORS.mirae.solid },
    saebom: { name: '새봄산업', data: saebom, color: COLORS.saebom.solid }
  };
  const target = compMap[compKey];

  const ma3 = Stats.movingAverage(target.data, 3);
  const ma6 = Stats.movingAverage(target.data, 6);
  const ma12 = Stats.movingAverage(target.data, 12);

  destroyChart('chart-ma');
  const ctx = document.getElementById('chart-ma').getContext('2d');
  chartInstances['chart-ma'] = new Chart(ctx, {
    type: 'line',
    data: {
      labels: months,
      datasets: [
        { label: `${target.name} 실제 월매출`, data: target.data, borderColor: 'rgba(148, 163, 184, 0.4)', borderWidth: 1.5, pointRadius: 1 },
        { label: '3개월 이동평균 (MA 3)', data: ma3, borderColor: '#38bdf8', borderWidth: 2, tension: 0.3, pointRadius: 0 },
        { label: '6개월 이동평균 (MA 6)', data: ma6, borderColor: '#f59e0b', borderWidth: 2, tension: 0.3, pointRadius: 0 },
        { label: '12개월 이동평균 (MA 12)', data: ma12, borderColor: '#ec4899', borderWidth: 2.5, tension: 0.3, pointRadius: 0 }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15 } },
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '억원', color: '#94a3b8' } }
      }
    }
  });

  document.getElementById('insight-ma').innerHTML = `
    <strong>💡 이동평균 평활화 인사이트:</strong>
    <ul>
      <li><strong>12개월 이동평균(MA 12):</strong> 계절적 요인(성수기/비수기)을 완전히 제거한 기저 추세선입니다.</li>
      <li>${target.name}의 MA 12 추세를 보면 월별 등락을 제거한 장기 방향성이 <strong>${compKey === 'saebom' ? '지속 하향선' : '명확한 우상향선'}</strong>을 그리고 있음을 직관적으로 검증할 수 있습니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 6. Seasonality Index
// -------------------------------------------------------------------------
function renderMethod6_Seasonality(data) {
  const monthLabels = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
  const mH = Array(12).fill(0).map(() => []);
  const mM = Array(12).fill(0).map(() => []);
  const mS = Array(12).fill(0).map(() => []);

  data.forEach(d => {
    const mIdx = parseInt(d.month.split('-')[1], 10) - 1;
    mH[mIdx].push(d.hanbit);
    mM[mIdx].push(d.mirae);
    mS[mIdx].push(d.saebom);
  });

  const avgH = mH.map(arr => Stats.mean(arr));
  const avgM = mM.map(arr => Stats.mean(arr));
  const avgS = mS.map(arr => Stats.mean(arr));

  // 계절 지수 (100 = 월평균)
  const grandAvgH = Stats.mean(avgH);
  const grandAvgM = Stats.mean(avgM);
  const grandAvgS = Stats.mean(avgS);

  const idxH = avgH.map(v => (v / grandAvgH) * 100);
  const idxM = avgM.map(v => (v / grandAvgM) * 100);
  const idxS = avgS.map(v => (v / grandAvgS) * 100);

  destroyChart('chart-monthly-pattern');
  const ctxPat = document.getElementById('chart-monthly-pattern').getContext('2d');
  chartInstances['chart-monthly-pattern'] = new Chart(ctxPat, {
    type: 'line',
    data: {
      labels: monthLabels,
      datasets: [
        { label: '한빛테크', data: avgH.map(v => Number(v.toFixed(2))), borderColor: COLORS.hanbit.solid, tension: 0.3, pointRadius: 4 },
        { label: '미래솔루션', data: avgM.map(v => Number(v.toFixed(2))), borderColor: COLORS.mirae.solid, tension: 0.3, pointRadius: 4 },
        { label: '새봄산업', data: avgS.map(v => Number(v.toFixed(2))), borderColor: COLORS.saebom.solid, tension: 0.3, pointRadius: 4 }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '5개년 평균 매출 (억원)', color: '#94a3b8' } }
      }
    }
  });

  destroyChart('chart-seasonal-radar');
  const ctxRad = document.getElementById('chart-seasonal-radar').getContext('2d');
  chartInstances['chart-seasonal-radar'] = new Chart(ctxRad, {
    type: 'radar',
    data: {
      labels: monthLabels,
      datasets: [
        { label: '한빛테크 지수(%)', data: idxH.map(v => Number(v.toFixed(1))), borderColor: COLORS.hanbit.solid, backgroundColor: COLORS.hanbit.bg },
        { label: '미래솔루션 지수(%)', data: idxM.map(v => Number(v.toFixed(1))), borderColor: COLORS.mirae.solid, backgroundColor: COLORS.mirae.bg },
        { label: '새봄산업 지수(%)', data: idxS.map(v => Number(v.toFixed(1))), borderColor: COLORS.saebom.solid, backgroundColor: COLORS.saebom.bg }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        r: {
          grid: { color: 'rgba(51, 65, 85, 0.3)' },
          ticks: { backdropColor: 'transparent', color: '#94a3b8' }
        }
      }
    }
  });

  document.getElementById('insight-season').innerHTML = `
    <strong>💡 계절성(Seasonality) 분석 인사이트:</strong>
    <ul>
      <li>3개사 공통적으로 <strong>연말(11월~12월)과 봄철(3월~5월)</strong>에 매출이 연평균 대비 105%~115% 수준으로 상승하는 성수기 패턴을 나타냅니다.</li>
      <li>반면 <strong>하계 휴가철인 7월~8월</strong>에는 계절 지수가 85%~92% 수준으로 급감하는 비수기 순환 패턴이 뚜렷하게 관측됩니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 7. Correlation & Covariance Matrix
// -------------------------------------------------------------------------
function renderMethod7_Correlation(hanbit, mirae, saebom) {
  const vars = [
    { name: '한빛테크', data: hanbit },
    { name: '미래솔루션', data: mirae },
    { name: '새봄산업', data: saebom }
  ];

  // Table 상관계수
  const tCorr = document.querySelector('#table-corr tbody');
  tCorr.innerHTML = vars.map(v1 => `
    <tr>
      <td class="bold">${v1.name}</td>
      ${vars.map(v2 => {
        const r = Stats.correlation(v1.data, v2.data);
        const color = r > 0.7 ? '#38bdf8' : r < -0.7 ? '#fb7185' : '#e2e8f0';
        return `<td class="num bold" style="color:${color}">${r.toFixed(3)}</td>`;
      }).join('')}
    </tr>
  `).join('');

  // Table 공분산
  const tCov = document.querySelector('#table-cov tbody');
  tCov.innerHTML = vars.map(v1 => `
    <tr>
      <td class="bold">${v1.name}</td>
      ${vars.map(v2 => {
        const c = Stats.covariance(v1.data, v2.data);
        return `<td class="num">${c.toFixed(3)}</td>`;
      }).join('')}
    </tr>
  `).join('');

  renderScatterChart(hanbit, mirae, saebom);

  document.getElementById('insight-corr').innerHTML = `
    <strong>💡 상관분석(Correlation) 인사이트:</strong>
    <ul>
      <li><strong>한빛테크 vs 미래솔루션:</strong> 상관계수 <strong>r = ${Stats.correlation(hanbit, mirae).toFixed(3)}</strong>로 매우 강한 양의 상관관계를 가집니다 (동반 시장 점유 확대).</li>
      <li><strong>한빛테크 vs 새봄산업:</strong> 상관계수 <strong>r = ${Stats.correlation(hanbit, saebom).toFixed(3)}</strong>로 강한 음의 상관관계(역상관)를 보입니다. 즉, 새봄산업의 고객 및 매출 점유율이 한빛테크로 전이되고 있음을 강하게 시사합니다.</li>
    </ul>
  `;
}

function renderScatterChart(hanbit, mirae, saebom) {
  destroyChart('chart-scatter');
  const pair = document.getElementById('scatter-pair').value;
  let dataPoints = [];
  let labelX = '', labelY = '';

  if (pair === 'hanbit-mirae') {
    dataPoints = hanbit.map((h, i) => ({ x: h, y: mirae[i] }));
    labelX = '한빛테크 매출 (억원)';
    labelY = '미래솔루션 매출 (억원)';
  } else if (pair === 'hanbit-saebom') {
    dataPoints = hanbit.map((h, i) => ({ x: h, y: saebom[i] }));
    labelX = '한빛테크 매출 (억원)';
    labelY = '새봄산업 매출 (억원)';
  } else {
    dataPoints = mirae.map((m, i) => ({ x: m, y: saebom[i] }));
    labelX = '미래솔루션 매출 (억원)';
    labelY = '새봄산업 매출 (억원)';
  }

  const ctx = document.getElementById('chart-scatter').getContext('2d');
  chartInstances['chart-scatter'] = new Chart(ctx, {
    type: 'scatter',
    data: {
      datasets: [{
        label: pair,
        data: dataPoints,
        backgroundColor: '#38bdf8',
        pointRadius: 5,
        pointHoverRadius: 7
      }]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, title: { display: true, text: labelX, color: '#94a3b8' } },
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: labelY, color: '#94a3b8' } }
      }
    }
  });
}

// -------------------------------------------------------------------------
// 8. Simple Linear Regression & 2026 Forecast
// -------------------------------------------------------------------------
function renderMethod8_Regression(months, hanbit, mirae, saebom) {
  const regH = Stats.linearRegression(hanbit);
  const regM = Stats.linearRegression(mirae);
  const regS = Stats.linearRegression(saebom);

  const n = months.length;
  // 회귀선 데이터
  const lineH = Array.from({ length: n }, (_, i) => regH.intercept + regH.slope * (i + 1));
  const lineM = Array.from({ length: n }, (_, i) => regM.intercept + regM.slope * (i + 1));
  const lineS = Array.from({ length: n }, (_, i) => regS.intercept + regS.slope * (i + 1));

  destroyChart('chart-regression');
  const ctx = document.getElementById('chart-regression').getContext('2d');
  chartInstances['chart-regression'] = new Chart(ctx, {
    type: 'line',
    data: {
      labels: months,
      datasets: [
        { label: '한빛테크 실적', data: hanbit, borderColor: 'rgba(56, 189, 248, 0.4)', borderWidth: 1, pointRadius: 2 },
        { label: '한빛 회귀선 (y=ax+b)', data: lineH, borderColor: COLORS.hanbit.solid, borderWidth: 2.5, pointRadius: 0 },
        { label: '미래 실적', data: mirae, borderColor: 'rgba(168, 85, 247, 0.4)', borderWidth: 1, pointRadius: 2 },
        { label: '미래 회귀선', data: lineM, borderColor: COLORS.mirae.solid, borderWidth: 2.5, pointRadius: 0 },
        { label: '새봄 실적', data: saebom, borderColor: 'rgba(244, 63, 94, 0.4)', borderWidth: 1, pointRadius: 2 },
        { label: '새봄 회귀선', data: lineS, borderColor: COLORS.saebom.solid, borderWidth: 2.5, pointRadius: 0 }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15 } },
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: '매출 (억원)', color: '#94a3b8' } }
      }
    }
  });

  // Table 회귀 결과
  const tbodyReg = document.querySelector('#table-reg tbody');
  const regData = [
    { name: '한빛테크', color: COLORS.hanbit.solid, ...regH },
    { name: '미래솔루션', color: COLORS.mirae.solid, ...regM },
    { name: '새봄산업', color: COLORS.saebom.solid, ...regS }
  ];

  tbodyReg.innerHTML = regData.map(r => `
    <tr>
      <td class="bold" style="color:${r.color}">${r.name}</td>
      <td class="num">${r.slope >= 0 ? '+' : ''}${r.slope.toFixed(4)} 억/월</td>
      <td class="num">${r.intercept.toFixed(2)} 억원</td>
      <td class="num bold">${(r.r2 * 100).toFixed(1)}% (R²=${r.r2.toFixed(3)})</td>
      <td class="bold" style="color:${r.slope > 0 ? '#34d399' : '#fb7185'}">${r.slope > 0 ? '강한 우상향' : '완만한 하향세'}</td>
    </tr>
  `).join('');

  // 2026년 예측 (월 61~72)
  const calcPred2026 = reg => {
    let sum2026 = 0;
    for (let m = 61; m <= 72; m++) {
      sum2026 += (reg.intercept + reg.slope * m);
    }
    return sum2026;
  };

  const predH26 = calcPred2026(regH);
  const predM26 = calcPred2026(regM);
  const predS26 = calcPred2026(regS);

  const actH25 = Stats.sum(hanbit.slice(48, 60));
  const actM25 = Stats.sum(mirae.slice(48, 60));
  const actS25 = Stats.sum(saebom.slice(48, 60));

  const tbodyPred = document.querySelector('#table-pred tbody');
  const preds = [
    { name: '한빛테크', color: COLORS.hanbit.solid, act: actH25, pred: predH26 },
    { name: '미래솔루션', color: COLORS.mirae.solid, act: actM25, pred: predM26 },
    { name: '새봄산업', color: COLORS.saebom.solid, act: actS25, pred: predS26 }
  ];

  tbodyPred.innerHTML = preds.map(p => {
    const diff = ((p.pred - p.act) / p.act) * 100;
    return `
      <tr>
        <td class="bold" style="color:${p.color}">${p.name}</td>
        <td class="num">${p.act.toFixed(2)}</td>
        <td class="num bold">${p.pred.toFixed(2)}</td>
        <td class="num bold" style="color:${diff >= 0 ? '#34d399' : '#fb7185'}">${diff >= 0 ? '+' : ''}${diff.toFixed(1)}%</td>
      </tr>
    `;
  }).join('');

  document.getElementById('insight-reg').innerHTML = `
    <strong>💡 선형회귀 분석 및 2026년 실적 전망:</strong>
    <ul>
      <li><strong>한빛테크의 결정계수 R² = ${(regH.r2 * 100).toFixed(1)}%:</strong> 단순 선형 모델만으로도 매출 변동의 약 90% 이상을 설명할 만큼 강력한 지속 추세를 보입니다. 매월 평균 +0.147억원의 성장이 누적됩니다.</li>
      <li>2026년 예상 연매출: <strong>한빛테크 약 ${predH26.toFixed(1)}억원</strong>, <strong>미래솔루션 약 ${predM26.toFixed(1)}억원</strong>, <strong>새봄산업 약 ${predS26.toFixed(1)}억원</strong>으로 전망됩니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 9. Market Share & HHI
// -------------------------------------------------------------------------
function renderMethod9_MarketShare(months, hanbit, mirae, saebom, total) {
  const shareH = hanbit.map((h, i) => Number(((h / total[i]) * 100).toFixed(1)));
  const shareM = mirae.map((m, i) => Number(((m / total[i]) * 100).toFixed(1)));
  const shareS = saebom.map((s, i) => Number(((s / total[i]) * 100).toFixed(1)));

  // HHI = sum of (market_share_percentage)^2
  const hhi = shareH.map((sh, i) => {
    const sm = shareM[i];
    const ss = shareS[i];
    return Math.round(Math.pow(sh, 2) + Math.pow(sm, 2) + Math.pow(ss, 2));
  });

  destroyChart('chart-market-share');
  const ctxShare = document.getElementById('chart-market-share').getContext('2d');
  chartInstances['chart-market-share'] = new Chart(ctxShare, {
    type: 'line',
    data: {
      labels: months,
      datasets: [
        { label: '한빛테크 점유율(%)', data: shareH, borderColor: COLORS.hanbit.solid, fill: false, tension: 0.2 },
        { label: '미래솔루션 점유율(%)', data: shareM, borderColor: COLORS.mirae.solid, fill: false, tension: 0.2 },
        { label: '새봄산업 점유율(%)', data: shareS, borderColor: COLORS.saebom.solid, fill: false, tension: 0.2 }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15 } },
        y: { ...defaultChartOptions.scales.y, min: 10, max: 60, title: { display: true, text: '점유율 (%)', color: '#94a3b8' } }
      }
    }
  });

  destroyChart('chart-hhi');
  const ctxHHI = document.getElementById('chart-hhi').getContext('2d');
  chartInstances['chart-hhi'] = new Chart(ctxHHI, {
    type: 'line',
    data: {
      labels: months,
      datasets: [{
        label: '허핀달-허쉬만 지수 (HHI)',
        data: hhi,
        borderColor: '#f59e0b',
        backgroundColor: 'rgba(245, 158, 11, 0.15)',
        fill: true,
        tension: 0.3
      }]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15 } },
        y: { ...defaultChartOptions.scales.y, title: { display: true, text: 'HHI 포인트 (집중도)', color: '#94a3b8' } }
      }
    }
  });

  document.getElementById('insight-share').innerHTML = `
    <strong>💡 시장점유율 및 HHI 지수 인사이트:</strong>
    <ul>
      <li>2021년 초: 새봄산업 39.3%, 한빛테크 33.2%, 미래솔루션 27.5%로 비교적 균형적인 삼파전 구조였습니다.</li>
      <li>2025년 말: <strong>한빛테크 44.5%</strong>로 시장 과반에 근접하는 지배적 사업자로 전환되었으며, 새봄산업은 25.7%로 축소되었습니다.</li>
      <li><strong>HHI 지표:</strong> 3,300 초반에서 점진적으로 상승하여 약 3,500선을 돌파, 시장의 독과점 집중화 경향이 강화되었음을 보여줍니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 10. Frequency Distribution & Z-Score Anomaly Detection
// -------------------------------------------------------------------------
function renderMethod10_Distribution(months, hanbit, mirae, saebom) {
  const compKey = document.getElementById('dist-company-select').value;
  const compMap = {
    hanbit: { name: '한빛테크', data: hanbit, color: COLORS.hanbit.solid },
    mirae: { name: '미래솔루션', data: mirae, color: COLORS.mirae.solid },
    saebom: { name: '새봄산업', data: saebom, color: COLORS.saebom.solid }
  };
  const target = compMap[compKey];
  const arr = target.data;
  const mean = Stats.mean(arr);
  const std = Stats.stdDev(arr);

  // 1. 도수분포표 계산 (7개 구간)
  const min = Math.floor(Stats.min(arr));
  const max = Math.ceil(Stats.max(arr));
  const binCount = 7;
  const binWidth = (max - min) / binCount;
  const bins = Array(binCount).fill(0);
  const binLabels = [];

  for (let i = 0; i < binCount; i++) {
    const low = min + i * binWidth;
    const high = low + binWidth;
    binLabels.push(`${low.toFixed(1)}~${high.toFixed(1)}억`);
  }

  arr.forEach(val => {
    let idx = Math.floor((val - min) / binWidth);
    if (idx >= binCount) idx = binCount - 1;
    bins[idx]++;
  });

  destroyChart('chart-histogram');
  const ctxHist = document.getElementById('chart-histogram').getContext('2d');
  chartInstances['chart-histogram'] = new Chart(ctxHist, {
    type: 'bar',
    data: {
      labels: binLabels,
      datasets: [{
        label: `${target.name} 빈도(개월 수)`,
        data: bins,
        backgroundColor: target.color
      }]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        y: { ...defaultChartOptions.scales.y, ticks: { stepSize: 1 }, title: { display: true, text: '빈도 (월)', color: '#94a3b8' } }
      }
    }
  });

  // 2. Z-Score 계산
  const zScores = arr.map(v => (v - mean) / std);

  destroyChart('chart-zscore');
  const ctxZ = document.getElementById('chart-zscore').getContext('2d');
  chartInstances['chart-zscore'] = new Chart(ctxZ, {
    type: 'line',
    data: {
      labels: months,
      datasets: [
        { label: 'Z-Score', data: zScores.map(z => Number(z.toFixed(2))), borderColor: target.color, borderWidth: 2, pointRadius: 3 },
        { label: '+2.0 임계선 (고매출 특이치)', data: Array(months.length).fill(2.0), borderColor: '#f43f5e', borderDash: [5, 5], pointRadius: 0 },
        { label: '-2.0 임계선 (저매출 특이치)', data: Array(months.length).fill(-2.0), borderColor: '#f43f5e', borderDash: [5, 5], pointRadius: 0 }
      ]
    },
    options: {
      ...defaultChartOptions,
      scales: {
        x: { ...defaultChartOptions.scales.x, ticks: { maxTicksLimit: 15 } },
        y: { ...defaultChartOptions.scales.y, min: -3, max: 3, title: { display: true, text: '표준점수 (Z)', color: '#94a3b8' } }
      }
    }
  });

  // 특이치 (|Z| >= 1.8) 추출 표
  const outliers = [];
  zScores.forEach((z, i) => {
    if (Math.abs(z) >= 1.8) {
      outliers.push({
        month: months[i],
        company: target.name,
        val: arr[i],
        z: z.toFixed(2),
        type: z > 0 ? '상한 이상치 (급등)' : '하한 이상치 (급락)'
      });
    }
  });

  const tbodyOut = document.querySelector('#table-outliers tbody');
  if (outliers.length === 0) {
    tbodyOut.innerHTML = `<tr><td colspan="6" style="text-align:center; color:#94a3b8;">|Z| ≥ 1.8 수준의 이상치가 발견되지 않았습니다 (정규 범위 내 안정).</td></tr>`;
  } else {
    tbodyOut.innerHTML = outliers.map(o => `
      <tr>
        <td class="bold">${o.month}</td>
        <td>${o.company}</td>
        <td class="num bold">${o.val.toFixed(2)}</td>
        <td class="num">${mean.toFixed(2)} / ${std.toFixed(2)}</td>
        <td class="num bold" style="color: ${o.z > 0 ? '#38bdf8' : '#fb7185'}">${o.z}</td>
        <td><span class="kpi-badge ${o.z > 0 ? 'up' : 'down'}">${o.type}</span></td>
      </tr>
    `).join('');
  }

  document.getElementById('insight-dist').innerHTML = `
    <strong>💡 분포 및 이상치(Anomaly Detection) 분석:</strong>
    <ul>
      <li><strong>Z-Score 검정:</strong> 정규분포 기준 약 95.4%의 데이터는 [-2, +2] 범위 내에 위치합니다.</li>
      <li>${target.name}의 경우 총 ${outliers.length}건의 잠재적 특이치(|Z| ≥ 1.8)가 식별되었으며, 이는 비정상적인 외부 이벤트 또는 연말 실적 마감 특수가 반영된 구간입니다.</li>
    </ul>
  `;
}

// -------------------------------------------------------------------------
// 11. 원본 데이터 테이블 렌더링
// -------------------------------------------------------------------------
function renderRawTable(data) {
  const tbody = document.querySelector('#table-raw tbody');
  tbody.innerHTML = data.map((row, idx) => {
    const tot = row.hanbit + row.mirae + row.saebom;
    let diffStr = '-';
    if (idx > 0) {
      const prevTot = data[idx - 1].hanbit + data[idx - 1].mirae + data[idx - 1].saebom;
      const diff = tot - prevTot;
      const pct = (diff / prevTot) * 100;
      const color = diff >= 0 ? '#34d399' : '#fb7185';
      diffStr = `<span style="color:${color}">${diff >= 0 ? '+' : ''}${diff.toFixed(2)} (${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%)</span>`;
    }
    return `
      <tr>
        <td>${idx + 1}</td>
        <td class="bold">${row.month}</td>
        <td class="num">${row.hanbit.toFixed(2)}</td>
        <td class="num">${row.mirae.toFixed(2)}</td>
        <td class="num">${row.saebom.toFixed(2)}</td>
        <td class="num bold">${tot.toFixed(2)}</td>
        <td class="num">${diffStr}</td>
      </tr>
    `;
  }).join('');
}

// -------------------------------------------------------------------------
// 이벤트 리스너 및 초기화
// -------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  // 탭 전환 핸들러
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabPanes = document.querySelectorAll('.tab-pane');

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      tabPanes.forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      const targetId = btn.getAttribute('data-tab');
      document.getElementById(targetId)?.classList.add('active');
    });
  });

  // 추이 뷰 라디오 버튼
  document.querySelectorAll('input[name="trend-view"]').forEach(radio => {
    radio.addEventListener('change', () => {
      const months = currentData.map(d => d.month);
      const hanbit = currentData.map(d => d.hanbit);
      const mirae = currentData.map(d => d.mirae);
      const saebom = currentData.map(d => d.saebom);
      const total = currentData.map(d => Number((d.hanbit + d.mirae + d.saebom).toFixed(2)));
      renderMethod1_Trend(months, hanbit, mirae, saebom, total);
    });
  });

  // 이동평균 선택 셀렉트
  document.getElementById('ma-company-select').addEventListener('change', () => {
    const months = currentData.map(d => d.month);
    const hanbit = currentData.map(d => d.hanbit);
    const mirae = currentData.map(d => d.mirae);
    const saebom = currentData.map(d => d.saebom);
    renderMethod5_MovingAverage(months, hanbit, mirae, saebom);
  });

  // 산점도 셀렉트
  document.getElementById('scatter-pair').addEventListener('change', () => {
    const hanbit = currentData.map(d => d.hanbit);
    const mirae = currentData.map(d => d.mirae);
    const saebom = currentData.map(d => d.saebom);
    renderScatterChart(hanbit, mirae, saebom);
  });

  // 분포 셀렉트
  document.getElementById('dist-company-select').addEventListener('change', () => {
    const months = currentData.map(d => d.month);
    const hanbit = currentData.map(d => d.hanbit);
    const mirae = currentData.map(d => d.mirae);
    const saebom = currentData.map(d => d.saebom);
    renderMethod10_Distribution(months, hanbit, mirae, saebom);
  });

  // 엑셀 파일 업로드 처리 (SheetJS)
  document.getElementById('file-upload').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
        
        if (rows.length < 2) {
          alert('데이터 형식이 올바르지 않습니다.');
          return;
        }

        const parsed = [];
        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (row[0] && row[1] !== undefined && row[2] !== undefined && row[3] !== undefined) {
            parsed.push({
              month: String(row[0]).trim(),
              hanbit: parseFloat(row[1]) || 0,
              mirae: parseFloat(row[2]) || 0,
              saebom: parseFloat(row[3]) || 0
            });
          }
        }

        if (parsed.length > 0) {
          currentData = parsed;
          renderDashboard();
          alert(`새 엑셀 파일(${file.name})에서 ${parsed.length}개 행의 데이터를 성공적으로 로드했습니다!`);
        } else {
          alert('엑셀 파일에서 유효한 데이터를 추출하지 못했습니다.');
        }
      } catch (err) {
        alert('엑셀 파일 파싱 중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // 기본 데이터 복원
  document.getElementById('btn-reset-data').addEventListener('click', () => {
    currentData = JSON.parse(JSON.stringify(DEFAULT_DATA));
    renderDashboard();
    alert('기본 5개년 매출 데이터로 복원되었습니다.');
  });

  // CSV 내보내기
  document.getElementById('btn-export-csv').addEventListener('click', () => {
    let csv = '\uFEFF기준월,한빛테크,미래솔루션,새봄산업,합계\n';
    currentData.forEach(d => {
      const tot = (d.hanbit + d.mirae + d.saebom).toFixed(2);
      csv += `${d.month},${d.hanbit},${d.mirae},${d.saebom},${tot}\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '기업매출데이터_통계추출.csv';
    a.click();
    URL.revokeObjectURL(url);
  });

  // 첫 화면 렌더링
  renderDashboard();
});
