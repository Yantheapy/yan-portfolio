/* ==========================================
   ARISHA V3.2 — Interactive Financial Model
   ========================================== */
const AED_USD = 0.2723;
let currencyMode = 'AED'; // 'AED' or 'USD'
function cx(val) { return currencyMode === 'USD' ? val * AED_USD : val; }
function cxFmt(val) { return fmt(cx(val)); }
function cxFmtM(val) { return fmtM(cx(val)); }
function cxLabel() { return currencyMode; }
function cxSymbol() { return currencyMode === 'USD' ? '$' : ''; }

const STORAGE_PREFIX='arishaDashboard';
const DEFAULT_PROJECT_ID='arisha';
const PROJECTS={
  arisha:{
    id:'arisha',
    label:'Arisha',
    title:'Arisha Medical Clinic',
    shortTitle:'Arisha',
    subtitle:{ru:'Финансовая модель · Март 2026',en:'Financial model · March 2026'},
    workbookPath:'Arisha_Enhanced_V3.xlsx',
    workbookLabel:'Arisha_Enhanced_V3.xlsx',
    benchmarkMarketLabel:{ru:'рынок Dubai',en:'Dubai market'},
    pdfSlug:'Arisha',
    assessmentLabel:{ru:'Март 2026',en:'March 2026'}
  }
};
const DASHBOARD_RUNTIME_CONFIG=
  typeof window!=='undefined'&&window.ARISHA_DASHBOARD_CONFIG&&typeof window.ARISHA_DASHBOARD_CONFIG==='object'
    ? window.ARISHA_DASHBOARD_CONFIG
    : {};
function resolveEnabledProjectIds(projectIds){
  if(!Array.isArray(projectIds))return [DEFAULT_PROJECT_ID];
  const resolved=[...new Set(projectIds.map(id=>String(id)).filter(id=>PROJECTS[id]))];
  return resolved.length?resolved:[DEFAULT_PROJECT_ID];
}
const ENABLED_PROJECT_IDS=resolveEnabledProjectIds(DASHBOARD_RUNTIME_CONFIG.allowedProjectIds);
const ACTIVE_DEFAULT_PROJECT_ID=(()=>{
  const requestedId=String(DASHBOARD_RUNTIME_CONFIG.defaultProjectId||'');
  if(ENABLED_PROJECT_IDS.includes(requestedId))return requestedId;
  return ENABLED_PROJECT_IDS.includes(DEFAULT_PROJECT_ID)?DEFAULT_PROJECT_ID:ENABLED_PROJECT_IDS[0];
})();
const I18N={
  ru:{
    navOverview:'Обзор',
    navRevenue:'Доходы',
    navPatients:'Пациенты',
    navSensitivity:'Чувствительность',
    navRisks:'Риски',
    filterPeriod:'📅 Период:',
    panelModel:'Модель',
    panelRealtime:'Параметры в реальном времени',
    groupScenario:'Сценарий',
    groupDrivers:'Параметры',
    scenarioBase:'Base',
    scenarioOptimistic:'Upside',
    scenarioPessimistic:'Downside',
    resetToBase:'↻ Сбросить к базовому',
    paybackBenchmark:'Benchmark: 3-5 лет',
    verdictHtml:'<strong>Инвестиционно привлекательный проект.</strong> 5Y IRR с учётом exit 49.4% существенно превышает WACC 12%, payback ~4.0 лет, 10Y NPV ≈ 15M AED.',
    overviewCapex:'Структура инвестиций (CapEx)',
    overviewPatients:'Рост пациентской базы',
    overviewSources:'Источники пациентов',
    overviewUtil:'Утилизация и ёмкость',
    benchmarkTitle:'Бенчмарк: {project} vs {market}',
    scenarioCompare:'📐 Сценарии: 620 sqm vs 450 sqm',
    scenarioMetric:'Метрика',
    scenarioDelta:'Δ Разница',
    scenarioPaybackBase:'4.0 лет',
    scenarioPaybackCompact:'—',
    scenarioPaybackDelta:'требуется пересчёт',
    revenueSection:'Доходность и прибыльность (10 лет)',
    chartAbsolute:'Абсолют',
    chartMargin:'Маржа %',
    monthlyPL:'Месячная P&L динамика (Y1)',
    mixEvolution:'Эволюция Revenue Mix по департаментам (10 лет)',
    plDetail:'Детализация P&L по годам',
    plMetric:'Показатель',
    showDetails:'Показать детали',
    hideDetails:'Скрыть детали',
    avgCheckDept:'Средний чек по департаменту {label}',
    irrByExit:'IRR по Exit Multiple',
    exitScenarios:'Сценарии Exit',
    riskConcentrationText:'Revenue-sharing @ 40% — крупнейший cost.',
    riskConcentrationMitigation:'Стандартный рыночный rate, low switching cost',
    riskUtilizationMitigation:'Рост после Y7 только через price increases',
    riskDubaiMitigation:'Модель учитывает seasonality',
    ratioTrends:'Ratio Trends (10 лет)',
    selectedPeriodToast:'📅 Период: {label}',
    costStructure:'Структура расходов {label}',
    sensitivityBreakevenDesc:'Клиника прибыльна даже при минимальной загрузке.',
    yearsUnit:'лет',
    monthsUnit:'мес',
    marketLabel:'Рынок',
    actualVsPlanTitle:'Факт vs План',
    patientEconomics:'Экономика привлечения (10 лет)',
    cashflowMain:'Накопленный и годовой Cash Flow',
    fundingBridge:'Структура финансирования',
    runwaySnapshot:'Сводка по runway',
    exitValuation:'Оценка при выходе (8× EBITDA)',
    investorModelAssumptions:'Параметры investor model',
    capexBreakdown:'Структура CAPEX',
    netBookValue:'Net Book Value',
    depreciationSchedule:'График амортизации (stacked)',
    amortizationTable:'Таблица амортизации',
    tornadoSectionTitle:'Tornado: влияние драйверов на 10Y NPV',
    scenarioLibrarySectionTitle:'Библиотека сценариев и сравнение',
    mcSimulationSectionTitle:'🎲 Monte Carlo Simulation',
    integrityChecksSectionTitle:'Проверки целостности',
    ratioRadarTitle:'Финансовые коэффициенты: Y1 vs Y5 vs Y10',
    assumptionsDefinitions:'Допущения и определения',
    benchmarkSources:'Источники бенчмарков',
    templateComplianceSectionTitle:'Проверка соответствия шаблону',
    importAuditSectionTitle:'Журнал импорта',
    excelChecklistSectionTitle:'Сверка с Excel',
    monthlyDrilldownTitle:'📅 Monthly Drill-down (Year 1)',
    monthlyDrilldownNote:'Помесячные данные показаны для Y1. Для Y2/Y3 используется линейная оценка.',
    equityCompareTableTitle:'Equity vs Revenue Share',
    mcLegendHtml:'NPV Distribution (слева) · 5Y IRR Distribution (справа) · <span style="color:#f43f5e">Красный</span> = ниже порога',
    scenarioNamePlaceholder:'Название сценария',
    saveCurrent:'Сохранить текущий',
    clearAll:'Очистить',
    timelineAnnual:'Annual',
    timelineCumulative:'Cumulative',
    scenarioNative:'Native',
    scenarioExitAligned:'Exit-aligned',
    modelMode:'Model Mode',
    modeCompare:'Compare',
    modeEquity:'Equity',
    modeRevenueShare:'Revenue Share',
    noData:'Нет данных'
  },
  en:{
    navOverview:'Overview',
    navRevenue:'Revenue',
    navPatients:'Patients',
    navSensitivity:'Sensitivity',
    navRisks:'Risks',
    filterPeriod:'📅 Period:',
    panelModel:'Model',
    panelRealtime:'Real-time parameters',
    groupScenario:'Scenario',
    groupDrivers:'Drivers',
    scenarioBase:'Base',
    scenarioOptimistic:'Upside',
    scenarioPessimistic:'Downside',
    resetToBase:'↻ Reset to base',
    paybackBenchmark:'Benchmark: 3-5 yrs',
    verdictHtml:'<strong>Attractive investment case.</strong> 5Y IRR with exit of 49.4% materially exceeds the 12% WACC, payback ~4.0 years, 10Y NPV ≈ AED 15M.',
    overviewCapex:'Investment Structure (CapEx)',
    overviewPatients:'Patient Base Growth',
    overviewSources:'Patient Sources',
    overviewUtil:'Utilization & Capacity',
    benchmarkTitle:'Benchmark: {project} vs {market}',
    scenarioCompare:'📐 Scenarios: 620 sqm vs 450 sqm',
    scenarioMetric:'Metric',
    scenarioDelta:'Δ Difference',
    scenarioPaybackBase:'4.0 yrs',
    scenarioPaybackCompact:'—',
    scenarioPaybackDelta:'recalc required',
    revenueSection:'Revenue & Profitability (10Y)',
    chartAbsolute:'Absolute',
    chartMargin:'Margin %',
    monthlyPL:'Monthly P&L Trend (Y1)',
    mixEvolution:'Revenue Mix Evolution by Department (10Y)',
    plDetail:'P&L Detail by Year',
    plMetric:'Metric',
    showDetails:'Show details',
    hideDetails:'Hide details',
    avgCheckDept:'Average Ticket by Department {label}',
    irrByExit:'IRR by Exit Multiple',
    exitScenarios:'Exit Scenarios',
    riskConcentrationText:'Revenue-sharing at 40% remains the largest cost bucket.',
    riskConcentrationMitigation:'In line with market rates and still operationally flexible.',
    riskUtilizationMitigation:'Growth beyond Y7 depends mainly on pricing and added capacity.',
    riskDubaiMitigation:'Seasonality is already reflected in the model.',
    ratioTrends:'Ratio Trends (10Y)',
    selectedPeriodToast:'📅 Period: {label}',
    costStructure:'Cost Structure {label}',
    sensitivityBreakevenDesc:'The clinic remains profitable even at a very low utilization threshold.',
    yearsUnit:'yrs',
    monthsUnit:'mo',
    marketLabel:'Market',
    actualVsPlanTitle:'Actual vs Plan',
    patientEconomics:'Acquisition Economics (10Y)',
    cashflowMain:'Cumulative & Annual Cash Flow',
    fundingBridge:'Funding Bridge',
    runwaySnapshot:'Runway Snapshot',
    exitValuation:'Exit Valuation (8× EBITDA)',
    investorModelAssumptions:'Investor Model Assumptions',
    capexBreakdown:'CAPEX Breakdown',
    netBookValue:'Net Book Value',
    depreciationSchedule:'Depreciation Schedule (Stacked)',
    amortizationTable:'Amortization Table',
    tornadoSectionTitle:'Tornado: Driver Impact on 10Y NPV',
    scenarioLibrarySectionTitle:'Scenario Library & Compare',
    mcSimulationSectionTitle:'🎲 Monte Carlo Simulation',
    integrityChecksSectionTitle:'Integrity Checks',
    ratioRadarTitle:'Financial Ratios: Y1 vs Y5 vs Y10',
    assumptionsDefinitions:'Assumptions & Definitions',
    benchmarkSources:'Benchmark Sources',
    templateComplianceSectionTitle:'Template Compliance Check',
    importAuditSectionTitle:'Import Audit Trail',
    excelChecklistSectionTitle:'Excel Reconciliation Checklist',
    monthlyDrilldownTitle:'📅 Monthly Drill-down (Year 1)',
    monthlyDrilldownNote:'Monthly data shown for Year 1. Y2/Y3 estimated via linear distribution.',
    equityCompareTableTitle:'Equity vs Revenue Share',
    mcLegendHtml:'NPV Distribution (left) · 5Y IRR Distribution (right) · <span style="color:#f43f5e">Red</span> = below threshold',
    scenarioNamePlaceholder:'Scenario name',
    saveCurrent:'Save current',
    clearAll:'Clear all',
    timelineAnnual:'Annual',
    timelineCumulative:'Cumulative',
    scenarioNative:'Native',
    scenarioExitAligned:'Exit-aligned',
    modelMode:'Model Mode',
    modeCompare:'Compare',
    modeEquity:'Equity',
    modeRevenueShare:'Revenue Share',
    noData:'No data'
  }
};

const DEFAULT_PRODUCT_CATALOG = [
  {department:'Physio Clinic',product:'Specialist Consultation',basePrice:700,monthlyQty:[93,147,202,225,248,280,310,310,310,310]},
  {department:'Physio Clinic',product:'Personal Rehabilitation Session',basePrice:600,monthlyQty:[447,708,969,1080,1192,1342,1490,1490,1490,1490]},
  {department:'Physio Clinic',product:'Follow Up Consultation',basePrice:350,monthlyQty:[51,80,110,123,135,152,169,169,169,169]},
  {department:'Physio Clinic',product:'Injections',basePrice:700,monthlyQty:[9,15,20,23,25,28,31,31,31,31]},
  {department:'Physio Clinic',product:'Orthotics',basePrice:1000,monthlyQty:[5,7,10,11,12,14,16,16,16,16]},
  {department:'Dental Clinic',product:'Hygiene',basePrice:350,monthlyQty:[39,61,84,94,103,116,125,125,125,125]},
  {department:'Dental Clinic',product:'Restorative',basePrice:500,monthlyQty:[52,82,112,125,138,155,167,167,167,167]},
  {department:'Dental Clinic',product:'Implantology',basePrice:3500,monthlyQty:[7,10,14,16,17,20,21,21,21,21]},
  {department:'Dental Clinic',product:'Aesthetic Dentistry',basePrice:2200,monthlyQty:[13,21,28,32,35,39,42,42,42,42]},
  {department:'Dental Clinic',product:'Follow Up Consultation',basePrice:250,monthlyQty:[20,31,42,47,52,59,63,63,63,63]},
  {department:'Polyclinic',product:'GP',basePrice:600,monthlyQty:[43,68,94,104,115,130,144,144,144,144]},
  {department:'Polyclinic',product:'Gynecology',basePrice:600,monthlyQty:[43,68,94,104,115,130,144,144,144,144]},
  {department:'Polyclinic',product:'Blood Tests',basePrice:300,monthlyQty:[35,55,75,84,92,104,115,115,115,115]},
  {department:'Polyclinic',product:'X-Ray',basePrice:500,monthlyQty:[3,4,6,7,7,8,9,9,9,9]},
  {department:'Polyclinic',product:'Ultrasound',basePrice:400,monthlyQty:[11,17,23,26,29,32,36,36,36,36]},
  {department:'Cosmetology',product:'Session',basePrice:450,monthlyQty:[242,383,525,585,645,696,696,696,696,696]},
  {department:'Wellness Packages',product:'Essential Package',basePrice:1200,monthlyQty:Array(10).fill(0)},
  {department:'Wellness Packages',product:'Premium Package',basePrice:2200,monthlyQty:Array(10).fill(0)},
  {department:'Wellness Packages',product:'Elite Package',basePrice:3800,monthlyQty:Array(10).fill(0)}
];

const DEFAULT_PA_MODEL = {
  uniquePatients:[2071.7983739837396,3284.1170731707307,4496.435772357722,5012.123577235771,5527.8113821138195,6224.442276422763,6912.026016260161,6912.02601626016,6912.026016260161,6912.026016260161],
  avgRevenuePerVisit:[600.0207195979142,620.2209005947325,645.6256793095021,672.7822774047711,698.3010907933984,722.3005761321455,734.438772403903,763.8163233000595,794.3689762320618,826.143735281344],
  avgRevenuePerPatient:[3690.1274255271724,3814.358538657605,3970.597927753438,4137.611006039342,4294.5517083794,4442.148543212695,4516.798450284004,4697.470388295366,4885.36920382718,5080.783971980266],
  marketingBudget:[1529040,1989935.3988682916,2585251.629229267,2297361.9106113818,2437952.6238356084,2774128.7226190898,3001678.262266111,2537580.107888715,2363582.1620199718,2186689.9623324694],
  pureCAC:[1200,1200,1200,1200,1200,1200,1200,1200,1200,1200],
  blendedCAC:[738.0254851054345,837.2790136992364,857.3782079087855,785.2141886151029,777.1534985258648,789.0429274502404,785.984268109041,718.9374263650417,686.1408022538106,650.8257279799586],
  ltv:[9225.318563817931,10037.78562804633,11029.438688203994,12169.444135409827,13420.474088685623,14807.161810708982,16131.423036728584,18067.19380113602,20355.705015946587,23094.472599910303],
  ltvCac:[7.687765469848276,8.364821356705274,9.191198906836663,10.141203446174856,11.18372840723802,12.339301508924152,13.442852530607153,15.055994834280018,16.963087513288823,19.245393833258586],
  cacPayback:[3.9023042674313104,3.7752088205813554,3.6266578137634577,3.4802691647381696,3.3530857183307754,3.241674577047234,3.1880988621696345,3.0654796751631084,2.9475766107337584,2.8342082795516914],
  romi:[3.021042949215009,2.0315247207485823,2.2153814105387584,1.9506066482845976,1.9907592486421446,2.14118799739428,2.169571614401986,1.8235162833138272,1.6603424218354483,1.43637093060702],
  lifetimeYears:[2.5,2.6315789473684212,2.7777777777777777,2.9411764705882346,3.1249999999999996,3.333333333333333,3.571428571428571,3.846153846153846,4.166666666666667,4.545454545454546],
  retainedVisits:[],
  organicVisits:[],
  paidVisits:[],
  acquisitionSpend:[1502158.0487804874,1271805.861307316,1730918.6692292674,1299662.8469365851,1356149.017607316,1604610.3095335353,1730169.5504178777,1182079.2718936135,963944.9304312917,741829.9018732737],
  retentionSpend:[0,181489.5375609756,296227.3599999999,417269.2396747966,478156.58926829253,541725.5154471544,618604.0983042972,676480.0379094076,693455.6015795586,710431.1652497096]
};

const DEFAULT_CAPEX_BREAKDOWN = {
  labels:[
    'Rent Deposit',
    'Renovation & Fit-out',
    'Physio Equipment',
    'Assessment Equipment',
    'Dental Equipment',
    'Laboratory & Medical Equipment',
    'Cosmetology & Beauty',
    'Office & IT',
    'Design / MEP / Engineering',
    'Licensing / Legal / Hiring / Launch',
    'Initial Inventory',
    'Pre-opening Salaries',
    'Working Capital',
    'Contingency'
  ],
  amounts:[204600,2604000,610000,270000,490000,60000,350000,212000,120000,770000,100000,132000,336600,312960],
  life:[0,7,7,7,7,7,7,7,7,0,0,0,0,0],
  colors:['#ec4899','#6366f1','#22d3ee','#3b82f6','#10b981','#14b8a6','#f59e0b','#8b5cf6','#a855f7','#64748b','#f97316','#fb7185','#f43f5e','#94a3b8']
};

/* ===== BASE DATA (immutable) ===== */
const BASE = {
  investment: 6572160,
  breakevenUtilPct:16.9,
  operationalBreakevenMonth:9,
  years: ['Y1','Y2','Y3','Y4','Y5','Y6','Y7','Y8','Y9','Y10'],
  revenue: [7645200,12526800,17853538.560000002,20738217.676800005,23739471.814656,27649897.190522887,31220228.39856661,32469037.534509283,33767799.035889655,35118510.997325234],
  ebitda:  [-321500.00000000047,1177869.241131708,2539293.4587707347,3352836.7105726223,4389966.636795429,5630139.664677785,6877942.756702806,7737225.751838965,8322215.932096815,8926540.05554898],
  netProfit:[-995214.2857142861,492531.00942985434,1731427.0474813688,2471751.406621086,3415539.63948384,4544097.094856785,5679597.908599554,7074625.434173458,7606966.498208102,8156901.450549573],
  cogs:[1106796,1808283.36,2576066.1888,2998426.2297600005,3426251.3342054402,3977360.003235841,4426586.993030333,4603650.472751548,4787796.49166161,4979308.351328074],
  margins:[-0.042052529691832846,0.09402794337993008,0.1422291413120702,0.16167429442711775,0.1849226752418814,0.20362244480994085,0.22030404995431063,0.2382955067151454,0.24645420103488708,0.25418332959016576],
  npRatio:[],cfRatio:[],
  annCF:[-6572160,-321500.00000000047,1166245.29514414,2405141.3331956547,3145465.692335372,4089253.9251981257,5217811.38057107,6353312.194313839,7074625.434173458,7606966.498208102,8156901.450549573],
  months:['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
  mRev:[111149.44615384615,234256.76923076934,298750.8923076923,396374.21538461535,459202.076923077,429307.3846153846,539280.6461538462,616320.7384615386,780692.5384615385,1083070,1269005.1846153846,1457292.7384615387],
  mEbitda:[-204321.43435897437,-164385.18205128206,-140451.65743589745,-107803.80820512827,-86820.43974358978,-83864.79897435897,-43803.16666666672,-15722.66666666657,20639.1417948717,110948.9871794871,171241.30948717936,232165.3917948717],
  retained:[0,907.447687804878,1481.1367999999998,2086.346198373983,2390.782946341463,2708.627577235772,3093.020491521486,3382.4001895470383,3467.278007897793,3552.1558262485482],
  referral:[0,326.6811676097561,533.2092479999999,751.0846314146338,860.6818606829266,975.1059278048779,1113.487376947735,1217.6640682369339,1248.2200828432053,1278.7760974494772],
  walkin:[820,990.15,1039.6575,1091.6403750000002,1146.2223937500003,1203.5335134375002,1263.7101891093753,1326.895698564844,1393.2404834930862,1462.9025076677406],
  marketing:[1251.7983739837396,1059.8382177560966,1442.4322243577228,1083.0523724471543,1130.12418133943,1337.1752579446127,1441.8079586815647,985.0660599113446,803.2874420260764,618.1915848943947],
  capacity:[28953.600000000002,30310.8,31668,38001.600000000006,41620.799999999996,44335.2,49311.600000000006,50668.8,50668.8,50668.8],
  util:[30,67,87,81,82,86,86,84,84,84],
  addFTE:[0,0,0,1,2,3,4,4,4,4],
  visits:[12741.56,20197.319999999996,27653.079999999994,30824.559999999994,33996.03999999999,38280.31999999999,42508.95999999999,42508.959999999985,42508.95999999999,42508.95999999999],
  patientsShort:[2072,3284,4496,5528,6912,6912],
  visitsShort:[12742,20197,27653,33996,42509,42509],
  patientYears:['Y1','Y2','Y3','Y5','Y7','Y10'],
  plYears:['Y1','Y2','Y3','Y5','Y7','Y10'],
  plIdx:[0,1,2,4,6,9],
  plBreakdown:{marketing:[],admin:[],rent:[],staff:[],rd:[]},
  costLabels:['Revenue Sharing','Marketing','Administrative','Rent','Fixed Staff','COGS','R&D','EBITDA'],
  benchmarks:[
    {label:'EBITDA Margin Y5',arisha:'18.5%',aw:46,market:'20-30%',mw:62},
    {label:'Revenue/sqm Y5',arisha:'38.3K AED',aw:64,market:'30-50K',mw:66},
    {label:'Payback Period',arisha:'4.0 лет',aw:65,market:'3-5 лет',mw:80},
    {label:'LTV : CAC Y5',arisha:'11.18×',aw:100,market:'3-5×',mw:40},
    {label:'CAC Payback Y5',arisha:'3.4 мес',aw:83,market:'6-12 мес',mw:60},
    {label:'Breakeven Util',arisha:'16.9%',aw:34,market:'20-35%',mw:55}
  ],
  checks:[
    {label:'Revenue ≡ PA × avg price',pass:true,delta:'Δ = 0'},
    {label:'OpEx total ≡ P&L costs',pass:true,delta:'Δ = 0'},
    {label:'Monthly ≈ Annual Rev',pass:false,delta:'Δ = 29,503 (0.4%)'},
    {label:'PA dept sum = total',pass:true,delta:'Δ = 0'},
    {label:'Staff util ≤ 100%',pass:true,delta:'86.5% (Physio, Y4)'},
    {label:'Room util ≤ 100%',pass:true,delta:'100.0% (Dental, Y7+)'},
    {label:'Seasonality avg = 1.0',pass:true,delta:'1.0001'},
    {label:'LTV:CAC ≥ 3×',pass:true,delta:'11.18×'}
  ],
  heatRows:['Rev -20%','Rev -10%','BASE','Rev +10%','Rev +20%'],
  heatOpexCols:['OPEX -20%','OPEX -10%','BASE','OPEX +10%','OPEX +20%'],
  heatFactors:[-0.2,-0.1,0,0.1,0.2],
  stressLabels:['Severe\n-30%','Downside\n-20%','Mild -10%','BASE','Mild +10%','Upside +20%','Strong +30%'],
  stressFactors:[-0.3,-0.2,-0.1,0,0.1,0.2,0.3],
  opexStressLabels:['Lean\n-20%','Lean\n-10%','Mild -5%','BASE','Mild +5%','Pressure +10%','Severe\n+20%'],
  opexStressFactors:[-0.2,-0.1,-0.05,0,0.05,0.1,0.2],
  capexLabels:DEFAULT_CAPEX_BREAKDOWN.labels.slice(),
  capexAmounts:DEFAULT_CAPEX_BREAKDOWN.amounts.slice(),
  capexLife:DEFAULT_CAPEX_BREAKDOWN.life.slice(),
  capexColors:DEFAULT_CAPEX_BREAKDOWN.colors.slice(),
  services:['Physio Clinic','Dental Clinic','Polyclinic','Cosmetology','Wellness Packages'],
  serviceColors:['#6366f1','#22d3ee','#10b981','#f59e0b','#a855f7'],
  serviceRevenue:[
    [4349400,7153536,10192565.760000002,11819845.939200003,13554683.191296,15884620.293734403,18339786.91788473,19073378.39460012,19836313.530384123,20629766.071599487],
    [1173000,1888224,2680204.8000000003,3152543.8464,3573449.957376001,4214242.323333121,4690284.537760974,4877895.919271413,5073011.75604227,5275932.226283961],
    [816000,1334112,1914432.0000000005,2212382.5152000003,2536721.3015040006,2978366.305075201,3434581.9438055432,3571965.2215577653,3714843.830420076,3863437.5836368795],
    [1306800,2150928,3066336.0000000005,3553445.376,4074617.364480001,4572668.268380161,4755574.999115367,4945797.999079983,5143629.919043182,5349375.11580491],
    [0,0,0,0,0,0,0,0,0,0]
  ],
  serviceVisits:[
    [7251.294308943089,11494.409756097559,15737.52520325203,17542.4325203252,19347.33983739837,21785.54796747967,24192.091056910565,24192.091056910558,24192.091056910565,24192.091056910565],
    [1553.8487804878048,2463.0878048780482,3372.326829268292,3759.0926829268283,4145.858536585365,4668.331707317072,5007.6,5007.6,5007.6,5007.6],
    [1035.8991869918698,1642.0585365853656,2248.2178861788612,2506.0617886178856,2763.90569105691,3112.2211382113814,3456.0130081300805,3456.01300813008,3456.0130081300805,3456.0130081300805],
    [2900.5177235772358,4597.763902439024,6295.010081300812,7016.97300813008,7738.935934959349,8346,8346,8346,8346,8346],
    [0,0,0,0,0,0,0,0,0,0]
  ],
  serviceMix:[
    [56.8905980222885,15.342960288808663,10.673363679171245,17.093078009731597,0],
    [57.10585305105853,15.073474470734743,10.650062266500622,17.170610211706103,0],
    [57.08989131620079,15.012176947402855,10.722983533859184,17.17494820253717,0],
    [56.99547629120968,15.201614215510785,10.668142024929213,17.134767468350308,0],
    [57.097661216404,15.052777859909527,10.685668667395999,17.163892256290476,0],
    [57.44911159806743,15.241439395939489,10.771708424637645,16.53774058135544,0],
    [58.74328234807773,15.023223013885173,11.00114291272524,15.232351725311868,0],
    [58.74328234807772,15.023223013885173,11.001142912725238,15.232351725311868,0],
    [58.74328234807771,15.023223013885173,11.001142912725237,15.232351725311868,0],
    [58.74328234807772,15.023223013885176,11.001142912725241,15.23235172531187,0]
  ],
  revPerVisit:[700.6,862.0,917.8,526.5,0],
  productCatalog:DEFAULT_PRODUCT_CATALOG.map(item=>({...item,monthlyQty:item.monthlyQty.slice()})),
  paModel:Object.fromEntries(Object.entries(DEFAULT_PA_MODEL).map(([key,series])=>[key,series.slice()])),
  scenarios:{
    '620 sqm':{capex:6572160,area:620,label:'Base'},
    '450 sqm':{capex:5555550,area:450,label:'Compact'}
  }
};

const PL_BREAKDOWN_KEYS=['marketing','admin','rent','staff','rd'];
const PA_MODEL_KEYS=['uniquePatients','avgRevenuePerVisit','avgRevenuePerPatient','marketingBudget','pureCAC','blendedCAC','ltv','ltvCac','cacPayback','romi','lifetimeYears','retainedVisits','organicVisits','paidVisits','acquisitionSpend','retentionSpend'];
const SERVICE_REVENUE_ROWS=['Physio Clinic','Dental Clinic','Polyclinic','Cosmetology','Wellness Packages'];
const TEMPLATE_SHEET_RULES=[
  {sheet:'P&L',severity:'fail',note:'Core operating model'},
  {sheet:'Cash Flow & DCF',severity:'fail',note:'DCF, investment, financing assumptions'},
  {sheet:'Patient Acquisition',severity:'fail',note:'Patient engine and CAC layer'},
  {sheet:'Revenue',severity:'fail',note:'Department revenue mix'},
  {sheet:'Capacity',severity:'fail',note:'Visits and utilization capacity'},
  {sheet:'Monthly P&L',severity:'warn',note:'Y1 monthly phasing'},
  {sheet:'Price List',severity:'optional',note:'Product mix drill-down'},
  {sheet:'CapEx',severity:'optional',note:'Capex breakdown and depreciation'},
  {sheet:'Summary',severity:'optional',note:'Cross-check summary layer'}
];
const TEMPLATE_ROW_RULES=[
  {sheet:'P&L',item:'Revenue',severity:'fail',match:{type:'starts',label:'Revenue'},note:'Required for core revenue series'},
  {sheet:'P&L',item:'COGS',severity:'fail',match:{type:'starts',label:'COGS'},note:'Required for cost waterfall'},
  {sheet:'P&L',item:'EBITDA',severity:'fail',match:{type:'starts',label:'EBITDA'},note:'Required for margin and return model'},
  {sheet:'P&L',item:'NET PROFIT',severity:'fail',match:{type:'starts',label:'NET PROFIT'},note:'Required for net income and dividends'},
  {sheet:'P&L',item:'Marketing',severity:'fail',match:{type:'starts',label:'Marketing'},note:'Required for OPEX split'},
  {sheet:'P&L',item:'Administrative',severity:'fail',match:{type:'starts',label:'Administrative'},note:'Required for OPEX split'},
  {sheet:'P&L',item:'Rent',severity:'fail',match:{type:'starts',label:'Rent'},note:'Required for OPEX split'},
  {sheet:'P&L',item:'R&D',severity:'fail',match:{type:'starts',label:'R&D'},note:'Required for OPEX split'},
  {sheet:'P&L',item:'Fixed Staff / Salary',severity:'fail',match:{type:'any',labels:['Fixed Staff','Salary']},note:'At least one payroll row must exist'},
  {sheet:'Cash Flow & DCF',item:'Capital Investment',severity:'fail',match:{type:'starts',label:'Capital Investment'},note:'Required for Y0 investment'},
  {sheet:'Cash Flow & DCF',item:'Operating Cash Flow',severity:'fail',match:{type:'starts',label:'Operating Cash Flow'},note:'Required for cash flow model'},
  {sheet:'Cash Flow & DCF',item:'Total Investment',severity:'fail',match:{type:'starts',label:'Total Investment'},note:'Required for investor capital base'},
  {sheet:'Cash Flow & DCF',item:'Investor Equity Stake',severity:'optional',match:{type:'starts',label:'Investor Equity Stake'},note:'Needed for equity model tie-out'},
  {sheet:'Cash Flow & DCF',item:'Dividend Payout Ratio',severity:'optional',match:{type:'starts',label:'Dividend Payout Ratio'},note:'Needed for equity cash yield tie-out'},
  {sheet:'Cash Flow & DCF',item:'Preferred Return',severity:'optional',match:{type:'starts',label:'Preferred Return'},note:'Needed for pref waterfall tie-out'},
  {sheet:'Cash Flow & DCF',item:'Revenue Share Rate',severity:'optional',match:{type:'starts',label:'Revenue Share Rate'},note:'Needed for investor rev-share model'},
  {sheet:'Cash Flow & DCF',item:'Target Return Multiple',severity:'optional',match:{type:'starts',label:'Target Return Multiple'},note:'Needed for rev-share cap'},
  {sheet:'Cash Flow & DCF',item:'Grace Period',severity:'optional',match:{type:'starts',label:'Grace Period'},note:'Needed for rev-share start delay'},
  {sheet:'Revenue',item:'Physio Clinic',severity:'fail',match:{type:'starts',label:'Physio Clinic'},note:'Department revenue row'},
  {sheet:'Revenue',item:'Dental Clinic',severity:'fail',match:{type:'starts',label:'Dental Clinic'},note:'Department revenue row'},
  {sheet:'Revenue',item:'Polyclinic',severity:'fail',match:{type:'starts',label:'Polyclinic'},note:'Department revenue row'},
  {sheet:'Revenue',item:'Cosmetology',severity:'fail',match:{type:'starts',label:'Cosmetology'},note:'Department revenue row'},
  {sheet:'Revenue',item:'Wellness Packages',severity:'fail',match:{type:'starts',label:'Wellness Packages'},note:'Department revenue row'},
  {sheet:'Patient Acquisition',item:'Total Visits',severity:'fail',match:{type:'starts',label:'Total Visits'},note:'Required for acquisition model'},
  {sheet:'Patient Acquisition',item:'Retained from Prior',severity:'fail',match:{type:'starts',label:'Retained from Prior'},note:'Required for retention bridge'},
  {sheet:'Patient Acquisition',item:'Referral',severity:'fail',match:{type:'starts',label:'Referral'},note:'Required for channel mix'},
  {sheet:'Patient Acquisition',item:'Walk-in',severity:'fail',match:{type:'starts',label:'Walk-in'},note:'Required for channel mix'},
  {sheet:'Patient Acquisition',item:'Marketing (paid)',severity:'fail',match:{type:'starts',label:'Marketing (paid'},note:'Required for paid acquisition'},
  {sheet:'Patient Acquisition',item:'Total Unique Patients',severity:'fail',match:{type:'starts',label:'Total Unique Patients'},note:'Required for patient base'},
  {sheet:'Patient Acquisition',item:'Annual Marketing Budget',severity:'warn',match:{type:'contains',label:'Annual Marketing Budget'},note:'Needed for budget tie-out'},
  {sheet:'Patient Acquisition',item:'CAC per Marketing Patient',severity:'warn',match:{type:'contains',label:'CAC per Marketing Patient'},note:'Needed for pure CAC tie-out'},
  {sheet:'Patient Acquisition',item:'Blended CAC',severity:'warn',match:{type:'contains',label:'Blended CAC'},note:'Needed for blended CAC tie-out'},
  {sheet:'Patient Acquisition',item:'LTV:CAC Ratio',severity:'warn',match:{type:'contains',label:'LTV:CAC Ratio'},note:'Needed for LTV:CAC tie-out'},
  {sheet:'Patient Acquisition',item:'CAC Payback',severity:'warn',match:{type:'contains',label:'CAC Payback'},note:'Needed for payback tie-out'},
  {sheet:'Patient Acquisition',item:'Patient Lifetime',severity:'warn',match:{type:'contains',label:'Patient Lifetime'},note:'Needed for LTV rebuild'},
  {sheet:'Patient Acquisition',item:'Patient LTV',severity:'warn',match:{type:'contains',label:'Patient LTV'},note:'Needed for economics cross-check'},
  {sheet:'Capacity',item:'Physio Clinic → Annual Visits',severity:'warn',match:{type:'annualVisitsAfter',label:'Physio Clinic'},note:'Needed for service utilization tie-out'},
  {sheet:'Capacity',item:'Dental Clinic → Annual Visits',severity:'warn',match:{type:'annualVisitsAfter',label:'Dental Clinic'},note:'Needed for service utilization tie-out'},
  {sheet:'Capacity',item:'GP → Annual Visits',severity:'warn',match:{type:'annualVisitsAfter',label:'GP'},note:'Needed for polyclinic visits tie-out'},
  {sheet:'Capacity',item:'Gynecology → Annual Visits',severity:'warn',match:{type:'annualVisitsAfter',label:'Gynecology'},note:'Needed for polyclinic visits tie-out'},
  {sheet:'Capacity',item:'Cosmetology → Annual Visits',severity:'warn',match:{type:'annualVisitsAfter',label:'Cosmetology'},note:'Needed for service utilization tie-out'},
  {sheet:'Monthly P&L',item:'Revenue',severity:'warn',match:{type:'starts',label:'Revenue'},note:'Needed for monthly phasing'},
  {sheet:'Monthly P&L',item:'EBITDA',severity:'warn',match:{type:'starts',label:'EBITDA'},note:'Needed for monthly phasing'},
  {sheet:'Price List',item:'Product rows',severity:'optional',match:{type:'priceListRows'},note:'Needed for product mix module'}
];
const DEFAULT_PL_BREAKDOWN_WEIGHTS=(()=>{
  const seed={marketing:2437953,admin:1799242,rent:957412,staff:758068,rd:510166};
  const total=Object.values(seed).reduce((sum,val)=>sum+val,0);
  return Object.fromEntries(Object.entries(seed).map(([key,val])=>[key,val/total]));
})();

function syncPLBreakdown(existing=BASE.plBreakdown){
  const next={marketing:[],admin:[],rent:[],staff:[],rd:[]};
  PL_BREAKDOWN_KEYS.forEach(key=>{
    const src=Array.isArray(existing?.[key])?existing[key]:[];
    next[key]=src.slice(0,10).map(v=>Math.abs(Number(v)||0));
  });

  for(let i=0;i<10;i++){
    const rev=Number(BASE.revenue[i])||0;
    const cogs=Math.abs(Number(BASE.cogs[i])||0);
    const ebitda=Number(BASE.ebitda[i])||0;
    const targetOther=Math.max(rev-(rev*0.40)-cogs-ebitda,0);
    const currentOther=PL_BREAKDOWN_KEYS.reduce((sum,key)=>sum+(next[key][i]||0),0);
    const scale=currentOther>0?targetOther/currentOther:1;

    PL_BREAKDOWN_KEYS.forEach(key=>{
      const fallback=targetOther*DEFAULT_PL_BREAKDOWN_WEIGHTS[key];
      next[key][i]=currentOther>0?(next[key][i]||0)*scale:fallback;
    });
  }

  BASE.plBreakdown=next;
}

function syncServiceRevenue(){
  const serviceCount=BASE.services.length;
  const revenueByService=Array.from({length:serviceCount},(_,svcIdx)=>{
    const src=Array.isArray(BASE.serviceRevenue?.[svcIdx])?BASE.serviceRevenue[svcIdx]:[];
    return Array.from({length:10},(_,yearIdx)=>Math.max(0,Number(src[yearIdx])||0));
  });
  const visitsByService=Array.from({length:serviceCount},(_,svcIdx)=>{
    const src=Array.isArray(BASE.serviceVisits?.[svcIdx])?BASE.serviceVisits[svcIdx]:[];
    return Array.from({length:10},(_,yearIdx)=>Math.max(0,Number(src[yearIdx])||0));
  });

  for(let yearIdx=0;yearIdx<10;yearIdx++){
    const serviceTotal=revenueByService.reduce((sum,row)=>sum+(row[yearIdx]||0),0);
    const modelTotal=Math.max(0,Number(BASE.revenue[yearIdx])||0);
    const fallbackMix=BASE.serviceMix?.[yearIdx]||[];
    if(serviceTotal<=0&&modelTotal>0){
      revenueByService.forEach((row,svcIdx)=>{
        row[yearIdx]=modelTotal*((Number(fallbackMix[svcIdx])||0)/100);
      });
    }else if(modelTotal>0&&serviceTotal>0&&Math.abs(serviceTotal-modelTotal)>1){
      const scale=modelTotal/serviceTotal;
      revenueByService.forEach(row=>{row[yearIdx]*=scale;});
    }
  }

  BASE.serviceRevenue=revenueByService;
  BASE.serviceVisits=visitsByService;
  BASE.serviceMix=Array.from({length:10},(_,yearIdx)=>{
    const total=revenueByService.reduce((sum,row)=>sum+(row[yearIdx]||0),0);
    return revenueByService.map(row=>total>0?(row[yearIdx]||0)/total*100:0);
  });
  BASE.revPerVisit=revenueByService.map((row,svcIdx)=>{
    const visit=row[4]&&visitsByService[svcIdx]?.[4]?visitsByService[svcIdx][4]:0;
    return visit>0?row[4]/visit:0;
  });
}

function preferPositiveValue(value,fallback){
  const num=Number(value);
  return Number.isFinite(num)&&num>0?num:fallback;
}

function syncPatientAcquisitionModel(existing=BASE.paModel){
  const next=Object.fromEntries(PA_MODEL_KEYS.map(key=>[
    key,
    Array.isArray(existing?.[key])?existing[key].slice(0,10).map(v=>Math.max(Number(v)||0,0)):[]
  ]));

  for(let i=0;i<10;i++){
    const retained=Math.max(0,Number(BASE.retained[i])||0);
    const referral=Math.max(0,Number(BASE.referral[i])||0);
    const walkin=Math.max(0,Number(BASE.walkin[i])||0);
    const marketing=Math.max(0,Number(BASE.marketing[i])||0);
    const uniqueFallback=retained+referral+walkin+marketing;
    const visits=Math.max(0,Number(BASE.visits[i])||0);
    const revenue=Math.max(0,Number(BASE.revenue[i])||0);
    const marketingBudgetFallback=Math.max(0,Number(BASE.plBreakdown.marketing?.[i])||0);
    const uniquePatients=preferPositiveValue(next.uniquePatients[i],uniqueFallback);
    const avgRevenuePerVisit=preferPositiveValue(next.avgRevenuePerVisit[i],visits>0?revenue/visits:0);
    const avgRevenuePerPatient=preferPositiveValue(next.avgRevenuePerPatient[i],uniquePatients>0?revenue/uniquePatients:0);
    const marketingBudget=preferPositiveValue(next.marketingBudget[i],marketingBudgetFallback);
    const pureCAC=preferPositiveValue(next.pureCAC[i],marketing>0?marketingBudget/marketing:0);
    const blendedCAC=preferPositiveValue(next.blendedCAC[i],(referral+walkin+marketing)>0?marketingBudget/(referral+walkin+marketing):0);
    const lifetimeYears=preferPositiveValue(next.lifetimeYears[i],3);
    const ltv=preferPositiveValue(next.ltv[i],avgRevenuePerPatient*lifetimeYears);
    const ltvCac=preferPositiveValue(next.ltvCac[i],pureCAC>0?ltv/pureCAC:(blendedCAC>0?ltv/blendedCAC:0));
    const marginFloor=Math.max(Number(BASE.margins[i])||0,0.08);
    const annualContribution=avgRevenuePerPatient*marginFloor;
    const cacPayback=preferPositiveValue(next.cacPayback[i],pureCAC>0&&annualContribution>0?pureCAC/(annualContribution/12):0);
    const romi=preferPositiveValue(next.romi[i],marketingBudget>0?((marketing*avgRevenuePerPatient)-marketingBudget)/marketingBudget:0);
    const retainedVisits=preferPositiveValue(next.retainedVisits[i],uniquePatients>0?visits*(retained/uniquePatients):0);
    const organicVisits=preferPositiveValue(next.organicVisits[i],uniquePatients>0?visits*((referral+walkin)/uniquePatients):0);
    const paidVisits=preferPositiveValue(next.paidVisits[i],uniquePatients>0?visits*(marketing/uniquePatients):0);
    const acquisitionSpend=preferPositiveValue(next.acquisitionSpend[i],marketingBudget);
    const retentionSpend=preferPositiveValue(next.retentionSpend[i],retained*avgRevenuePerPatient*0.05);

    next.uniquePatients[i]=uniquePatients;
    next.avgRevenuePerVisit[i]=avgRevenuePerVisit;
    next.avgRevenuePerPatient[i]=avgRevenuePerPatient;
    next.marketingBudget[i]=marketingBudget;
    next.pureCAC[i]=pureCAC;
    next.blendedCAC[i]=blendedCAC;
    next.lifetimeYears[i]=lifetimeYears;
    next.ltv[i]=ltv;
    next.ltvCac[i]=ltvCac;
    next.cacPayback[i]=cacPayback;
    next.romi[i]=romi;
    next.retainedVisits[i]=retainedVisits;
    next.organicVisits[i]=organicVisits;
    next.paidVisits[i]=paidVisits;
    next.acquisitionSpend[i]=acquisitionSpend;
    next.retentionSpend[i]=retentionSpend;
  }

  BASE.paModel=next;
  BASE.patientsShort=BASE.plIdx.map(i=>Math.round(next.uniquePatients[i]||0));
  BASE.visitsShort=BASE.plIdx.map(i=>Math.round(Number(BASE.visits[i])||0));
  return next;
}
function syncProductCatalog(existing=BASE.productCatalog){
  const source=Array.isArray(existing)&&existing.length?existing:DEFAULT_PRODUCT_CATALOG;
  BASE.productCatalog=source.map(item=>({
    department:String(item?.department||'').trim(),
    product:String(item?.product||'').trim(),
    basePrice:Math.max(0,Number(item?.basePrice)||0),
    monthlyQty:Array.from({length:10},(_,yearIdx)=>Math.max(0,Number(item?.monthlyQty?.[yearIdx])||0))
  })).filter(item=>item.department&&item.product);
  return BASE.productCatalog;
}

syncPLBreakdown();
syncServiceRevenue();
syncPatientAcquisitionModel();
syncProductCatalog();

// Compute ratios from base
for(let i=0;i<10;i++){
  BASE.npRatio[i]=BASE.netProfit[i]/BASE.ebitda[i];
  BASE.cfRatio[i]=BASE.annCF[i+1]/BASE.ebitda[i];
}
// opexRate = non-revshare cost rate
BASE.opexRate=BASE.margins.map(m=>1-m-0.40);

/* ===== STATE ===== */
const STATE={revGrowth:0,revShare:40,opexAdj:0,staffAdj:0,wacc:12,exitMult:8,scenario:'base'};
const SCENARIOS={
  base:{revGrowth:0,revShare:40,opexAdj:0,staffAdj:0,wacc:12,exitMult:8},
  optimistic:{revGrowth:15,revShare:38,opexAdj:-10,staffAdj:-5,wacc:10,exitMult:10},
  pessimistic:{revGrowth:-15,revShare:45,opexAdj:10,staffAdj:10,wacc:15,exitMult:6}
};
const UI_STORAGE_KEYS={
  panelCollapsed:`${STORAGE_PREFIX}.panelCollapsed`,
  panelGroups:`${STORAGE_PREFIX}.panelGroups`,
  templateComplianceDetails:`${STORAGE_PREFIX}.templateComplianceDetails`,
  language:`${STORAGE_PREFIX}.language`
};
const APP_STORAGE_KEYS={
  selectedProject:`${STORAGE_PREFIX}.selectedProject`,
  scenarioLibrary:'scenarioLibrary',
  importAudit:'importAudit',
  equityModel:'equityModelV2'
};
const MC_PROFILES={
  balanced:{
    label:{ru:'База',en:'Base'},
    sims:10000,
    macroSd:0.08,
    executionSd:0.06,
    pricingSd:0.04,
    costSd:0.035,
    yearlySd:0.025,
    discountSd:0.010,
    exitSd:0.35,
    downsideShift:0,
    launchDelayProb:0.10,
    launchDelayMonths:[1,4],
    capexOverrunProb:0.12,
    capexOverrunRange:[0.05,0.18],
    occupancyShortfallProb:0.14,
    occupancyShortfallRange:[0.08,0.24],
    pricePressureProb:0.12,
    pricePressureRange:[0.03,0.10],
    description:{ru:'Коррелированные макро- и операционные шоки вокруг текущего состояния модели.',en:'Correlated macro + execution shocks around the current model state.'}
  },
  cautious:{
    label:{ru:'Осторожный',en:'Cautious'},
    sims:10000,
    macroSd:0.11,
    executionSd:0.08,
    pricingSd:0.05,
    costSd:0.045,
    yearlySd:0.030,
    discountSd:0.013,
    exitSd:0.45,
    downsideShift:-0.02,
    launchDelayProb:0.18,
    launchDelayMonths:[2,6],
    capexOverrunProb:0.18,
    capexOverrunRange:[0.08,0.24],
    occupancyShortfallProb:0.20,
    occupancyShortfallRange:[0.12,0.32],
    pricePressureProb:0.18,
    pricePressureRange:[0.05,0.14],
    description:{ru:'Более тяжелые downside-хвосты, более сильное давление на затраты и слабее условия выхода.',en:'Heavier downside tails, wider cost pressure and weaker exit conditions.'}
  },
  wide:{
    label:{ru:'Широкий',en:'Wide'},
    sims:10000,
    macroSd:0.14,
    executionSd:0.10,
    pricingSd:0.06,
    costSd:0.055,
    yearlySd:0.040,
    discountSd:0.016,
    exitSd:0.60,
    downsideShift:0,
    launchDelayProb:0.15,
    launchDelayMonths:[1,6],
    capexOverrunProb:0.15,
    capexOverrunRange:[0.04,0.25],
    occupancyShortfallProb:0.18,
    occupancyShortfallRange:[0.08,0.30],
    pricePressureProb:0.16,
    pricePressureRange:[0.03,0.14],
    description:{ru:'Широкий комитетный диапазон с максимальной дисперсией по выручке, затратам и мультипликатору выхода.',en:'Broad committee-style range with the widest revenue, cost and multiple dispersion.'}
  }
};
const FOCUS_STATE={card:null,placeholder:null,trigger:null,hideTimer:null};
let monteCarloProfile='balanced';
let investorTimelineMode='annual';
let investorScenarioMode='native';
let currentProjectId=ACTIVE_DEFAULT_PROJECT_ID;
let currentProjectLoadToken=0;
let currentLanguage='ru';
let appReady=false;
let monteCarloLastResults=null;
let scenarioLibrary=[];
let importAudit={fileName:'No workbook loaded',importedAt:null,sheets:[],mappedMetrics:[],actualSheets:[],actualsStatus:'No actuals connected'};
let excelChecklist={fileName:'No workbook loaded',wacc:12,checks:[]};
let templateCompliance={fileName:'No workbook checked',checkedAt:null,ready:false,summary:{ok:0,fail:0,warn:0,optional:0},checks:[]};
let templateComplianceDetailsExpanded=false;
let templateComplianceFilter='all';
let runtimeModelSource=createModelSourceState();
let patientEngineView='trend';
let patientTrendMode='blend';
let equityState={
  totalInvestment:6572160,
  returnMode:'compare',
  exitYear:5,
  equityStakePct:30,
  dividendPayoutPct:80,
  preferredReturnPct:8,
  revenueShareRatePct:10,
  targetReturnMultiple:2.5,
  graceMonths:6
};
let ADJ={revenue:[],ebitda:[],netProfit:[],cf:[],cumCF:[],discCF:[],discCumCF:[],margins:[],npv:0,irr5:0,irr10:0,payback:0,exitEV5:0,moic5:0,exitIRR5:0};

function resolveProjectId(projectId){
  return ENABLED_PROJECT_IDS.includes(projectId)?projectId:ACTIVE_DEFAULT_PROJECT_ID;
}
function isProjectEnabled(projectId){
  return ENABLED_PROJECT_IDS.includes(projectId);
}
function resolveLanguage(lang){
  return I18N[lang]?lang:'ru';
}
function tr(key,vars={}){
  const dict=I18N[currentLanguage]||I18N.ru;
  const fallback=I18N.ru;
  const template=(dict&&dict[key]!==undefined?dict[key]:fallback[key])??key;
  return String(template).replace(/\{(\w+)\}/g,(_,name)=>vars[name]??`{${name}}`);
}
function localizeProjectValue(value){
  if(value&&typeof value==='object'){
    return value[currentLanguage]??value.ru??value.en??'';
  }
  return value;
}
function formatPaybackValue(value,withApprox=false){
  if(value===null||value===undefined||Number.isNaN(Number(value)))return 'N/A';
  const prefix=withApprox?'~':'';
  return `${prefix}${Number(value).toFixed(1)} ${tr('yearsUnit')}`;
}
function localizeDurationText(value){
  if(value===null||value===undefined)return value;
  if(currentLanguage==='ru')return String(value);
  return String(value)
    .replace(/лет/g,tr('yearsUnit'))
    .replace(/мес/g,tr('monthsUnit'));
}
function formatDateTimeLocalized(value){
  if(!value)return '—';
  return new Date(value).toLocaleString(currentLanguage==='ru'?'ru-RU':'en-GB',{hour12:false});
}
function setPlaceholder(id,value){
  const el=document.getElementById(id);
  if(el)el.placeholder=value;
}
function setNodeListText(selector,values){
  document.querySelectorAll(selector).forEach((el,idx)=>{
    if(values[idx]!==undefined)el.textContent=values[idx];
  });
}
function getReturnModeLabel(mode){
  const labels={
    compare:tr('modeCompare'),
    equity:tr('modeEquity'),
    'revenue-share':tr('modeRevenueShare')
  };
  return labels[mode]||mode;
}
function getWinnerLabel(value){
  if(value==='Equity')return tr('modeEquity');
  if(value==='Revenue Share')return tr('modeRevenueShare');
  if(value==='Tie')return currentLanguage==='ru'?'Паритет':'Tie';
  return value||'—';
}
function getLocalizedActualsStatus(){
  if(importAudit.actualSheets?.length){
    return currentLanguage==='ru'
      ? `Найдены листы факта (${importAudit.actualSheets.length}): ${importAudit.actualSheets.join(', ')}`
      : `Detected ${importAudit.actualSheets.length} actuals-oriented sheet(s): ${importAudit.actualSheets.join(', ')}`;
  }
  const normalized=String(importAudit.actualsStatus||'').toLowerCase();
  if(!normalized||normalized.includes('awaiting'))return currentLanguage==='ru'?'Ожидается маппинг факта':'Awaiting actuals mapping';
  if(normalized.includes('no actuals connected'))return currentLanguage==='ru'?'Факт не подключен':'No actuals connected';
  if(normalized.includes('no actuals sheet detected'))return currentLanguage==='ru'?'В workbook не найден лист факта':'No actuals sheet detected in workbook';
  return importAudit.actualsStatus||'—';
}
function getMonteCarloProfileLabel(profileKey=monteCarloProfile){
  return localizeProjectValue(MC_PROFILES[profileKey]?.label)||profileKey;
}
function getMonteCarloProfileDescription(profileKey=monteCarloProfile){
  return localizeProjectValue(MC_PROFILES[profileKey]?.description)||'';
}
function updateLanguageSwitcher(){
  document.querySelectorAll('#languageSwitcher [data-lang]').forEach(btn=>{
    btn.classList.toggle('active',btn.dataset.lang===currentLanguage);
  });
}
function applyStaticTranslations(){
  document.documentElement.lang=currentLanguage==='en'?'en':'ru';
  updateLanguageSwitcher();
  const textMap=[
    ['.nav-btn[data-section="overview"]','navOverview'],
    ['.nav-btn[data-section="revenue"]','navRevenue'],
    ['.nav-btn[data-section="patients"]','navPatients'],
    ['.nav-btn[data-section="sensitivity"]','navSensitivity'],
    ['.nav-btn[data-section="risks"]','navRisks'],
    ['#filterLabel','filterPeriod'],
    ['.panel-title','panelModel'],
    ['.panel-sub','panelRealtime'],
    ['.panel-group[data-panel-group="scenario"] .panel-group-title','groupScenario'],
    ['.panel-group[data-panel-group="drivers"] .panel-group-title','groupDrivers'],
    ['.sc-btn[data-scenario="base"]','scenarioBase'],
    ['.sc-btn[data-scenario="optimistic"]','scenarioOptimistic'],
    ['.sc-btn[data-scenario="pessimistic"]','scenarioPessimistic'],
    ['#btnReset','resetToBase'],
    ['#kpi-payback-sub','paybackBenchmark'],
    ['#overviewCapexTitle','overviewCapex'],
    ['#overviewPatientsTitle','overviewPatients'],
    ['#overviewSourcesTitle','overviewSources'],
    ['#overviewUtilTitle','overviewUtil'],
    ['#scenarioCompareTitle','scenarioCompare'],
    ['#scenarioMetricTh','scenarioMetric'],
    ['#scenarioDeltaTh','scenarioDelta'],
    ['#scenarioPaybackBase','scenarioPaybackBase'],
    ['#scenarioPaybackCompact','scenarioPaybackCompact'],
    ['#scenarioPaybackDelta','scenarioPaybackDelta'],
    ['#revenueSectionTitle','revenueSection'],
    ['[data-chart="revenue-abs"]','chartAbsolute'],
    ['[data-chart="revenue-margin"]','chartMargin'],
    ['#monthlyPlTitle','monthlyPL'],
    ['#actualVsPlanTitle','actualVsPlanTitle'],
    ['#mixEvolutionTitle','mixEvolution'],
    ['#plDetailTitle','plDetail'],
    ['#plMetricHeader','plMetric'],
    ['#patientEconomicsTitle','patientEconomics'],
    ['#cashflowMainTitle','cashflowMain'],
    ['#fundingBridgeTitle','fundingBridge'],
    ['#runwaySnapshotTitle','runwaySnapshot'],
    ['#exitValuationTitle','exitValuation'],
    ['#investorModelAssumptionsTitle','investorModelAssumptions'],
    ['#equityCompareTableTitle','equityCompareTableTitle'],
    ['#capexBreakdownTitle','capexBreakdown'],
    ['#netBookValueTitle','netBookValue'],
    ['#depreciationScheduleTitle','depreciationSchedule'],
    ['#amortizationTableTitle','amortizationTable'],
    ['#tornadoTitle','tornadoSectionTitle'],
    ['#scenarioLibraryTitle','scenarioLibrarySectionTitle'],
    ['#mcSimulationTitle','mcSimulationSectionTitle'],
    ['#integrityChecksTitle','integrityChecksSectionTitle'],
    ['#ratioRadarTitle','ratioRadarTitle'],
    ['#assumptionsDefinitionsTitle','assumptionsDefinitions'],
    ['#benchmarkSourcesTitle','benchmarkSources'],
    ['#templateComplianceSectionTitle','templateComplianceSectionTitle'],
    ['#importAuditSectionTitle','importAuditSectionTitle'],
    ['#excelChecklistSectionTitle','excelChecklistSectionTitle'],
    ['#monthlyDrilldownTitle','monthlyDrilldownTitle'],
    ['#monthlyDrilldownNote','monthlyDrilldownNote'],
    ['#irrByExitTitle','irrByExit'],
    ['#exitScenariosTitle','exitScenarios'],
    ['#riskConcentrationText','riskConcentrationText'],
    ['#riskConcentrationMitigation','riskConcentrationMitigation'],
    ['#riskUtilizationMitigation','riskUtilizationMitigation'],
    ['#riskDubaiMitigation','riskDubaiMitigation'],
    ['#ratioTrendTitle','ratioTrends']
  ];
  textMap.forEach(([selector,key])=>{
    const el=document.querySelector(selector);
    if(el)el.textContent=tr(key);
  });
  const verdict=document.getElementById('verdictText');
  if(verdict)verdict.innerHTML=tr('verdictHtml');
  const mcLegend=document.getElementById('mcLegend');
  if(mcLegend)mcLegend.innerHTML=tr('mcLegendHtml');
  setPlaceholder('scenarioNameInput',tr('scenarioNamePlaceholder'));
  setText('scenarioSaveBtn',tr('saveCurrent'));
  setText('scenarioClearBtn',tr('clearAll'));
  const mcRunBtn=document.getElementById('mcRunBtn');
  if(mcRunBtn&&mcRunBtn.dataset.running!=='true'){
    mcRunBtn.textContent=currentLanguage==='ru'?'▶ Запустить 10K симуляций':'▶ Run 10K Simulations';
  }
  if(document.getElementById('mcStatus')&&mcRunBtn&&mcRunBtn.dataset.running!=='true'){
    markMonteCarloStale(false);
  }
  setNodeListText('#returnModeAssumptionTabs [data-return-mode]',[tr('modeCompare'),tr('modeEquity'),tr('modeRevenueShare')]);
  setNodeListText('#investorTimelineTabs [data-timeline-mode]',[tr('timelineAnnual'),tr('timelineCumulative')]);
  setNodeListText('#investorScenarioTabs [data-scenario-mode]',[tr('scenarioNative'),tr('scenarioExitAligned')]);
  setNodeListText('#mcProfileTabs [data-mc-profile]',[
    getMonteCarloProfileLabel('balanced'),
    getMonteCarloProfileLabel('cautious'),
    getMonteCarloProfileLabel('wide')
  ]);
  setNodeListText('#equityControls .equity-mode-control .equity-label',[tr('modelMode')]);
  setNodeListText('#patientAcqTable thead th',currentLanguage==='ru'
    ? ['Канал','Пациенты','Доля','Визиты','Revenue proxy','Затраты','Эффективность']
    : ['Channel','Patients','Share','Visits','Revenue Proxy','Cost','Efficiency']);
  setNodeListText('#patientBaseTable thead th',currentLanguage==='ru'
    ? ['Год','База на начало','Новые платные','Новые органические','Удержанные','Отток','База на конец','Retention %','Churn %']
    : ['Year','Opening Base','New Paid','New Organic','Retained','Churned','Closing Base','Retention %','Churn %']);
  setNodeListText('#equityTable thead th',currentLanguage==='ru'
    ? ['Метрика','Equity','Revenue Share','Победитель']
    : ['Metric','Equity Model','Revenue Share','Winner']);
  setNodeListText('#scenarioLibraryTable thead th',currentLanguage==='ru'
    ? ['Название','Рост выручки','Comp провайдера','OPEX','Fixed Staff','WACC','Exit','NPV','10Y IRR','Payback','Действия']
    : ['Name','Rev','Provider Comp','OPEX','Fixed Staff','WACC','Exit','NPV','10Y IRR','Payback','Actions']);
}
function initLanguageSwitcher(){
  const switcher=document.getElementById('languageSwitcher');
  if(!switcher)return;
  switcher.querySelectorAll('[data-lang]').forEach(btn=>{
    if(btn.dataset.bound)return;
    btn.addEventListener('click',()=>setLanguage(btn.dataset.lang));
    btn.dataset.bound='true';
  });
  updateLanguageSwitcher();
}
function setLanguage(lang,{persist=true}={}){
  currentLanguage=resolveLanguage(lang);
  if(persist)writeStoredUI(UI_STORAGE_KEYS.language,currentLanguage);
  applyStaticTranslations();
  updateProjectBranding();
  syncPatientTrendModeTabs();
  syncPatientEngineViewTabs();
  applyPatientEngineView();
  if(appReady){
    updateAll();
    if(monteCarloLastResults)renderMCResults(monteCarloLastResults);
  }
}
function getProjectConfig(projectId=currentProjectId){
  return PROJECTS[resolveProjectId(projectId)];
}
function getProjectStorageKey(key,projectId=currentProjectId){
  return `${STORAGE_PREFIX}.${resolveProjectId(projectId)}.${key}`;
}
function createImportAudit(project=getProjectConfig()){
  return{
    fileName:project.workbookLabel,
    importedAt:null,
    sheets:[],
    mappedMetrics:[],
    actualSheets:[],
    actualsStatus:'No actuals connected'
  };
}
function createExcelChecklist(project=getProjectConfig()){
  return{fileName:project.workbookLabel,wacc:STATE.wacc,checks:[]};
}
function createTemplateCompliance(project=getProjectConfig()){
  return{
    fileName:project.workbookLabel,
    checkedAt:null,
    ready:false,
    summary:{ok:0,fail:0,warn:0,optional:0},
    checks:[]
  };
}
function createModelSourceState(overrides={}){
  return{
    type:'fallback',
    reason:'pending',
    origin:'embedded',
    fileName:null,
    loadedAt:null,
    ...overrides
  };
}
function setModelSourceState(nextState={},options={}){
  const {rerender=false}=options;
  runtimeModelSource={...runtimeModelSource,...nextState};
  if(rerender){
    renderSourcesPanel();
    renderImportAudit();
    updateFooterCurrency();
  }
}
function describeModelSource(){
  const project=getProjectConfig();
  if(runtimeModelSource.type==='workbook'){
    const autoLoaded=runtimeModelSource.origin==='auto';
    return{
      sourceLabel:runtimeModelSource.fileName||project.workbookLabel,
      statusLabel:currentLanguage==='ru'
        ? (autoLoaded?'Workbook загружен автоматически':'Workbook загружен вручную')
        : (autoLoaded?'Workbook auto-loaded':'Workbook imported manually'),
      note:currentLanguage==='ru'
        ? `Текущая модель дашборда построена из workbook ${runtimeModelSource.fileName||project.workbookLabel}. Live-сценарии применяются поверх этого baseline.`
        : `The current dashboard model is built from workbook ${runtimeModelSource.fileName||project.workbookLabel}. Live scenarios are applied on top of that baseline.`,
      tone:'positive'
    };
  }
  const fallbackLabel=currentLanguage==='ru'?'Embedded workbook-aligned baseline':'Embedded workbook-aligned baseline';
  const reasonNotes={
    pending:currentLanguage==='ru'
      ? 'Workbook этой сессии еще не загружен. Дашборд работает на встроенном baseline, синхронизированном с workbook на момент сборки.'
      : 'No workbook has been loaded in this session yet. The dashboard is running on the embedded baseline aligned to the workbook at build time.',
    file_protocol:currentLanguage==='ru'
      ? 'Auto-load недоступен при запуске через file://. Дашборд продолжает работать на встроенном workbook-aligned baseline; для live-import поднимите локальный сервер.'
      : 'Auto-load is unavailable when the dashboard is opened via file://. The dashboard stays on the embedded workbook-aligned baseline; start a local server for live import.',
    fetch_failed:currentLanguage==='ru'
      ? `Не удалось автоматически загрузить ${project.workbookLabel}. Дашборд оставлен на встроенном fallback-наборе.`
      : `Could not auto-load ${project.workbookLabel}. The dashboard remains on the embedded fallback dataset.`,
    workbook_missing:currentLanguage==='ru'
      ? `Workbook ${project.workbookLabel} не найден. Дашборд оставлен на встроенном fallback-наборе.`
      : `Workbook ${project.workbookLabel} was not found. The dashboard remains on the embedded fallback dataset.`,
    template_blocked:currentLanguage==='ru'
      ? `Workbook не прошел template compliance. Текущая модель не была заменена.`
      : 'Workbook failed template compliance. The current model was not replaced.'
  };
  return{
    sourceLabel:fallbackLabel,
    statusLabel:currentLanguage==='ru'?'Fallback active':'Fallback active',
    note:reasonNotes[runtimeModelSource.reason]||reasonNotes.pending,
    tone:'warning'
  };
}
function updateProjectBranding(){
  const project=getProjectConfig();
  document.title=`${project.title} — Financial Dashboard`;
  const titleEl=document.getElementById('projectTitle');
  const subtitleEl=document.getElementById('projectSubtitle');
  const benchmarkTitle=document.getElementById('benchmarkTitle');
  const dropHint=document.getElementById('dropOverlayHint');
  if(titleEl)titleEl.textContent=project.title;
  if(subtitleEl)subtitleEl.textContent=localizeProjectValue(project.subtitle);
  if(benchmarkTitle)benchmarkTitle.textContent=tr('benchmarkTitle',{project:project.shortTitle,market:localizeProjectValue(project.benchmarkMarketLabel)});
  if(dropHint)dropHint.textContent=project.workbookLabel;
  const switcher=document.getElementById('projectSwitcher');
  if(switcher){
    let visibleCount=0;
    switcher.querySelectorAll('[data-project]').forEach(btn=>{
      const enabled=isProjectEnabled(btn.dataset.project);
      btn.hidden=!enabled;
      btn.classList.toggle('active',enabled&&btn.dataset.project===project.id);
      if(enabled)visibleCount+=1;
    });
    switcher.hidden=visibleCount<=1;
    switcher.setAttribute('aria-hidden',String(visibleCount<=1));
  }
}
function setImportButtonState(state,label=''){
  const importBtn=document.getElementById('importBtn');
  if(!importBtn)return;
  importBtn.disabled=state==='loading';
  importBtn.classList.toggle('loaded',state==='loaded');
  if(state==='loading'){
    importBtn.textContent=`⏳ ${label||'Loading workbook'}`;
    return;
  }
  if(state==='loaded'){
    importBtn.textContent=`✅ ${label}`;
    return;
  }
  importBtn.textContent='📊 Import .xlsx';
}
function loadProjectScopedState(){
  importAudit=readStoredUI(APP_STORAGE_KEYS.importAudit,createImportAudit(),'project');
  scenarioLibrary=readStoredUI(APP_STORAGE_KEYS.scenarioLibrary,[],'project');
  equityState=normalizeEquityState(readStoredUI(APP_STORAGE_KEYS.equityModel,getDefaultEquityState(),'project'));
  setModelSourceState(createModelSourceState({reason:'pending'}));
  if(!scenarioLibrary.length){
    scenarioLibrary=buildPresetScenarioLibrary();
    writeStoredUI(APP_STORAGE_KEYS.scenarioLibrary,scenarioLibrary,'project');
  }
}
function initProjectSelector(){
  const switcher=document.getElementById('projectSwitcher');
  if(!switcher)return;
  switcher.querySelectorAll('[data-project]').forEach(btn=>{
    if(btn.dataset.bound)return;
    btn.addEventListener('click',()=>switchProject(btn.dataset.project));
    btn.dataset.bound='true';
  });
  updateProjectBranding();
}
async function switchProject(projectId,{force=false,announce=true}={}){
  const nextProjectId=resolveProjectId(projectId);
  if(!force&&nextProjectId===currentProjectId)return;
  currentProjectId=nextProjectId;
  writeStoredUI(APP_STORAGE_KEYS.selectedProject,currentProjectId);
  loadProjectScopedState();
  monteCarloLastResults=null;
  templateCompliance=createTemplateCompliance();
  excelChecklist=createExcelChecklist();
  templateComplianceFilter='all';
  setModelSourceState(createModelSourceState({reason:'pending'}));
  updateProjectBranding();
  setImportButtonState('loading',`Loading ${getProjectConfig().label}`);
  renderScenarioLibrary();
  initEquityModelModule();
  renderEquityModel();
  renderAssumptionsPanel();
  renderSourcesPanel();
  renderImportAudit();
  renderExcelChecklist();
  renderTemplateCompliance();
  updateFooterCurrency();
  if(announce)showToast(`📁 ${getProjectConfig().title}`,'');
  await loadProjectWorkbook();
}

/* ===== FINANCIAL ENGINE ===== */
function computeModelSnapshot(sourceState=STATE){
  const input={...STATE,...sourceState};
  const snapshot={revenue:[],ebitda:[],netProfit:[],cf:[],cumCF:[],discCF:[],discCumCF:[],margins:[],npv:0,irr5:0,irr10:0,payback:null,exitEV5:0,moic5:0,exitIRR5:0};
  const rm=1+input.revGrowth/100;
  for(let y=0;y<10;y++){
    const revenue=BASE.revenue[y]*rm;
    const adjustedOpex=BASE.opexRate[y]*(1+input.opexAdj/100);
    const staffDelta=0.03*input.staffAdj/100;
    const margin=Math.max(1-(input.revShare/100)-adjustedOpex-staffDelta,-0.5);
    const ebitda=revenue*margin;
    snapshot.revenue[y]=revenue;
    snapshot.ebitda[y]=ebitda;
    snapshot.margins[y]=margin;
    snapshot.netProfit[y]=ebitda*BASE.npRatio[y];
  }
  snapshot.cf=[-BASE.investment];
  snapshot.cumCF=[-BASE.investment];
  let cumulative=-BASE.investment;
  for(let y=0;y<10;y++){
    const cf=snapshot.ebitda[y]*BASE.cfRatio[y];
    snapshot.cf.push(cf);
    cumulative+=cf;
    snapshot.cumCF.push(cumulative);
  }
  const discountExponent=t=>t===0?0:t-0.5; // Workbook DCF uses mid-year convention.
  snapshot.npv=snapshot.cf.reduce((sum,cf,t)=>sum+cf/Math.pow(1+input.wacc/100,discountExponent(t)),0);
  let discountedCumulative=0;
  snapshot.cf.forEach((cf,t)=>{
    const discounted=cf/Math.pow(1+input.wacc/100,discountExponent(t));
    snapshot.discCF[t]=discounted;
    discountedCumulative+=discounted;
    snapshot.discCumCF[t]=discountedCumulative;
  });
  snapshot.irr10=calcIRR(snapshot.cf);
  snapshot.irr5=calcIRR(snapshot.cf.slice(0,6));
  snapshot.payback=calcPayback(snapshot.cf);
  snapshot.exitEV5=snapshot.ebitda[4]*input.exitMult;
  const exitCF=[...snapshot.cf.slice(0,6)];
  const netExit5=calcCompanyNetExit(snapshot.exitEV5).net;
  exitCF[5]+=netExit5;
  snapshot.exitIRR5=calcIRR(exitCF);
  const totalReturn=snapshot.cumCF[5]+netExit5;
  snapshot.moic5=totalReturn/BASE.investment;
  return snapshot;
}
function recalculate(){
  ADJ=computeModelSnapshot(STATE);
}
function calcIRR(cfs){
  let r=0.2;
  for(let i=0;i<200;i++){
    let npv=0,dnpv=0;
    cfs.forEach((c,t)=>{const d=Math.pow(1+r,t);npv+=c/d;dnpv-=t*c/Math.pow(1+r,t+1);});
    if(Math.abs(npv)<100)return r;
    if(Math.abs(dnpv)<1e-10)break;
    r-=npv/dnpv;
    if(r<-0.5||r>10)return NaN;
  }
  return r;
}

/* ===== UTILS ===== */
function fmt(n){if(Math.abs(n)>=1e6)return(n/1e6).toFixed(1)+'M';if(Math.abs(n)>=1e3)return(n/1e3).toFixed(0)+'K';return Math.round(n).toLocaleString();}
function fmtM(n){return(n/1e6).toFixed(1);}
function pct(n){return(n*100).toFixed(1)+'%';}
function fmtCount(n){return Math.round(Number(n)||0).toLocaleString();}
function fmtAuditValue(kind,value){
  const num=Number(value)||0;
  if(kind==='count')return fmtCount(num);
  if(kind==='pct')return (num*100).toFixed(1)+'%';
  if(kind==='multiple')return num.toFixed(2)+'×';
  if(kind==='months')return num.toFixed(2)+' mo';
  if(Math.abs(num)>=1e6)return (num/1e6).toFixed(3)+'M AED';
  if(Math.abs(num)>=1e3)return Math.round(num).toLocaleString()+' AED';
  return num.toFixed(2)+' AED';
}
function fmtAuditDelta(kind,delta){
  const num=Number(delta)||0;
  const sign=num>0?'+':'';
  if(kind==='count')return sign+fmtCount(num);
  if(kind==='pct')return sign+(num*100).toFixed(2)+'pp';
  if(kind==='multiple')return sign+num.toFixed(3)+'×';
  if(kind==='months')return sign+num.toFixed(2)+' mo';
  if(Math.abs(num)>=1e6)return sign+(num/1e6).toFixed(3)+'M';
  return sign+Math.round(num).toLocaleString();
}
function calcChecklistNpv(cashFlows,wacc=SCENARIOS.base.wacc){
  return (cashFlows||[]).reduce((sum,cf,t)=>sum+(Number(cf)||0)/Math.pow(1+wacc/100,t===0?0:Math.max(t-.5,0)),0);
}
function getChecklistStatus(delta,tolerance=1){
  const abs=Math.abs(Number(delta)||0);
  if(abs<=tolerance)return'ok';
  if(abs<=tolerance*10)return'warn';
  return'fail';
}
function makeChecklistRecord({metric,source,kind,workbookValue,dashboardValue,tolerance=1}){
  const delta=(Number(dashboardValue)||0)-(Number(workbookValue)||0);
  return{
    metric,
    source,
    kind,
    workbookValue:Number(workbookValue)||0,
    dashboardValue:Number(dashboardValue)||0,
    delta,
    tolerance,
    status:getChecklistStatus(delta,tolerance)
  };
}
function buildExcelChecklistPayload(raw={},filename='Workbook'){
  const checks=[];
  const hasSeriesValue=(series,yearIdx)=>Array.isArray(series)&&series[yearIdx]!==null&&series[yearIdx]!==undefined&&!Number.isNaN(Number(series[yearIdx]));
  const serviceRevenueTotalAt=yearIdx=>(raw.serviceRevenue||[]).reduce((sum,row)=>sum+(Number(row?.[yearIdx])||0),0);
  const revenueBridgeY5=(Number(raw.revenue?.[4])||0)+(Number(raw.cogs?.[4])||0)+(Number(raw.marketingCost?.[4])||0)+(Number(raw.salary?.[4])||0)+(Number(raw.fixedStaff?.[4])||0)+(Number(raw.admin?.[4])||0)+(Number(raw.rd?.[4])||0)+(Number(raw.rent?.[4])||0);
  const annualizedMonthlyRevenueY1=(raw.monthlyRevenue||[]).reduce((sum,val)=>sum+(Number(val)||0),0);
  const totalInvestmentRebuild=raw.totalCapex!==null&&raw.totalCapex!==undefined
    ? (Number(raw.totalCapex)||0)
    : null;
  const rebuiltNpv=(raw.annCF||[]).length?calcChecklistNpv(raw.annCF,raw.wacc||SCENARIOS.base.wacc):null;
  const uniquePatientsY5=(Number(raw.retained?.[4])||0)+(Number(raw.referral?.[4])||0)+(Number(raw.walkin?.[4])||0)+(Number(raw.marketing?.[4])||0);
  const totalVisitsY5=hasSeriesValue(raw.retainedVisits,4)&&hasSeriesValue(raw.organicVisits,4)&&hasSeriesValue(raw.paidVisits,4)
    ? (Number(raw.retainedVisits?.[4])||0)+(Number(raw.organicVisits?.[4])||0)+(Number(raw.paidVisits?.[4])||0)
    : null;
  const pureCacY5=(Number(raw.acquisitionSpend?.[4])||0)>0&&(Number(raw.marketing?.[4])||0)>0
    ? (Number(raw.acquisitionSpend?.[4])||0)/(Number(raw.marketing?.[4])||1)
    : null;
  const blendedCacY5=(Number(raw.marketingBudget?.[4])||0)>0&&((Number(raw.referral?.[4])||0)+(Number(raw.walkin?.[4])||0)+(Number(raw.marketing?.[4])||0))>0
    ? (Number(raw.marketingBudget?.[4])||0)/((Number(raw.referral?.[4])||0)+(Number(raw.walkin?.[4])||0)+(Number(raw.marketing?.[4])||0))
    : null;
  const ltvY5=hasSeriesValue(raw.avgRevenuePerPatient,4)&&hasSeriesValue(raw.lifetimeYears,4)
    ? (Number(raw.avgRevenuePerPatient?.[4])||0)*(Number(raw.lifetimeYears?.[4])||0)
    : null;
  const ltvCacY5=ltvY5!==null&&pureCacY5!==null&&pureCacY5>0
    ? ltvY5/pureCacY5
    : null;
  const pushCheck=config=>{
    if(config.workbookValue===null||config.workbookValue===undefined||Number.isNaN(Number(config.workbookValue)))return;
    if(config.dashboardValue===null||config.dashboardValue===undefined||Number.isNaN(Number(config.dashboardValue)))return;
    checks.push(makeChecklistRecord(config));
  };

  pushCheck({metric:'Revenue Y1',source:'P&L → Revenue vs Revenue sheet departments',kind:'amount',workbookValue:raw.revenue?.[0],dashboardValue:serviceRevenueTotalAt(0),tolerance:1});
  pushCheck({metric:'Revenue Y5',source:'P&L → Revenue vs Revenue sheet departments',kind:'amount',workbookValue:raw.revenue?.[4],dashboardValue:serviceRevenueTotalAt(4),tolerance:1});
  pushCheck({metric:'Revenue Y10',source:'P&L → Revenue vs Revenue sheet departments',kind:'amount',workbookValue:raw.revenue?.[9],dashboardValue:serviceRevenueTotalAt(9),tolerance:1});
  pushCheck({metric:'EBITDA Y5',source:'P&L → Revenue + cost lines',kind:'amount',workbookValue:raw.ebitda?.[4],dashboardValue:revenueBridgeY5,tolerance:1});
  pushCheck({metric:'Revenue Y1 vs Monthly P&L',source:'P&L Revenue vs sum of Monthly P&L revenue',kind:'amount',workbookValue:raw.revenue?.[0],dashboardValue:annualizedMonthlyRevenueY1,tolerance:50000});
  pushCheck({metric:'Capital Investment',source:'Cash Flow & DCF vs CapEx total',kind:'amount',workbookValue:raw.capitalInvestment,dashboardValue:raw.totalCapex,tolerance:1});
  pushCheck({metric:'Total Investment',source:'Cash Flow & DCF vs CapEx total (working capital already included)',kind:'amount',workbookValue:raw.totalInvestment,dashboardValue:totalInvestmentRebuild,tolerance:1});
  pushCheck({metric:`10Y NPV @ ${(raw.wacc||SCENARIOS.base.wacc).toFixed(0)}%`,source:'Cash Flow & DCF → NPV vs rebuilt from operating cash flow',kind:'amount',workbookValue:raw.npv10,dashboardValue:rebuiltNpv,tolerance:100});
  pushCheck({metric:'Unique Patients Y5',source:'Patient Acquisition → channel bridge',kind:'count',workbookValue:raw.uniquePatients?.[4],dashboardValue:uniquePatientsY5,tolerance:1});
  pushCheck({metric:'Total Visits Y5',source:'Patient Acquisition → visit-source bridge',kind:'count',workbookValue:raw.totalVisits?.[4],dashboardValue:totalVisitsY5,tolerance:1});
  pushCheck({metric:'Pure CAC Y5',source:'Patient Acquisition → acquisition spend / paid patients',kind:'amount',workbookValue:raw.pureCAC?.[4],dashboardValue:pureCacY5,tolerance:1});
  pushCheck({metric:'Blended CAC Y5',source:'Patient Acquisition → total marketing / all new patients',kind:'amount',workbookValue:raw.blendedCAC?.[4],dashboardValue:blendedCacY5,tolerance:1});
  pushCheck({metric:'Patient LTV Y5',source:'Patient Acquisition → avg revenue per patient × lifetime',kind:'amount',workbookValue:raw.ltv?.[4],dashboardValue:ltvY5,tolerance:1});
  pushCheck({metric:'LTV:CAC Y5',source:'Patient Acquisition → LTV / pure CAC',kind:'multiple',workbookValue:raw.ltvCac?.[4],dashboardValue:ltvCacY5,tolerance:.01});

  excelChecklist={
    fileName:filename||'Workbook',
    wacc:raw.wacc||SCENARIOS.base.wacc,
    checks
  };
}
function isBaseScenarioActive(){
  return Object.entries(SCENARIOS.base).every(([key,val])=>Math.abs((Number(STATE[key])||0)-Number(val))<1e-9);
}
function fmtCx(n){return fmt(cx(n));}
function fmtMCx(n){return fmtM(cx(n));}
function clamp(n,min,max){return Math.min(Math.max(n,min),max);}
function randBetween(min,max){return min+Math.random()*(max-min);}
function equalitySum(arr){return arr.reduce((sum,val)=>sum+(Number(val)||0),0);}
function calcPayback(cfs){
  let cumulative=0;
  for(let t=0;t<cfs.length;t++){
    cumulative+=cfs[t];
    if(cumulative>=0){
      const previous=cumulative-cfs[t];
      return t>0?(t-1+Math.abs(previous)/(cfs[t]||1)):0;
    }
  }
  return null;
}
function calcMonthlyPayback(cfs){
  const annualPayback=calcPayback(cfs);
  return annualPayback===null?null:annualPayback*12;
}
function percentileFromSorted(sorted,p){
  if(!sorted.length)return NaN;
  const idx=clamp(Math.floor((sorted.length-1)*p),0,sorted.length-1);
  return sorted[idx];
}
function formatProb(value){
  if(!Number.isFinite(value))return'—';
  const pctValue=value*100;
  return pctValue>=10?pctValue.toFixed(0)+'%':pctValue.toFixed(1)+'%';
}
function getRiskTone(probability){
  if(probability>=0.30)return'high';
  if(probability>=0.12)return'medium';
  return'low';
}
function normalizeTemplateLabel(value){
  return String(value??'').trim().toLowerCase().replace(/[—–]/g,'-').replace(/\s+/g,' ');
}
function findTemplateRow(sheet,label,mode='starts'){
  const target=normalizeTemplateLabel(label);
  if(!sheet)return -1;
  for(let r=0;r<sheet.length;r++){
    const cell=normalizeTemplateLabel(sheet[r]?.[0]);
    if(mode==='contains'&&cell.includes(target))return r;
    if(mode==='starts'&&cell.startsWith(target))return r;
  }
  return -1;
}
function findAnnualVisitsAfter(sheet,label){
  const start=findTemplateRow(sheet,label,'starts');
  if(start<0)return -1;
  for(let r=start+1;r<Math.min(start+6,sheet.length);r++){
    if(normalizeTemplateLabel(sheet[r]?.[0]).includes('annual visits'))return r;
  }
  return -1;
}
function countPriceListRows(sheet){
  if(!sheet)return 0;
  let currentDepartment='';
  let count=0;
  for(let r=0;r<sheet.length;r++){
    const row=sheet[r]||[];
    const first=String(row[0]||'').trim();
    const second=String(row[1]||'').trim();
    if(SERVICE_REVENUE_ROWS.includes(first)){
      currentDepartment=first;
      continue;
    }
    if(!currentDepartment||!second)continue;
    const basePrice=Math.max(0,Number(row[2])||0);
    if(basePrice>0)count++;
  }
  return count;
}
function evaluateTemplateMatch(sheet,match){
  if(!sheet||!match)return false;
  if(match.type==='starts')return findTemplateRow(sheet,match.label,'starts')>=0;
  if(match.type==='contains')return findTemplateRow(sheet,match.label,'contains')>=0;
  if(match.type==='any')return (match.labels||[]).some(label=>findTemplateRow(sheet,label,'starts')>=0);
  if(match.type==='annualVisitsAfter')return findAnnualVisitsAfter(sheet,match.label)>=0;
  if(match.type==='priceListRows')return countPriceListRows(sheet)>0;
  return false;
}
function makeTemplateCheck({area,item,severity='warn',ok=false,detail=''}) {
  return {
    area,
    item,
    severity,
    ok,
    status:ok?'ok':severity==='fail'?'fail':'warn',
    detail
  };
}
function assessWorkbookTemplate(wb,filename='Workbook'){
  const checks=[];
  const sheetPresence=new Map();

  TEMPLATE_SHEET_RULES.forEach(rule=>{
    const present=Boolean(wb?.Sheets?.[rule.sheet]);
    sheetPresence.set(rule.sheet,present);
    checks.push(makeTemplateCheck({
      area:'Sheet',
      item:rule.sheet,
      severity:rule.severity,
      ok:present,
      detail:present?rule.note:`Missing sheet. ${rule.note}`
    }));
  });

  TEMPLATE_ROW_RULES.forEach(rule=>{
    const sheet=getSheet(wb,rule.sheet);
    const sheetPresent=sheetPresence.get(rule.sheet);
    if(!sheetPresent)return;
    const ok=evaluateTemplateMatch(sheet,rule.match);
    checks.push(makeTemplateCheck({
      area:rule.sheet,
      item:rule.item,
      severity:rule.severity,
      ok,
      detail:ok?rule.note:`Missing template row. ${rule.note}`
    }));
  });

  const summary=checks.reduce((acc,check)=>{
    if(check.ok)acc.ok++;
    else acc[check.severity]=(acc[check.severity]||0)+1;
    return acc;
  },{ok:0,fail:0,warn:0,optional:0});
  const ready=(summary.fail||0)===0;
  return{
    fileName:filename,
    checkedAt:new Date().toISOString(),
    ready,
    summary,
    checks
  };
}

/* ===== CHART.JS DEFAULTS ===== */
Chart.defaults.color='#94a3b8';
Chart.defaults.font.family="'Inter',sans-serif";
Chart.defaults.font.size=10;
Chart.defaults.plugins.legend.labels.usePointStyle=true;
Chart.defaults.plugins.legend.labels.pointStyleWidth=7;
Chart.defaults.plugins.tooltip.backgroundColor='rgba(17,24,39,.95)';
Chart.defaults.plugins.tooltip.borderColor='rgba(99,102,241,.3)';
Chart.defaults.plugins.tooltip.borderWidth=1;
Chart.defaults.plugins.tooltip.cornerRadius=8;
Chart.defaults.plugins.tooltip.padding=8;
Chart.defaults.scale.grid.color='rgba(255,255,255,.04)';
Chart.defaults.scale.border.color='rgba(255,255,255,.06)';
const C={indigo:'#6366f1',purple:'#a855f7',cyan:'#22d3ee',emerald:'#10b981',amber:'#f59e0b',rose:'#f43f5e',blue:'#3b82f6',gray:'#64748b'};

/* ===== CHARTS REGISTRY ===== */
const charts={};

/* ===== UPDATE ALL ===== */
function updateAll(){
  recalculate();
  updateKPIs();
  updateLiveMetrics();
  updateActiveCharts();
  renderYearSelection(false);
  updateFooterCurrency();
  markMonteCarloStale(true);
  renderScenarioLibrary();
  renderServiceLineEconomics();
  renderProductRevenueModule();
  updateRevenueMixCharts();
  renderPatientAcquisitionModel();
  renderPatientBaseBridge();
  renderFundingMetrics();
  renderEquityModel();
  renderAssumptionsPanel();
  renderSourcesPanel();
  renderTemplateCompliance();
  renderImportAudit();
  renderExcelChecklist();
  renderActualVsPlan();
}

function updateKPIs(){
  setText('kpi-invest',fmtMCx(BASE.investment)+'M');
  setText('kpi-irr5',pct(ADJ.irr5));
  setText('kpi-irr10',pct(ADJ.irr10));
  setText('kpi-irr5-sub',`vs ${STATE.wacc}% WACC`);
  setText('kpi-payback',formatPaybackValue(ADJ.payback));
  setText('kpi-exitIrr',isNaN(ADJ.exitIRR5)?'N/A':pct(ADJ.exitIRR5));
  setText('kpi-moic',ADJ.moic5.toFixed(2)+'×');
  setText('kpi-margin5',(ADJ.margins[4]*100).toFixed(1)+'%');
  setText('kpi-rev10',fmtMCx(ADJ.revenue[9])+'M');
  setText('kpi-np10',fmtMCx(ADJ.netProfit[9])+'M');
  const investorSnap=getEquityModelSnapshot();
  setText('kpi-investorNetGain',fmtMCx(investorSnap.totalReturnSelected)+'M');
  const y1PatientEconomics=getPatientAcquisitionSnapshot('0');
  const baselineLtvCac=Number(BASE.paModel?.ltvCac?.[0])||0;
  const kpiLtvCac=isBaseScenarioActive()&&baselineLtvCac>0?baselineLtvCac:y1PatientEconomics.ltvCac;
  setText('kpi-ltvCac',kpiLtvCac?`${kpiLtvCac.toFixed(1)}×`:'—');
  const breakevenMonth=getOperationalBreakevenMonth();
  setText('kpi-operationalBreakeven',breakevenMonth
    ? (currentLanguage==='ru'?`Месяц ${breakevenMonth} (Y1)`:`Month ${breakevenMonth} (Y1)`)
    : '—');
  const verdict=document.getElementById('verdictText');
  if(verdict){
    verdict.innerHTML=currentLanguage==='ru'
      ? `<strong>${ADJ.irr5>=STATE.wacc/100?'Инвестиционный кейс устойчив.':'Кейс требует пересмотра допущений.'}</strong> 5-летний IRR ${pct(ADJ.irr5)} против WACC ${STATE.wacc}%. Payback ${formatPaybackValue(ADJ.payback)}.`
      : `<strong>${ADJ.irr5>=STATE.wacc/100?'The investment case remains defensible.':'The case needs assumption review.'}</strong> 5-year IRR is ${pct(ADJ.irr5)} versus ${STATE.wacc}% WACC. Payback is ${formatPaybackValue(ADJ.payback)}.`;
  }
  setText('riskCashText',currentLanguage==='ru'
    ? `Y1 EBITDA: ${fmtCx(ADJ.ebitda[0])}. Общая потребность: ${fmtCx(BASE.investment)}.`
    : `Y1 EBITDA: ${fmtCx(ADJ.ebitda[0])}. Total need: ${fmtCx(BASE.investment)}.`);
  const workingCapitalIncluded=getCapexAmountByLabels(['Working Capital']);
  setText('riskCashMitigation',currentLanguage==='ru'
    ? `Буфер working capital ${fmtCx(workingCapitalIncluded)} уже включен в CapEx plan.`
    : `Working capital buffer of ${fmtCx(workingCapitalIncluded)} is already included in the capex plan.`);
  // Update capex badge
  const capBadge=document.querySelector('.chart-badge');
  if(capBadge&&capBadge.closest('#section-overview'))capBadge.textContent=fmtCx(BASE.investment)+' '+cxLabel();
}

function updateLiveMetrics(){
  setLM('lm-npv',fmtCx(ADJ.npv),ADJ.npv>=0);
  setLM('lm-irr',pct(ADJ.irr10),ADJ.irr10>STATE.wacc/100);
  setLM('lm-ebitda',fmtCx(ADJ.ebitda[4]),ADJ.ebitda[4]>0);
  setLM('lm-payback',formatPaybackValue(ADJ.payback),ADJ.payback!==null&&ADJ.payback<5);
  setLM('lm-exitEV',fmtCx(ADJ.exitEV5),true);
  setLM('lm-moic',ADJ.moic5.toFixed(2)+'×',ADJ.moic5>1);
}

function setText(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}
function escapeAttr(v){
  return String(v).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function setLM(id,v,pos){const e=document.getElementById(id);if(!e)return;e.textContent=v;e.className='lm-val '+(pos?'positive':'negative');}

function getCapexAmountByLabels(labels=[]){
  return labels.reduce((sum,label)=>{
    const idx=BASE.capexLabels.indexOf(label);
    return sum+(idx>=0?(Number(BASE.capexAmounts[idx])||0):0);
  },0);
}

function getCapexOverviewGroups(){
  const groups=[
    {label:'Renovation & Fit-out',items:['Renovation & Fit-out'],color:C.indigo},
    {label:'Pre-operational',items:['Design / MEP / Engineering','Licensing / Legal / Hiring / Launch'],color:C.purple},
    {label:'Physio Equipment',items:['Physio Equipment'],color:C.cyan},
    {label:'Dental Equipment',items:['Dental Equipment'],color:C.emerald},
    {label:'Cosmetology & Beauty',items:['Cosmetology & Beauty'],color:C.amber},
    {label:'Assessment + Lab + Office/IT',items:['Assessment Equipment','Laboratory & Medical Equipment','Office & IT'],color:C.blue},
    {label:'Working Capital + Deposit + Inventory',items:['Rent Deposit','Initial Inventory','Pre-opening Salaries','Working Capital'],color:C.rose},
    {label:'Contingency',items:['Contingency'],color:C.gray}
  ];
  return groups.map(group=>({
    ...group,
    amount:getCapexAmountByLabels(group.items)
  })).filter(group=>group.amount>0);
}

function getOperationalBreakevenMonth(){
  const monthlyScale=BASE.ebitda[0]!==0?(ADJ.ebitda[0]/BASE.ebitda[0]):1;
  const monthlySeries=BASE.mEbitda.map(v=>v*monthlyScale);
  const breakevenIndex=monthlySeries.findIndex(v=>v>=0);
  return breakevenIndex>=0?breakevenIndex+1:BASE.operationalBreakevenMonth;
}

function getScenarioPatientFlowSeries(){
  const revenueScale=BASE.revenue.map((revenue,idx)=>{
    const baseRevenue=Math.max(0,Number(revenue)||0);
    return baseRevenue>0?Math.max((Number(ADJ.revenue[idx])||0)/baseRevenue,0):1;
  });
  const series={
    scale:[],
    effectiveScale:[],
    openingBase:[],
    uniquePatients:[],
    visits:[],
    util:[],
    retained:[],
    referral:[],
    walkin:[],
    marketing:[],
    retainedVisits:[],
    organicVisits:[],
    paidVisits:[],
    patientsShort:[],
    visitsShort:[]
  };
  let priorClosing=0;
  for(let i=0;i<10;i++){
    const baseVisits=Math.max(0,Number(BASE.visits[i])||0);
    const capacity=Math.max(0,Number(BASE.capacity[i])||0);
    const rawScale=revenueScale[i];
    const effectiveScale=baseVisits>0&&capacity>0?Math.min(rawScale,capacity/baseVisits):rawScale;
    const safeScale=Math.max(effectiveScale,0);

    const baseOpening=i===0?0:Math.max(0,Number(BASE.paModel.uniquePatients?.[i-1])||0);
    const baseClosing=Math.max(0,Number(BASE.paModel.uniquePatients?.[i])||0);
    const baseRetained=Math.max(0,Number(BASE.retained[i])||0);
    const baseReferral=Math.max(0,Number(BASE.referral[i])||0);
    const baseWalkin=Math.max(0,Number(BASE.walkin[i])||0);
    const baseMarketing=Math.max(0,Number(BASE.marketing[i])||0);
    const baseOrganic=baseReferral+baseWalkin;
    const baseNewTotal=baseOrganic+baseMarketing;

    const opening=i===0?0:priorClosing;
    const retentionRate=baseOpening>0?clamp(baseRetained/baseOpening,0,1):0;
    const retained=Math.round(opening*retentionRate);
    const targetClosing=Math.max(0,Math.round(baseClosing*safeScale));
    const targetNewTotal=Math.max(targetClosing-retained,0);

    const paidShare=baseNewTotal>0?baseMarketing/baseNewTotal:0;
    const organicTarget=Math.max(targetNewTotal-Math.round(targetNewTotal*paidShare),0);
    let marketing=Math.round(targetNewTotal*paidShare);
    let walkin=Math.round(organicTarget*(baseOrganic>0?baseWalkin/baseOrganic:0.5));
    let referral=Math.max(targetNewTotal-marketing-walkin,0);
    let closing=retained+marketing+walkin+referral;
    const closingGap=targetClosing-closing;
    if(closingGap!==0){
      referral=Math.max(referral+closingGap,0);
      closing=retained+marketing+walkin+referral;
    }

    const targetVisits=Math.max(0,Math.round(capacity>0?Math.min(baseVisits*safeScale,capacity):baseVisits*safeScale));
    const baseRetainedVisits=Math.max(0,Number(BASE.paModel.retainedVisits?.[i])||0);
    const baseOrganicVisits=Math.max(0,Number(BASE.paModel.organicVisits?.[i])||0);
    const basePaidVisits=Math.max(0,Number(BASE.paModel.paidVisits?.[i])||0);
    const retainedVpp=baseRetained>0?baseRetainedVisits/baseRetained:1;
    const organicVpp=baseOrganic>0?baseOrganicVisits/baseOrganic:1;
    const paidVpp=baseMarketing>0?basePaidVisits/baseMarketing:1;
    const rawRetainedVisits=retained*retainedVpp;
    const rawOrganicVisits=(referral+walkin)*organicVpp;
    const rawPaidVisits=marketing*paidVpp;
    const rawVisitTotal=rawRetainedVisits+rawOrganicVisits+rawPaidVisits;
    const visitScale=rawVisitTotal>0?targetVisits/rawVisitTotal:0;
    let retainedVisits=Math.round(rawRetainedVisits*visitScale);
    let organicVisits=Math.round(rawOrganicVisits*visitScale);
    let paidVisits=Math.round(rawPaidVisits*visitScale);
    let visits=retainedVisits+organicVisits+paidVisits;
    const visitGap=targetVisits-visits;
    if(visitGap!==0){
      if(paidVisits>0||marketing>0)paidVisits=Math.max(paidVisits+visitGap,0);
      else if(organicVisits>0||walkin>0||referral>0)organicVisits=Math.max(organicVisits+visitGap,0);
      else retainedVisits=Math.max(retainedVisits+visitGap,0);
      visits=retainedVisits+organicVisits+paidVisits;
    }

    series.scale[i]=rawScale;
    series.effectiveScale[i]=safeScale;
    series.openingBase[i]=opening;
    series.uniquePatients[i]=closing;
    series.visits[i]=visits;
    series.util[i]=capacity>0?(visits/capacity)*100:0;
    series.retained[i]=retained;
    series.referral[i]=referral;
    series.walkin[i]=walkin;
    series.marketing[i]=marketing;
    series.retainedVisits[i]=retainedVisits;
    series.organicVisits[i]=organicVisits;
    series.paidVisits[i]=paidVisits;
    priorClosing=closing;
  }
  series.patientsShort=BASE.plIdx.map(i=>Math.round(series.uniquePatients[i]||0));
  series.visitsShort=BASE.plIdx.map(i=>Math.round(series.visits[i]||0));
  return series;
}

function updateActiveCharts(){
  const sec=document.querySelector('.dashboard-section.active');
  if(!sec)return;
  const id=sec.id.replace('section-','');
  if(id==='overview')updateOverviewCharts();
  if(id==='revenue')updateRevenueCharts();
  if(id==='patients'){renderPatientAcquisitionModel();renderPatientBaseBridge();updatePatientAcquisitionCharts();}
  if(id==='cashflow')updateCashflowCharts();
  if(id==='sensitivity')updateSensitivityCharts();
}

function flushChartsWithin(root){
  if(!root||typeof Chart==='undefined'||typeof Chart.getChart!=='function')return;
  root.querySelectorAll('canvas').forEach(canvas=>{
    const chart=Chart.getChart(canvas);
    if(chart){
      chart.resize();
      chart.update('none');
    }
  });
}
function resizeChartsWithin(root){
  if(!root||typeof Chart==='undefined'||typeof Chart.getChart!=='function')return;
  const resize=()=>root.querySelectorAll('canvas').forEach(canvas=>{
    const chart=Chart.getChart(canvas);
    if(chart){
      chart.resize();
      chart.update('none');
    }
  });
  requestAnimationFrame(resize);
  [80,220,420].forEach(delay=>setTimeout(resize,delay));
}
function hydrateSection(section){
  const sec=typeof section==='string'?document.getElementById(`section-${section}`):section;
  if(!sec)return null;
  const id=sec.id.replace('section-','');
  sec.querySelectorAll('[data-animate]').forEach(el=>el.classList.add('visible'));

  if(id==='overview'){
    initOverviewCharts();
    updateOverviewCharts();
    initBenchmarks();
    initChecks();
  }
  if(id==='revenue'){
    initRevenueCharts();
    updateRevenueCharts();
    initRevenueMixCharts();
    initMonthlyDrilldown();
  }
  if(id==='patients'){
    renderPatientAcquisitionModel();
    renderPatientBaseBridge();
    initPatientEngineViewTabs();
    initPatientTrendControls();
    initPatientAcquisitionCharts();
    updatePatientAcquisitionCharts();
  }
  if(id==='cashflow'){
    renderFundingMetrics();
    renderEquityModel();
    initCashflowCharts();
    updateCashflowCharts();
  }
  if(id==='capex')initCapexCharts();
  if(id==='sensitivity'){
    renderScenarioLibrary();
    initSensitivityCharts();
    initMonteCarloCharts();
  }
  if(id==='risks'){
    renderAssumptionsPanel();
    renderSourcesPanel();
    renderTemplateCompliance();
    renderImportAudit();
    renderExcelChecklist();
    renderActualVsPlan();
    initRatioCharts();
  }

  flushChartsWithin(sec);
  resizeChartsWithin(sec);
  return sec;
}
function getFocusCardTitle(card){
  return card?.querySelector('.chart-header h3')?.textContent?.trim()||'Chart';
}
function setFocusOverlayOpen(isOpen){
  const overlay=document.getElementById('cardFocusOverlay');
  if(!overlay)return;
  clearTimeout(FOCUS_STATE.hideTimer);
  if(isOpen){
    overlay.hidden=false;
    overlay.setAttribute('aria-hidden','false');
    requestAnimationFrame(()=>overlay.classList.add('is-open'));
    return;
  }
  overlay.classList.remove('is-open');
  overlay.setAttribute('aria-hidden','true');
  FOCUS_STATE.hideTimer=setTimeout(()=>{
    if(!FOCUS_STATE.card)overlay.hidden=true;
  },180);
}
function openChartFocus(card,trigger=null){
  const overlay=document.getElementById('cardFocusOverlay');
  const body=document.getElementById('cardFocusBody');
  const title=document.getElementById('cardFocusTitle');
  const closeBtn=document.getElementById('cardFocusCloseBtn');
  if(!overlay||!body||!title||!card)return;
  if(FOCUS_STATE.card===card)return;
  if(FOCUS_STATE.card)closeChartFocus(false);

  const placeholder=document.createElement('div');
  placeholder.className='chart-focus-placeholder';
  placeholder.style.height=card.getBoundingClientRect().height+'px';
  card.parentNode.insertBefore(placeholder,card);

  FOCUS_STATE.card=card;
  FOCUS_STATE.placeholder=placeholder;
  FOCUS_STATE.trigger=trigger;
  title.textContent=getFocusCardTitle(card);
  card.classList.add('is-focus-active');
  body.scrollTop=0;
  body.appendChild(card);
  document.body.classList.add('focus-mode-open');
  setFocusOverlayOpen(true);
  closeBtn?.focus();
  resizeChartsWithin(card);
}
function closeChartFocus(restoreFocus=true){
  const body=document.getElementById('cardFocusBody');
  const {card,placeholder,trigger}=FOCUS_STATE;
  if(card&&placeholder){
    placeholder.replaceWith(card);
    card.classList.remove('is-focus-active');
    resizeChartsWithin(card);
    setTimeout(()=>resizeChartsWithin(card),180);
  }
  if(body)body.scrollTop=0;
  FOCUS_STATE.card=null;
  FOCUS_STATE.placeholder=null;
  FOCUS_STATE.trigger=null;
  document.body.classList.remove('focus-mode-open');
  setFocusOverlayOpen(false);
  if(restoreFocus&&trigger instanceof HTMLElement&&trigger.isConnected)trigger.focus();
}
function initCardFocusMode(){
  const overlay=document.getElementById('cardFocusOverlay');
  const closeBtn=document.getElementById('cardFocusCloseBtn');
  const backdrop=document.getElementById('cardFocusBackdrop');
  if(!overlay)return;

  if(overlay.dataset.bound!=='true'){
    closeBtn?.addEventListener('click',()=>closeChartFocus());
    backdrop?.addEventListener('click',()=>closeChartFocus(false));
    overlay.addEventListener('click',e=>{
      if(e.target===overlay)closeChartFocus(false);
    });
    document.addEventListener('keydown',e=>{
      if(e.key==='Escape'&&FOCUS_STATE.card){
        e.preventDefault();
        closeChartFocus();
      }
    });
    overlay.dataset.bound='true';
  }

  document.querySelectorAll('.chart-card').forEach(card=>{
    if(card.dataset.focusReady==='true')return;
    const header=card.querySelector('.chart-header');
    if(!header)return;
    const button=document.createElement('button');
    button.type='button';
    button.className='chart-focus-btn';
    button.setAttribute('aria-label',`Open ${getFocusCardTitle(card)} in focus mode`);
    button.innerHTML='<span aria-hidden="true">⤢</span>';
    button.addEventListener('click',e=>{
      e.preventDefault();
      e.stopPropagation();
      openChartFocus(card,button);
    });
    header.appendChild(button);
    card.dataset.focusReady='true';
  });
}

/* ===== OVERVIEW CHARTS ===== */
function initOverviewCharts(){
  if(charts.capex)return;
  const patientSeries=getScenarioPatientFlowSeries();
  const capexOverview=getCapexOverviewGroups();
  const capexTotal=capexOverview.reduce((sum,item)=>sum+item.amount,0);
  const cx=document.getElementById('chart-capex').getContext('2d');
  charts.capex=new Chart(cx,{
    type:'doughnut',
    data:{
      labels:capexOverview.map(item=>`${item.label} (${capexTotal>0?Math.round(item.amount/capexTotal*100):0}%)`),
      datasets:[{
        data:capexOverview.map(item=>item.amount),
        backgroundColor:capexOverview.map(item=>item.color+'cc'),
        borderColor:capexOverview.map(item=>item.color),
        borderWidth:0
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      cutout:'62%',
      plugins:{
        legend:{position:'right',labels:{font:{size:9},padding:4}},
        tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmtCx(ctx.raw)}`}}
      }
    }
  });
  const px=document.getElementById('chart-patients').getContext('2d');
  charts.patients=new Chart(px,{type:'bar',data:{labels:BASE.patientYears,datasets:[{label:'Unique Patients',data:patientSeries.patientsShort,backgroundColor:'rgba(99,102,241,.5)',borderRadius:4,order:2},{label:'Annual Visits',data:patientSeries.visitsShort,type:'line',borderColor:C.cyan,backgroundColor:'transparent',pointBackgroundColor:C.cyan,pointRadius:3,borderWidth:2,tension:.4,yAxisID:'y1',order:1}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,title:{display:true,text:'Patients'}},y1:{position:'right',beginAtZero:true,grid:{display:false}}}}});
  const cohx=document.getElementById('chart-cohort').getContext('2d');
  charts.cohort=new Chart(cohx,{type:'line',data:{labels:BASE.years,datasets:[{label:'Marketing',data:patientSeries.marketing,backgroundColor:'rgba(239,68,68,.3)',borderColor:C.rose,fill:true,pointRadius:2,borderWidth:1.5,tension:.3},{label:'Walk-in',data:patientSeries.walkin,backgroundColor:'rgba(245,158,11,.3)',borderColor:C.amber,fill:true,pointRadius:2,borderWidth:1.5,tension:.3},{label:'Referral',data:patientSeries.referral,backgroundColor:'rgba(16,185,129,.3)',borderColor:C.emerald,fill:true,pointRadius:2,borderWidth:1.5,tension:.3},{label:'Retained',data:patientSeries.retained,backgroundColor:'rgba(99,102,241,.35)',borderColor:C.indigo,fill:true,pointRadius:2,borderWidth:1.5,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index'},scales:{y:{stacked:true,beginAtZero:true}}}});
  const ux=document.getElementById('chart-util').getContext('2d');
  charts.util=new Chart(ux,{type:'bar',data:{labels:BASE.years,datasets:[{label:'Capacity',data:BASE.capacity,backgroundColor:'rgba(99,102,241,.25)',borderRadius:3,order:2},{label:'Visits',data:patientSeries.visits,backgroundColor:'rgba(168,85,247,.3)',borderRadius:3,order:1},{label:'Util %',data:patientSeries.util,type:'line',borderColor:C.amber,backgroundColor:'transparent',pointBackgroundColor:patientSeries.util.map(v=>v>80?C.rose:v>60?C.amber:C.emerald),pointRadius:3,borderWidth:2,tension:.3,yAxisID:'y1',order:0}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{y:{beginAtZero:true,ticks:{callback:v=>fmt(v)}},y1:{position:'right',beginAtZero:true,max:100,grid:{display:false}}}}});
  initBenchmarks();
}
function updateOverviewCharts(){
  const patientSeries=getScenarioPatientFlowSeries();
  if(!charts.capex||!charts.patients||!charts.cohort||!charts.util){
    initOverviewCharts();
    return;
  }
  const capexOverview=getCapexOverviewGroups();
  const capexTotal=capexOverview.reduce((sum,item)=>sum+item.amount,0);
  charts.capex.data.labels=capexOverview.map(item=>`${item.label} (${capexTotal>0?Math.round(item.amount/capexTotal*100):0}%)`);
  charts.capex.data.datasets[0].data=capexOverview.map(item=>item.amount);
  charts.capex.data.datasets[0].backgroundColor=capexOverview.map(item=>item.color+'cc');
  charts.capex.data.datasets[0].borderColor=capexOverview.map(item=>item.color);
  charts.capex.options.plugins.tooltip.callbacks.label=ctx=>`${ctx.label}: ${fmtCx(ctx.raw)}`;
  charts.capex.update('active');
  charts.patients.data.datasets[0].data=patientSeries.patientsShort;
  charts.patients.data.datasets[1].data=patientSeries.visitsShort;
  charts.patients.update('active');
  charts.cohort.data.datasets[0].data=patientSeries.marketing;
  charts.cohort.data.datasets[1].data=patientSeries.walkin;
  charts.cohort.data.datasets[2].data=patientSeries.referral;
  charts.cohort.data.datasets[3].data=patientSeries.retained;
  charts.cohort.update('active');
  charts.util.data.datasets[1].data=patientSeries.visits;
  charts.util.data.datasets[2].data=patientSeries.util;
  charts.util.data.datasets[2].pointBackgroundColor=patientSeries.util.map(v=>v>80?C.rose:v>60?C.amber:C.emerald);
  charts.util.update('active');
}
function initBenchmarks(){
  const g=document.getElementById('benchmarkGrid');
  const project=getProjectConfig();
  g.innerHTML=BASE.benchmarks.map(b=>{
    const spread=b.aw-b.mw;
    const tone=spread>6?'positive':spread<-6?'negative':'neutral';
    const deltaLabel=spread>6?`Lead +${Math.round(spread)} pts`:spread<-6?`Gap ${Math.round(spread)} pts`:'Near market';
    return `<article class="bench-item bench-${tone}">
      <div class="bench-head">
        <div class="bench-label">${b.label}</div>
        <div class="bench-delta bench-${tone}">${deltaLabel}</div>
      </div>
      <div class="bench-stats">
        <div class="bench-stat">
          <span class="bench-stat-key">${project.shortTitle}</span>
          <span class="bench-stat-val">${localizeDurationText(b.arisha)}</span>
        </div>
        <div class="bench-stat">
          <span class="bench-stat-key">${tr('marketLabel')}</span>
          <span class="bench-stat-val">${localizeDurationText(b.market)}</span>
        </div>
      </div>
      <div class="bench-bars">
        <div class="bench-bar bench-arisha" data-width="${b.aw}%"></div>
        <div class="bench-bar bench-market" data-width="${b.mw}%"></div>
      </div>
    </article>`;
  }).join('');
}

/* ===== REVENUE CHARTS ===== */
function initRevenueCharts(){
  if(charts.revenue)return;
  const adjM=BASE.plIdx.map(i=>fmtM(ADJ.revenue[i]));
  const adjEM=BASE.plIdx.map(i=>fmtM(ADJ.ebitda[i]));
  const adjNM=BASE.plIdx.map(i=>fmtM(ADJ.netProfit[i]));
  const rx=document.getElementById('chart-revenue').getContext('2d');
  charts.revenue=new Chart(rx,{type:'bar',data:{labels:BASE.plYears,datasets:[{label:'Revenue(M)',data:BASE.plIdx.map(i=>ADJ.revenue[i]/1e6),backgroundColor:'rgba(99,102,241,.5)',borderRadius:5},{label:'EBITDA(M)',data:BASE.plIdx.map(i=>ADJ.ebitda[i]/1e6),backgroundColor:'rgba(16,185,129,.5)',borderRadius:5},{label:'Net Profit(M)',data:BASE.plIdx.map(i=>ADJ.netProfit[i]/1e6),type:'line',borderColor:C.amber,backgroundColor:'transparent',pointBackgroundColor:C.amber,pointRadius:4,borderWidth:2,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{y:{title:{display:true,text:'AED (M)'}}}}});
  document.querySelectorAll('[data-chart]').forEach(t=>{
    if(t.dataset.bound)return;
    t.addEventListener('click',()=>{
    t.closest('.chart-tabs')?.querySelectorAll('[data-chart]').forEach(b=>b.classList.remove('active'));t.classList.add('active');
    if(t.dataset.chart==='revenue-margin'){
      charts.revenue.data.datasets=[{label:'EBITDA %',data:BASE.plIdx.map(i=>ADJ.margins[i]*100),borderColor:C.emerald,backgroundColor:'rgba(16,185,129,.1)',fill:true,type:'line',pointRadius:4,borderWidth:2,tension:.3},{label:'NP %',data:BASE.plIdx.map(i=>(ADJ.netProfit[i]/ADJ.revenue[i])*100),borderColor:C.amber,backgroundColor:'rgba(245,158,11,.1)',fill:true,type:'line',pointRadius:4,borderWidth:2,tension:.3}];
      charts.revenue.options.scales.y.title.text='%';
    }else{
      charts.revenue.data.datasets=[{label:'Revenue(M)',data:BASE.plIdx.map(i=>ADJ.revenue[i]/1e6),backgroundColor:'rgba(99,102,241,.5)',borderRadius:5},{label:'EBITDA(M)',data:BASE.plIdx.map(i=>ADJ.ebitda[i]/1e6),backgroundColor:'rgba(16,185,129,.5)',borderRadius:5},{label:'Net Profit(M)',data:BASE.plIdx.map(i=>ADJ.netProfit[i]/1e6),type:'line',borderColor:C.amber,backgroundColor:'transparent',pointBackgroundColor:C.amber,pointRadius:4,borderWidth:2,tension:.3}];
      charts.revenue.options.scales.y.title.text='AED (M)';
    }
    charts.revenue.update('active');
    });
    t.dataset.bound='true';
  });
  // Cost doughnut
  const pl=getCurrentPLBreakdown(getRevenueFocusYear());
  const cx=document.getElementById('chart-costs').getContext('2d');
  charts.costs=new Chart(cx,{type:'doughnut',data:{labels:['COGS + OpEx','Provider Compensation','EBITDA'],datasets:[{data:[pl.otherOpex,pl.revShare,pl.ebitda],backgroundColor:['rgba(139,92,246,.6)',C.rose,C.emerald],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{position:'right',labels:{font:{size:9}}},tooltip:{callbacks:{label:c=>{const raw=Number(c.raw)||0;const total=(c.dataset.data||[]).reduce((sum,val)=>sum+(Number(val)||0),0);const share=total>0?raw/total*100:0;return`${c.label}: ${fmtCx(raw)} ${cxLabel()} (${share.toFixed(1)}%)`;}}}}}});
  // Monthly
  const mx=document.getElementById('chart-monthly').getContext('2d');
  charts.monthly=new Chart(mx,{type:'bar',data:{labels:BASE.months,datasets:[{label:'Revenue',data:BASE.mRev,backgroundColor:'rgba(99,102,241,.4)',borderRadius:3},{label:'EBITDA',data:BASE.mEbitda,type:'line',segment:{borderColor:ctx=>ctx.p0.parsed.y<0&&ctx.p1.parsed.y<0?C.rose:ctx.p0.parsed.y>=0&&ctx.p1.parsed.y>=0?C.emerald:C.amber},pointBackgroundColor:BASE.mEbitda.map(v=>v>=0?C.emerald:C.rose),pointRadius:3,borderWidth:2,tension:.3,fill:false}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
  // Waterfall
  buildWaterfall();
  initPLTable();
  renderServiceLineEconomics();
  renderProductRevenueModule();
}
function normalizeYearSelection(selection=selectedYear){
  if(Array.isArray(selection))return selection.map(v=>Math.max(0,Math.min(9,Number(v)||0)));
  if(selection==='all')return Array.from({length:10},(_,i)=>i);
  if(typeof selection==='number')return [Math.max(0,Math.min(9,selection))];
  if(typeof selection==='string'&&selection.includes(':')){
    const [rawStart,rawEnd]=selection.split(':');
    const start=Math.max(0,Math.min(9,parseInt(rawStart,10)||0));
    const end=Math.max(0,Math.min(9,parseInt(rawEnd,10)||0));
    const from=Math.min(start,end);
    const to=Math.max(start,end);
    return Array.from({length:to-from+1},(_,i)=>from+i);
  }
  const idx=Math.max(0,Math.min(9,parseInt(selection,10)||0));
  return [idx];
}
function getSelectedYearBounds(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  return [years[0],years[years.length-1]];
}
function getSelectedYearLabel(selection=selectedYear){
  const [start,end]=getSelectedYearBounds(selection);
  if(start===0&&end===9)return'Y1-Y10';
  return start===end?`Y${start+1}`:`Y${start+1}-Y${end+1}`;
}
function getSensitivityAnchor(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  const followsFilter=years.length===1;
  const yearIdx=followsFilter?years[0]:4;
  const yearLabel=`Y${yearIdx+1}`;
  const contextLabel=followsFilter?yearLabel:'Y5 anchor';
  return {yearIdx,yearLabel,contextLabel,followsFilter};
}
function getSelectionSummary(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  const revenue=years.reduce((sum,i)=>sum+(ADJ.revenue[i]||0),0);
  const ebitda=years.reduce((sum,i)=>sum+(ADJ.ebitda[i]||0),0);
  const netProfit=years.reduce((sum,i)=>sum+(ADJ.netProfit[i]||0),0);
  const margin=revenue!==0?ebitda/revenue:0;
  return {years,revenue,ebitda,netProfit,margin};
}
function getRevenueFocusYear(){
  return selectedYear;
}
function getCurrentPLBreakdown(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  if(years.length>1){
    return years.map(i=>getCurrentPLBreakdown(i)).reduce((sum,pl)=>({
      revenue:sum.revenue+pl.revenue,
      revShare:sum.revShare+pl.revShare,
      cogs:sum.cogs+pl.cogs,
      marketing:sum.marketing+pl.marketing,
      admin:sum.admin+pl.admin,
      rent:sum.rent+pl.rent,
      staff:sum.staff+pl.staff,
      rd:sum.rd+pl.rd,
      otherOpex:sum.otherOpex+pl.otherOpex,
      ebitda:sum.ebitda+pl.ebitda
    }),{revenue:0,revShare:0,cogs:0,marketing:0,admin:0,rent:0,staff:0,rd:0,otherOpex:0,ebitda:0});
  }
  const yearIdx=years[0];
  const rm=1+STATE.revGrowth/100;
  const opexMult=1+STATE.opexAdj/100;
  const scale=rm*opexMult;
  const revenue=ADJ.revenue[yearIdx];
  const revShare=revenue*STATE.revShare/100;
  const cogs=(BASE.cogs[yearIdx]||0)*scale;
  const marketing=(BASE.plBreakdown.marketing[yearIdx]||0)*scale;
  const admin=(BASE.plBreakdown.admin[yearIdx]||0)*scale;
  const rent=(BASE.plBreakdown.rent[yearIdx]||0)*scale;
  const staffBase=(BASE.plBreakdown.staff[yearIdx]||0)*scale;
  const staff=staffBase+(revenue*0.03*STATE.staffAdj/100);
  const rd=(BASE.plBreakdown.rd[yearIdx]||0)*scale;
  const otherOpex=cogs+marketing+admin+rent+staff+rd;
  const ebitda=revenue-revShare-otherOpex;
  return {revenue,revShare,cogs,marketing,admin,rent,staff,rd,otherOpex,ebitda};
}
function getPLTableYearIdx(){
  return normalizeYearSelection();
}
function updateSensitivityContextUI(anchor=getSensitivityAnchor()){
  setText('sensitivityStressTitle',`Revenue Stress Test → EBITDA (${anchor.contextLabel})`);
  setText('sensitivityOpexStressTitle',`OPEX Stress Test → EBITDA (${anchor.contextLabel})`);
  setText('sensitivityHeatmapTitle',`Two-Way Heatmap: Revenue × OPEX (${anchor.contextLabel})`);
  setText('sensitivityStaffingTitle',`Staffing Scenarios Impact (${anchor.contextLabel})`);
  setText('sensitivitySummaryTitle',`Sensitivity Summary (${anchor.contextLabel})`);
}
function buildWaterfall(){
  if(charts.waterfall){charts.waterfall.destroy();charts.waterfall=null;}
  const focusYear=getRevenueFocusYear();
  const yearLabel=getSelectedYearLabel(focusYear);
  const pl=getCurrentPLBreakdown(focusYear);
  setText('costsChartTitle',tr('costStructure',{label:yearLabel}));
  const title='Operating Waterfall: Revenue → EBITDA ('+yearLabel+')';
  const labels=['Revenue','COGS','Provider Compensation','Marketing','Admin','Rent','Fixed Staff','R&D','EBITDA'];
  const vals=[pl.revenue,-pl.cogs,-pl.revShare,-pl.marketing,-pl.admin,-pl.rent,-pl.staff,-pl.rd,pl.ebitda];
  const isTotal=[true,false,false,false,false,false,false,false,true];
  setText('waterfallTitle',title);
  let run=0;
  const barData=vals.map((v,i)=>{if(isTotal[i]){run=v;return[0,v];}const s=run;run+=v;return v<0?[run,s]:[s,run];});
  const wx=document.getElementById('chart-waterfall').getContext('2d');
  charts.waterfall=new Chart(wx,{type:'bar',data:{labels,datasets:[{data:barData,backgroundColor:vals.map((v,i)=>isTotal[i]?'rgba(99,102,241,.5)':v<0?'rgba(244,63,94,.35)':'rgba(16,185,129,.4)'),borderColor:vals.map((v,i)=>isTotal[i]?C.indigo:v<0?C.rose:C.emerald),borderWidth:1,borderRadius:3,borderSkipped:false}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{ticks:{callback:v=>fmtCx(v)}}},plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>(vals[c.dataIndex]>=0?'+':'')+Math.round(cx(vals[c.dataIndex])).toLocaleString()+' '+cxLabel()}}}}});
}
function updateRevenueCharts(){
  if(!charts.revenue)return;
  charts.revenue.data.datasets.forEach(ds=>{
    if(ds.label==='Revenue(M)')ds.data=BASE.plIdx.map(i=>ADJ.revenue[i]/1e6);
    if(ds.label==='EBITDA(M)')ds.data=BASE.plIdx.map(i=>ADJ.ebitda[i]/1e6);
    if(ds.label==='Net Profit(M)')ds.data=BASE.plIdx.map(i=>ADJ.netProfit[i]/1e6);
    if(ds.label==='EBITDA %')ds.data=BASE.plIdx.map(i=>ADJ.margins[i]*100);
    if(ds.label==='NP %')ds.data=BASE.plIdx.map(i=>(ADJ.netProfit[i]/ADJ.revenue[i])*100);
  });
  charts.revenue.update('active');
  if(charts.costs){
    const pl=getCurrentPLBreakdown(getRevenueFocusYear());
    charts.costs.data.labels=['COGS + OpEx','Provider Compensation','EBITDA'];
    charts.costs.data.datasets[0].data=[pl.otherOpex,pl.revShare,pl.ebitda];
    charts.costs.update('active');
  }
  buildWaterfall();
  updatePLTable();
  renderServiceLineEconomics();
  renderProductRevenueModule();
  updateRevenueMixCharts();
}
const PL_ROW_DETAILS={
  ru:{
    'Revenue (M)':'Выручка по выбранным годам, с учетом текущего Revenue Growth драйвера.',
    'COGS (M)':'Прямые затраты: материалы, департаментские COGS и прочие variable clinical costs. В presentation-логике показываются до provider compensation.',
    'Provider Compensation (M)':'Переменная выплата провайдерам/врачам. Сейчас считается от revenue, как отдельный operating driver, а не как investor Revenue Share.',
    'Marketing (M)':'Маркетинг и acquisition budget, связанный с Patient Acquisition там, где данные доступны из workbook.',
    'Admin (M)':'Административные OpEx из P&L/OpEx breakdown.',
    'Rent (M)':'Аренда и premises-related operating cost.',
    'Fixed Staff (M)':'Фиксированный payroll плюс текущий Fixed Staff Cost adjustment.',
    'R&D (M)':'R&D / development-related operating cost.'
  },
  en:{
    'Revenue (M)':'Revenue across the selected years, adjusted for the current Revenue Growth driver.',
    'COGS (M)':'Direct costs: materials, department COGS and other variable clinical costs. In the presentation logic they are shown before provider compensation.',
    'Provider Compensation (M)':'Variable compensation paid to providers/doctors. It is modeled off revenue as an operating driver, not as investor revenue share.',
    'Marketing (M)':'Marketing and acquisition budget tied to Patient Acquisition where workbook data is available.',
    'Admin (M)':'Administrative OpEx from the P&L / OpEx breakdown.',
    'Rent (M)':'Rent and premises-related operating cost.',
    'Fixed Staff (M)':'Fixed payroll plus the current Fixed Staff Cost adjustment.',
    'R&D (M)':'R&D / development-related operating cost.'
  }
};
function initPLTable(){
  const yearIdxs=getPLTableYearIdx();
  const head=document.querySelector('#plTable thead tr');
  if(head)head.innerHTML=`<th>${tr('plMetric')}</th>`+yearIdxs.map(i=>`<th data-year="${i}">Y${i+1}</th>`).join('');
  const tbody=document.querySelector('#plTable tbody');tbody.innerHTML='';
  const rows=[
    {l:'Revenue (M)',v:yearIdxs.map(i=>ADJ.revenue[i]),h:false,pct:false},
    {l:'COGS (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).cogs),h:false,pct:false},
    {l:'Provider Compensation (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).revShare),h:false,pct:false},
    {l:'Marketing (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).marketing),h:false,pct:false},
    {l:'Admin (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).admin),h:false,pct:false},
    {l:'Rent (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).rent),h:false,pct:false},
    {l:'Fixed Staff (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).staff),h:false,pct:false},
    {l:'R&D (M)',v:yearIdxs.map(i=>-getCurrentPLBreakdown(i).rd),h:false,pct:false},
    {l:'EBITDA (M)',v:yearIdxs.map(i=>ADJ.ebitda[i]),h:true,pct:false},
    {l:'EBITDA %',v:yearIdxs.map(i=>ADJ.margins[i]*100),h:false,pct:true},
    {l:'Net Profit (M)',v:yearIdxs.map(i=>ADJ.netProfit[i]),h:true,pct:false},
    {l:'NP %',v:yearIdxs.map(i=>(ADJ.netProfit[i]/ADJ.revenue[i])*100),h:false,pct:true}
  ];
  rows.forEach(r=>{
    const tr=document.createElement('tr');
    if(r.h)tr.classList.add('row-highlight');
    const detail=(PL_ROW_DETAILS[currentLanguage]||PL_ROW_DETAILS.ru)[r.l]||'';
    if(detail){
      tr.classList.add('pl-row-has-detail');
      tr.title=detail;
    }
    tr.innerHTML=`<td style="font-weight:600" title="${escapeAttr(detail)}">${r.l}</td>`+r.v.map((v,idx)=>{
      const cls=v<0?'neg':v>0?'pos':'';
      const val=r.pct?v.toFixed(1)+'%':fmtMCx(v);
      return `<td class="${cls}" data-year="${yearIdxs[idx]}" title="${escapeAttr(detail)}">${val}</td>`;
    }).join('');
    tbody.appendChild(tr);
  });
}
function updatePLTable(){initPLTable();}

/* ===== CASHFLOW CHARTS ===== */
function initCashflowCharts(){
  if(charts.cashflow){
    buildDcfWaterfall();
    return;
  }
  const cfx=document.getElementById('chart-cashflow').getContext('2d');
  charts.cashflow=new Chart(cfx,{type:'line',data:{labels:['Y0',...BASE.years],datasets:[{label:'Cumulative CF',data:ADJ.cumCF,borderColor:C.indigo,backgroundColor:'rgba(99,102,241,.12)',fill:true,pointRadius:4,pointBackgroundColor:ADJ.cumCF.map(v=>v>=0?C.emerald:C.rose),borderWidth:2.5,tension:.3},{label:`Discounted CF @ ${STATE.wacc}% WACC`,data:ADJ.discCumCF,borderColor:C.amber,backgroundColor:'transparent',borderDash:[6,4],pointRadius:3,pointBackgroundColor:C.amber,borderWidth:2,tension:.3},{label:'Annual CF',data:ADJ.cf,type:'bar',backgroundColor:ADJ.cf.map(v=>v>=0?'rgba(16,185,129,.35)':'rgba(244,63,94,.35)'),borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{y:{ticks:{callback:v=>fmt(v)}}}},plugins:[{id:'pbLine',afterDraw(ch){const yS=ch.scales.y;const y0=yS.getPixelForValue(0);ch.ctx.save();ch.ctx.setLineDash([6,4]);ch.ctx.strokeStyle='rgba(255,255,255,.12)';ch.ctx.lineWidth=1;ch.ctx.beginPath();ch.ctx.moveTo(ch.chartArea.left,y0);ch.ctx.lineTo(ch.chartArea.right,y0);ch.ctx.stroke();ch.ctx.restore();}}]});
  // Exit
  const ex5=ADJ.ebitda[4]*STATE.exitMult/1e6;
  const exx=document.getElementById('chart-exit').getContext('2d');
  charts.exit=new Chart(exx,{type:'bar',data:{labels:['Y3','Y5','Y7','Y10'],datasets:[{label:'EBITDA(M)',data:[ADJ.ebitda[2]/1e6,ADJ.ebitda[4]/1e6,ADJ.ebitda[6]/1e6,ADJ.ebitda[9]/1e6],backgroundColor:'rgba(99,102,241,.5)',borderRadius:3},{label:'EV(M)',data:[ADJ.ebitda[2]*STATE.exitMult/1e6,ADJ.ebitda[4]*STATE.exitMult/1e6,ADJ.ebitda[6]*STATE.exitMult/1e6,ADJ.ebitda[9]*STATE.exitMult/1e6],backgroundColor:'rgba(168,85,247,.5)',borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{y:{beginAtZero:true,title:{display:true,text:'AED (M)'}}}}});
  // IRR Radar
  const irr3=calcExitIRR(3);const irr5a=calcExitIRR(5);const irr7=calcExitIRR(7);
  const rrx=document.getElementById('chart-irr-radar').getContext('2d');
  charts.irrRadar=new Chart(rrx,{type:'radar',data:{labels:['Y3','Y5','Y7'],datasets:[{label:STATE.exitMult+'× Multiple',data:[irr3*100,irr5a*100,irr7*100],borderColor:C.indigo,backgroundColor:'rgba(99,102,241,.1)',pointRadius:4,borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,scales:{r:{beginAtZero:true,angleLines:{color:'rgba(255,255,255,.06)'},grid:{color:'rgba(255,255,255,.06)'}}}}});
  initExitTable();
  initFundingBridgeChart();
  initEquityChart();
  initFounderChart();
  initInvestorTimelineChart();
  buildDcfWaterfall();
}
function buildDcfWaterfall(){
  const canvas=document.getElementById('chart-dcf-waterfall');
  if(!canvas)return;
  if(charts.dcfWaterfall){charts.dcfWaterfall.destroy();charts.dcfWaterfall=null;}
  setText('dcfWaterfallTitle',`DCF Waterfall: Cash Flows → NPV (${STATE.wacc}% WACC)`);
  const labels=['Y0 Investment',...BASE.years,'NPV'];
  const vals=[ADJ.discCF[0]||0,...ADJ.discCF.slice(1,11),ADJ.npv||0];
  const isTotal=vals.map((_,i)=>i===0||i===vals.length-1);
  let run=0;
  const barData=vals.map((v,i)=>{
    if(isTotal[i]){
      run=v;
      return v<0?[v,0]:[0,v];
    }
    const start=run;
    run+=v;
    return v<0?[run,start]:[start,run];
  });
  charts.dcfWaterfall=new Chart(canvas.getContext('2d'),{
    type:'bar',
    data:{
      labels,
      datasets:[{
        data:barData,
        backgroundColor:vals.map((v,i)=>isTotal[i]?'rgba(99,102,241,.5)':v<0?'rgba(244,63,94,.35)':'rgba(16,185,129,.4)'),
        borderColor:vals.map((v,i)=>isTotal[i]?C.indigo:v<0?C.rose:C.emerald),
        borderWidth:1,
        borderRadius:3,
        borderSkipped:false
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{y:{ticks:{callback:v=>fmtCx(v)}}},
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>{
          const v=vals[c.dataIndex]||0;
          return `${v>=0?'+':''}${fmtCx(v)} ${cxLabel()} discounted`;
        }}}
      }
    }
  });
}
function calcExitIRR(yr){
  const cfs=ADJ.cf.slice(0,yr+1);
  cfs[yr]+=calcCompanyNetExit(ADJ.ebitda[yr-1]*STATE.exitMult).net;
  return calcIRR(cfs);
}
function updateCashflowCharts(){
  if(!charts.cashflow){
    initCashflowCharts();
    return;
  }
  charts.cashflow.data.datasets[0].data=ADJ.cumCF;
  charts.cashflow.data.datasets[0].pointBackgroundColor=ADJ.cumCF.map(v=>v>=0?C.emerald:C.rose);
  charts.cashflow.data.datasets[1].data=ADJ.discCumCF;
  charts.cashflow.data.datasets[1].label=`Discounted CF @ ${STATE.wacc}% WACC`;
  charts.cashflow.data.datasets[2].data=ADJ.cf;
  charts.cashflow.data.datasets[2].backgroundColor=ADJ.cf.map(v=>v>=0?'rgba(16,185,129,.35)':'rgba(244,63,94,.35)');
  charts.cashflow.update('active');
  buildDcfWaterfall();
  if(charts.exit){
    const yrs=[2,4,6,9];
    charts.exit.data.datasets[0].data=yrs.map(i=>ADJ.ebitda[i]/1e6);
    charts.exit.data.datasets[1].data=yrs.map(i=>ADJ.ebitda[i]*STATE.exitMult/1e6);
    charts.exit.update('active');
  }
  if(charts.irrRadar){
    const irr3=calcExitIRR(3);const irr5a=calcExitIRR(5);const irr7=calcExitIRR(7);
    charts.irrRadar.data.datasets=[{label:STATE.exitMult+'× Multiple',data:[irr3*100,irr5a*100,irr7*100],borderColor:C.indigo,backgroundColor:'rgba(99,102,241,.1)',pointRadius:4,borderWidth:2}];
    charts.irrRadar.update('active');
  }
  updateExitTable();
  updateFundingBridgeChart();
  updateEquityChart();
  updateFounderChart();
  updateInvestorTimelineChart();
}
function initExitTable(){
  const tbody=document.querySelector('#exitTable tbody');tbody.innerHTML='';
  [3,5,7,10].forEach(yr=>{
    const y=yr-1;const ebitda=ADJ.ebitda[y];const ev=ebitda*STATE.exitMult;const net=calcCompanyNetExit(ev).net;
    const cfs=ADJ.cf.slice(0,yr+1);cfs[yr]+=net;const irr=calcIRR(cfs);
    const tr=document.createElement('tr');
    tr.innerHTML=`<td style="font-weight:700">Y${yr}</td><td>${fmtM(ebitda)}M</td><td>${fmtM(ev)}M</td><td>${fmtM(net)}M</td><td class="${irr>.3?'pos':''}">${isNaN(irr)?'—':pct(irr)}</td>`;
    tbody.appendChild(tr);
  });
}
function updateExitTable(){initExitTable();}

/* ===== SENSITIVITY ===== */
function calcSensitivityPoint({revenueFactor=0,opexFactor=0,staffFactor=0}={}){
  const anchor=getSensitivityAnchor();
  const yearIdx=anchor.yearIdx;
  const revenue=ADJ.revenue[yearIdx]*(1+revenueFactor);
  const baseOpex=BASE.opexRate[yearIdx]*(1+STATE.opexAdj/100);
  const effectiveOpex=baseOpex*(1+opexFactor);
  const effectiveStaff=(0.03*STATE.staffAdj/100)+staffFactor;
  const margin=1-STATE.revShare/100-effectiveOpex-effectiveStaff;
  const ebitda=revenue*margin;
  const npv10=ebitda*5.65;
  return{revenue,margin,ebitda,npv10};
}
function getSensitivityPalette(values,baseIndex=3){
  const baseValue=values[baseIndex]??values[Math.floor(values.length/2)]??0;
  return values.map((value,index)=>{
    if(index===baseIndex)return{bg:'rgba(99,102,241,.7)',border:C.indigo};
    const ratio=baseValue!==0?value/baseValue:0;
    if(ratio<0.82)return{bg:'rgba(244,63,94,.5)',border:C.rose};
    if(ratio<0.95)return{bg:'rgba(245,158,11,.5)',border:C.amber};
    return{bg:'rgba(16,185,129,.5)',border:C.emerald};
  });
}
function calcStressVals(){
  return BASE.stressFactors.map(f=>calcSensitivityPoint({revenueFactor:f}).ebitda);
}
function calcOpexStressVals(){
  return BASE.opexStressFactors.map(f=>calcSensitivityPoint({opexFactor:f}).ebitda);
}
function calcStaffingVals(){
  const yearIdx=getSensitivityAnchor().yearIdx;
  const rev5=ADJ.revenue[yearIdx];
  // Provider compensation +5pp impact
  const rs5=rev5*0.05;
  // Salaries +25%
  const sal25=(BASE.plBreakdown.staff[yearIdx]||758068)*(1+STATE.revGrowth/100)*(1+STATE.opexAdj/100)*0.25;
  // Staff +2 FTE (~150K each)
  const staff2=300000;
  // Efficiency -1 FTE (~150K saved)
  const eff1=150000;
  return [-(rs5/1e6),-(sal25/1e6),-(staff2/1e6),(eff1/1e6)];
}
function renderSensitivitySummary(){
  const wrap=document.getElementById('sensitivitySummary');
  if(!wrap)return;
  const anchor=getSensitivityAnchor();
  const anchorDesc=anchor.followsFilter?`${anchor.yearLabel} selected in the period filter.`:'Multi-year selection active, so sensitivity stays anchored to Y5.';
  const worstRevenue=calcSensitivityPoint({revenueFactor:Math.min(...BASE.stressFactors)});
  const worstOpex=calcSensitivityPoint({opexFactor:Math.max(...BASE.opexStressFactors)});
  const worstTwoWay=calcSensitivityPoint({revenueFactor:Math.min(...BASE.heatFactors),opexFactor:Math.max(...BASE.heatFactors)});
  const cardTone=value=>value>=0?'scenario-safe':'scenario-warn';
  const cardValue=value=>`${fmtMCx(value)}M EBITDA`;
  wrap.innerHTML=`
    <div class="scenario-item scenario-safe">
      <div class="scenario-label">Breakeven Utilization</div>
      <div class="scenario-value">${BASE.breakevenUtilPct.toFixed(1)}%</div>
      <div class="scenario-desc">${tr('sensitivityBreakevenDesc')}</div>
    </div>
    <div class="scenario-item ${cardTone(worstRevenue.ebitda)}">
      <div class="scenario-label">Worst Revenue Case</div>
      <div class="scenario-value">${cardValue(worstRevenue.ebitda)}</div>
      <div class="scenario-desc">Revenue -30% vs current model base in ${anchor.yearLabel}. ${anchorDesc}</div>
    </div>
    <div class="scenario-item ${cardTone(worstOpex.ebitda)}">
      <div class="scenario-label">Worst OPEX Case</div>
      <div class="scenario-value">${cardValue(worstOpex.ebitda)}</div>
      <div class="scenario-desc">Non-revenue OPEX +20% vs current cost base in ${anchor.yearLabel}. ${anchorDesc}</div>
    </div>
    <div class="scenario-item ${cardTone(worstTwoWay.ebitda)}">
      <div class="scenario-label">Worst 2-Way</div>
      <div class="scenario-value">${cardValue(worstTwoWay.ebitda)}</div>
      <div class="scenario-desc">Revenue -20% with OPEX +20% in ${anchor.yearLabel}. ${anchorDesc}</div>
    </div>
  `;
}
function initSensitivityCharts(){
  const anchor=getSensitivityAnchor();
  updateSensitivityContextUI(anchor);
  // Destroy existing chart instances
  if(charts.stress){charts.stress.destroy();delete charts.stress;}
  if(charts.opexStress){charts.opexStress.destroy();delete charts.opexStress;}
  if(charts.staffing){charts.staffing.destroy();delete charts.staffing;}
  
  // Replace canvas elements with fresh ones to avoid any Chart.js residual state
  ['chart-stress','chart-opex-stress','chart-staffing'].forEach(id=>{
    const old=document.getElementById(id);
    if(old){
      const fresh=document.createElement('canvas');
      fresh.id=id;
      old.parentNode.replaceChild(fresh,old);
    }
  });
  
  const stressVals=calcStressVals();
  const stressPalette=getSensitivityPalette(stressVals,3);
  
  charts.stress=new Chart(document.getElementById('chart-stress'),{
    type:'bar',
    data:{
      labels:BASE.stressLabels,
      datasets:[{
        label:`${anchor.yearLabel} EBITDA`,
        data:stressVals,
        backgroundColor:stressPalette.map(t=>t.bg),
        borderColor:stressPalette.map(t=>t.border),
        borderWidth:1,
        borderRadius:4,
        barThickness:18
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      scales:{
        x:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)},grid:{color:'rgba(255,255,255,.04)'}},
        y:{grid:{display:false}}
      },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}
      }
    }
  });

  const opexStressVals=calcOpexStressVals();
  const opexPalette=getSensitivityPalette(opexStressVals,3);
  charts.opexStress=new Chart(document.getElementById('chart-opex-stress'),{
    type:'bar',
    data:{
      labels:BASE.opexStressLabels,
      datasets:[{
        label:`${anchor.yearLabel} EBITDA`,
        data:opexStressVals,
        backgroundColor:opexPalette.map(t=>t.bg),
        borderColor:opexPalette.map(t=>t.border),
        borderWidth:1,
        borderRadius:4,
        barThickness:18
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      scales:{
        x:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)},grid:{color:'rgba(255,255,255,.04)'}},
        y:{grid:{display:false}}
      },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}
      }
    }
  });
  
  initHeatmap();
  renderSensitivitySummary();
  
  const staffVals=calcStaffingVals();
  const staffBg=staffVals.map(v=>v>=0?'rgba(16,185,129,.5)':'rgba(244,63,94,.45)');
  const staffBorder=staffVals.map(v=>v>=0?C.emerald:C.rose);
  
  charts.staffing=new Chart(document.getElementById('chart-staffing'),{
    type:'bar',
    data:{
      labels:['Provider Comp +5pp','Salaries +25%','Staff +2 FTE','Efficiency -1 FTE'],
      datasets:[{
        label:`${anchor.yearLabel} EBITDA Δ`,
        data:staffVals,
        backgroundColor:staffBg,
        borderColor:staffBorder,
        borderWidth:1,
        borderRadius:4,
        barThickness:18
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      scales:{
        x:{title:{display:true,text:cxLabel()+' M'},grid:{color:'rgba(255,255,255,.04)'}},
        y:{grid:{display:false}}
      },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>(c.raw>=0?'+':'')+c.raw.toFixed(2)+'M '+cxLabel()}}
      }
    }
  });
  initTornadoChart();
}
function updateSensitivityCharts(){
  const anchor=getSensitivityAnchor();
  updateSensitivityContextUI(anchor);
  if(!charts.stress){
    initSensitivityCharts();
    return;
  }
  const stressVals=calcStressVals();
  const stressPalette=getSensitivityPalette(stressVals,3);
  charts.stress.data.datasets[0].data=stressVals;
  charts.stress.data.datasets[0].label=`${anchor.yearLabel} EBITDA`;
  charts.stress.data.datasets[0].backgroundColor=stressPalette.map(t=>t.bg);
  charts.stress.data.datasets[0].borderColor=stressPalette.map(t=>t.border);
  charts.stress.options.scales.x.ticks.callback=v=>fmtCx(v);
  charts.stress.update('active');
  if(charts.opexStress){
    const opexStressVals=calcOpexStressVals();
    const opexPalette=getSensitivityPalette(opexStressVals,3);
    charts.opexStress.data.datasets[0].data=opexStressVals;
    charts.opexStress.data.datasets[0].label=`${anchor.yearLabel} EBITDA`;
    charts.opexStress.data.datasets[0].backgroundColor=opexPalette.map(t=>t.bg);
    charts.opexStress.data.datasets[0].borderColor=opexPalette.map(t=>t.border);
    charts.opexStress.options.scales.x.ticks.callback=v=>fmtCx(v);
    charts.opexStress.update('active');
  }
  if(charts.staffing){
    const staffVals=calcStaffingVals();
    charts.staffing.data.datasets[0].data=staffVals;
    charts.staffing.data.datasets[0].label=`${anchor.yearLabel} EBITDA Δ`;
    charts.staffing.data.datasets[0].backgroundColor=staffVals.map(v=>v>=0?'rgba(16,185,129,.5)':'rgba(244,63,94,.45)');
    charts.staffing.data.datasets[0].borderColor=staffVals.map(v=>v>=0?C.emerald:C.rose);
    charts.staffing.options.scales.x.title.text=cxLabel()+' M';
    charts.staffing.update('active');
  }
  initHeatmap();
  renderSensitivitySummary();
  updateTornadoChart();
}
function initHeatmap(){
  const anchor=getSensitivityAnchor();
  const w=document.getElementById('heatmapWrap');
  const factors=BASE.heatFactors;
  let html='<table class="heatmap-table"><thead><tr><th></th>';
  BASE.heatOpexCols.forEach(c=>{html+='<th>'+c+'</th>';});
  html+='</tr></thead><tbody>';
  factors.forEach((rf,ri)=>{
    html+='<tr><td class="heatmap-label">'+BASE.heatRows[ri]+'</td>';
    factors.forEach((of,ci)=>{
      const point=calcSensitivityPoint({revenueFactor:rf,opexFactor:of});
      const isBase=ri===2&&ci===2;
      html+='<td class="heatmap-cell'+(isBase?' heatmap-base':'')+'" data-value="'+point.ebitda+'" data-rf="'+rf+'" data-of="'+of+'" data-npv="'+point.npv10+'" data-margin="'+(point.margin*100).toFixed(1)+'">'+fmtCx(point.ebitda)+'</td>';
    });
    html+='</tr>';
  });
  html+='</tbody></table>';
  html+='<div class="heatmap-detail" id="heatmapDetail"></div>';
  w.innerHTML=html;
  colorHeatmap();

  // Interactive click
  w.querySelectorAll('.heatmap-cell').forEach(cell=>{
    cell.addEventListener('click',()=>{
      w.querySelectorAll('.heatmap-cell').forEach(c=>c.classList.remove('heatmap-selected'));
      cell.classList.add('heatmap-selected');
      const rf=(parseFloat(cell.dataset.rf)*100).toFixed(0);
      const of=(parseFloat(cell.dataset.of)*100).toFixed(0);
      const ebitda=fmtCx(parseFloat(cell.dataset.value));
      const margin=cell.dataset.margin;
      const detail=document.getElementById('heatmapDetail');
      detail.innerHTML=`<strong>Revenue ${rf>=0?'+':''}${rf}% / OPEX ${of>=0?'+':''}${of}%</strong> → ${anchor.yearLabel} EBITDA: ${ebitda} ${cxLabel()} | Margin: ${margin}%`;
      detail.classList.add('visible');
    });
  });
}
function colorHeatmap(){
  const cells=document.querySelectorAll('.heatmap-cell');if(!cells.length)return;
  const vals=Array.from(cells).map(c=>parseFloat(c.dataset.value));
  const mn=Math.min(...vals),mx=Math.max(...vals);
  cells.forEach(c=>{const v=parseFloat(c.dataset.value);const t=(v-mn)/(mx-mn);
    const r=Math.round(244-t*200),g=Math.round(63+t*122),b=Math.round(94+t*37);
    c.style.backgroundColor=`rgba(${r},${g},${b},.25)`;c.style.color=`rgb(${r},${g},${b})`;
  });
}

/* ===== DECISION / OPS / GOVERNANCE MODULES ===== */
function getOperationalAnchor(selection=selectedYear){
  return getSensitivityAnchor(selection);
}
function renderServiceLineEconomics(){
  const anchor=getOperationalAnchor();
  const tbody=document.querySelector('#serviceEconomicsTable tbody');
  if(!tbody)return;
  setText('serviceEconomicsTitle',`Department Economics (${anchor.contextLabel})`);
  const yearIdx=anchor.yearIdx;
  const mix=BASE.serviceMix[yearIdx]||BASE.serviceMix[4];
  const totalRevenue=ADJ.revenue[yearIdx]||0;
  const baseRevenue=BASE.revenue[yearIdx]||0;
  const revenueScale=baseRevenue>0?totalRevenue/baseRevenue:1;
  const margin=ADJ.margins[yearIdx]||0;
  tbody.innerHTML=BASE.services.map((service,i)=>{
    const share=(mix[i]||0)/100;
    const revenue=(BASE.serviceRevenue[i]?.[yearIdx]||totalRevenue*share)*revenueScale;
    const visits=BASE.serviceVisits[i]?.[yearIdx]||0;
    const avgPrice=visits>0?revenue/visits:(BASE.revPerVisit[i]||0)*revenueScale;
    const ebitda=revenue*margin;
    return `<tr>
      <td style="font-weight:600">${service}</td>
      <td>${(share*100).toFixed(1)}%</td>
      <td>${fmtCx(avgPrice)}</td>
      <td>${fmt(visits)}</td>
      <td class="pos">${fmtMCx(revenue)}M</td>
      <td class="${ebitda>=0?'pos':'neg'}">${fmtMCx(ebitda)}M</td>
    </tr>`;
  }).join('');
}
function getPatientAcquisitionSeries(){
  const flow=getScenarioPatientFlowSeries();
  return{
    uniquePatients:flow.uniquePatients.slice(),
    avgRevenuePerVisit:flow.visits.map((visits,idx)=>visits>0?(Number(ADJ.revenue[idx])||0)/visits:Math.max(0,Number(BASE.paModel.avgRevenuePerVisit[idx])||0)),
    avgRevenuePerPatient:flow.uniquePatients.map((patients,idx)=>patients>0?(Number(ADJ.revenue[idx])||0)/patients:Math.max(0,Number(BASE.paModel.avgRevenuePerPatient[idx])||0)),
    marketingBudget:BASE.paModel.marketingBudget.map((v,idx)=>{
      const importedBudget=Math.max(0,Number(v)||0);
      const baseBudget=Math.max(importedBudget,Number(BASE.plBreakdown.marketing[idx])||0);
      const currentBudget=Math.max(0,Number(getCurrentPLBreakdown(idx).marketing)||0);
      const marketingScale=baseBudget>0&&currentBudget>0?currentBudget/baseBudget:1;
      return baseBudget*marketingScale;
    }),
    pureCAC:[],
    blendedCAC:[],
    ltv:[],
    ltvCac:[],
    cacPayback:[],
    romi:[],
    lifetimeYears:BASE.paModel.lifetimeYears.map(v=>preferPositiveValue(v,3)),
    retainedVisits:flow.retainedVisits.slice(),
    organicVisits:flow.organicVisits.slice(),
    paidVisits:flow.paidVisits.slice(),
    acquisitionSpend:[],
    retentionSpend:[]
  };
}
function finalizePatientAcquisitionSeries(series,flow=getScenarioPatientFlowSeries()){
  for(let i=0;i<10;i++){
    const retained=Math.max(0,Number(flow.retained[i])||0);
    const referral=Math.max(0,Number(flow.referral[i])||0);
    const walkin=Math.max(0,Number(flow.walkin[i])||0);
    const marketing=Math.max(0,Number(flow.marketing[i])||0);
    const unique=Math.max(0,Number(series.uniquePatients[i])||0);
    const newPatients=referral+walkin+marketing;
    const marketingBudget=Math.max(0,Number(series.marketingBudget[i])||0);
    const baseTotalBudget=Math.max(0,Number(BASE.paModel.marketingBudget[i])||0);
    const budgetScale=baseTotalBudget>0&&marketingBudget>0?marketingBudget/baseTotalBudget:1;
    const acquisitionSpend=Math.max(0,(Number(BASE.paModel.acquisitionSpend[i])||0)*budgetScale);
    const retentionSpend=Math.max(0,(Number(BASE.paModel.retentionSpend[i])||0)*budgetScale);
    const fallbackAcquisitionSpend=Math.max(0,marketing*(Number(BASE.paModel.pureCAC[i])||0));
    const effectiveAcquisitionSpend=acquisitionSpend>0?acquisitionSpend:(fallbackAcquisitionSpend>0?fallbackAcquisitionSpend:marketingBudget);
    const effectiveRetentionSpend=retentionSpend>0?retentionSpend:0;
    const effectiveTotalBudget=marketingBudget>0?marketingBudget:(effectiveAcquisitionSpend+effectiveRetentionSpend);
    const pureCAC=marketing>0?effectiveAcquisitionSpend/marketing:0;
    const blendedCAC=newPatients>0?effectiveTotalBudget/newPatients:0;
    const avgRevenuePerPatient=Math.max(0,Number(series.avgRevenuePerPatient[i])||0);
    const lifetimeYears=preferPositiveValue(series.lifetimeYears[i],3);
    const ltv=Math.max(0,avgRevenuePerPatient*lifetimeYears);
    const ltvCacBase=pureCAC>0?ltv/pureCAC:(blendedCAC>0?ltv/blendedCAC:0);
    const contributionPerPatient=avgRevenuePerPatient*Math.max(Number(ADJ.margins[i])||0,0.08);
    const cacPaybackBase=pureCAC>0&&contributionPerPatient>0?pureCAC/(contributionPerPatient/12):0;
    const romiBase=effectiveTotalBudget>0?(marketing*avgRevenuePerPatient)/effectiveTotalBudget:0;
    const retainedVisits=Math.max(0,Number(flow.retainedVisits[i])||0);
    const organicVisits=Math.max(0,Number(flow.organicVisits[i])||0);
    const paidVisits=Math.max(0,Number(flow.paidVisits[i])||0);

    series.marketingBudget[i]=effectiveTotalBudget;
    series.pureCAC[i]=pureCAC;
    series.blendedCAC[i]=blendedCAC;
    series.ltv[i]=ltv;
    series.ltvCac[i]=ltvCacBase;
    series.cacPayback[i]=cacPaybackBase;
    series.romi[i]=romiBase;
    series.acquisitionSpend[i]=effectiveAcquisitionSpend;
    series.retentionSpend[i]=effectiveRetentionSpend;
    series.retainedVisits[i]=retainedVisits>0?retainedVisits:(unique>0?(Number(flow.visits[i])||0)*(retained/unique):0);
    series.organicVisits[i]=organicVisits>0?organicVisits:(unique>0?(Number(flow.visits[i])||0)*((referral+walkin)/unique):0);
    series.paidVisits[i]=paidVisits>0?paidVisits:(unique>0?(Number(flow.visits[i])||0)*(marketing/unique):0);
  }
  return series;
}
function sumSeriesRange(series,years){
  return years.reduce((sum,yearIdx)=>sum+(Number(series?.[yearIdx])||0),0);
}
function avgSeriesRange(series,years,weightSeries=null){
  if(!Array.isArray(years)||!years.length)return 0;
  if(weightSeries){
    const weighted=years.reduce((sum,yearIdx)=>sum+(Number(series?.[yearIdx])||0)*(Number(weightSeries?.[yearIdx])||0),0);
    const weightTotal=years.reduce((sum,yearIdx)=>sum+(Number(weightSeries?.[yearIdx])||0),0);
    return weightTotal>0?weighted/weightTotal:0;
  }
  return years.reduce((sum,yearIdx)=>sum+(Number(series?.[yearIdx])||0),0)/years.length;
}
function getPatientTrendDatasetLabels(){
  return{
    retained:currentLanguage==='ru'?'Удержанные':'Retained',
    referral:currentLanguage==='ru'?'Реферал':'Referral',
    walkin:'Walk-in',
    paid:currentLanguage==='ru'?'Платный':'Paid',
    visits:currentLanguage==='ru'?'Всего визитов':'Total Visits',
    patients:currentLanguage==='ru'?'Всего пациентов':'Total Patients'
  };
}
function getPatientEngineViewLabels(){
  return{
    trend:currentLanguage==='ru'?'Тренд':'Trend',
    bridge:currentLanguage==='ru'?'База':'Base Bridge',
    table:currentLanguage==='ru'?'Детали':'Detail'
  };
}
function syncPatientEngineViewTabs(){
  const labels=getPatientEngineViewLabels();
  document.querySelectorAll('#patientEngineViewTabs [data-pa-view-tab]').forEach(btn=>{
    const view=btn.dataset.paViewTab;
    btn.classList.toggle('active',view===patientEngineView);
    btn.textContent=labels[view]||view;
  });
}
function applyPatientEngineView(){
  const activeView=['trend','bridge','table'].includes(patientEngineView)?patientEngineView:'trend';
  document.querySelectorAll('[data-pa-view]').forEach(panel=>{
    const isActive=panel.dataset.paView===activeView;
    panel.hidden=!isActive;
    panel.classList.toggle('is-active',isActive);
    panel.setAttribute('aria-hidden',String(!isActive));
  });
  syncPatientEngineViewTabs();
  const activePanel=document.querySelector(`[data-pa-view="${activeView}"]`);
  if(activePanel)resizeChartsWithin(activePanel);
}
function setPatientEngineView(view){
  const nextView=['trend','bridge','table'].includes(view)?view:'trend';
  if(nextView===patientEngineView){
    applyPatientEngineView();
    return;
  }
  patientEngineView=nextView;
  applyPatientEngineView();
}
function initPatientEngineViewTabs(){
  const tabs=document.getElementById('patientEngineViewTabs');
  if(!tabs||tabs.dataset.bound)return;
  tabs.querySelectorAll('[data-pa-view-tab]').forEach(btn=>{
    btn.addEventListener('click',()=>setPatientEngineView(btn.dataset.paViewTab));
  });
  tabs.dataset.bound='true';
  applyPatientEngineView();
}
function getPatientTrendModeLabels(){
  return{
    blend:currentLanguage==='ru'?'Смешанный':'Blend',
    patients:currentLanguage==='ru'?'Пациенты':'Patients',
    visits:currentLanguage==='ru'?'Визиты':'Visits'
  };
}
function syncPatientTrendModeTabs(){
  const labels=getPatientTrendModeLabels();
  document.querySelectorAll('#patientTrendModeTabs [data-pa-trend-mode]').forEach(btn=>{
    const mode=btn.dataset.paTrendMode;
    btn.classList.toggle('active',mode===patientTrendMode);
    btn.textContent=labels[mode]||mode;
  });
}
function setPatientTrendMode(mode,{announce=true}={}){
  const nextMode=['blend','patients','visits'].includes(mode)?mode:'blend';
  if(nextMode===patientTrendMode){
    syncPatientTrendModeTabs();
    return;
  }
  patientTrendMode=nextMode;
  syncPatientTrendModeTabs();
  updatePatientAcquisitionCharts();
  if(announce){
    const labels=getPatientTrendModeLabels();
    showToast(currentLanguage==='ru'?`📈 Режим тренда: ${labels[nextMode]}`:`📈 Trend mode: ${labels[nextMode]}`,'success');
  }
}
function initPatientTrendControls(){
  const tabs=document.getElementById('patientTrendModeTabs');
  if(!tabs||tabs.dataset.bound)return;
  tabs.querySelectorAll('[data-pa-trend-mode]').forEach(btn=>{
    btn.addEventListener('click',()=>setPatientTrendMode(btn.dataset.paTrendMode));
  });
  tabs.dataset.bound='true';
  syncPatientTrendModeTabs();
}
function getPatientTrendSelection(){
  const years=normalizeYearSelection();
  const allSelected=years.length===10&&years[0]===0&&years[years.length-1]===9;
  return{years,allSelected,label:getSelectedYearLabel()};
}
function buildTrendColorArray(rgb,selection,activeAlpha,mutedAlpha=Math.max(activeAlpha*0.32,0.12)){
  return BASE.years.map((_,idx)=>{
    const alpha=selection.allSelected||selection.years.includes(idx)?activeAlpha:mutedAlpha;
    return `rgba(${rgb[0]},${rgb[1]},${rgb[2]},${alpha})`;
  });
}
function buildTrendNumericArray(selection,activeValue,mutedValue){
  return BASE.years.map((_,idx)=>(selection.allSelected||selection.years.includes(idx)?activeValue:mutedValue));
}
function getPatientTrendChartState(){
  const flow=getScenarioPatientFlowSeries();
  const series=finalizePatientAcquisitionSeries(getPatientAcquisitionSeries(),flow);
  const trendLabels=getPatientTrendDatasetLabels();
  const selection=getPatientTrendSelection();
  const referralVisitSeries=BASE.years.map((_,idx)=>{
    const organicPatients=(Number(flow.referral[idx])||0)+(Number(flow.walkin[idx])||0);
    return organicPatients>0?(Number(series.organicVisits[idx])||0)*((Number(flow.referral[idx])||0)/organicPatients):0;
  });
  const walkinVisitSeries=BASE.years.map((_,idx)=>{
    const organicPatients=(Number(flow.referral[idx])||0)+(Number(flow.walkin[idx])||0);
    return organicPatients>0?(Number(series.organicVisits[idx])||0)*((Number(flow.walkin[idx])||0)/organicPatients):0;
  });

  let bars;
  let lineData;
  let yTitle;
  let y1Title;
  let lineAxis='y1';
  if(patientTrendMode==='patients'){
    bars={
      retained:flow.retained.slice(),
      referral:flow.referral.slice(),
      walkin:flow.walkin.slice(),
      paid:flow.marketing.slice()
    };
    lineData=series.uniquePatients.slice();
    yTitle=currentLanguage==='ru'?'Пациенты':'Patients';
    y1Title=trendLabels.patients;
    lineAxis='y';
  }else if(patientTrendMode==='visits'){
    bars={
      retained:series.retainedVisits.slice(),
      referral:referralVisitSeries,
      walkin:walkinVisitSeries,
      paid:series.paidVisits.slice()
    };
    lineData=flow.visits.slice();
    yTitle=currentLanguage==='ru'?'Визиты':'Visits';
    y1Title=trendLabels.visits;
    lineAxis='y';
  }else{
    bars={
      retained:flow.retained.slice(),
      referral:flow.referral.slice(),
      walkin:flow.walkin.slice(),
      paid:flow.marketing.slice()
    };
    lineData=flow.visits.slice();
    yTitle=currentLanguage==='ru'?'Пациенты':'Patients';
    y1Title=trendLabels.visits;
    lineAxis='y1';
  }

  return{
    labels:BASE.years,
    selection,
    lineAxis,
    yTitle,
    y1Title,
    datasets:[
      {type:'bar',label:trendLabels.retained,data:bars.retained,backgroundColor:buildTrendColorArray([99,102,241],selection,0.55),borderColor:buildTrendColorArray([99,102,241],selection,0.9,0.2),borderWidth:buildTrendNumericArray(selection,1,0),borderRadius:6,stack:'patients'},
      {type:'bar',label:trendLabels.referral,data:bars.referral,backgroundColor:buildTrendColorArray([16,185,129],selection,0.52),borderColor:buildTrendColorArray([16,185,129],selection,0.88,0.18),borderWidth:buildTrendNumericArray(selection,1,0),borderRadius:6,stack:'patients'},
      {type:'bar',label:trendLabels.walkin,data:bars.walkin,backgroundColor:buildTrendColorArray([245,158,11],selection,0.5),borderColor:buildTrendColorArray([245,158,11],selection,0.86,0.18),borderWidth:buildTrendNumericArray(selection,1,0),borderRadius:6,stack:'patients'},
      {type:'bar',label:trendLabels.paid,data:bars.paid,backgroundColor:buildTrendColorArray([34,211,238],selection,0.52),borderColor:buildTrendColorArray([34,211,238],selection,0.88,0.18),borderWidth:buildTrendNumericArray(selection,1,0),borderRadius:6,stack:'patients'},
      {type:'line',label:lineAxis==='y'?y1Title:trendLabels.visits,data:lineData,borderColor:'rgba(168,85,247,.95)',backgroundColor:'transparent',pointBackgroundColor:buildTrendColorArray([168,85,247],selection,0.95,0.28),pointBorderColor:buildTrendColorArray([168,85,247],selection,1,0.35),pointRadius:buildTrendNumericArray(selection,3.5,2),pointHoverRadius:buildTrendNumericArray(selection,5,3),borderWidth:2.5,tension:.3,yAxisID:lineAxis}
    ]
  };
}
function getPatientAcquisitionSnapshot(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  const label=getSelectedYearLabel(selection);
  const anchor=getOperationalAnchor(selection);
  const yearIdx=anchor.yearIdx;
  const flow=getScenarioPatientFlowSeries();
  const series=finalizePatientAcquisitionSeries(getPatientAcquisitionSeries(),flow);
  if(years.length>1){
    const retained=sumSeriesRange(flow.retained,years);
    const referral=sumSeriesRange(flow.referral,years);
    const walkin=sumSeriesRange(flow.walkin,years);
    const marketing=sumSeriesRange(flow.marketing,years);
    const unique=sumSeriesRange(series.uniquePatients,years);
    const visits=sumSeriesRange(flow.visits,years);
    const organicNew=referral+walkin;
    const totalRevenue=years.reduce((sum,i)=>sum+(Number(ADJ.revenue[i])||0),0);
    const totalEbitda=years.reduce((sum,i)=>sum+(Number(ADJ.ebitda[i])||0),0);
    const marketingBudget=sumSeriesRange(series.marketingBudget,years);
    const acquisitionSpend=sumSeriesRange(series.acquisitionSpend,years);
    const retentionSpend=sumSeriesRange(series.retentionSpend,years);
    const retainedVisits=sumSeriesRange(series.retainedVisits,years);
    const organicVisits=sumSeriesRange(series.organicVisits,years);
    const paidVisits=sumSeriesRange(series.paidVisits,years);
    const avgRevenuePerVisit=visits>0?totalRevenue/visits:avgSeriesRange(series.avgRevenuePerVisit,years);
    const avgRevenuePerPatient=unique>0?totalRevenue/unique:avgSeriesRange(series.avgRevenuePerPatient,years);
    const lifetimeYears=avgSeriesRange(series.lifetimeYears,years,series.uniquePatients);
    const pureCAC=marketing>0?acquisitionSpend/marketing:0;
    const blendedCAC=organicNew+marketing>0?marketingBudget/(organicNew+marketing):0;
    const ltv=avgRevenuePerPatient*lifetimeYears;
    const ltvCac=pureCAC>0?ltv/pureCAC:(blendedCAC>0?ltv/blendedCAC:0);
    const aggregateMargin=totalRevenue>0?Math.max(totalEbitda/totalRevenue,0.08):0.08;
    const contributionPerPatient=avgRevenuePerPatient*aggregateMargin;
    const cacPayback=pureCAC>0&&contributionPerPatient>0?pureCAC/(contributionPerPatient/12):0;
    const romi=marketingBudget>0?(marketing*avgRevenuePerPatient)/marketingBudget:0;
    return{
      anchor:{...anchor,contextLabel:currentLanguage==='ru'?`${label} итог`:`${label} total`,yearIdx:null,yearLabel:label},
      years,
      isAggregate:true,
      yearIdx:null,
      flow,
      series,
      retained,
      referral,
      walkin,
      marketing,
      organicNew,
      unique,
      visits,
      paidShare:unique>0?marketing/unique:0,
      organicShare:unique>0?organicNew/unique:0,
      avgRevenuePerVisit,
      avgRevenuePerPatient,
      marketingBudget,
      pureCAC,
      blendedCAC,
      ltv,
      ltvCac,
      cacPayback,
      romi,
      retainedVisits,
      organicVisits,
      paidVisits,
      acquisitionSpend,
      retentionSpend,
      otherMarketingSpend:Math.max(marketingBudget-acquisitionSpend-retentionSpend,0)
    };
  }
  const retained=Math.max(0,Number(flow.retained[yearIdx])||0);
  const referral=Math.max(0,Number(flow.referral[yearIdx])||0);
  const walkin=Math.max(0,Number(flow.walkin[yearIdx])||0);
  const marketing=Math.max(0,Number(flow.marketing[yearIdx])||0);
  const unique=Math.max(0,Number(series.uniquePatients[yearIdx])||0);
  const organicNew=referral+walkin;
  return{
    anchor,
    years,
    isAggregate:false,
    yearIdx,
    flow,
    series,
    retained,
    referral,
    walkin,
    marketing,
    organicNew,
    unique,
    paidShare:unique>0?marketing/unique:0,
    organicShare:unique>0?organicNew/unique:0,
    avgRevenuePerVisit:series.avgRevenuePerVisit[yearIdx]||0,
    avgRevenuePerPatient:series.avgRevenuePerPatient[yearIdx]||0,
    marketingBudget:series.marketingBudget[yearIdx]||0,
    pureCAC:series.pureCAC[yearIdx]||0,
    blendedCAC:series.blendedCAC[yearIdx]||0,
    ltv:series.ltv[yearIdx]||0,
    ltvCac:series.ltvCac[yearIdx]||0,
    cacPayback:series.cacPayback[yearIdx]||0,
    romi:series.romi[yearIdx]||0,
    retainedVisits:series.retainedVisits[yearIdx]||0,
    organicVisits:series.organicVisits[yearIdx]||0,
    paidVisits:series.paidVisits[yearIdx]||0,
    acquisitionSpend:series.acquisitionSpend[yearIdx]||0,
    retentionSpend:series.retentionSpend[yearIdx]||0,
    otherMarketingSpend:Math.max((series.marketingBudget[yearIdx]||0)-(series.acquisitionSpend[yearIdx]||0)-(series.retentionSpend[yearIdx]||0),0)
  };
}
function renderPatientAcquisitionModel(){
  const metrics=document.getElementById('patientAcqMetrics');
  const tbody=document.querySelector('#patientAcqTable tbody');
  if(!metrics||!tbody)return;
  const snapshot=getPatientAcquisitionSnapshot();
  const importedWorkbook=Boolean(importAudit.sheets?.includes('Patient Acquisition')&&importAudit.mappedMetrics?.some(label=>['Pure CAC','Blended CAC','LTV:CAC Ratio','Annual Marketing Budget'].includes(label)));
  const revenueAdjusted=BASE.revenue.some((value,idx)=>Math.abs((Number(ADJ.revenue[idx])||0)-(Number(value)||0))>1e-6);
  const marginAdjusted=BASE.margins.some((value,idx)=>Math.abs((Number(ADJ.margins[idx])||0)-(Number(value)||0))>1e-6);
  const marketingAdjusted=snapshot.series.marketingBudget.some((value,idx)=>Math.abs((Number(value)||0)-(Number(BASE.paModel.marketingBudget[idx])||0))>1e-6);
  const scenarioAdjusted=Boolean(importedWorkbook&&(revenueAdjusted||marginAdjusted||marketingAdjusted));
  const romiLabel=snapshot.romi>=1?snapshot.romi.toFixed(1)+'×':(snapshot.romi>=0?Math.round(snapshot.romi*100)+'%':Math.round(snapshot.romi*100)+'%');
  const contextLabel=currentLanguage==='ru'&&snapshot.anchor.contextLabel==='Y5 anchor'?'якорь Y5':snapshot.anchor.contextLabel;
  const patientCountLabel=snapshot.isAggregate
    ? (currentLanguage==='ru'?'Пациенты за период':'Patients in period')
    : (currentLanguage==='ru'?'Unique Patients':'Unique Patients');
  const trendSelection=getPatientTrendSelection();
  setText('patientEngineModuleTitle','Patient Engine');
  setText('patientAcqTitle',currentLanguage==='ru'?`Модель привлечения пациентов (${contextLabel})`:`Patient Acquisition Model (${contextLabel})`);
  setText('patientEngineTitle',currentLanguage==='ru'?`Patient Engine (${contextLabel})`:`Patient Engine (${contextLabel})`);
  setText('patientEngineTrendTitle',currentLanguage==='ru'?'Динамика Patient Engine (Y1-Y10)':'Patient Engine Trend (Y1-Y10)');
  setText('patientEngineTrendBadge',trendSelection.allSelected
    ? (currentLanguage==='ru'?'Полная история':'Full history')
    : (currentLanguage==='ru'?`Подсвечен ${trendSelection.label}`:`Highlight ${trendSelection.label}`));
  syncPatientEngineViewTabs();
  syncPatientTrendModeTabs();
  setText('patientChannelTitle',currentLanguage==='ru'?`Экономика каналов (${contextLabel})`:`Channel Economics (${contextLabel})`);
  setText('patientAcqBadge',importedWorkbook?(scenarioAdjusted?(currentLanguage==='ru'?'Workbook baseline + сценарий':'Workbook baseline + scenario'):(currentLanguage==='ru'?'Workbook baseline':'Workbook baseline')):(currentLanguage==='ru'?'Расчетная модель':'Model-derived'));

  const cards=[
    {label:patientCountLabel,value:fmt(snapshot.unique),tone:'positive'},
    {label:currentLanguage==='ru'?'Платный share (total)':'Paid Share (total)',value:(snapshot.paidShare*100).toFixed(1)+'%',tone:snapshot.paidShare<=0.25?'positive':'warning'},
    {label:currentLanguage==='ru'?'Маркетинговый бюджет':'Marketing Budget',value:fmtCx(snapshot.marketingBudget),tone:'warning'},
    {label:'Blended CAC',value:fmtCx(snapshot.blendedCAC),tone:snapshot.blendedCAC<=snapshot.avgRevenuePerPatient*0.4?'positive':'warning'},
    {label:'LTV:CAC',value:snapshot.ltvCac?snapshot.ltvCac.toFixed(1)+'×':'—',tone:snapshot.ltvCac>=3?'positive':snapshot.ltvCac>=2?'warning':'negative'},
    {label:currentLanguage==='ru'?'CAC Payback':'CAC Payback',value:snapshot.cacPayback?snapshot.cacPayback.toFixed(1)+' '+(currentLanguage==='ru'?'мес':'mo'):'—',tone:snapshot.cacPayback<=12?'positive':snapshot.cacPayback<=18?'warning':'negative'}
  ];
  metrics.innerHTML=cards.map(card=>`<div class="module-metric-card ${card.tone||''}"><div class="module-metric-label">${card.label}</div><div class="module-metric-value">${card.value}</div></div>`).join('');

  const callout=document.getElementById('patientAcqCallout');
  if(callout){
    callout.innerHTML=currentLanguage==='ru'
      ? `Органический приток новых пациентов: <strong>${fmt(snapshot.organicNew)}</strong> (${(snapshot.organicShare*100).toFixed(1)}% от общего потока). Средняя выручка на пациента: <strong>${fmtCx(snapshot.avgRevenuePerPatient)}</strong>, pure CAC: <strong>${fmtCx(snapshot.pureCAC)}</strong>, ROMI: <strong>${romiLabel}</strong>. ${snapshot.isAggregate?'Показатели выше агрегированы по выбранному диапазону лет. ':''}${scenarioAdjusted?'Baseline импортирован из workbook, затем CAC/LTV/payback пересчитываются от текущих параметров сценария. Поток пациентов также масштабируется сценарием спроса, но visits ограничены capacity.':'Значения выровнены с workbook baseline, а разрез бюджета сходится с годовым маркетинговым бюджетом.'}`
      : `Organic new patients: <strong>${fmt(snapshot.organicNew)}</strong> (${(snapshot.organicShare*100).toFixed(1)}% of total). Avg revenue per patient: <strong>${fmtCx(snapshot.avgRevenuePerPatient)}</strong>, pure CAC: <strong>${fmtCx(snapshot.pureCAC)}</strong>, ROMI: <strong>${romiLabel}</strong>. ${snapshot.isAggregate?'Values above are aggregated across the selected year range. ':''}${scenarioAdjusted?'Baseline is imported from the workbook; CAC/LTV/payback are then recalculated from the current scenario inputs. Patient flow also scales with scenario demand, while visits remain capped by capacity.':'Values are aligned to the workbook baseline and the budget split reconciles to the annual marketing budget.'}`;
  }

  const organicReferralVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.referral/snapshot.organicNew):0;
  const organicWalkinVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.walkin/snapshot.organicNew):0;
  const rows=[
    {label:currentLanguage==='ru'?'Paid Marketing':'Paid Marketing',patients:snapshot.marketing,share:snapshot.unique>0?snapshot.marketing/snapshot.unique:0,visits:snapshot.paidVisits,revenue:snapshot.marketing*snapshot.avgRevenuePerPatient,cost:snapshot.acquisitionSpend,efficiency:snapshot.ltvCac?`LTV:CAC ${snapshot.ltvCac.toFixed(1)}×`:`CAC ${fmtCx(snapshot.pureCAC)}`},
    {label:'Walk-in',patients:snapshot.walkin,share:snapshot.unique>0?snapshot.walkin/snapshot.unique:0,visits:organicWalkinVisits,revenue:snapshot.walkin*snapshot.avgRevenuePerPatient,cost:0,efficiency:currentLanguage==='ru'?'Органический спрос':'Organic demand'},
    {label:currentLanguage==='ru'?'Referral':'Referral',patients:snapshot.referral,share:snapshot.unique>0?snapshot.referral/snapshot.unique:0,visits:organicReferralVisits,revenue:snapshot.referral*snapshot.avgRevenuePerPatient,cost:0,efficiency:currentLanguage==='ru'?'Word of mouth':'Word of mouth'},
    {label:'Retained',patients:snapshot.retained,share:snapshot.unique>0?snapshot.retained/snapshot.unique:0,visits:snapshot.retainedVisits,revenue:snapshot.retained*snapshot.avgRevenuePerPatient,cost:snapshot.retentionSpend,efficiency:`${currentLanguage==='ru'?'Повторные':'Repeat'} ${(snapshot.retainedVisits/Math.max(snapshot.retained,1)).toFixed(1)}×`},
    {label:currentLanguage==='ru'?'Budget Cap / Other':'Budget Cap / Other',patients:0,share:0,visits:0,revenue:0,cost:snapshot.otherMarketingSpend,efficiency:currentLanguage==='ru'?'Остаток лимита бюджета':'Capped budget remainder'}
  ];
  tbody.innerHTML=rows.map(row=>`
    <tr>
      <td style="font-weight:600">${row.label}</td>
      <td>${fmtCount(row.patients)}</td>
      <td>${(row.share*100).toFixed(1)}%</td>
      <td>${fmtCount(row.visits)}</td>
      <td class="pos">${fmtCx(row.revenue)}</td>
      <td>${row.cost>0?fmtCx(row.cost):'—'}</td>
      <td>${row.efficiency}</td>
    </tr>
  `).join('');
}
function getPatientBaseBridgeRows(){
  const flow=getScenarioPatientFlowSeries();
  return BASE.years.map((year,yearIdx)=>{
    const opening=yearIdx===0?0:Math.max(0,Number(flow.uniquePatients?.[yearIdx-1])||0);
    const retained=Math.max(0,Number(flow.retained[yearIdx])||0);
    const newPaid=Math.max(0,Number(flow.marketing[yearIdx])||0);
    const newOrganic=Math.max(0,(Number(flow.referral[yearIdx])||0)+(Number(flow.walkin[yearIdx])||0));
    const closing=Math.max(0,Number(flow.uniquePatients?.[yearIdx])||0);
    const churn=Math.max(opening-retained,0);
    const retentionRate=opening>0?retained/opening:0;
    const churnRate=opening>0?churn/opening:0;
    return {year,opening,newPaid,newOrganic,retained,churn,closing,retentionRate,churnRate};
  });
}
function renderPatientBaseBridge(){
  const tbody=document.querySelector('#patientBaseTable tbody');
  const callout=document.getElementById('patientBaseCallout');
  if(!tbody)return;
  const years=normalizeYearSelection();
  const label=getSelectedYearLabel();
  setText('patientBaseTitle',currentLanguage==='ru'?`Детализация patient base (${label})`:`Patient Base Detail (${label})`);
  setText('patientBaseChartTitle',currentLanguage==='ru'?'Patient Base Bridge (10Y)':'Patient Base Bridge (10Y)');
  setText('patientBaseChartBadge','opening → closing');
  setText('patientBaseDetailBadge',label);
  if(callout){
    callout.innerHTML=currentLanguage==='ru'
      ? 'Opening base равен closing base предыдущего года. Churn рассчитывается как <strong>opening base минус retained из предыдущего года</strong>, потому что в workbook есть retained patients, но нет отдельной строки lost patients.'
      : 'Opening base equals prior-year closing base. Churn is derived as <strong>opening base minus retained from prior year</strong>, because the workbook provides retained patients but not an explicit lost-patient row.';
  }
  const rows=getPatientBaseBridgeRows().filter((_,yearIdx)=>years.includes(yearIdx));
  tbody.innerHTML=rows.map(row=>`<tr>
    <td style="font-weight:600">${row.year}</td>
    <td>${fmtCount(row.opening)}</td>
    <td>${fmtCount(row.newPaid)}</td>
    <td>${fmtCount(row.newOrganic)}</td>
    <td>${fmtCount(row.retained)}</td>
    <td class="${row.churn>0?'neg':''}">${fmtCount(row.churn)}</td>
    <td class="pos">${fmtCount(row.closing)}</td>
    <td>${row.opening>0?(row.retentionRate*100).toFixed(1)+'%':'—'}</td>
    <td>${row.opening>0?(row.churnRate*100).toFixed(1)+'%':'—'}</td>
  </tr>`).join('');
}
function initPatientAcquisitionCharts(){
  const snapshot=getPatientAcquisitionSnapshot();
  const series=snapshot.series;
  const trendState=getPatientTrendChartState();
  const engineLabels=[
    snapshot.isAggregate?(currentLanguage==='ru'?'Пациенты за период':'Patients in period'):'Unique Patients',
    snapshot.isAggregate?(currentLanguage==='ru'?'Визиты за период':'Visits in period'):'Visits'
  ];
  const organicReferralVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.referral/snapshot.organicNew):0;
  const organicWalkinVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.walkin/snapshot.organicNew):0;
  const patientBaseRows=getPatientBaseBridgeRows();
  const engineCtx=document.getElementById('chart-pa-engine')?.getContext('2d');
  const trendCtx=document.getElementById('chart-pa-trend')?.getContext('2d');
  const econCtx=document.getElementById('chart-pa-econ')?.getContext('2d');
  const baseCtx=document.getElementById('chart-patient-base')?.getContext('2d');
  if(engineCtx&&!charts.paEngine){
    charts.paEngine=new Chart(engineCtx,{
      type:'bar',
      data:{
        labels:engineLabels,
        datasets:[
          {label:'Retained',data:[snapshot.retained,snapshot.retainedVisits],backgroundColor:'rgba(99,102,241,.62)',borderRadius:6,borderSkipped:false},
          {label:'Referral',data:[snapshot.referral,organicReferralVisits],backgroundColor:'rgba(16,185,129,.58)',borderRadius:6,borderSkipped:false},
          {label:'Walk-in',data:[snapshot.walkin,organicWalkinVisits],backgroundColor:'rgba(245,158,11,.58)',borderRadius:6,borderSkipped:false},
          {label:'Paid',data:[snapshot.marketing,snapshot.paidVisits],backgroundColor:'rgba(34,211,238,.58)',borderRadius:6,borderSkipped:false}
        ]
      },
      options:{
        indexAxis:'y',
        responsive:true,
        maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        scales:{
          x:{stacked:true,ticks:{callback:v=>fmt(v)},title:{display:true,text:'Count'}},
          y:{stacked:true}
        },
        plugins:{
          tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmt(ctx.parsed.x)}`}}
        }
      }
    });
  }
  if(trendCtx&&!charts.paTrend){
    charts.paTrend=new Chart(trendCtx,{
      data:{
        labels:trendState.labels,
        datasets:trendState.datasets
      },
      options:{
        responsive:true,
        maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        scales:{
          x:{stacked:true},
          y:{stacked:true,beginAtZero:true,ticks:{callback:v=>fmtCount(v)},title:{display:true,text:trendState.yTitle}},
          y1:{display:trendState.lineAxis==='y1',position:'right',beginAtZero:true,grid:{display:false},ticks:{callback:v=>fmtCount(v)},title:{display:trendState.lineAxis==='y1',text:trendState.y1Title}}
        },
        plugins:{
          legend:{position:'top',labels:{padding:12,usePointStyle:true}},
          tooltip:{callbacks:{label:ctx=>{
            const isVisits=ctx.dataset.yAxisID==='y1';
            return `${ctx.dataset.label}: ${fmtCount(ctx.parsed.y)}${isVisits?(currentLanguage==='ru'?' визитов':' visits'):''}`;
          }}}
        }
      }
    });
  }
  if(econCtx&&!charts.paEcon){
    charts.paEcon=new Chart(econCtx,{
      type:'bar',
      data:{
        labels:BASE.years,
        datasets:[
          {label:'Blended CAC',data:series.blendedCAC.map(v=>cx(v)),backgroundColor:'rgba(99,102,241,.45)',borderColor:C.indigo,borderWidth:1,borderRadius:6,yAxisID:'y'},
          {label:'Pure CAC',data:series.pureCAC.map(v=>cx(v)),type:'line',borderColor:C.cyan,backgroundColor:'transparent',pointBackgroundColor:C.cyan,pointRadius:3,borderWidth:2,tension:.3,yAxisID:'y'},
          {label:'CAC Payback (mo)',data:series.cacPayback,type:'line',borderColor:C.amber,backgroundColor:'transparent',pointBackgroundColor:C.amber,pointRadius:3,borderWidth:2,tension:.3,yAxisID:'y1'}
        ]
      },
      options:{
        responsive:true,
        maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        scales:{
          y:{beginAtZero:true,title:{display:true,text:`CAC (${cxLabel()})`},ticks:{callback:v=>fmt(v)}},
          y1:{position:'right',beginAtZero:true,grid:{display:false},title:{display:true,text:'Months'}}
        },
        plugins:{
          tooltip:{callbacks:{label:ctx=>ctx.dataset.yAxisID==='y1'?`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1)} mo`:`${ctx.dataset.label}: ${fmt(ctx.parsed.y)} ${cxLabel()}`}}
        }
      }
    });
  }
  if(baseCtx&&!charts.paBridge){
    charts.paBridge=new Chart(baseCtx,{
      data:{
        labels:patientBaseRows.map(row=>row.year),
        datasets:[
          {type:'bar',label:'Retained',data:patientBaseRows.map(row=>row.retained),backgroundColor:'rgba(99,102,241,.58)',borderColor:C.indigo,borderWidth:1,borderRadius:5,stack:'flows'},
          {type:'bar',label:'New Organic',data:patientBaseRows.map(row=>row.newOrganic),backgroundColor:'rgba(16,185,129,.55)',borderColor:C.emerald,borderWidth:1,borderRadius:5,stack:'flows'},
          {type:'bar',label:'New Paid',data:patientBaseRows.map(row=>row.newPaid),backgroundColor:'rgba(34,211,238,.55)',borderColor:C.cyan,borderWidth:1,borderRadius:5,stack:'flows'},
          {type:'bar',label:'Churned',data:patientBaseRows.map(row=>-row.churn),backgroundColor:'rgba(244,63,94,.4)',borderColor:C.rose,borderWidth:1,borderRadius:5,stack:'flows'},
          {type:'line',label:'Opening Base',data:patientBaseRows.map(row=>row.opening),borderColor:'rgba(148,163,184,.9)',backgroundColor:'transparent',borderDash:[5,4],pointBackgroundColor:'rgba(148,163,184,.9)',pointRadius:2,borderWidth:2,tension:.25,yAxisID:'y'},
          {type:'line',label:'Closing Base',data:patientBaseRows.map(row=>row.closing),borderColor:C.amber,backgroundColor:'transparent',pointBackgroundColor:C.amber,pointRadius:3,borderWidth:2.5,tension:.25,yAxisID:'y'}
        ]
      },
      options:{
        responsive:true,
        maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        scales:{
          x:{stacked:true},
          y:{stacked:true,beginAtZero:true,ticks:{callback:v=>fmtCount(v)},title:{display:true,text:'Patients'}}
        },
        plugins:{
          legend:{position:'bottom',labels:{font:{size:10},padding:8,usePointStyle:true}},
          tooltip:{callbacks:{label:ctx=>{
            const row=patientBaseRows[ctx.dataIndex];
            if(ctx.dataset.label==='Opening Base')return`Opening Base: ${fmtCount(row.opening)}`;
            if(ctx.dataset.label==='Closing Base')return`Closing Base: ${fmtCount(row.closing)}`;
            if(ctx.dataset.label==='Churned')return`Churned: ${fmtCount(row.churn)}`;
            return`${ctx.dataset.label}: ${fmtCount(ctx.parsed.y)}`;
          }}}
        }
      }
    });
  }
}
function updatePatientAcquisitionCharts(){
  const snapshot=getPatientAcquisitionSnapshot();
  const series=snapshot.series;
  const trendState=getPatientTrendChartState();
  const engineLabels=[
    snapshot.isAggregate?(currentLanguage==='ru'?'Пациенты за период':'Patients in period'):'Unique Patients',
    snapshot.isAggregate?(currentLanguage==='ru'?'Визиты за период':'Visits in period'):'Visits'
  ];
  const organicReferralVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.referral/snapshot.organicNew):0;
  const organicWalkinVisits=snapshot.organicNew>0?snapshot.organicVisits*(snapshot.walkin/snapshot.organicNew):0;
  const patientBaseRows=getPatientBaseBridgeRows();
  if(charts.paEngine){
    charts.paEngine.data.labels=engineLabels;
    charts.paEngine.data.datasets[0].data=[snapshot.retained,snapshot.retainedVisits];
    charts.paEngine.data.datasets[1].data=[snapshot.referral,organicReferralVisits];
    charts.paEngine.data.datasets[2].data=[snapshot.walkin,organicWalkinVisits];
    charts.paEngine.data.datasets[3].data=[snapshot.marketing,snapshot.paidVisits];
    charts.paEngine.update('active');
  }
  if(charts.paTrend){
    charts.paTrend.data.labels=trendState.labels;
    charts.paTrend.data.datasets=trendState.datasets;
    charts.paTrend.options.scales.y.title.text=trendState.yTitle;
    charts.paTrend.options.scales.y1.display=trendState.lineAxis==='y1';
    charts.paTrend.options.scales.y1.title.display=trendState.lineAxis==='y1';
    charts.paTrend.options.scales.y1.title.text=trendState.y1Title;
    charts.paTrend.options.plugins.tooltip.callbacks.label=ctx=>{
      const isVisits=ctx.dataset.yAxisID==='y1';
      return `${ctx.dataset.label}: ${fmtCount(ctx.parsed.y)}${isVisits?(currentLanguage==='ru'?' визитов':' visits'):''}`;
    };
    charts.paTrend.update('active');
  }
  if(charts.paEcon){
    charts.paEcon.options.scales.y.title.text=`CAC (${cxLabel()})`;
    charts.paEcon.data.datasets[0].data=series.blendedCAC.map(v=>cx(v));
    charts.paEcon.data.datasets[1].data=series.pureCAC.map(v=>cx(v));
    charts.paEcon.data.datasets[2].data=series.cacPayback;
    charts.paEcon.update('active');
  }
  if(charts.paBridge){
    charts.paBridge.data.labels=patientBaseRows.map(row=>row.year);
    charts.paBridge.data.datasets[0].data=patientBaseRows.map(row=>row.retained);
    charts.paBridge.data.datasets[1].data=patientBaseRows.map(row=>row.newOrganic);
    charts.paBridge.data.datasets[2].data=patientBaseRows.map(row=>row.newPaid);
    charts.paBridge.data.datasets[3].data=patientBaseRows.map(row=>-row.churn);
    charts.paBridge.data.datasets[4].data=patientBaseRows.map(row=>row.opening);
    charts.paBridge.data.datasets[5].data=patientBaseRows.map(row=>row.closing);
    charts.paBridge.options.plugins.tooltip.callbacks.label=ctx=>{
      const row=patientBaseRows[ctx.dataIndex];
      if(ctx.dataset.label==='Opening Base')return`Opening Base: ${fmtCount(row.opening)}`;
      if(ctx.dataset.label==='Closing Base')return`Closing Base: ${fmtCount(row.closing)}`;
      if(ctx.dataset.label==='Churned')return`Churned: ${fmtCount(row.churn)}`;
      return`${ctx.dataset.label}: ${fmtCount(ctx.parsed.y)}`;
    };
    charts.paBridge.update('active');
  }
}
function renderActualVsPlan(){
  const panel=document.getElementById('actualVsPlanPanel');
  if(!panel)return;
  const hasActuals=Boolean(importAudit.actualSheets?.length);
  const importedAt=formatDateTimeLocalized(importAudit.importedAt);
  const statusText=getLocalizedActualsStatus();
  panel.innerHTML=`
    <div class="module-list">
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Источник факта':'Actual data source'}</span><span class="module-val ${hasActuals?'positive':'negative'}">${hasActuals?importAudit.actualSheets.join(', '):(currentLanguage==='ru'?'Не подключен':'Not connected')}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Последний workbook':'Last workbook'}</span><span class="module-val">${importAudit.fileName||'—'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Импортировано':'Imported at'}</span><span class="module-val">${importedAt}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Статус':'Status'}</span><span class="module-val">${statusText}</span></div>
    </div>
    <div class="module-callout ${hasActuals?'positive':'warning'}">${hasActuals
      ? (currentLanguage==='ru'
        ? 'Листы факта найдены. Следующий шаг: расширить mapping для variance-reporting.'
        : 'Actual sheets detected. Mapping layer can be extended for variance reporting.')
      : (currentLanguage==='ru'
        ? 'Панель работает как readiness-слой. Variance charts появятся после маппинга листа факта.'
        : 'This component is live as a readiness layer. Variance charts will populate once an actuals sheet is mapped.')}</div>
  `;
}
function getFundingBridgeData(){
  const overviewGroups=getCapexOverviewGroups();
  const amountByOverviewLabel=label=>overviewGroups.find(group=>group.label===label)?.amount||0;
  const fitOut=amountByOverviewLabel('Renovation & Fit-out')
    +amountByOverviewLabel('Physio Equipment')
    +amountByOverviewLabel('Dental Equipment')
    +amountByOverviewLabel('Cosmetology & Beauty')
    +amountByOverviewLabel('Assessment + Lab + Office/IT');
  const preOpening=amountByOverviewLabel('Pre-operational');
  const workingCapital=amountByOverviewLabel('Working Capital + Deposit + Inventory');
  const contingency=amountByOverviewLabel('Contingency');
  const raiseTarget=fitOut+preOpening+workingCapital+contingency;
  const operatingBurn=Math.max(...ADJ.cf.slice(1).map(v=>Math.max(-v,0)),0);
  const breakevenMonth=getOperationalBreakevenMonth()||'—';
  return{
    labels:currentLanguage==='ru'
      ? ['Fit-out и оборудование','Pre-opening','Working capital','Contingency']
      : ['Fit-out & equipment','Pre-opening','Working capital','Contingency'],
    values:[fitOut,preOpening,workingCapital,contingency],
    metrics:[
      {label:currentLanguage==='ru'?'Целевой раунд':'Raise target',value:fmtCx(raiseTarget),tone:'positive'},
      {label:currentLanguage==='ru'?'Операционный burn':'Operating burn',value:operatingBurn>0?fmtCx(operatingBurn):(currentLanguage==='ru'?'Без внешнего burn':'Self-funded'),tone:operatingBurn>0?'negative':'positive'},
      {label:currentLanguage==='ru'?'Месяц breakeven':'Breakeven month',value:breakevenMonth==='—'?'—':'M'+breakevenMonth,tone:'positive'},
      {label:currentLanguage==='ru'?'Payback':'Payback',value:formatPaybackValue(ADJ.payback),tone:ADJ.payback!==null&&ADJ.payback<5?'positive':'warning'}
    ]
  };
}
function renderFundingMetrics(){
  const wrap=document.getElementById('fundingMetrics');
  if(!wrap)return;
  const bridge=getFundingBridgeData();
  wrap.innerHTML=bridge.metrics.map(item=>`<div class="module-metric-card ${item.tone||''}"><div class="module-metric-label">${item.label}</div><div class="module-metric-value">${item.value}</div></div>`).join('');
}
function getDefaultEquityState(){
  // Total CapEx already includes Working Capital (336,600) — see DEFAULT_CAPEX_BREAKDOWN line 12.
  // Excel `Cash Flow & DCF!Total Investment` = 6,908,760 double-counts WC; do not follow that.
  const totalInvestment=BASE.investment;
  return{
    totalInvestment,
    returnMode:'compare',
    exitYear:5,
    equityStakePct:30,
    dividendPayoutPct:80,
    preferredReturnPct:8,
    revenueShareRatePct:10,
    targetReturnMultiple:2.5,
    graceMonths:6
  };
}
function normalizeEquityState(raw=equityState){
  const defaults=getDefaultEquityState();
  const mode=['compare','equity','revenue-share'].includes(raw.returnMode)?raw.returnMode:defaults.returnMode;
  const totalInvestment=Number(raw.totalInvestment)||defaults.totalInvestment;
  return{
    totalInvestment:clamp(totalInvestment,1000000,50000000),
    returnMode:mode,
    exitYear:clamp(parseInt(raw.exitYear,10)||defaults.exitYear,3,10),
    equityStakePct:clamp(Number(raw.equityStakePct)||defaults.equityStakePct,1,100),
    dividendPayoutPct:clamp(Number(raw.dividendPayoutPct)||defaults.dividendPayoutPct,0,100),
    preferredReturnPct:clamp(Number(raw.preferredReturnPct)||defaults.preferredReturnPct,0,30),
    revenueShareRatePct:clamp(Number(raw.revenueShareRatePct)||defaults.revenueShareRatePct,0,30),
    targetReturnMultiple:clamp(Number(raw.targetReturnMultiple)||defaults.targetReturnMultiple,1,5),
    graceMonths:clamp(Number(raw.graceMonths)||defaults.graceMonths,0,24)
  };
}
function persistEquityState(){
  equityState=normalizeEquityState(equityState);
  writeStoredUI(APP_STORAGE_KEYS.equityModel,equityState,'project');
}
function setEquityReturnMode(mode){
  if(!['compare','equity','revenue-share'].includes(mode))return;
  equityState.returnMode=mode;
  persistEquityState();
  syncEquityInputs();
  renderEquityModel();
  updateEquityChart();
  updateFounderChart();
  updateInvestorTimelineChart();
}
function updateReturnModeTabs(){
  const state=normalizeEquityState(equityState);
  document.querySelectorAll('#returnModeAssumptionTabs [data-return-mode]').forEach(btn=>{
    btn.classList.toggle('active',btn.dataset.returnMode===state.returnMode);
  });
}
function updateReturnModeVisibility(mode=normalizeEquityState(equityState).returnMode){
  document.querySelectorAll('[data-return-field]').forEach(field=>{
    const fieldMode=field.dataset.returnField;
    const shouldHide=mode!=='compare'&&fieldMode!==mode;
    field.hidden=shouldHide;
    field.classList.toggle('is-hidden',shouldHide);
    field.setAttribute('aria-hidden',String(shouldHide));
  });
  const compareCard=document.getElementById('equityCompareCard');
  if(compareCard){
    const shouldHide=mode!=='compare';
    compareCard.hidden=shouldHide;
    compareCard.classList.toggle('is-hidden',shouldHide);
    compareCard.setAttribute('aria-hidden',String(shouldHide));
  }
}
function syncEquityInputs(){
  const state=normalizeEquityState(equityState);
  const mapping=[
    ['eq-totalInvestment','totalInvestment'],
    ['eq-exitYear','exitYear'],
    ['eq-equityStakePct','equityStakePct'],
    ['eq-dividendPayoutPct','dividendPayoutPct'],
    ['eq-preferredReturnPct','preferredReturnPct'],
    ['eq-revenueShareRatePct','revenueShareRatePct'],
    ['eq-targetReturnMultiple','targetReturnMultiple'],
    ['eq-graceMonths','graceMonths']
  ];
  mapping.forEach(([id,key])=>{
    const el=document.getElementById(id);
    if(!el)return;
    if(document.activeElement===el)return;
    el.value=state[key];
  });
  updateReturnModeTabs();
  updateReturnModeVisibility(state.returnMode);
}
function calcCompanyNetExit(grossExit,taxBasis=BASE.investment){
  const taxableGain=Math.max(grossExit-taxBasis-375000,0);
  const tax=taxableGain*0.09;
  return {gross:grossExit,tax,net:Math.max(grossExit-tax,0)};
}
function getEquityModelSnapshot(){
  const state=normalizeEquityState(equityState);
  const exitYear=Math.max(3,Math.min(10,parseInt(state.exitYear,10)||5));
  const exitIdx=exitYear-1;
  const capital=state.totalInvestment;
  const equityStake=state.equityStakePct/100;
  const founderStake=Math.max(1-equityStake,0);
  const payoutRatio=state.dividendPayoutPct/100;
  const preferredReturn=state.preferredReturnPct/100;
  const revenueShareRate=state.revenueShareRatePct/100;
  const capTarget=capital*state.targetReturnMultiple;
  const graceFactor=Math.max(0,(12-state.graceMonths)/12);
  const dividendsTotal=ADJ.netProfit.map(v=>Math.max(v,0)*payoutRatio);
  const equityDividends=[];
  const founderDividends=[];
  const prefCatchup=[];
  let prefAccrued=0;
  dividendsTotal.forEach(totalDiv=>{
    prefAccrued+=capital*preferredReturn;
    const prefPaid=Math.min(totalDiv,prefAccrued);
    prefAccrued=Math.max(prefAccrued-prefPaid,0);
    const residual=Math.max(totalDiv-prefPaid,0);
    const investorDividend=prefPaid+(residual*equityStake);
    const founderDividend=residual*founderStake;
    prefCatchup.push(prefPaid);
    equityDividends.push(investorDividend);
    founderDividends.push(founderDividend);
  });
  const exitGrossSelected=ADJ.ebitda[exitIdx]*STATE.exitMult;
  const exitGross10=ADJ.ebitda[9]*STATE.exitMult;
  const companyExitSelected=calcCompanyNetExit(exitGrossSelected);
  const companyExit10=calcCompanyNetExit(exitGross10);
  // Investor exit uses investor-level tax (9% on gain above 375K, basis = investor capital) to match Excel DCF
  const calcInvestorNetExit=(grossEV,stake,investorCapital)=>{
    const grossShare=grossEV*stake;
    const taxableGain=Math.max(grossShare-investorCapital-375000,0);
    const tax=taxableGain*0.09;
    return{gross:grossShare,tax,net:Math.max(grossShare-tax,0)};
  };
  const investorExitSelected=calcInvestorNetExit(exitGrossSelected,equityStake,capital);
  const investorExit10=calcInvestorNetExit(exitGross10,equityStake,capital);
  const founderExitSelected=companyExitSelected.net*founderStake;
  const founderExit10=companyExit10.net*founderStake;
  const equityCashOnly=[-capital,...equityDividends];
  const equityCumulative=[];
  equityCashOnly.reduce((sum,val,idx)=>equityCumulative[idx]=sum+val,0);
  const equityCfSelected=[-capital,...equityDividends.slice(0,exitIdx),equityDividends[exitIdx]+investorExitSelected.net];
  const equityCf10=[-capital,...equityDividends.slice(0,9),equityDividends[9]+investorExit10.net];
  const totalReturnSelected=equalitySum(equityDividends.slice(0,exitYear))+investorExitSelected.net-capital;
  const totalReturn10=equalitySum(equityDividends)+investorExit10.net-capital;
  const grossMoicSelected=(equalitySum(equityDividends.slice(0,exitYear))+investorExitSelected.net)/capital;
  const grossMoic10=(equalitySum(equityDividends)+investorExit10.net)/capital;
  const equityIrrSelected=calcIRR(equityCfSelected);
  const equityIrr10=calcIRR(equityCf10);
  const scheduledSharePayments=ADJ.revenue.map((rev,idx)=>rev*revenueShareRate*(idx===0?graceFactor:1));
  const cappedSharePayments=[];
  let remainingCap=capTarget;
  scheduledSharePayments.forEach(payment=>{
    const capped=Math.max(Math.min(payment,remainingCap),0);
    cappedSharePayments.push(capped);
    remainingCap=Math.max(remainingCap-capped,0);
  });
  const revenueShareCf=[-capital,...cappedSharePayments];
  const revenueShareCumulative=[];
  revenueShareCf.reduce((sum,val,idx)=>revenueShareCumulative[idx]=sum+val,0);
  const revenueSharePaybackMonths=calcMonthlyPayback(revenueShareCf);
  const founderRevenueShareDividends=dividendsTotal.map((dividend,idx)=>Math.max(dividend-(cappedSharePayments[idx]||0),0));
  const founderEquityReturnSelected=equalitySum(founderDividends.slice(0,exitYear))+founderExitSelected;
  const founderEquityReturn10=equalitySum(founderDividends)+founderExit10;
  const founderRevShareReturnSelected=equalitySum(founderRevenueShareDividends.slice(0,exitYear))+companyExitSelected.net;
  const founderRevShareReturn10=equalitySum(founderRevenueShareDividends)+companyExit10.net;
  const activeInvestorReturn=state.returnMode==='revenue-share'?revenueShareCumulative[exitYear]:state.returnMode==='equity'?totalReturnSelected:Math.max(totalReturnSelected,revenueShareCumulative[exitYear]);
  const activeFounderReturn=state.returnMode==='revenue-share'?founderRevShareReturnSelected:state.returnMode==='equity'?founderEquityReturnSelected:Math.max(founderEquityReturnSelected,founderRevShareReturnSelected);
  const annualCashYieldY3Equity=capital>0?(equityDividends[2]||0)/capital:0;
  const annualCashYieldY3RevShare=capital>0?(cappedSharePayments[2]||0)/capital:0;
  const winnerLabel=(a,b,higherBetter=true)=>{
    if(a===null||b===null||Number.isNaN(a)||Number.isNaN(b))return'—';
    if(Math.abs(a-b)<1e-9)return'Tie';
    return higherBetter?(a>b?'Equity':'Revenue Share'):(a<b?'Equity':'Revenue Share');
  };
  const comparisonRows=[
    {label:currentLanguage==='ru'?'Капитал инвестора':'Investor Capital',equity:capital,revenueShare:capital,winner:'—',fmt:'money'},
    {label:currentLanguage==='ru'?'Доля инвестора':'Investor % Ownership',equity:equityStake,revenueShare:0,winner:'Equity',fmt:'pct'},
    {label:currentLanguage==='ru'?'Y3 annual cash yield':'Y3 Annual Cash Yield',equity:annualCashYieldY3Equity,revenueShare:annualCashYieldY3RevShare,winner:winnerLabel(annualCashYieldY3Equity,annualCashYieldY3RevShare,true),fmt:'pct'},
    {label:currentLanguage==='ru'?`Доход инвестора Y${exitYear}`:`Y${exitYear} Total Return`,equity:totalReturnSelected,revenueShare:revenueShareCumulative[exitYear],winner:winnerLabel(totalReturnSelected,revenueShareCumulative[exitYear],true),fmt:'money'},
    {label:currentLanguage==='ru'?'Доход инвестора 10Y':'10Y Total Return',equity:totalReturn10,revenueShare:revenueShareCumulative[10],winner:winnerLabel(totalReturn10,revenueShareCumulative[10],true),fmt:'money'},
    {label:currentLanguage==='ru'?`Gross MOIC Y${exitYear}`:`Y${exitYear} Gross MOIC`,equity:grossMoicSelected,revenueShare:(capital+revenueShareCumulative[exitYear])/capital,winner:winnerLabel(grossMoicSelected,(capital+revenueShareCumulative[exitYear])/capital,true),fmt:'multiple'},
    {label:currentLanguage==='ru'?'Gross MOIC 10Y':'10Y Gross MOIC',equity:grossMoic10,revenueShare:(capital+revenueShareCumulative[10])/capital,winner:winnerLabel(grossMoic10,(capital+revenueShareCumulative[10])/capital,true),fmt:'multiple'},
    {label:currentLanguage==='ru'?'Payback (месяцы)':'Payback (months)',equity:calcMonthlyPayback(equityCashOnly),revenueShare:revenueSharePaybackMonths,winner:winnerLabel(calcMonthlyPayback(equityCashOnly),revenueSharePaybackMonths,false),fmt:'months'},
    {label:currentLanguage==='ru'?'Всего выплачено инвестору (10Y)':'Total Cash Paid to Investor (10Y)',equity:equalitySum(equityDividends)+investorExit10.net,revenueShare:equalitySum(cappedSharePayments),winner:winnerLabel(equalitySum(equityDividends)+investorExit10.net,equalitySum(cappedSharePayments),false),fmt:'money'},
    {label:currentLanguage==='ru'?'Доля founder retained':'Founder Ownership Retained',equity:founderStake,revenueShare:1,winner:'Revenue Share',fmt:'pct'}
  ];
  return{
    state,
    exitYear,
    exitIdx,
    capital,
    equityStake,
    founderStake,
    payoutRatio,
    preferredReturn,
    revenueShareRate,
    capTarget,
    graceFactor,
    equityDividends,
    founderDividends,
    prefCatchup,
    prefArrears:prefAccrued,
    companyExitSelected,
    companyExit10,
    investorExitSelected,
    investorExit10,
    founderExitSelected,
    founderExit10,
    equityCashOnly,
    equityCumulative,
    equityAnnualSelected:[-capital,...equityDividends.map((val,idx)=>idx===exitIdx?val+investorExitSelected.net:val)],
    equityAnnual10:[-capital,...equityDividends.slice(0,9),equityDividends[9]+investorExit10.net],
    totalReturnSelected,
    totalReturn10,
    grossMoicSelected,
    grossMoic10,
    equityIrrSelected,
    equityIrr10,
    cappedSharePayments,
    revenueShareCf,
    revenueShareCumulative,
    revenueSharePaybackMonths,
    founderRevenueShareDividends,
    founderEquityReturnSelected,
    founderEquityReturn10,
    founderRevShareReturnSelected,
    founderRevShareReturn10,
    activeInvestorReturn,
    activeFounderReturn,
    annualCashYieldY3Equity,
    annualCashYieldY3RevShare,
    comparisonRows
  };
}
function renderEquityModel(){
  const metrics=document.getElementById('equityMetrics');
  const founderMetrics=document.getElementById('founderMetrics');
  const tbody=document.querySelector('#equityTable tbody');
  const note=document.getElementById('equityAssumptions');
  const comparisonTitle=document.getElementById('investorComparisonTitle');
  const returnModeBadge=document.getElementById('investorReturnModeBadge');
  const founderTitle=document.getElementById('founderComparisonTitle');
  const timelineTitle=document.getElementById('investorTimelineTitle');
  const exitYearLabel=document.getElementById('eq-exitYearLabel');
  const returnSpecificFieldsLabel=document.getElementById('returnSpecificFieldsLabel');
  if(!metrics||!tbody||!note)return;
  const snap=getEquityModelSnapshot();
  syncEquityInputs();
  setNodeListText('#equityControls > label:not([data-return-field]) .equity-label',[
    currentLanguage==='ru'?'Капитал инвестора':'Investor Capital',
    snap.state.returnMode==='revenue-share'
      ? (currentLanguage==='ru'?'Горизонт':'Horizon Year')
      : (currentLanguage==='ru'?'Год выхода / горизонт':'Exit / Horizon Year')
  ]);
  setNodeListText('#equityControls label[data-return-field="equity"] .equity-label',currentLanguage==='ru'
    ? ['Equity Stake %','Dividend Payout %','Preferred Return %']
    : ['Equity Stake %','Dividend Payout %','Preferred Return %']);
  setNodeListText('#equityControls label[data-return-field="revenue-share"] .equity-label',currentLanguage==='ru'
    ? ['Investor Revenue Share %','Target Return ×','Grace Period (months)']
    : ['Investor Revenue Share %','Target Return ×','Grace Period (months)']);
  const modeLabel=getReturnModeLabel(snap.state.returnMode);
  if(comparisonTitle)comparisonTitle.textContent=snap.state.returnMode==='compare'
    ? (currentLanguage==='ru'?`Сравнение доходности инвестора (Y${snap.exitYear} / Y10)`:`Investor Return Comparison (Y${snap.exitYear} / Y10)`)
    : (currentLanguage==='ru'?`Investor Return — ${modeLabel} (Y${snap.exitYear} / Y10)`:`Investor Return — ${modeLabel} (Y${snap.exitYear} / Y10)`);
  if(returnModeBadge)returnModeBadge.textContent=modeLabel;
  if(founderTitle)founderTitle.textContent=snap.state.returnMode==='compare'
    ? (currentLanguage==='ru'?`Сравнение доходности founder (Y${snap.exitYear} / Y10)`:`Founder Return Comparison (Y${snap.exitYear} / Y10)`)
    : (currentLanguage==='ru'?`Founder Return — ${modeLabel} (Y${snap.exitYear} / Y10)`:`Founder Return — ${modeLabel} (Y${snap.exitYear} / Y10)`);
  if(timelineTitle)timelineTitle.textContent=snap.state.returnMode==='compare'
    ? (currentLanguage==='ru'?'Таймлайн cash flow инвестора':'Investor Cash Flow Timeline')
    : (currentLanguage==='ru'?`Таймлайн cash flow инвестора — ${modeLabel}`:`Investor Cash Flow Timeline — ${modeLabel}`);
  if(exitYearLabel)exitYearLabel.textContent=snap.state.returnMode==='revenue-share'
    ? (currentLanguage==='ru'?'Горизонт':'Horizon Year')
    : (currentLanguage==='ru'?'Год выхода / горизонт':'Exit / Horizon Year');
  if(returnSpecificFieldsLabel) returnSpecificFieldsLabel.textContent=snap.state.returnMode==='equity'
    ? (currentLanguage==='ru'?'Параметры equity-модели':'Equity model inputs')
    : snap.state.returnMode==='revenue-share'
      ? (currentLanguage==='ru'?'Параметры revenue-share модели':'Revenue-share model inputs')
      : (currentLanguage==='ru'?'Параметры по модели':'Model-specific inputs');
  const metricCards=[
    `<div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Активный режим':'Active Return Mode'}</div><div class="module-metric-value">${modeLabel}</div></div>`
  ];
  if(snap.state.returnMode==='compare'){
    metricCards.push(`<div class="module-metric-card positive"><div class="module-metric-label">${currentLanguage==='ru'?`Лучший доход инвестора @ Y${snap.exitYear}`:`Best Investor Return @ Y${snap.exitYear}`}</div><div class="module-metric-value">${fmtCx(snap.activeInvestorReturn)}</div></div>`);
  }
  if(snap.state.returnMode==='equity'||snap.state.returnMode==='compare'){
    metricCards.push(
      `<div class="module-metric-card positive"><div class="module-metric-label">${currentLanguage==='ru'?`Equity доход / IRR @ Y${snap.exitYear}`:`Equity Y${snap.exitYear} Return / IRR`}</div><div class="module-metric-value">${fmtCx(snap.totalReturnSelected)} / ${isNaN(snap.equityIrrSelected)?tr('noData'):pct(snap.equityIrrSelected)}</div></div>`,
      `<div class="module-metric-card ${snap.grossMoic10>=2?'positive':'warning'}"><div class="module-metric-label">${currentLanguage==='ru'?'Equity 10Y Gross MOIC':'Equity 10Y Gross MOIC'}</div><div class="module-metric-value">${snap.grossMoic10.toFixed(2)}×</div></div>`,
      `<div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?`Выход инвестора / founder @ Y${snap.exitYear}`:`Investor Exit @ Y${snap.exitYear} / Founder Exit @ Y${snap.exitYear}`}</div><div class="module-metric-value">${fmtCx(snap.investorExitSelected.net)} / ${fmtCx(snap.founderExitSelected)}</div></div>`,
      `<div class="module-metric-card ${snap.prefArrears>0?'warning':'positive'}"><div class="module-metric-label">${currentLanguage==='ru'?'Задолженность по preferred return':'Preferred Return Arrears'}</div><div class="module-metric-value">${fmtCx(snap.prefArrears)}</div></div>`,
      `<div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?`Налог при exit @ Y${snap.exitYear} / Y10`:`Exit Tax @ Y${snap.exitYear} / Y10`}</div><div class="module-metric-value">${fmtCx(snap.companyExitSelected.tax)} / ${fmtCx(snap.companyExit10.tax)}</div></div>`
    );
  }
  if(snap.state.returnMode==='revenue-share'||snap.state.returnMode==='compare'){
    metricCards.push(
      `<div class="module-metric-card positive"><div class="module-metric-label">${currentLanguage==='ru'?`Investor rev share @ Y${snap.exitYear}`:`Investor Rev Share @ Y${snap.exitYear}`}</div><div class="module-metric-value">${fmtCx(snap.revenueShareCumulative[snap.exitYear])}</div></div>`,
      `<div class="module-metric-card positive"><div class="module-metric-label">${currentLanguage==='ru'?'Investor rev share @ Y10':'Investor Rev Share @ Y10'}</div><div class="module-metric-value">${fmtCx(snap.revenueShareCumulative[10])}</div></div>`,
      `<div class="module-metric-card ${snap.revenueSharePaybackMonths!==null&&snap.revenueSharePaybackMonths<60?'positive':'warning'}"><div class="module-metric-label">${currentLanguage==='ru'?'Payback / cap по revenue share':'Investor Rev Share Payback / Cap'}</div><div class="module-metric-value">${snap.revenueSharePaybackMonths!==null?snap.revenueSharePaybackMonths.toFixed(1)+(currentLanguage==='ru'?' мес':'m'):tr('noData')} / ${fmtCx(snap.capTarget)}</div></div>`
    );
  }
  metrics.innerHTML=metricCards.join('');
  if(founderMetrics){
    const founderCards=[];
    if(snap.state.returnMode==='equity'||snap.state.returnMode==='compare'){
      founderCards.push(
        `<div class="module-metric-card ${snap.founderEquityReturnSelected>=0?'positive':'warning'}"><div class="module-metric-label">${currentLanguage==='ru'?`Founder @ Y${snap.exitYear} (equity)`:`Founder Y${snap.exitYear} Equity`}</div><div class="module-metric-value">${fmtCx(snap.founderEquityReturnSelected)}</div></div>`,
        `<div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Founder 10Y (equity)':'Founder 10Y Equity'}</div><div class="module-metric-value">${fmtCx(snap.founderEquityReturn10)}</div></div>`
      );
    }
    if(snap.state.returnMode==='revenue-share'||snap.state.returnMode==='compare'){
      founderCards.push(
        `<div class="module-metric-card ${snap.founderRevShareReturnSelected>=0?'positive':'warning'}"><div class="module-metric-label">${currentLanguage==='ru'?`Founder @ Y${snap.exitYear} (rev share)`:`Founder Y${snap.exitYear} under Rev Share`}</div><div class="module-metric-value">${fmtCx(snap.founderRevShareReturnSelected)}</div></div>`,
        `<div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Founder 10Y (rev share)':'Founder 10Y under Rev Share'}</div><div class="module-metric-value">${fmtCx(snap.founderRevShareReturn10)}</div></div>`
      );
    }
    founderMetrics.innerHTML=founderCards.join('');
  }
  const fmtCell=(value,kind)=>{
    if(value===null||value===undefined||Number.isNaN(value))return'—';
    if(kind==='money')return fmtCx(value);
    if(kind==='pct')return pct(value);
    if(kind==='multiple')return value.toFixed(2)+'×';
    if(kind==='months')return value.toFixed(1);
    return String(value);
  };
  tbody.innerHTML=snap.comparisonRows.map(row=>`
    <tr>
      <td style="font-weight:600">${row.label}</td>
      <td>${fmtCell(row.equity,row.fmt)}</td>
      <td>${fmtCell(row.revenueShare,row.fmt)}</td>
      <td class="${row.winner==='Equity'?'pos':row.winner==='Revenue Share'?'neg':''}">${getWinnerLabel(row.winner)}</td>
    </tr>
  `).join('');
  const equityNote=currentLanguage==='ru'
    ? `Equity-модель включает initial cash out в Year 0, ежегодные дивиденды с catch-up waterfall по preferred return ${snap.state.preferredReturnPct.toFixed(1)}%, exit proceeds в Y${snap.exitYear}/Y10 и налог 9% на прибыль при продаже сверх AED 375K на уровне компании.`
    : `Equity model includes initial cash out at Year 0, annual dividends with an ${snap.state.preferredReturnPct.toFixed(1)}% preferred-return catch-up waterfall, exit proceeds at Y${snap.exitYear}/Y10, and 9% exit tax on gains above AED 375K at company level.`;
  const revShareNote=currentLanguage==='ru'
    ? `Revenue-share модель включает initial cash out в Year 0 и выплаты инвестору в размере ${snap.state.revenueShareRatePct.toFixed(1)}% от выручки, с cap на уровне ${snap.state.targetReturnMultiple.toFixed(1)}× капитала после grace period ${snap.state.graceMonths.toFixed(0)} мес.`
    : `Revenue-share model includes initial cash out at Year 0 and investor revenue-share payments at ${snap.state.revenueShareRatePct.toFixed(1)}% of revenue, capped at ${snap.state.targetReturnMultiple.toFixed(1)}× capital after a ${snap.state.graceMonths.toFixed(0)}-month grace period.`;
  note.textContent=snap.state.returnMode==='equity'?equityNote:snap.state.returnMode==='revenue-share'?revShareNote:`${equityNote} ${revShareNote}`;
}
function initEquityModelModule(){
  equityState=normalizeEquityState(readStoredUI(APP_STORAGE_KEYS.equityModel,getDefaultEquityState(),'project'));
  const wrap=document.getElementById('equityControls');
  const tabs=document.getElementById('investorTimelineTabs');
  const scenarioTabs=document.getElementById('investorScenarioTabs');
  const assumptionReturnTabs=document.getElementById('returnModeAssumptionTabs');
  if(wrap&&!wrap.dataset.bound){
    wrap.querySelectorAll('.equity-input').forEach(input=>{
      input.addEventListener('change',()=>{
        const keyMap={
          'eq-totalInvestment':'totalInvestment',
          'eq-exitYear':'exitYear',
          'eq-equityStakePct':'equityStakePct',
          'eq-dividendPayoutPct':'dividendPayoutPct',
          'eq-preferredReturnPct':'preferredReturnPct',
          'eq-revenueShareRatePct':'revenueShareRatePct',
          'eq-targetReturnMultiple':'targetReturnMultiple',
          'eq-graceMonths':'graceMonths'
        };
        const key=keyMap[input.id];
        if(!key)return;
        equityState[key]=key==='returnMode'?input.value:parseFloat(input.value);
        if(key==='returnMode'){
          setEquityReturnMode(input.value);
        }else{
          persistEquityState();
          renderEquityModel();
          updateEquityChart();
          updateFounderChart();
          updateInvestorTimelineChart();
        }
      });
    });
    wrap.dataset.bound='true';
  }
  if(assumptionReturnTabs&&!assumptionReturnTabs.dataset.bound){
    assumptionReturnTabs.querySelectorAll('[data-return-mode]').forEach(btn=>{
      btn.addEventListener('click',()=>setEquityReturnMode(btn.dataset.returnMode||'compare'));
    });
    assumptionReturnTabs.dataset.bound='true';
  }
  if(tabs&&!tabs.dataset.bound){
    tabs.querySelectorAll('[data-timeline-mode]').forEach(btn=>{
      btn.addEventListener('click',()=>{
        investorTimelineMode=btn.dataset.timelineMode||'annual';
        updateInvestorTimelineChart();
      });
    });
    tabs.dataset.bound='true';
  }
  if(scenarioTabs&&!scenarioTabs.dataset.bound){
    scenarioTabs.querySelectorAll('[data-scenario-mode]').forEach(btn=>{
      btn.addEventListener('click',()=>{
        investorScenarioMode=btn.dataset.scenarioMode||'native';
        updateInvestorTimelineChart();
      });
    });
    scenarioTabs.dataset.bound='true';
  }
  syncEquityInputs();
  renderEquityModel();
}
function initFundingBridgeChart(){
  const el=document.getElementById('chart-funding-bridge');
  if(!el)return;
  if(charts.fundingBridge){charts.fundingBridge.destroy();delete charts.fundingBridge;}
  const bridge=getFundingBridgeData();
  charts.fundingBridge=new Chart(el,{
    type:'bar',
    data:{
      labels:bridge.labels,
      datasets:[{
        label:'AED',
        data:bridge.values,
        backgroundColor:['rgba(99,102,241,.55)','rgba(34,211,238,.55)','rgba(245,158,11,.55)','rgba(168,85,247,.55)'],
        borderColor:[C.indigo,C.cyan,C.amber,C.purple],
        borderWidth:1,
        borderRadius:4,
        barThickness:18
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      scales:{x:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)}},y:{grid:{display:false}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}}
    }
  });
  renderFundingMetrics();
}
function updateFundingBridgeChart(){
  const bridge=getFundingBridgeData();
  if(!charts.fundingBridge){
    initFundingBridgeChart();
    return;
  }
  charts.fundingBridge.data.labels=bridge.labels;
  charts.fundingBridge.data.datasets[0].data=bridge.values;
  charts.fundingBridge.options.scales.x.ticks.callback=v=>fmtCx(v);
  charts.fundingBridge.options.plugins.tooltip.callbacks.label=c=>fmtCx(c.raw)+' '+cxLabel();
  charts.fundingBridge.update('active');
  renderFundingMetrics();
}
function filterReturnModeDatasets(mode,datasets){
  if(mode==='equity')return datasets.filter(ds=>ds.mode==='equity').map(({mode,...ds})=>ds);
  if(mode==='revenue-share')return datasets.filter(ds=>ds.mode==='revenue-share').map(({mode,...ds})=>ds);
  return datasets.map(({mode,...ds})=>ds);
}
function initEquityChart(){
  const el=document.getElementById('chart-equity-proceeds');
  if(!el)return;
  const snap=getEquityModelSnapshot();
  const datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:getReturnModeLabel('equity'),data:[snap.totalReturnSelected,snap.totalReturn10],backgroundColor:'rgba(99,102,241,.55)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:getReturnModeLabel('revenue-share'),data:[snap.revenueShareCumulative[snap.exitYear],snap.revenueShareCumulative[10]],backgroundColor:'rgba(16,185,129,.45)',borderColor:C.emerald,borderWidth:1,borderRadius:4}
  ]);
  if(charts.equityProceeds){charts.equityProceeds.destroy();delete charts.equityProceeds;}
  charts.equityProceeds=new Chart(el,{
    type:'bar',
    data:{
      labels:[currentLanguage==='ru'?`Доход Y${snap.exitYear}`:`Y${snap.exitYear} Total Return`,currentLanguage==='ru'?'Доход 10Y':'10Y Total Return'],
      datasets
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{y:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)}},x:{grid:{display:false}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtCx(c.raw)} ${cxLabel()}`}}}
    }
  });
}
function updateEquityChart(){
  const snap=getEquityModelSnapshot();
  if(!charts.equityProceeds){
    initEquityChart();
    return;
  }
  const datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:getReturnModeLabel('equity'),data:[snap.totalReturnSelected,snap.totalReturn10],backgroundColor:'rgba(99,102,241,.55)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:getReturnModeLabel('revenue-share'),data:[snap.revenueShareCumulative[snap.exitYear],snap.revenueShareCumulative[10]],backgroundColor:'rgba(16,185,129,.45)',borderColor:C.emerald,borderWidth:1,borderRadius:4}
  ]);
  charts.equityProceeds.data.labels=[currentLanguage==='ru'?`Доход Y${snap.exitYear}`:`Y${snap.exitYear} Total Return`,currentLanguage==='ru'?'Доход 10Y':'10Y Total Return'];
  charts.equityProceeds.data.datasets=datasets;
  charts.equityProceeds.options.scales.y.ticks.callback=v=>fmtCx(v);
  charts.equityProceeds.update('active');
}
function initFounderChart(){
  const el=document.getElementById('chart-founder-proceeds');
  if(!el)return;
  const snap=getEquityModelSnapshot();
  const datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:currentLanguage==='ru'?'Founder under Equity':'Founder under Equity',data:[snap.founderEquityReturnSelected,snap.founderEquityReturn10],backgroundColor:'rgba(99,102,241,.45)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:currentLanguage==='ru'?'Founder under Revenue Share':'Founder under Revenue Share',data:[snap.founderRevShareReturnSelected,snap.founderRevShareReturn10],backgroundColor:'rgba(245,158,11,.45)',borderColor:C.amber,borderWidth:1,borderRadius:4}
  ]);
  if(charts.founderProceeds){charts.founderProceeds.destroy();delete charts.founderProceeds;}
  charts.founderProceeds=new Chart(el,{
    type:'bar',
    data:{labels:[currentLanguage==='ru'?`Доход founder Y${snap.exitYear}`:`Y${snap.exitYear} Founder Return`,currentLanguage==='ru'?'Доход founder 10Y':'10Y Founder Return'],datasets},
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{y:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)}},x:{grid:{display:false}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtCx(c.raw)} ${cxLabel()}`}}}
    }
  });
}
function updateFounderChart(){
  const snap=getEquityModelSnapshot();
  if(!charts.founderProceeds){
    initFounderChart();
    return;
  }
  charts.founderProceeds.data.labels=[currentLanguage==='ru'?`Доход founder Y${snap.exitYear}`:`Y${snap.exitYear} Founder Return`,currentLanguage==='ru'?'Доход founder 10Y':'10Y Founder Return'];
  charts.founderProceeds.data.datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:currentLanguage==='ru'?'Founder under Equity':'Founder under Equity',data:[snap.founderEquityReturnSelected,snap.founderEquityReturn10],backgroundColor:'rgba(99,102,241,.45)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:currentLanguage==='ru'?'Founder under Revenue Share':'Founder under Revenue Share',data:[snap.founderRevShareReturnSelected,snap.founderRevShareReturn10],backgroundColor:'rgba(245,158,11,.45)',borderColor:C.amber,borderWidth:1,borderRadius:4}
  ]);
  charts.founderProceeds.options.scales.y.ticks.callback=v=>fmtCx(v);
  charts.founderProceeds.update('active');
}
function getInvestorTimelineSeries(){
  const snap=getEquityModelSnapshot();
  const labels=['Y0',...BASE.years];
  const equityAnnual=[-snap.capital,...snap.equityDividends.map((value,idx)=>idx+1>snap.exitYear?0:value)];
  equityAnnual[snap.exitYear]+=snap.investorExitSelected.net;
  const revenueAnnual=[
    -snap.capital,
    ...snap.cappedSharePayments.map((value,idx)=>{
      if(investorScenarioMode==='exit-aligned'&&idx+1>snap.exitYear)return 0;
      return value;
    })
  ];
  const cumulative=(series)=>{
    let running=0;
    return series.map(value=>{
      running+=value;
      return running;
    });
  };
  return{
    labels,
    annual:{equity:equityAnnual,revenueShare:revenueAnnual},
    cumulative:{equity:cumulative(equityAnnual),revenueShare:cumulative(revenueAnnual)}
  };
}
function updateInvestorTimelineTabs(){
  document.querySelectorAll('#investorTimelineTabs [data-timeline-mode]').forEach(btn=>{
    btn.classList.toggle('active',btn.dataset.timelineMode===investorTimelineMode);
  });
  document.querySelectorAll('#investorScenarioTabs [data-scenario-mode]').forEach(btn=>{
    btn.classList.toggle('active',btn.dataset.scenarioMode===investorScenarioMode);
  });
}
function initInvestorTimelineChart(){
  const el=document.getElementById('chart-investor-timeline');
  if(!el)return;
  const series=getInvestorTimelineSeries();
  const snap=getEquityModelSnapshot();
  const activeSeries=series[investorTimelineMode];
  const datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:getReturnModeLabel('equity'),data:activeSeries.equity,backgroundColor:'rgba(99,102,241,.5)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:getReturnModeLabel('revenue-share'),data:activeSeries.revenueShare,backgroundColor:'rgba(16,185,129,.42)',borderColor:C.emerald,borderWidth:1,borderRadius:4}
  ]);
  if(charts.investorTimeline){charts.investorTimeline.destroy();delete charts.investorTimeline;}
  charts.investorTimeline=new Chart(el,{
    type:'bar',
    data:{
      labels:series.labels,
      datasets
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      scales:{y:{ticks:{callback:v=>fmtCx(v)}},x:{grid:{display:false}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtCx(c.raw)} ${cxLabel()}`}}}
    }
  });
  updateInvestorTimelineTabs();
}
function updateInvestorTimelineChart(){
  const series=getInvestorTimelineSeries();
  const snap=getEquityModelSnapshot();
  const activeSeries=series[investorTimelineMode];
  if(!charts.investorTimeline){
    initInvestorTimelineChart();
    return;
  }
  charts.investorTimeline.data.labels=series.labels;
  charts.investorTimeline.data.datasets=filterReturnModeDatasets(snap.state.returnMode,[
    {mode:'equity',label:getReturnModeLabel('equity'),data:activeSeries.equity,backgroundColor:'rgba(99,102,241,.5)',borderColor:C.indigo,borderWidth:1,borderRadius:4},
    {mode:'revenue-share',label:getReturnModeLabel('revenue-share'),data:activeSeries.revenueShare,backgroundColor:'rgba(16,185,129,.42)',borderColor:C.emerald,borderWidth:1,borderRadius:4}
  ]);
  charts.investorTimeline.options.scales.y.ticks.callback=v=>fmtCx(v);
  charts.investorTimeline.update('active');
  updateInvestorTimelineTabs();
}
function getTornadoSeries(){
  const baseline=ADJ.npv;
  const cases=[
    {label:'Revenue Growth ±10pp',low:{revGrowth:STATE.revGrowth-10},high:{revGrowth:STATE.revGrowth+10}},
    {label:'Provider Comp ±5pp',low:{revShare:STATE.revShare-5},high:{revShare:STATE.revShare+5}},
    {label:'OPEX ±10%',low:{opexAdj:STATE.opexAdj-10},high:{opexAdj:STATE.opexAdj+10}},
    {label:'Staff ±10%',low:{staffAdj:STATE.staffAdj-10},high:{staffAdj:STATE.staffAdj+10}},
    {label:'WACC ±2pp',low:{wacc:STATE.wacc-2},high:{wacc:STATE.wacc+2}},
    {label:'Exit Multiple ±1.5x',low:{exitMult:STATE.exitMult-1.5},high:{exitMult:STATE.exitMult+1.5}}
  ];
  return cases.map(item=>{
    const lowNpv=computeModelSnapshot(item.low).npv;
    const highNpv=computeModelSnapshot(item.high).npv;
    return{
      label:item.label,
      downside:Math.min(lowNpv-baseline,highNpv-baseline)/1e6,
      upside:Math.max(lowNpv-baseline,highNpv-baseline)/1e6
    };
  }).sort((a,b)=>Math.max(Math.abs(b.downside),Math.abs(b.upside))-Math.max(Math.abs(a.downside),Math.abs(a.upside)));
}
function initTornadoChart(){
  const el=document.getElementById('chart-tornado');
  if(!el)return;
  const series=getTornadoSeries();
  if(charts.tornado){charts.tornado.destroy();delete charts.tornado;}
  charts.tornado=new Chart(el,{
    type:'bar',
    data:{
      labels:series.map(s=>s.label),
      datasets:[
        {label:'Downside Δ NPV',data:series.map(s=>s.downside),backgroundColor:'rgba(244,63,94,.5)',borderColor:C.rose,borderWidth:1,borderRadius:4,barThickness:16},
        {label:'Upside Δ NPV',data:series.map(s=>s.upside),backgroundColor:'rgba(16,185,129,.45)',borderColor:C.emerald,borderWidth:1,borderRadius:4,barThickness:16}
      ]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      scales:{x:{ticks:{callback:v=>fmtMCx(v*1e6)+'M'},title:{display:true,text:'Δ 10Y NPV'}},y:{grid:{display:false}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${(c.raw>=0?'+':'')+c.raw.toFixed(1)}M ${cxLabel()}`}}}
    }
  });
}
function updateTornadoChart(){
  const series=getTornadoSeries();
  if(!charts.tornado){
    initTornadoChart();
    return;
  }
  charts.tornado.data.labels=series.map(s=>s.label);
  charts.tornado.data.datasets[0].data=series.map(s=>s.downside);
  charts.tornado.data.datasets[1].data=series.map(s=>s.upside);
  charts.tornado.update('active');
}
function getScenarioRecord(name,stateValues,source='custom'){
  const snapshot=computeModelSnapshot(stateValues);
  return{
    id:`scn-${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
    name,
    source,
    values:sliderMap.reduce((acc,item)=>{acc[item.key]=Number(stateValues[item.key]);return acc;},{}),
    metrics:{npv:snapshot.npv,irr10:snapshot.irr10,payback:snapshot.payback},
    createdAt:new Date().toISOString()
  };
}
function buildPresetScenarioLibrary(){
  return[
    getScenarioRecord('Base',SCENARIOS.base,'preset'),
    getScenarioRecord('Optimistic',SCENARIOS.optimistic,'preset'),
    getScenarioRecord('Pessimistic',SCENARIOS.pessimistic,'preset')
  ];
}
function persistScenarioLibrary(){
  writeStoredUI(APP_STORAGE_KEYS.scenarioLibrary,scenarioLibrary,'project');
}
function refreshScenarioLibraryMetrics(persist=true){
  if(!scenarioLibrary.length)return;
  scenarioLibrary=scenarioLibrary.map(record=>{
    const snapshot=computeModelSnapshot(record.values||STATE);
    return{
      ...record,
      metrics:{npv:snapshot.npv,irr10:snapshot.irr10,payback:snapshot.payback}
    };
  });
  if(persist)persistScenarioLibrary();
}
function renderScenarioLibrary(){
  const tbody=document.querySelector('#scenarioLibraryTable tbody');
  if(!tbody)return;
  if(!scenarioLibrary.length){
    tbody.innerHTML=`<tr><td colspan="11" class="text-muted">${currentLanguage==='ru'?'Сохраненных сценариев пока нет.':'No saved scenarios yet.'}</td></tr>`;
    return;
  }
  tbody.innerHTML=scenarioLibrary.map(record=>`
    <tr>
      <td style="font-weight:600">${record.name}</td>
      <td>${record.values.revGrowth>0?'+':''}${record.values.revGrowth}%</td>
      <td>${record.values.revShare}%</td>
      <td>${record.values.opexAdj>0?'+':''}${record.values.opexAdj}%</td>
      <td>${record.values.staffAdj>0?'+':''}${record.values.staffAdj}%</td>
      <td>${record.values.wacc}%</td>
      <td>${record.values.exitMult}×</td>
      <td class="${record.metrics.npv>=0?'pos':'neg'}">${fmtCx(record.metrics.npv)}</td>
      <td class="${record.metrics.irr10>0.12?'pos':'neg'}">${pct(record.metrics.irr10)}</td>
      <td>${formatPaybackValue(record.metrics.payback)}</td>
      <td class="scenario-row-actions">
        <button class="mini-action-btn" data-scenario-action="load" data-scenario-id="${record.id}">${currentLanguage==='ru'?'Загрузить':'Load'}</button>
        ${record.source==='preset'
          ? `<span class="text-muted">${currentLanguage==='ru'?'Preset':'Preset'}</span>`
          : `<button class="mini-action-btn mini-action-btn-danger" data-scenario-action="delete" data-scenario-id="${record.id}">${currentLanguage==='ru'?'Удалить':'Delete'}</button>`}
      </td>
    </tr>
  `).join('');
}
function saveCurrentScenario(){
  const input=document.getElementById('scenarioNameInput');
  const baseName=input?.value.trim();
  const name=baseName||`Scenario ${scenarioLibrary.filter(r=>r.source!=='preset').length+1}`;
  scenarioLibrary.unshift(getScenarioRecord(name,STATE,'saved'));
  persistScenarioLibrary();
  renderScenarioLibrary();
  if(input)input.value='';
  showToast(currentLanguage==='ru'?`💾 ${name} сохранен`:`💾 ${name} saved`,'success');
}
function loadScenarioRecord(id){
  const record=scenarioLibrary.find(item=>item.id===id);
  if(!record)return;
  Object.assign(STATE,record.values);
  syncSliders();
  updateScenarioHighlight();
  updateAll();
  showToast(currentLanguage==='ru'?`↺ ${record.name} загружен`:`↺ ${record.name} loaded`,'success');
}
function deleteScenarioRecord(id){
  scenarioLibrary=scenarioLibrary.filter(item=>item.id!==id||item.source==='preset');
  persistScenarioLibrary();
  renderScenarioLibrary();
}
function clearSavedScenarios(){
  scenarioLibrary=scenarioLibrary.filter(item=>item.source==='preset');
  persistScenarioLibrary();
  renderScenarioLibrary();
}
function renderAssumptionsPanel(){
  const panel=document.getElementById('assumptionsPanel');
  if(!panel)return;
  const anchor=getSensitivityAnchor();
  const anchorLabel=currentLanguage==='ru'&&anchor.contextLabel==='Y5 anchor'?'якорь Y5':anchor.contextLabel;
  panel.innerHTML=`
    <div class="module-list">
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Рост выручки':'Revenue growth'}</span><span class="module-val">${STATE.revGrowth>0?'+':''}${STATE.revGrowth}%</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Comp провайдеров':'Provider compensation'}</span><span class="module-val">${STATE.revShare}%</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Коррекция OPEX':'OPEX adj.'}</span><span class="module-val">${STATE.opexAdj>0?'+':''}${STATE.opexAdj}%</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Коррекция fixed staff':'Fixed staff adj.'}</span><span class="module-val">${STATE.staffAdj>0?'+':''}${STATE.staffAdj}%</span></div>
      <div class="module-row"><span class="module-key">WACC</span><span class="module-val">${STATE.wacc}%</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Exit multiple':'Exit multiple'}</span><span class="module-val">${STATE.exitMult}×</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Якорь sensitivity':'Sensitivity anchor'}</span><span class="module-val">${anchorLabel}</span></div>
      <div class="module-row"><span class="module-key">MC profile</span><span class="module-val">${getMonteCarloProfileLabel()}</span></div>
    </div>
    <div class="module-callout">${currentLanguage==='ru'
      ? 'Определения: EBITDA учитывает текущую compensation провайдеров, non-revenue OPEX и fixed staff adjustment. NPV и IRR считаются из того же snapshot модели, который используется в sensitivity и scenario compare.'
      : 'Definitions: EBITDA uses current provider compensation, non-revenue OPEX, and fixed staff adjustment. NPV and IRR are computed from the same model snapshot used across sensitivity and scenario compare.'}</div>
  `;
}
function renderSourcesPanel(){
  const panel=document.getElementById('sourcesPanel');
  if(!panel)return;
  const sourceInfo=describeModelSource();
  panel.innerHTML=`
    <div class="module-tags">
      <span class="module-tag">${currentLanguage==='ru'?'Model workbook':'Model workbook'}</span>
      <span class="module-tag">P&L / Monthly P&L</span>
      <span class="module-tag">Cash Flow & DCF</span>
      <span class="module-tag">${currentLanguage==='ru'?'Привлечение пациентов':'Patient Acquisition'}</span>
      <span class="module-tag">Capacity</span>
    </div>
    <div class="module-list">
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Benchmark band':'Benchmark band'}</span><span class="module-val">${currentLanguage==='ru'?'Dubai outpatient market, внутренний management benchmark':'Dubai outpatient market, internal management benchmark'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Service mix':'Service mix'}</span><span class="module-val">${currentLanguage==='ru'?'Внутренняя модель по годам':'Internal model assumption by year'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Текущий источник модели':'Current model source'}</span><span class="module-val">${sourceInfo.sourceLabel}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Статус сессии':'Session status'}</span><span class="module-val">${sourceInfo.statusLabel}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Последний импорт workbook':'Last workbook import'}</span><span class="module-val">${importAudit.fileName||(currentLanguage==='ru'?'Нет workbook':'No workbook')}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Листы последнего импорта':'Sheets in last import'}</span><span class="module-val">${importAudit.sheets?.length||0} ${currentLanguage==='ru'?'обнаружено':'detected'}</span></div>
    </div>
    <div class="module-callout ${sourceInfo.tone}">${sourceInfo.note}</div>
    <div class="module-callout warning">${currentLanguage==='ru'
      ? 'Benchmark cards используются как decision support, а не как аудитированные рыночные данные. Если дашборд идет во внешние investor materials, добавьте датированные внешние источники.'
      : 'Benchmark cards are decision support, not audited market data. Add dated external source links if this dashboard is used in investor materials.'}</div>
  `;
}
function renderImportAudit(){
  const panel=document.getElementById('importAuditPanel');
  if(!panel)return;
  const importedAt=formatDateTimeLocalized(importAudit.importedAt);
  const sourceInfo=describeModelSource();
  panel.innerHTML=`
    <div class="module-audit-grid">
      <div class="module-metric-card ${sourceInfo.tone}"><div class="module-metric-label">${currentLanguage==='ru'?'Текущая модель':'Current model'}</div><div class="module-metric-value">${sourceInfo.sourceLabel}</div></div>
      <div class="module-metric-card ${sourceInfo.tone}"><div class="module-metric-label">${currentLanguage==='ru'?'Статус сессии':'Session status'}</div><div class="module-metric-value">${sourceInfo.statusLabel}</div></div>
      <div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Последний импорт':'Last import'}</div><div class="module-metric-value">${importedAt}</div></div>
      <div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Смаппленные метрики':'Mapped metrics'}</div><div class="module-metric-value">${importAudit.mappedMetrics?.length||0}</div></div>
    </div>
    <div class="module-callout ${sourceInfo.tone}">${sourceInfo.note}</div>
    <div class="module-list">
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Последний импортированный файл':'Last imported workbook'}</span><span class="module-val">${importAudit.fileName||(currentLanguage==='ru'?'Нет workbook':'No workbook')}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Листы последнего импорта':'Detected sheets from last import'}</span><span class="module-val">${(importAudit.sheets||[]).join(', ')||'—'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Смаппленные поля':'Mapped fields'}</span><span class="module-val">${(importAudit.mappedMetrics||[]).join(', ')||'—'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Готовность факта':'Actuals readiness'}</span><span class="module-val">${getLocalizedActualsStatus()}</span></div>
    </div>
  `;
}
function renderTemplateCompliance(){
  const panel=document.getElementById('templateCompliancePanel');
  if(!panel)return;
  const checks=Array.isArray(templateCompliance.checks)?templateCompliance.checks:[];
  if(!checks.length){
    panel.innerHTML=`<div class="module-callout warning">${currentLanguage==='ru'?'Скан шаблона еще не выполнен. Загрузите Excel workbook, чтобы проверить соответствие обязательному project template.':'No template scan yet. Load an Excel workbook to run compliance checks against the required project template.'}</div>`;
    return;
  }
  const checkedAt=formatDateTimeLocalized(templateCompliance.checkedAt);
  const summary=templateCompliance.summary||{ok:0,fail:0,warn:0,optional:0};
  const severityTotals=checks.reduce((acc,check)=>{
    acc[check.severity]=(acc[check.severity]||0)+1;
    return acc;
  },{fail:0,warn:0,optional:0});
  const severityRank={fail:0,warn:1,optional:2};
  const orderedChecks=[...checks].sort((a,b)=>{
    const severityDelta=(severityRank[a.severity]??9)-(severityRank[b.severity]??9);
    if(severityDelta!==0)return severityDelta;
    const missingDelta=Number(a.ok)-Number(b.ok);
    if(missingDelta!==0)return missingDelta;
    const areaDelta=String(a.area).localeCompare(String(b.area));
    if(areaDelta!==0)return areaDelta;
    return String(a.item).localeCompare(String(b.item));
  });
  const tone=templateCompliance.ready?(summary.warn||summary.optional?'warning':'positive'):'negative';
  const callout=templateCompliance.ready
    ? (summary.warn||summary.optional
      ? (currentLanguage==='ru'
        ? 'Workbook можно импортировать, но часть неблокирующих элементов шаблона отсутствует. Базовая модель загрузится, а вторичные модули могут использовать fallback.'
        : 'Workbook is importable, but some non-blocking template items are missing. Core model will load; secondary modules may rely on fallbacks.')
      : (currentLanguage==='ru'
        ? 'Workbook полностью соответствует требуемому шаблону. Импорт может идти без структурных предупреждений.'
        : 'Workbook is fully compliant with the required template. Import can proceed without structural warnings.'))
    : (currentLanguage==='ru'
      ? 'Импорт заблокирован: отсутствует один или несколько критичных листов или строк шаблона. Дашборд сохранит предыдущую загруженную модель, пока workbook не будет приведен к шаблону.'
      : 'Import is blocked: one or more critical template sheets or rows are missing. Dashboard will keep the previous loaded model until the workbook matches the template.');
  const isExpanded=templateComplianceDetailsExpanded;
  const activeFilter=['all','fail','warn','optional'].includes(templateComplianceFilter)?templateComplianceFilter:'all';
  const filteredChecks=orderedChecks.filter(check=>activeFilter==='all'||check.severity===activeFilter);
  const filterTabs=[
    {key:'all',label:`${currentLanguage==='ru'?'Все':'All'} ${checks.length}`},
    {key:'fail',label:`${currentLanguage==='ru'?'Critical':'Critical'} ${severityTotals.fail||0}`},
    {key:'warn',label:`${currentLanguage==='ru'?'Warning':'Warning'} ${severityTotals.warn||0}`},
    {key:'optional',label:`${currentLanguage==='ru'?'Optional':'Optional'} ${severityTotals.optional||0}`}
  ];
  panel.innerHTML=`
    <div class="module-audit-grid">
      <div class="module-metric-card ${templateCompliance.ready?'positive':'negative'}"><div class="module-metric-label">${currentLanguage==='ru'?'Статус':'Status'}</div><div class="module-metric-value">${templateCompliance.ready?(currentLanguage==='ru'?'Готов':'Ready'):(currentLanguage==='ru'?'Blocked':'Blocked')}</div></div>
      <div class="module-metric-card ${summary.fail?'negative':'positive'}"><div class="module-metric-label">${currentLanguage==='ru'?'Critical gaps':'Critical Gaps'}</div><div class="module-metric-value">${summary.fail||0}</div></div>
      <div class="module-metric-card ${summary.warn?'warning':'positive'}"><div class="module-metric-label">${currentLanguage==='ru'?'Warnings':'Warnings'}</div><div class="module-metric-value">${summary.warn||0}</div></div>
      <div class="module-metric-card ${summary.optional?'warning':'positive'}"><div class="module-metric-label">${currentLanguage==='ru'?'Optional gaps':'Optional Gaps'}</div><div class="module-metric-value">${summary.optional||0}</div></div>
    </div>
    <div class="module-list">
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Проверенный файл':'Checked file'}</span><span class="module-val">${templateCompliance.fileName||'Workbook'}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Проверено':'Checked at'}</span><span class="module-val">${checkedAt}</span></div>
      <div class="module-row"><span class="module-key">${currentLanguage==='ru'?'Детали':'Details'}</span><button class="module-btn-ghost" id="templateComplianceToggle" type="button" aria-expanded="${String(isExpanded)}">${isExpanded?tr('hideDetails'):tr('showDetails')}</button></div>
    </div>
    <div class="module-callout ${tone}">${callout}</div>
    <div id="templateComplianceDetails" ${isExpanded?'':'hidden'}>
    <div class="chart-tabs template-filter-tabs" id="templateComplianceFilters">
      ${filterTabs.map(tab=>`<button class="chart-tab ${tab.key===activeFilter?'active':''}" type="button" data-template-filter="${tab.key}">${tab.label}</button>`).join('')}
    </div>
    <div class="table-scroll"><table class="data-table">
      <thead><tr><th>${currentLanguage==='ru'?'Область':'Area'}</th><th>${currentLanguage==='ru'?'Требование':'Requirement'}</th><th>${currentLanguage==='ru'?'Severity':'Severity'}</th><th>${currentLanguage==='ru'?'Статус':'Status'}</th><th>${currentLanguage==='ru'?'Комментарий':'Detail'}</th></tr></thead>
      <tbody>
        ${filteredChecks.map(check=>{
          const severityLabel=check.severity==='fail'?(currentLanguage==='ru'?'Critical':'Critical'):check.severity==='warn'?(currentLanguage==='ru'?'Warning':'Warning'):(currentLanguage==='ru'?'Optional':'Optional');
          const statusLabel=check.ok?(currentLanguage==='ru'?'Есть':'Present'):(currentLanguage==='ru'?'Missing':'Missing');
          const severityTone=check.severity==='fail'?'fail':'warn';
          return `
            <tr>
              <td style="font-weight:600">${check.area}</td>
              <td>${check.item}</td>
              <td><span class="audit-status ${severityTone}">${severityLabel}</span></td>
              <td><span class="audit-status ${check.ok?'ok':check.status}">${statusLabel}</span></td>
              <td>${check.detail}</td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table></div></div>
  `;
  panel.querySelector('#templateComplianceToggle')?.addEventListener('click',()=>setTemplateComplianceDetailsExpanded(!templateComplianceDetailsExpanded));
  panel.querySelectorAll('[data-template-filter]').forEach(btn=>{
    btn.addEventListener('click',()=>setTemplateComplianceFilter(btn.dataset.templateFilter));
  });
}
function renderExcelChecklist(){
  const panel=document.getElementById('excelChecklistPanel');
  if(!panel)return;
  const checks=Array.isArray(excelChecklist.checks)?excelChecklist.checks:[];
  if(!checks.length){
    panel.innerHTML=`<div class="module-callout warning">${currentLanguage==='ru'?'Workbook checklist пока недоступен. Загрузите Excel-модель, чтобы сформировать reconciliation checks.':'No workbook checklist available yet. Load an Excel model to generate reconciliation checks.'}</div>`;
    return;
  }
  const summary=checks.reduce((acc,check)=>{acc[check.status]=(acc[check.status]||0)+1;return acc;},{ok:0,warn:0,fail:0});
  const scenarioNote=currentLanguage==='ru'
    ? 'Панель ниже не сравнивает workbook с самим собой. Она пересобирает метрики из независимых листов и формул workbook и игнорирует live what-if слайдеры.'
    : 'The panel below does not compare the workbook to itself. It rebuilds metrics from independent workbook sheets and formulas, and intentionally ignores live what-if sliders.';
  panel.innerHTML=`
    <div class="module-audit-grid">
      <div class="module-metric-card positive"><div class="module-metric-label">OK</div><div class="module-metric-value">${summary.ok||0}</div></div>
      <div class="module-metric-card warning"><div class="module-metric-label">${currentLanguage==='ru'?'Minor':'Minor'}</div><div class="module-metric-value">${summary.warn||0}</div></div>
      <div class="module-metric-card ${summary.fail?'negative':'positive'}"><div class="module-metric-label">${currentLanguage==='ru'?'Mismatch':'Mismatch'}</div><div class="module-metric-value">${summary.fail||0}</div></div>
      <div class="module-metric-card"><div class="module-metric-label">${currentLanguage==='ru'?'Baseline':'Baseline'}</div><div class="module-metric-value">${excelChecklist.fileName||'Workbook'} @ ${excelChecklist.wacc}% WACC</div></div>
    </div>
    <div class="module-callout ${summary.fail?'warning':'positive'}">${scenarioNote}</div>
    <div class="table-scroll"><table class="data-table">
      <thead><tr><th>${currentLanguage==='ru'?'Метрика':'Metric'}</th><th>${currentLanguage==='ru'?'Источник':'Source'}</th><th>Workbook</th><th>${currentLanguage==='ru'?'Пересборка':'Rebuilt'}</th><th>Δ</th><th>${currentLanguage==='ru'?'Статус':'Status'}</th></tr></thead>
      <tbody>
        ${checks.map(check=>{
          const deltaClass=check.delta===0?'':(check.delta>0?'pos':'neg');
          const statusLabel=check.status==='ok'?'OK':check.status==='warn'?(currentLanguage==='ru'?'Minor':'Minor'):(currentLanguage==='ru'?'Mismatch':'Mismatch');
          return `
            <tr>
              <td style="font-weight:600">${check.metric}</td>
              <td>${check.source}</td>
              <td>${fmtAuditValue(check.kind,check.workbookValue)}</td>
              <td>${fmtAuditValue(check.kind,check.dashboardValue)}</td>
              <td class="${deltaClass}">${fmtAuditDelta(check.kind,check.delta)}</td>
              <td><span class="audit-status ${check.status}">${statusLabel}</span></td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table></div>
  `;
}
function initScenarioLibraryModule(){
  scenarioLibrary=readStoredUI(APP_STORAGE_KEYS.scenarioLibrary,[],'project');
  if(!scenarioLibrary.length){
    scenarioLibrary=buildPresetScenarioLibrary();
    persistScenarioLibrary();
  }
  refreshScenarioLibraryMetrics(false);
  const saveBtn=document.getElementById('scenarioSaveBtn');
  const clearBtn=document.getElementById('scenarioClearBtn');
  const table=document.getElementById('scenarioLibraryTable');
  if(saveBtn&&!saveBtn.dataset.bound){
    saveBtn.addEventListener('click',saveCurrentScenario);
    saveBtn.dataset.bound='true';
  }
  if(clearBtn&&!clearBtn.dataset.bound){
    clearBtn.addEventListener('click',clearSavedScenarios);
    clearBtn.dataset.bound='true';
  }
  if(table&&!table.dataset.bound){
    table.addEventListener('click',e=>{
      const btn=e.target.closest('[data-scenario-action]');
      if(!btn)return;
      const id=btn.dataset.scenarioId;
      if(btn.dataset.scenarioAction==='load')loadScenarioRecord(id);
      if(btn.dataset.scenarioAction==='delete')deleteScenarioRecord(id);
    });
    table.dataset.bound='true';
  }
  renderScenarioLibrary();
}

/* ===== INTEGRITY CHECKS ===== */
function initChecks(){
  document.getElementById('checksGrid').innerHTML=BASE.checks.map(c=>`<div class="check-item ${c.pass?'check-pass':'check-warn'}"><span class="check-icon">${c.pass?'✓':'⚠'}</span><span>${c.label}</span><span class="check-delta">${c.delta}</span></div>`).join('');
}

/* ===== CAPEX + AMORTIZATION ===== */
function calcDepreciation(){
  // Returns 10-year depreciation schedule: depSchedule[year][category]
  const dep=[];
  for(let y=0;y<10;y++){
    dep[y]=BASE.capexAmounts.map((amt,i)=>{
      const life=BASE.capexLife[i];
      if(life===0||y>=life)return 0;
      return amt/life;
    });
  }
  return dep;
}
function calcNBV(){
  // Net Book Value per year
  const dep=calcDepreciation();
  const nbv=[BASE.investment];
  for(let y=0;y<10;y++){
    const totalDep=dep[y].reduce((s,v)=>s+v,0);
    nbv.push(nbv[y]-totalDep);
  }
  return nbv;
}
function initCapexCharts(){
  if(charts.capexBar)return;
  // 1. CAPEX Breakdown horizontal bar
  charts.capexBar=new Chart(document.getElementById('chart-capex-bar'),{
    type:'bar',
    data:{
      labels:BASE.capexLabels,
      datasets:[{
        data:BASE.capexAmounts,
        backgroundColor:BASE.capexColors.map(c=>c+'cc'),
        borderColor:BASE.capexColors,
        borderWidth:1,
        borderRadius:4,
        barThickness:16
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,indexAxis:'y',
      scales:{
        x:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)},grid:{color:'rgba(255,255,255,.04)'}},
        y:{grid:{display:false}}
      },
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}}
    }
  });
  // 2. Depreciation stacked area
  const dep=calcDepreciation();
  const depDatasets=BASE.capexLabels.map((label,i)=>{
    if(BASE.capexLife[i]===0)return null;
    return {
      label:label,
      data:dep.map(y=>cx(y[i])),
      backgroundColor:BASE.capexColors[i]+'40',
      borderColor:BASE.capexColors[i],
      borderWidth:1.5,
      fill:true,
      pointRadius:2,
      tension:.3
    };
  }).filter(Boolean);
  charts.depArea=new Chart(document.getElementById('chart-dep-area'),{
    type:'line',
    data:{labels:BASE.years,datasets:depDatasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:'index'},
      scales:{
        y:{stacked:true,beginAtZero:true,ticks:{callback:v=>fmtCx(v)}}
      },
      plugins:{legend:{position:'bottom',labels:{font:{size:9},padding:6,usePointStyle:true}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmtCx(c.raw)+' '+cxLabel()}}}
    }
  });
  // 3. Net Book Value line
  const nbv=calcNBV();
  charts.nbvLine=new Chart(document.getElementById('chart-nbv'),{
    type:'line',
    data:{
      labels:['Y0',...BASE.years],
      datasets:[{
        label:'Net Book Value',
        data:nbv.map(v=>cx(v)),
        borderColor:C.cyan,
        backgroundColor:'rgba(34,211,238,.15)',
        fill:true,
        pointBackgroundColor:C.cyan,
        pointRadius:4,
        borderWidth:2,
        tension:.3
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      scales:{y:{beginAtZero:true,ticks:{callback:v=>fmtCx(v)}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}}
    }
  });
  // 4. Amortization table
  buildAmortTable();
}
function buildAmortTable(){
  const dep=calcDepreciation();
  const tbody=document.querySelector('#amortTable tbody');
  if(!tbody)return;
  let html='';
  BASE.capexLabels.forEach((label,i)=>{
    if(BASE.capexLife[i]===0)return;
    html+=`<tr><td>${label}</td><td>${fmtCx(BASE.capexAmounts[i])}</td><td>${BASE.capexLife[i]}y</td>`;
    for(let y=0;y<10;y++){
      const v=dep[y][i];
      html+=`<td class="${v>0?'':'text-muted'}">${v>0?fmtCx(v):'-'}</td>`;
    }
    html+='</tr>';
  });
  // Total row
  html+='<tr class="total-row"><td><strong>Total D&A</strong></td><td><strong>'+fmtCx(BASE.capexAmounts.reduce((s,v,i)=>s+(BASE.capexLife[i]>0?v:0),0))+'</strong></td><td></td>';
  for(let y=0;y<10;y++){
    const total=dep[y].reduce((s,v)=>s+v,0);
    html+=`<td><strong>${fmtCx(total)}</strong></td>`;
  }
  html+='</tr>';
  tbody.innerHTML=html;
}

/* ===== REVENUE MIX ===== */
function getServiceSelectionSnapshot(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  const label=getSelectedYearLabel(selection);
  const revenueByService=BASE.services.map((_,svcIdx)=>years.reduce((sum,yearIdx)=>{
    const baseTotal=Number(BASE.revenue[yearIdx])||0;
    const currentTotal=Number(ADJ.revenue[yearIdx])||0;
    const revenueScale=baseTotal>0?currentTotal/baseTotal:1;
    const mixShare=((BASE.serviceMix?.[yearIdx]?.[svcIdx]||0)/100);
    const baseDeptRevenue=Number(BASE.serviceRevenue?.[svcIdx]?.[yearIdx])||baseTotal*mixShare;
    return sum+baseDeptRevenue*revenueScale;
  },0));
  const visitsByService=BASE.services.map((_,svcIdx)=>years.reduce((sum,yearIdx)=>sum+(Number(BASE.serviceVisits?.[svcIdx]?.[yearIdx])||0),0));
  const totalRevenue=revenueByService.reduce((sum,value)=>sum+value,0);
  const mix=revenueByService.map(value=>totalRevenue>0?value/totalRevenue*100:0);
  const avgCheck=revenueByService.map((value,svcIdx)=>{
    const visits=visitsByService[svcIdx]||0;
    if(visits>0)return value/visits;
    const selectionBaseRevenue=years.reduce((sum,yearIdx)=>sum+(Number(BASE.serviceRevenue?.[svcIdx]?.[yearIdx])||0),0);
    const selectionScale=selectionBaseRevenue>0?value/selectionBaseRevenue:1;
    return (Number(BASE.revPerVisit?.[svcIdx])||0)*selectionScale;
  });
  return {years,label,revenueByService,visitsByService,totalRevenue,mix,avgCheck};
}
function getProductLabel(row){
  const dept=row.department.replace(' Clinic','');
  return `${dept} • ${row.product}`;
}
function getProductSelectionSnapshot(selection=selectedYear){
  const years=normalizeYearSelection(selection);
  const label=getSelectedYearLabel(selection);
  const catalog=Array.isArray(BASE.productCatalog)?BASE.productCatalog:[];
  const departmentBaseTotals=new Map();

  catalog.forEach(item=>{
    const deptTotals=departmentBaseTotals.get(item.department)||Array(10).fill(0);
    item.monthlyQty.forEach((qty,yearIdx)=>{
      deptTotals[yearIdx]+=Math.max(0,Number(qty)||0)*12*Math.max(0,Number(item.basePrice)||0);
    });
    departmentBaseTotals.set(item.department,deptTotals);
  });

  const rows=catalog.map(item=>{
    const deptIdx=BASE.services.indexOf(item.department);
    let units=0;
    let revenue=0;

    years.forEach(yearIdx=>{
      const monthlyQty=Math.max(0,Number(item.monthlyQty?.[yearIdx])||0);
      const yearlyUnits=monthlyQty*12;
      units+=yearlyUnits;

      const deptRevenueBase=Math.max(0,Number(BASE.serviceRevenue?.[deptIdx]?.[yearIdx])||0);
      const baseTotalRevenue=Math.max(0,Number(BASE.revenue[yearIdx])||0);
      const currentTotalRevenue=Math.max(0,Number(ADJ.revenue[yearIdx])||0);
      const revenueScale=baseTotalRevenue>0?currentTotalRevenue/baseTotalRevenue:1;
      const deptRevenueCurrent=deptRevenueBase*revenueScale;
      const deptPriceListBase=(departmentBaseTotals.get(item.department)?.[yearIdx])||0;
      const baseProductRevenue=yearlyUnits*Math.max(0,Number(item.basePrice)||0);
      const deptScale=deptPriceListBase>0?deptRevenueCurrent/deptPriceListBase:1;
      revenue+=baseProductRevenue*deptScale;
    });

    return {
      department:item.department,
      product:item.product,
      label:getProductLabel(item),
      units,
      revenue,
      avgPrice:units>0?revenue/units:0,
      color:BASE.serviceColors[Math.max(BASE.services.indexOf(item.department),0)]||C.indigo
    };
  }).filter(row=>row.units>0||row.revenue>0);

  const totalRevenue=rows.reduce((sum,row)=>sum+row.revenue,0);
  rows.forEach(row=>{row.mix=totalRevenue>0?row.revenue/totalRevenue*100:0;});
  rows.sort((a,b)=>b.revenue-a.revenue);
  return {years,label,rows,totalRevenue};
}
function renderProductRevenueModule(){
  const tbody=document.querySelector('#productEconomicsTable tbody');
  if(!tbody)return;
  const snapshot=getProductSelectionSnapshot();
  setText('productMixTitle',`Revenue by Product ${snapshot.label}`);
  setText('productEconomicsTitle',`Product Economics (${snapshot.label})`);
  if(!snapshot.rows.length){
    tbody.innerHTML='<tr><td colspan="6" class="text-muted">No product-level rows available.</td></tr>';
    return;
  }
  tbody.innerHTML=snapshot.rows.map(row=>`<tr>
    <td style="font-weight:600">${row.product}</td>
    <td>${row.department}</td>
    <td>${row.mix.toFixed(1)}%</td>
    <td>${fmt(row.units)}</td>
    <td>${fmtCx(row.avgPrice)}</td>
    <td class="pos">${fmtMCx(row.revenue)}M</td>
  </tr>`).join('');
}
function initRevenueMixCharts(){
  if(charts.svcDonut)return;
  const serviceSelection=getServiceSelectionSnapshot();
  const productSelection=getProductSelectionSnapshot();
  setText('svcMixTitle',`Revenue Mix ${serviceSelection.label}`);
  setText('svcAvgTitle',tr('avgCheckDept',{label:serviceSelection.label}));
  setText('productMixTitle',`Revenue by Product ${productSelection.label}`);
  // 1. Doughnut for selected period
  charts.svcDonut=new Chart(document.getElementById('chart-svc-donut'),{
    type:'doughnut',
    data:{
      labels:BASE.services,
      datasets:[{data:serviceSelection.mix,backgroundColor:BASE.serviceColors.map(c=>c+'cc'),borderWidth:0}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{position:'right',labels:{font:{size:10},padding:6}},
        tooltip:{callbacks:{label:c=>c.label+': '+(Number(c.raw)||0).toFixed(1)+'%'}}}
    }
  });
  // 2. Stacked area for mix evolution
  const mixDatasets=BASE.services.map((svc,i)=>({
    label:svc,
    data:BASE.serviceMix.map(y=>y[i]),
    backgroundColor:BASE.serviceColors[i]+'40',
    borderColor:BASE.serviceColors[i],
    borderWidth:1.5,
    fill:true,
    pointRadius:2,
    tension:.3
  }));
  charts.svcArea=new Chart(document.getElementById('chart-svc-area'),{
    type:'line',
    data:{labels:BASE.years,datasets:mixDatasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:'index'},
      scales:{y:{stacked:true,max:100,ticks:{callback:v=>v+'%'}}},
      plugins:{legend:{position:'bottom',labels:{font:{size:9},padding:6,usePointStyle:true}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+(Number(c.raw)||0).toFixed(1)+'%'}}}
    }
  });
  // 3. Revenue per visit horizontal bar for selected period
  charts.svcRevBar=new Chart(document.getElementById('chart-svc-rev'),{
    type:'bar',
    data:{
      labels:BASE.services,
      datasets:[{
        data:serviceSelection.avgCheck.map(v=>cx(v)),
        backgroundColor:BASE.serviceColors.map(c=>c+'99'),
        borderColor:BASE.serviceColors,
        borderWidth:1,
        borderRadius:4,
        barThickness:16
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,indexAxis:'y',
      scales:{
        x:{beginAtZero:true,ticks:{callback:v=>fmt(v)},title:{display:true,text:'Avg '+cxLabel()+'/visit'}},
        y:{grid:{display:false}}
      },
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmt(c.raw)+' '+cxLabel()}}}
    }
  });
  charts.productRevenue=new Chart(document.getElementById('chart-product-revenue'),{
    type:'bar',
    data:{
      labels:productSelection.rows.map(row=>row.label),
      datasets:[{
        data:productSelection.rows.map(row=>cx(row.revenue)),
        backgroundColor:productSelection.rows.map(row=>row.color+'99'),
        borderColor:productSelection.rows.map(row=>row.color),
        borderWidth:1,
        borderRadius:4,
        barThickness:14
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,indexAxis:'y',
      scales:{
        x:{beginAtZero:true,ticks:{callback:v=>fmt(v)},title:{display:true,text:'Revenue ('+cxLabel()+')'}},
        y:{grid:{display:false},ticks:{autoSkip:false}}
      },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>{
          const row=productSelection.rows[c.dataIndex];
          return `${fmt(c.raw)} ${cxLabel()} | ${row.mix.toFixed(1)}% mix | ${fmt(row.units)} units @ ${fmtCx(row.avgPrice)}`;
        }}}
      }
    }
  });
}
function updateRevenueMixCharts(){
  if(!charts.svcDonut||!charts.svcRevBar||!charts.productRevenue)return;
  const serviceSelection=getServiceSelectionSnapshot();
  const productSelection=getProductSelectionSnapshot();
  setText('svcMixTitle',`Revenue Mix ${serviceSelection.label}`);
  setText('svcAvgTitle',tr('avgCheckDept',{label:serviceSelection.label}));
  setText('productMixTitle',`Revenue by Product ${productSelection.label}`);
  charts.svcDonut.data.datasets[0].data=serviceSelection.mix;
  charts.svcDonut.update('active');
  charts.svcRevBar.data.datasets[0].data=serviceSelection.avgCheck.map(v=>cx(v));
  charts.svcRevBar.options.scales.x.title.text='Avg '+cxLabel()+'/visit';
  charts.svcRevBar.update('active');
  charts.productRevenue.data.labels=productSelection.rows.map(row=>row.label);
  charts.productRevenue.data.datasets[0].data=productSelection.rows.map(row=>cx(row.revenue));
  charts.productRevenue.data.datasets[0].backgroundColor=productSelection.rows.map(row=>row.color+'99');
  charts.productRevenue.data.datasets[0].borderColor=productSelection.rows.map(row=>row.color);
  charts.productRevenue.options.scales.x.title.text='Revenue ('+cxLabel()+')';
  charts.productRevenue.options.plugins.tooltip.callbacks.label=c=>{
    const row=productSelection.rows[c.dataIndex];
    return `${fmt(c.raw)} ${cxLabel()} | ${row.mix.toFixed(1)}% mix | ${fmt(row.units)} units @ ${fmtCx(row.avgPrice)}`;
  };
  charts.productRevenue.update('active');
  renderProductRevenueModule();
}

/* ===== MONTHLY DRILL-DOWN ===== */
const seasonality=[1.13,1.20,1.02,1.01,0.94,0.73,0.79,0.79,0.89,1.11,1.18,1.24];
function initMonthlyDrilldown(){
  if(charts.mdRev)return;
  const yr=parseInt(document.querySelector('.md-yr-btn.active')?.dataset.mdyr||'0');
  renderMonthlyCharts(yr);

  document.querySelectorAll('.md-yr-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      document.querySelectorAll('.md-yr-btn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      const y=parseInt(btn.dataset.mdyr);
      if(charts.mdRev){charts.mdRev.destroy();delete charts.mdRev;}
      if(charts.mdEbitda){charts.mdEbitda.destroy();delete charts.mdEbitda;}
      renderMonthlyCharts(y);
      document.querySelector('.monthly-drilldown h3').textContent='📅 Monthly Drill-down (Year '+(y+1)+')';
    });
  });
}
function getMonthlyData(yearIdx){
  if(yearIdx===0)return{rev:BASE.mRev,ebitda:BASE.mEbitda};
  // Estimate Y2/Y3 by distributing annual total with seasonality
  const annualRev=ADJ.revenue[yearIdx];
  const annualEbitda=ADJ.ebitda[yearIdx];
  const totalSeason=seasonality.reduce((s,v)=>s+v,0);
  const rev=seasonality.map(s=>annualRev*s/totalSeason);
  const margin=ADJ.margins[yearIdx];
  const ebitda=rev.map(r=>r*margin);
  return{rev,ebitda};
}
function renderMonthlyCharts(yearIdx){
  const data=getMonthlyData(yearIdx);
  charts.mdRev=new Chart(document.getElementById('chart-md-rev'),{
    type:'bar',
    data:{
      labels:BASE.months,
      datasets:[{label:'Revenue',data:data.rev.map(cx),backgroundColor:data.rev.map(v=>v>0?'rgba(99,102,241,.6)':'rgba(244,63,94,.5)'),borderRadius:4}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}},
      scales:{y:{ticks:{callback:v=>fmtM(v)}}}
    }
  });
  charts.mdEbitda=new Chart(document.getElementById('chart-md-ebitda'),{
    type:'bar',
    data:{
      labels:BASE.months,
      datasets:[{label:'EBITDA',data:data.ebitda.map(cx),backgroundColor:data.ebitda.map(v=>v>=0?'rgba(16,185,129,.6)':'rgba(244,63,94,.5)'),borderRadius:4}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtCx(c.raw)+' '+cxLabel()}}},
      scales:{y:{ticks:{callback:v=>fmtM(v)}}}
    }
  });
}

/* ===== FINANCIAL RATIOS ===== */
function calcRatios(){
  const ratios={
    roic:[],assetTurn:[],ebitdaMargin:[],opLeverage:[],cashConv:[]
  };
  const nbv=calcNBV();
  for(let y=0;y<10;y++){
    ratios.roic[y]=ADJ.netProfit[y]/BASE.investment*100;
    ratios.assetTurn[y]=ADJ.revenue[y]/(nbv[y+1]||1);
    ratios.ebitdaMargin[y]=ADJ.margins[y]*100;
    ratios.opLeverage[y]=y>0?(ADJ.ebitda[y]/ADJ.ebitda[y-1]-1)/(ADJ.revenue[y]/ADJ.revenue[y-1]-1):1;
    ratios.cashConv[y]=ADJ.cf[y+1]/ADJ.ebitda[y]*100;
  }
  return ratios;
}
function initRatioCharts(){
  if(charts.ratioRadar)return;
  const r=calcRatios();
  // Radar: Y1 vs Y5 vs Y10
  const radarLabels=['ROIC %','Asset Turnover','EBITDA Margin %','Op Leverage','Cash Conv %'];
  const normalize=(arr,max)=>arr.map(v=>Math.min(Math.max(v/max*100,0),100));
  const y1=[r.roic[0],r.assetTurn[0],r.ebitdaMargin[0],Math.max(r.opLeverage[0],0),r.cashConv[0]];
  const y5=[r.roic[4],r.assetTurn[4],r.ebitdaMargin[4],Math.max(r.opLeverage[4],0),r.cashConv[4]];
  const y10=[r.roic[9],r.assetTurn[9],r.ebitdaMargin[9],Math.max(r.opLeverage[9],0),r.cashConv[9]];
  const maxVals=[Math.max(...y1,...y5,...y10)*1.2||100,Math.max(y1[1],y5[1],y10[1])*1.2||10,50,5,150];
  charts.ratioRadar=new Chart(document.getElementById('chart-ratio-radar'),{
    type:'radar',
    data:{
      labels:radarLabels,
      datasets:[
        {label:'Y1',data:normalize(y1,100),borderColor:C.rose,backgroundColor:'rgba(244,63,94,.1)',pointBackgroundColor:C.rose,borderWidth:2},
        {label:'Y5',data:normalize(y5,100),borderColor:C.indigo,backgroundColor:'rgba(99,102,241,.15)',pointBackgroundColor:C.indigo,borderWidth:2},
        {label:'Y10',data:normalize(y10,100),borderColor:C.emerald,backgroundColor:'rgba(16,185,129,.1)',pointBackgroundColor:C.emerald,borderWidth:2}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      scales:{r:{angleLines:{color:'rgba(255,255,255,.06)'},grid:{color:'rgba(255,255,255,.06)'},pointLabels:{color:'#94a3b8',font:{size:10}},ticks:{display:false}}},
      plugins:{legend:{position:'bottom',labels:{font:{size:10},padding:8}}}
    }
  });
  // Multi-line ratio trends
  charts.ratioTrend=new Chart(document.getElementById('chart-ratio-trend'),{
    type:'line',
    data:{
      labels:BASE.years,
      datasets:[
        {label:'ROIC %',data:r.roic,borderColor:C.indigo,backgroundColor:'rgba(99,102,241,.1)',fill:false,pointRadius:3,borderWidth:2,tension:.3},
        {label:'EBITDA Margin %',data:r.ebitdaMargin,borderColor:C.emerald,backgroundColor:'rgba(16,185,129,.1)',fill:false,pointRadius:3,borderWidth:2,tension:.3},
        {label:'Cash Conv %',data:r.cashConv,borderColor:C.cyan,backgroundColor:'rgba(34,211,238,.1)',fill:false,pointRadius:3,borderWidth:2,tension:.3}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:'index'},
      scales:{y:{ticks:{callback:v=>v.toFixed(0)+'%'}}},
      plugins:{legend:{position:'bottom',labels:{font:{size:10},padding:8,usePointStyle:true}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+c.raw.toFixed(1)+'%'}}}
    }
  });
  // KPI cards for key ratios
  const ratioCards=document.getElementById('ratioCards');
  if(ratioCards){
    ratioCards.innerHTML=`
      <div class="ratio-card"><div class="ratio-label">ROIC (Y5)</div><div class="ratio-value" style="color:${C.indigo}">${r.roic[4].toFixed(1)}%</div><div class="ratio-bar"><div style="width:${Math.min(r.roic[4],100)}%;background:${C.indigo}"></div></div></div>
      <div class="ratio-card"><div class="ratio-label">Asset Turnover (Y5)</div><div class="ratio-value" style="color:${C.cyan}">${r.assetTurn[4].toFixed(1)}×</div><div class="ratio-bar"><div style="width:${Math.min(r.assetTurn[4]*10,100)}%;background:${C.cyan}"></div></div></div>
      <div class="ratio-card"><div class="ratio-label">EBITDA Margin (Y5)</div><div class="ratio-value" style="color:${C.emerald}">${r.ebitdaMargin[4].toFixed(1)}%</div><div class="ratio-bar"><div style="width:${Math.min(r.ebitdaMargin[4]*2,100)}%;background:${C.emerald}"></div></div></div>
      <div class="ratio-card"><div class="ratio-label">Cash Conversion (Y5)</div><div class="ratio-value" style="color:${C.amber}">${r.cashConv[4].toFixed(0)}%</div><div class="ratio-bar"><div style="width:${Math.min(r.cashConv[4],100)}%;background:${C.amber}"></div></div></div>
    `;
  }
}

/* ===== MONTE CARLO SIMULATION ===== */
function sampleMonteCarloCase(profile){
  const macro=clamp(randNorm()*profile.macroSd+profile.downsideShift,-0.28,0.28);
  const execution=clamp(macro*0.35+randNorm()*profile.executionSd+profile.downsideShift*0.6,-0.24,0.24);
  const pricing=clamp(macro*0.25+randNorm()*profile.pricingSd,-0.12,0.12);
  const costPressure=clamp(-0.4*macro+0.55*execution+randNorm()*profile.costSd+Math.max(-profile.downsideShift,0)*0.3,-0.10,0.18);
  const discountRate=clamp(STATE.wacc/100+randNorm()*profile.discountSd+(-macro)*0.015+Math.max(costPressure,0)*0.08,0.08,0.22);
  const downsideBias=clamp(Math.max(-macro,0)+Math.max(-execution,0)+Math.max(costPressure,0)*0.8,0,0.45);
  const tails=sampleTailEvents(profile,downsideBias);
  const exitMultBase=clamp(STATE.exitMult+macro*1.6-Math.max(costPressure,0)*2.2+randNorm()*profile.exitSd,4,12);
  const exitTailPenalty=tails.launchDelayMonths*0.12+tails.occupancyShortfall*3.2+tails.pricePressure*5.5;
  const exitMult=clamp(exitMultBase-exitTailPenalty,4,12);
  const revShareBase=clamp(STATE.revShare+costPressure*12+randNorm()*0.8,34,52);
  const opexBaseAdj=clamp(STATE.opexAdj+costPressure*70+execution*20,-25,35);
  const staffBaseAdj=clamp(STATE.staffAdj+costPressure*50+execution*18,-20,30);
  const delayFrac=tails.launchDelayMonths/12;
  const initialInvestment=BASE.investment*(1+tails.capexOverrun+tails.launchDelayMonths*0.0045);
  const cfs=[-initialInvestment];
  let demandCarry=clamp(macro*0.9+execution*0.65+profile.downsideShift*0.6,-0.32,0.32);
  let y5Margin=0;
  let y5Ebitda=0;

  for(let y=0;y<10;y++){
    const centeredRevenue=BASE.revenue[y]*(1+STATE.revGrowth/100);
    const yearNoise=clamp(randNorm()*profile.yearlySd,-0.08,0.08);
    demandCarry=clamp(demandCarry*0.78+macro*0.16+yearNoise,-0.36,0.42);
    const revenueShock=clamp(demandCarry+pricing*0.35+(y/9)*macro*0.05,-0.45,0.50);
    let revenue=centeredRevenue*Math.max(0.40,1+revenueShock);
    if(tails.launchDelayMonths){
      if(y===0)revenue*=Math.max(0.05,1-delayFrac*1.2);
      else if(y===1)revenue*=Math.max(0.55,1-delayFrac*0.35);
    }
    if(tails.occupancyShortfall){
      const occupancyPenalty=tails.occupancyShortfall*Math.max(0.24,1-y*0.18);
      revenue*=Math.max(0.35,1-occupancyPenalty);
    }
    let pricePenalty=0;
    if(tails.pricePressure){
      pricePenalty=tails.pricePressure*Math.min(1,0.55+y*0.08);
      revenue*=Math.max(0.45,1-pricePenalty);
    }
    const yearlyRevShare=clamp(revShareBase+yearNoise*6-pricing*2,34,52);
    const yearlyOpexAdj=clamp(opexBaseAdj+yearNoise*18+Math.max(-revenueShock,0)*12,-25,35);
    const yearlyStaffAdj=clamp(staffBaseAdj+yearNoise*12+Math.max(-revenueShock,0)*10,-20,30);
    const adjustedOpex=BASE.opexRate[y]*(1+yearlyOpexAdj/100);
    const staffDelta=0.03*yearlyStaffAdj/100;
    const occupancyDeleverage=tails.occupancyShortfall?tails.occupancyShortfall*Math.max(0.02,0.085-y*0.012):0;
    const priceMarginPenalty=pricePenalty*0.18;
    const launchPenalty=tails.launchDelayMonths&&y<2?delayFrac*(y===0?0.075:0.025):0;
    const deleveragePenalty=revenueShock<0?Math.min(Math.abs(revenueShock)*0.09,0.045):0;
    const totalCostRate=yearlyRevShare/100+adjustedOpex+staffDelta+deleveragePenalty+occupancyDeleverage+priceMarginPenalty+launchPenalty;
    const margin=clamp(1-totalCostRate,-0.55,0.60);
    const ebitda=revenue*margin;
    let cf=ebitda*BASE.cfRatio[y];
    if(tails.launchDelayMonths){
      if(y===0)cf-=BASE.investment*delayFrac*0.07;
      else if(y===1&&tails.launchDelayMonths>=4)cf-=BASE.investment*delayFrac*0.025;
    }
    cfs.push(cf);
    if(y===4){
      y5Margin=margin;
      y5Ebitda=ebitda;
    }
  }

  const npv=cfs.reduce((sum,cf,t)=>sum+cf/Math.pow(1+discountRate,t),0);
  const irr5=calcIRR(cfs.slice(0,6));
  const irr10=calcIRR(cfs);
  const payback=calcPayback(cfs);
  const exitCF=[...cfs.slice(0,6)];
  exitCF[5]+=calcCompanyNetExit(y5Ebitda*exitMult,initialInvestment).net;
  const exitIRR5=calcIRR(exitCF);
  return{npv,irr5,irr10,payback,exitIRR5,hurdle:discountRate,y5Margin,exitMult,tails};
}
function runMonteCarlo(profileKey=monteCarloProfile,nSims=MC_PROFILES[profileKey]?.sims||10000){
  const profile=MC_PROFILES[profileKey]||MC_PROFILES.balanced;
  const results={
    npv:[],
    irr:[],
    payback:[],
    exitIrr:[],
    summary:{lossRisk:0,belowHurdle:0,slowPayback:0,marginSqueeze:0},
    meta:{profileKey,profileLabel:profile.label,total:nSims,marginFloor:0.20}
  };

  for(let s=0;s<nSims;s++){
    const outcome=sampleMonteCarloCase(profile);
    if(Number.isFinite(outcome.npv))results.npv.push(outcome.npv);
    if(Number.isFinite(outcome.irr5)&&outcome.irr5>-0.5&&outcome.irr5<5)results.irr.push(outcome.irr5*100);
    if(Number.isFinite(outcome.exitIRR5)&&outcome.exitIRR5>-0.5&&outcome.exitIRR5<5)results.exitIrr.push(outcome.exitIRR5*100);
    if(Number.isFinite(outcome.payback))results.payback.push(outcome.payback);
    if(!Number.isFinite(outcome.npv)||outcome.npv<0)results.summary.lossRisk++;
    if(!Number.isFinite(outcome.irr5)||outcome.irr5<outcome.hurdle)results.summary.belowHurdle++;
    if(outcome.payback===null||outcome.payback>5)results.summary.slowPayback++;
    if(!Number.isFinite(outcome.y5Margin)||outcome.y5Margin<results.meta.marginFloor)results.summary.marginSqueeze++;
  }

  results.npv.sort((a,b)=>a-b);
  results.irr.sort((a,b)=>a-b);
  results.exitIrr.sort((a,b)=>a-b);
  results.payback.sort((a,b)=>a-b);
  Object.keys(results.summary).forEach(key=>{
    results.summary[key]/=nSims;
  });
  return results;
}
function randNorm(){
  let u=0,v=0;
  while(u===0)u=Math.random();
  while(v===0)v=Math.random();
  return Math.sqrt(-2*Math.log(u))*Math.cos(2*Math.PI*v);
}
function buildHistogram(data,nBins=40){
  if(!data.length)return{labels:[],bins:[],min:0,max:0};
  const min=data[0],max=data[data.length-1];
  const span=Math.max(max-min,1);
  const binW=span/nBins;
  const bins=new Array(nBins).fill(0);
  const labels=Array.from({length:nBins},(_,i)=>min+i*binW);
  data.forEach(v=>{
    const idx=Math.min(nBins-1,Math.floor((v-min)/binW));
    bins[Math.max(0,idx)]++;
  });
  return {labels,bins,min,max};
}
function sampleTailEvents(profile,downsideBias){
  const biasedProb=baseProb=>clamp(baseProb+downsideBias*0.45,0,0.7);
  const trigger=baseProb=>Math.random()<biasedProb(baseProb);
  return{
    launchDelayMonths:trigger(profile.launchDelayProb)?Math.round(randBetween(profile.launchDelayMonths[0],profile.launchDelayMonths[1])):0,
    capexOverrun:trigger(profile.capexOverrunProb)?randBetween(profile.capexOverrunRange[0],profile.capexOverrunRange[1]):0,
    occupancyShortfall:trigger(profile.occupancyShortfallProb)?randBetween(profile.occupancyShortfallRange[0],profile.occupancyShortfallRange[1]):0,
    pricePressure:trigger(profile.pricePressureProb)?randBetween(profile.pricePressureRange[0],profile.pricePressureRange[1]):0
  };
}
function renderMonteCarloAssumptions(profile){
  const el=document.getElementById('mcAssumptions');
  if(!el)return;
  const revGrowthLabel=`${STATE.revGrowth>0?'+':''}${STATE.revGrowth}%`;
  el.textContent=currentLanguage==='ru'
    ? `${getMonteCarloProfileDescription(monteCarloProfile)} Центрировано на текущих параметрах: рост выручки ${revGrowthLabel}, provider rev-share ${STATE.revShare}%, WACC ${STATE.wacc}%, exit ${STATE.exitMult}×. В tail events включены: launch delay ${Math.round(profile.launchDelayProb*100)}% (до ${profile.launchDelayMonths[1]} мес), CAPEX overrun ${Math.round(profile.capexOverrunProb*100)}%, occupancy shortfall ${Math.round(profile.occupancyShortfallProb*100)}%, price pressure ${Math.round(profile.pricePressureProb*100)}%.`
    : `${getMonteCarloProfileDescription(monteCarloProfile)} Centered on current inputs: revenue growth ${revGrowthLabel}, provider rev-share ${STATE.revShare}%, WACC ${STATE.wacc}%, exit ${STATE.exitMult}×. Tail events included: launch delay ${Math.round(profile.launchDelayProb*100)}% (up to ${profile.launchDelayMonths[1]}m), CAPEX overrun ${Math.round(profile.capexOverrunProb*100)}%, occupancy shortfall ${Math.round(profile.occupancyShortfallProb*100)}%, price pressure ${Math.round(profile.pricePressureProb*100)}%.`;
}
function setMonteCarloProfile(profileKey,markDirty=false){
  monteCarloProfile=MC_PROFILES[profileKey]?profileKey:'balanced';
  document.querySelectorAll('.mc-profile-btn').forEach(btn=>{
    btn.classList.toggle('active',btn.dataset.mcProfile===monteCarloProfile);
  });
  renderMonteCarloAssumptions(MC_PROFILES[monteCarloProfile]);
  if(markDirty)markMonteCarloStale(true);
  else markMonteCarloStale(false);
}
function markMonteCarloStale(force=false){
  const runBtn=document.getElementById('mcRunBtn');
  const statusEl=document.getElementById('mcStatus');
  if(!runBtn||!statusEl||runBtn.dataset.running==='true')return;
  const profile=MC_PROFILES[monteCarloProfile]||MC_PROFILES.balanced;
  const hasResults=runBtn.dataset.hasResults==='true';
  const profileLabel=getMonteCarloProfileLabel(monteCarloProfile);
  if(force&&hasResults){
    runBtn.dataset.hasResults='false';
    statusEl.textContent=currentLanguage==='ru'
      ? `${profileLabel} профиль · модель изменилась, перезапустите симуляцию`
      : `${profileLabel} profile · model changed, rerun simulation`;
  }else if(force){
    statusEl.textContent=currentLanguage==='ru'
      ? `${profileLabel} профиль · готово к запуску`
      : `${profileLabel} profile · ready to simulate`;
  }else if(hasResults){
    statusEl.textContent=currentLanguage==='ru'
      ? `${profileLabel} профиль · загружен последний прогон`
      : `${profileLabel} profile · last run loaded`;
  }else{
    statusEl.textContent=currentLanguage==='ru'
      ? `${profileLabel} профиль · готово к запуску`
      : `${profileLabel} profile · ready to simulate`;
  }
}
function initMonteCarloCharts(){
  const runBtn=document.getElementById('mcRunBtn');
  const statusEl=document.getElementById('mcStatus');
  if(!runBtn)return;
  if(runBtn.dataset.bound==='true'){
    setMonteCarloProfile(monteCarloProfile,false);
    return;
  }
  runBtn.dataset.bound='true';
  runBtn.dataset.hasResults='false';
  setMonteCarloProfile(monteCarloProfile,false);

  document.querySelectorAll('.mc-profile-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      if(btn.dataset.mcProfile===monteCarloProfile)return;
      setMonteCarloProfile(btn.dataset.mcProfile,true);
    });
  });

  runBtn.addEventListener('click',()=>{
    const profile=MC_PROFILES[monteCarloProfile]||MC_PROFILES.balanced;
    const profileLabel=getMonteCarloProfileLabel(monteCarloProfile);
    runBtn.disabled=true;
    runBtn.dataset.running='true';
    runBtn.textContent=currentLanguage==='ru'?'⏳ Симуляция...':'⏳ Simulating...';
    statusEl.textContent=currentLanguage==='ru'
      ? `${profileLabel} профиль · выполняется ${profile.sims.toLocaleString()} сценариев...`
      : `${profileLabel} profile · running ${profile.sims.toLocaleString()} scenarios...`;
    setTimeout(()=>{
      const results=runMonteCarlo(monteCarloProfile,profile.sims);
      renderMCResults(results);
      runBtn.disabled=false;
      runBtn.dataset.running='false';
      runBtn.dataset.hasResults='true';
      runBtn.textContent=currentLanguage==='ru'?'▶ Запустить 10K симуляций':'▶ Run 10K Simulations';
      statusEl.textContent=currentLanguage==='ru'
        ? `${profileLabel} профиль · ${profile.sims.toLocaleString()} сценариев завершено`
        : `${profileLabel} profile · ${profile.sims.toLocaleString()} scenarios complete`;
    },50);
  });
}
function renderMCResults(results){
  monteCarloLastResults=results;
  const fmtOrDash=value=>Number.isFinite(value)?fmtCx(value):'—';
  const pctOrDash=value=>Number.isFinite(value)?value.toFixed(1)+'%':'—';
  const totalNpv=results.npv.length||1;
  const totalIrr=results.irr.length||1;
  const riskWrap=document.getElementById('mcRiskSummary');
  if(riskWrap){
    riskWrap.innerHTML=`
      <div class="mc-risk-card ${getRiskTone(results.summary.lossRisk)}">
        <div class="mc-risk-label">${currentLanguage==='ru'?'Риск убытка':'Loss Risk'}</div>
        <div class="mc-risk-value">${formatProb(results.summary.lossRisk)}</div>
        <div class="mc-risk-note">${currentLanguage==='ru'?'Вероятность того, что 10Y NPV уйдет ниже нуля после моделирования discount rate.':'Probability that 10Y NPV falls below zero after the simulated discount rate.'}</div>
      </div>
      <div class="mc-risk-card ${getRiskTone(results.summary.belowHurdle)}">
        <div class="mc-risk-label">${currentLanguage==='ru'?'Ниже hurdle':'Below Hurdle'}</div>
        <div class="mc-risk-value">${formatProb(results.summary.belowHurdle)}</div>
        <div class="mc-risk-note">${currentLanguage==='ru'?'Вероятность того, что 5Y IRR окажется ниже сценарной hurdle rate.':'Probability that 5Y IRR lands below the scenario-specific hurdle rate.'}</div>
      </div>
      <div class="mc-risk-card ${getRiskTone(results.summary.slowPayback)}">
        <div class="mc-risk-label">${currentLanguage==='ru'?'Медленный payback':'Slow Payback'}</div>
        <div class="mc-risk-value">${formatProb(results.summary.slowPayback)}</div>
        <div class="mc-risk-note">${currentLanguage==='ru'?'Payback позже 5 года или отсутствие полного возврата начальных инвестиций.':'Payback later than year 5 or no full recovery of the initial investment.'}</div>
      </div>
      <div class="mc-risk-card ${getRiskTone(results.summary.marginSqueeze)}">
        <div class="mc-risk-label">${currentLanguage==='ru'?'Сжатие маржи':'Margin Squeeze'}</div>
        <div class="mc-risk-value">${formatProb(results.summary.marginSqueeze)}</div>
        <div class="mc-risk-note">${currentLanguage==='ru'?'EBITDA margin в Y5 ниже 20%, то есть у нижней границы Dubai benchmark band.':'Y5 EBITDA margin below 20%, the lower edge of the Dubai benchmark band.'}</div>
      </div>
    `;
  }

  if(charts.mcNpv){charts.mcNpv.destroy();delete charts.mcNpv;}
  if(charts.mcIrr){charts.mcIrr.destroy();delete charts.mcIrr;}
  const npvH=buildHistogram(results.npv,40);
  const p10npv=percentileFromSorted(results.npv,.1);
  const p50npv=percentileFromSorted(results.npv,.5);
  const p90npv=percentileFromSorted(results.npv,.9);
  charts.mcNpv=new Chart(document.getElementById('chart-mc-npv'),{
    type:'bar',
    data:{
      labels:npvH.labels.map(v=>fmtCx(v)),
      datasets:[{data:npvH.bins,backgroundColor:npvH.labels.map(v=>v<0?'rgba(244,63,94,.5)':'rgba(99,102,241,.5)'),borderWidth:0,barPercentage:1,categoryPercentage:1}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      scales:{x:{display:false},y:{beginAtZero:true,title:{display:true,text:currentLanguage==='ru'?'Частота':'Frequency'}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${currentLanguage==='ru'?'Count':'Count'}: ${c.raw} (${((c.raw/totalNpv)*100).toFixed(1)}%)`}}}
    }
  });
  const irrH=buildHistogram(results.irr,40);
  charts.mcIrr=new Chart(document.getElementById('chart-mc-irr'),{
    type:'bar',
    data:{
      labels:irrH.labels.map(v=>v.toFixed(0)+'%'),
      datasets:[{data:irrH.bins,backgroundColor:irrH.labels.map(v=>v<STATE.wacc?'rgba(244,63,94,.5)':'rgba(16,185,129,.5)'),borderWidth:0,barPercentage:1,categoryPercentage:1}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      scales:{x:{display:false},y:{beginAtZero:true,title:{display:true,text:currentLanguage==='ru'?'Частота':'Frequency'}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${currentLanguage==='ru'?'Count':'Count'}: ${c.raw} (${((c.raw/totalIrr)*100).toFixed(1)}%)`}}}
    }
  });
  const p10irr=percentileFromSorted(results.irr,.1);
  const p50irr=percentileFromSorted(results.irr,.5);
  const p90irr=percentileFromSorted(results.irr,.9);
  const confCards=document.getElementById('mcConfidence');
  if(confCards){
    confCards.innerHTML=`
      <div class="mc-card"><div class="mc-label">NPV P10</div><div class="mc-value" style="color:${C.rose}">${fmtOrDash(p10npv)}</div></div>
      <div class="mc-card"><div class="mc-label">${currentLanguage==='ru'?'NPV Медиана':'NPV Median'}</div><div class="mc-value" style="color:${C.indigo}">${fmtOrDash(p50npv)}</div></div>
      <div class="mc-card"><div class="mc-label">NPV P90</div><div class="mc-value" style="color:${C.emerald}">${fmtOrDash(p90npv)}</div></div>
      <div class="mc-card"><div class="mc-label">IRR P10</div><div class="mc-value" style="color:${C.rose}">${pctOrDash(p10irr)}</div></div>
      <div class="mc-card"><div class="mc-label">${currentLanguage==='ru'?'IRR Медиана':'IRR Median'}</div><div class="mc-value" style="color:${C.indigo}">${pctOrDash(p50irr)}</div></div>
      <div class="mc-card"><div class="mc-label">IRR P90</div><div class="mc-value" style="color:${C.emerald}">${pctOrDash(p90irr)}</div></div>
    `;
  }
  renderMonteCarloAssumptions(MC_PROFILES[results.meta.profileKey]||MC_PROFILES.balanced);
}

/* ===== NAVIGATION ===== */
document.querySelectorAll('.nav-btn').forEach(btn=>{
  btn.addEventListener('click',()=>{
    closeChartFocus(false);
    document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    document.querySelectorAll('.dashboard-section').forEach(s=>s.classList.remove('active'));
    const secEl=document.getElementById('section-'+btn.dataset.section);
    secEl.classList.add('active');
    animateVisible();
    hydrateSection(secEl);
  });
});

/* ===== SLIDER BINDINGS ===== */
const sliderMap=[
  {id:'sl-revGrowth',key:'revGrowth',valId:'val-revGrowth',fmt:v=>v+'%',dflt:0},
  {id:'sl-revShare',key:'revShare',valId:'val-revShare',fmt:v=>v+'%',dflt:40},
  {id:'sl-opex',key:'opexAdj',valId:'val-opex',fmt:v=>(v>0?'+':'')+v+'%',dflt:0},
  {id:'sl-staff',key:'staffAdj',valId:'val-staff',fmt:v=>(v>0?'+':'')+v+'%',dflt:0},
  {id:'sl-wacc',key:'wacc',valId:'val-wacc',fmt:v=>v+'%',dflt:12},
  {id:'sl-exitMult',key:'exitMult',valId:'val-exitMult',fmt:v=>v+'×',dflt:8}
];
sliderMap.forEach(s=>{
  const el=document.getElementById(s.id);
  el.addEventListener('input',()=>{
    const v=parseFloat(el.value);
    STATE[s.key]=v;
    const valEl=document.getElementById(s.valId);
    valEl.textContent=s.fmt(v);
    valEl.classList.toggle('changed',v!==s.dflt);
    updateScenarioHighlight();
    updateAll();
  });
});

/* ===== SCENARIO BUTTONS ===== */
document.querySelectorAll('.sc-btn').forEach(btn=>{
  btn.addEventListener('click',()=>{
    const sc=SCENARIOS[btn.dataset.scenario];if(!sc)return;
    STATE.scenario=btn.dataset.scenario;
    Object.assign(STATE,sc);
    syncSliders();
    document.querySelectorAll('.sc-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    updateAll();
  });
});
document.getElementById('btnReset').addEventListener('click',()=>{
  Object.assign(STATE,SCENARIOS.base);STATE.scenario='base';
  syncSliders();
  document.querySelectorAll('.sc-btn').forEach(b=>b.classList.remove('active'));
  document.querySelector('.sc-btn[data-scenario="base"]').classList.add('active');
  updateAll();
});
function syncSliders(){
  sliderMap.forEach(s=>{
    const el=document.getElementById(s.id);el.value=STATE[s.key];
    const valEl=document.getElementById(s.valId);valEl.textContent=s.fmt(STATE[s.key]);
    valEl.classList.toggle('changed',STATE[s.key]!==s.dflt);
  });
}
function updateScenarioHighlight(){
  const match=Object.entries(SCENARIOS).find(([k,sc])=>sliderMap.every(s=>STATE[s.key]==sc[s.key]));
  document.querySelectorAll('.sc-btn').forEach(b=>b.classList.remove('active'));
  if(match)document.querySelector(`.sc-btn[data-scenario="${match[0]}"]`)?.classList.add('active');
}

function readStoredUI(key,fallback,scope='global'){
  try{
    const storageKey=scope==='project'?getProjectStorageKey(key):key;
    const raw=localStorage.getItem(storageKey);
    return raw===null?fallback:JSON.parse(raw);
  }catch(_){
    return fallback;
  }
}
function writeStoredUI(key,value,scope='global'){
  try{
    const storageKey=scope==='project'?getProjectStorageKey(key):key;
    localStorage.setItem(storageKey,JSON.stringify(value));
  }catch(_){}
}
function setTemplateComplianceDetailsExpanded(expanded,persist=true){
  templateComplianceDetailsExpanded=Boolean(expanded);
  if(persist)writeStoredUI(UI_STORAGE_KEYS.templateComplianceDetails,templateComplianceDetailsExpanded);
  renderTemplateCompliance();
}
function setTemplateComplianceFilter(nextFilter){
  templateComplianceFilter=['all','fail','warn','optional'].includes(nextFilter)?nextFilter:'all';
  renderTemplateCompliance();
}
function setControlPanelCollapsed(collapsed,persist=true){
  const panel=document.getElementById('controlPanel');
  const layout=document.querySelector('.dashboard-layout');
  const toggle=document.getElementById('panelToggleBtn');
  const titleCopy=panel?.querySelector('.panel-title-copy');
  const groups=panel?.querySelector('.panel-groups');
  if(!panel||!layout||!toggle||!groups||!titleCopy)return;
  panel.classList.toggle('collapsed',collapsed);
  layout.classList.toggle('panel-collapsed',collapsed);
  titleCopy.hidden=collapsed;
  groups.hidden=collapsed;
  toggle.setAttribute('aria-expanded',String(!collapsed));
  toggle.setAttribute('aria-label',collapsed?'Expand model panel':'Collapse model panel');
  toggle.textContent=collapsed?'›':'‹';
  if(persist)writeStoredUI(UI_STORAGE_KEYS.panelCollapsed,collapsed);
}
function setPanelGroupOpen(group,isOpen,persist=true){
  if(!group)return;
  group.classList.toggle('is-open',isOpen);
  const toggle=group.querySelector('.panel-group-toggle');
  const body=group.querySelector('.panel-group-body');
  if(toggle)toggle.setAttribute('aria-expanded',String(isOpen));
  if(body)body.hidden=!isOpen;
  if(persist){
    const current=readStoredUI(UI_STORAGE_KEYS.panelGroups,{});
    current[group.dataset.panelGroup]=isOpen;
    writeStoredUI(UI_STORAGE_KEYS.panelGroups,current);
  }
}
function initControlPanel(){
  const panel=document.getElementById('controlPanel');
  const toggle=document.getElementById('panelToggleBtn');
  if(!panel||!toggle)return;

  setControlPanelCollapsed(Boolean(readStoredUI(UI_STORAGE_KEYS.panelCollapsed,false)),false);
  toggle.addEventListener('click',()=>setControlPanelCollapsed(!panel.classList.contains('collapsed')));

  const storedGroups=readStoredUI(UI_STORAGE_KEYS.panelGroups,{});
  document.querySelectorAll('.panel-group').forEach(group=>{
    const defaultOpen=group.classList.contains('is-open');
    const isOpen=storedGroups[group.dataset.panelGroup]??defaultOpen;
    setPanelGroupOpen(group,isOpen,false);
    group.querySelector('.panel-group-toggle')?.addEventListener('click',()=>{
      setPanelGroupOpen(group,!group.classList.contains('is-open'));
    });
  });
}

/* ===== CURRENCY TOGGLE ===== */
document.getElementById('currencyToggle').addEventListener('change',e=>{
  currencyMode = e.target.checked ? 'USD' : 'AED';
  // Destroy all charts so they reinit with new currency values
  Object.keys(charts).forEach(k=>{
    if(charts[k]&&charts[k].destroy){charts[k].destroy();}
    delete charts[k];
  });
  updateAll();
  // Re-init active section charts
  const sec=document.querySelector('.dashboard-section.active');
  if(sec){
    hydrateSection(sec);
  }
  // Update footer currency text
  updateFooterCurrency();
});
function updateFooterCurrency(){
  const footer=document.querySelector('#main-footer p');
  if(footer){
    const project=getProjectConfig();
    const sourceInfo=describeModelSource();
    footer.innerHTML=`Project: ${project.title} | Source: ${sourceInfo.sourceLabel} | Cohort + Step-Function growth | Demand-driven hiring<br>Assessment: ${localizeProjectValue(project.assessmentLabel)} | Currency: ${cxLabel()} | Discount rate: ${STATE.wacc}% WACC`;
  }
}

/* ===== ANIMATIONS ===== */
function animateVisible(){
  document.querySelectorAll('.dashboard-section.active [data-animate]').forEach((el,i)=>{setTimeout(()=>el.classList.add('visible'),i*60);});
  setTimeout(()=>{document.querySelectorAll('.bench-bar').forEach(b=>{b.style.width=b.dataset.width;});},300);
  colorHeatmap();
}

/* ===== YEAR FILTER ===== */
let selectedYear='all';
let yearRangeAnchor=null;
function syncYearPills(){
  const [start,end]=getSelectedYearBounds();
  document.querySelectorAll('.year-pill').forEach(btn=>{
    const year=btn.dataset.year;
    btn.classList.remove('active','in-range','range-start','range-end');
    if(year==='all'){
      if(selectedYear==='all')btn.classList.add('active');
      return;
    }
    const idx=parseInt(year,10);
    if(selectedYear!=='all'&&idx>=start&&idx<=end){
      btn.classList.add('in-range');
      if(idx===start||idx===end)btn.classList.add('active');
      if(idx===start)btn.classList.add('range-start');
      if(idx===end)btn.classList.add('range-end');
    }
  });
}
function updateSelectedPeriodUI(summary=getSelectionSummary()){
  const label=getSelectedYearLabel();
  setText('yearFilterSummary',label);
  setText('lm-ebitda-label',label+' EBITDA');
  setText('kpi-margin-label',label+' EBITDA MARGIN');
  setText('kpi-rev-label',label+' REVENUE');
  setText('kpi-np-label',label+' NET PROFIT');
  setText('kpi-margin5',(summary.margin*100).toFixed(1)+'%');
  setText('kpi-rev10',fmtMCx(summary.revenue)+'M');
  setText('kpi-np10',fmtMCx(summary.netProfit)+'M');
  setLM('lm-ebitda',fmtCx(summary.ebitda),summary.ebitda>=0);
}
function renderYearSelection(announce=false){
  updateSelectedPeriodUI(getSelectionSummary());
  updateSensitivityContextUI(getSensitivityAnchor());

  if(charts.revenue)updateRevenueCharts();
  else updatePLTable();
  if(charts.stress)updateSensitivityCharts();
  if(charts.paEngine||charts.paEcon){renderPatientAcquisitionModel();renderPatientBaseBridge();updatePatientAcquisitionCharts();}

  document.querySelectorAll('#plTable th, #plTable td').forEach(cell=>{
    cell.classList.remove('year-highlight');
  });
  if(selectedYear!=='all'){
    normalizeYearSelection().forEach(yearIdx=>{
      document.querySelectorAll(`#plTable [data-year="${yearIdx}"]`).forEach(cell=>cell.classList.add('year-highlight'));
    });
  }

  if(announce)showToast(tr('selectedPeriodToast',{label:getSelectedYearLabel()}),'');
}
function initYearFilter(){
  syncYearPills();
  document.querySelectorAll('.year-pill').forEach(btn=>{
    btn.addEventListener('click',()=>{
      if(btn.dataset.year==='all'){
        selectedYear='all';
        yearRangeAnchor=null;
      }else{
        const yearIdx=Math.max(0,Math.min(9,parseInt(btn.dataset.year,10)||0));
        if(selectedYear==='all'||yearRangeAnchor===null){
          selectedYear=String(yearIdx);
          yearRangeAnchor=yearIdx;
        }else if(!selectedYear.includes(':')&&yearIdx!==yearRangeAnchor){
          const start=Math.min(yearRangeAnchor,yearIdx);
          const end=Math.max(yearRangeAnchor,yearIdx);
          selectedYear=start===end?String(start):`${start}:${end}`;
        }else{
          selectedYear=String(yearIdx);
          yearRangeAnchor=yearIdx;
        }
      }
      syncYearPills();
      applyYearFilter();
    });
  });
}
function applyYearFilter(announce=true){
  updateKPIs();
  updateLiveMetrics();
  renderYearSelection(announce);
}

/* ===== INIT ===== */
window.addEventListener('DOMContentLoaded',()=>{
  currentProjectId=resolveProjectId(readStoredUI(APP_STORAGE_KEYS.selectedProject,ACTIVE_DEFAULT_PROJECT_ID));
  currentLanguage=resolveLanguage(readStoredUI(UI_STORAGE_KEYS.language,'ru'));
  initLanguageSwitcher();
  applyStaticTranslations();
  initProjectSelector();
  loadProjectScopedState();
  templateComplianceDetailsExpanded=Boolean(readStoredUI(UI_STORAGE_KEYS.templateComplianceDetails,false));
  recalculate();
  updateKPIs();
  updateLiveMetrics();
  initCardFocusMode();
  animateVisible();
  initOverviewCharts();
  initControlPanel();
  initChecks();
  initExcelImport();
  initYearFilter();
  initScenarioLibraryModule();
  initEquityModelModule();
  initPatientEngineViewTabs();
  initPatientTrendControls();
  renderYearSelection(false);
  renderServiceLineEconomics();
  renderPatientAcquisitionModel();
  renderPatientBaseBridge();
  renderFundingMetrics();
  renderEquityModel();
  renderAssumptionsPanel();
  renderSourcesPanel();
  renderTemplateCompliance();
  renderImportAudit();
  renderExcelChecklist();
  renderActualVsPlan();
  initPdfExport();
  appReady=true;
  setImportButtonState('loading',`Loading ${getProjectConfig().label}`);
  loadProjectWorkbook();
});

/* ===== PDF EXPORT ===== */
function initPdfExport(){
  const btn=document.getElementById('exportPdfBtn');
  if(!btn)return;
  btn.addEventListener('click',async()=>{
    btn.disabled=true;
    btn.textContent='⏳ Generating...';
    showToast('📄 Generating PDF report...','');
    try{
      await generatePDF();
      btn.textContent='📄 Export PDF';
      btn.disabled=false;
      showToast('✅ PDF saved!','success');
    }catch(e){
      console.error('PDF error:',e);
      btn.textContent='📄 Export PDF';
      btn.disabled=false;
      showToast('❌ PDF error: '+e.message,'error');
    }
  });
}
async function generatePDF(){
  closeChartFocus(false);
  const {jsPDF}=window.jspdf;
  const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4',compress:true});
  const project=getProjectConfig();
  const sections=document.querySelectorAll('.dashboard-section');
  const pageW=pdf.internal.pageSize.getWidth();
  const pageH=pdf.internal.pageSize.getHeight();
  const marginX=10;
  const headerH=14;
  const footerH=8;
  const contentW=pageW-(marginX*2);
  const contentH=pageH-headerH-footerH-6;
  const exportDate=new Date().toLocaleDateString('en-GB');
  const exportPeriod=getSelectedYearLabel();
  const body=document.body;
  const originalActive=document.querySelector('.dashboard-section.active');
  const originalScrollY=window.scrollY;
  let firstPage=true;

  const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
  const getSectionTitle=sec=>{
    const key=sec.id.replace('section-','');
    return({
      overview:'OVERVIEW',
      revenue:'REVENUE',
      patients:'PATIENTS',
      cashflow:'CASH FLOW',
      capex:'CAPEX',
      sensitivity:'SENSITIVITY',
      risks:'RISKS'
    })[key]||key.toUpperCase();
  };
  const addCanvasPages=(canvas,sectionTitle)=>{
    const sliceHeightPx=Math.max(1,Math.floor((contentH*canvas.width)/contentW));
    let offsetY=0;
    let part=1;
    while(offsetY<canvas.height){
      const pageCanvas=document.createElement('canvas');
      pageCanvas.width=canvas.width;
      pageCanvas.height=Math.min(sliceHeightPx,canvas.height-offsetY);
      const pageCtx=pageCanvas.getContext('2d');
      pageCtx.drawImage(canvas,0,offsetY,canvas.width,pageCanvas.height,0,0,canvas.width,pageCanvas.height);
      const renderedH=pageCanvas.height*(contentW/pageCanvas.width);
      const imgData=pageCanvas.toDataURL('image/png');

      if(!firstPage)pdf.addPage();
      pdf.setFillColor(10,14,26);
      pdf.rect(0,0,pageW,pageH,'F');
      pdf.setDrawColor(99,102,241);
      pdf.setLineWidth(.4);
      pdf.line(marginX,headerH,pageW-marginX,headerH);
      pdf.setFontSize(12);
      pdf.setTextColor(241,245,249);
      pdf.text(`${project.title.toUpperCase()} — ${sectionTitle}`,marginX,8);
      pdf.setFontSize(8);
      pdf.setTextColor(148,163,184);
      pdf.text(`Period: ${exportPeriod} | Exported: ${exportDate} | Part ${part}`,marginX,12);
      pdf.addImage(imgData,'PNG',marginX,headerH+4,contentW,renderedH,undefined,'FAST');
      firstPage=false;
      offsetY+=pageCanvas.height;
      part+=1;
    }
  };

  body.classList.add('pdf-export-mode');
  updateAll();
  await wait(120);
  for(const sec of sections){
    sections.forEach(s=>s.classList.remove('active'));
    sec.classList.add('active');
    hydrateSection(sec);
    await wait(160);
    flushChartsWithin(sec);
    await wait(220);

    const canvas=await html2canvas(sec,{
      backgroundColor:'#0a0e1a',
      scale:2,
      useCORS:true,
      logging:false,
      scrollX:-window.scrollX,
      scrollY:-window.scrollY,
      windowWidth:1600,
      windowHeight:Math.max(sec.scrollHeight,sec.offsetHeight)
    });
    addCanvasPages(canvas,getSectionTitle(sec));
  }

  const totalPages=pdf.getNumberOfPages();
  for(let i=1;i<=totalPages;i++){
    pdf.setPage(i);
    pdf.setFontSize(8);
    pdf.setTextColor(100,116,139);
    pdf.text(`Page ${i} / ${totalPages}`,pageW-marginX,pageH-4,{align:'right'});
  }

  sections.forEach(s=>s.classList.remove('active'));
  body.classList.remove('pdf-export-mode');
  if(originalActive){
    originalActive.classList.add('active');
    hydrateSection(originalActive);
  }else{
    resizeChartsWithin(document.querySelector('.dashboard-section.active'));
  }
  window.scrollTo({top:originalScrollY,left:0,behavior:'auto'});

  pdf.save(`${project.pdfSlug}_Dashboard_Report_${exportPeriod}_${new Date().toISOString().slice(0,10)}.pdf`);
}

/* ===== EXCEL IMPORT (SheetJS) ===== */
function initExcelImport(){
  const overlay=document.getElementById('dropOverlay');
  const importBtn=document.getElementById('importBtn');
  const fileInput=document.getElementById('xlsxFileInput');
  if(!overlay||!importBtn)return;

  // Click import button → open file picker
  importBtn.addEventListener('click',()=>fileInput.click());
  fileInput.addEventListener('change',e=>{
    if(e.target.files.length)processExcelFile(e.target.files[0]);
  });

  // Drag & Drop
  let dragCounter=0;
  document.addEventListener('dragenter',e=>{
    e.preventDefault();
    dragCounter++;
    overlay.classList.add('active');
  });
  document.addEventListener('dragleave',e=>{
    e.preventDefault();
    dragCounter--;
    if(dragCounter<=0){overlay.classList.remove('active');dragCounter=0;}
  });
  document.addEventListener('dragover',e=>e.preventDefault());
  document.addEventListener('drop',e=>{
    e.preventDefault();
    overlay.classList.remove('active');
    dragCounter=0;
    const files=e.dataTransfer.files;
    if(files.length&&/\.xlsx?$/i.test(files[0].name)){
      processExcelFile(files[0]);
    } else {
      showToast('⚠ Please drop an .xlsx file','error');
    }
  });
}

function processExcelFile(file){
  showToast('📊 Reading '+file.name+'...','');
  setImportButtonState('loading',`Reading ${file.name}`);
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      templateCompliance=assessWorkbookTemplate(wb,file.name);
      templateComplianceFilter=templateCompliance.summary.fail>0?'fail':'all';
      if(!templateCompliance.ready&&!templateComplianceDetailsExpanded){
        templateComplianceDetailsExpanded=true;
      }
      renderTemplateCompliance();
      if(!templateCompliance.ready){
        setImportButtonState('idle');
        showToast(`❌ Template mismatch: ${templateCompliance.summary.fail||0} critical gap(s). Previous model kept.`,`error`);
        return;
      }
      parseWorkbook(wb,file.name);
      showToast('✅ '+file.name+' loaded! Dashboard updated.','success');
      setImportButtonState('loaded',file.name);
    }catch(err){
      console.error('Excel parse error:',err);
      setImportButtonState('idle');
      showToast('❌ Error: '+err.message,'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

async function loadProjectWorkbook(projectId=currentProjectId){
  if(typeof XLSX==='undefined'||typeof fetch!=='function')return;
  const targetProjectId=resolveProjectId(projectId);
  const project=getProjectConfig(targetProjectId);
  const loadToken=++currentProjectLoadToken;
  try{
    if(window.location?.protocol==='file:'){
      setModelSourceState(createModelSourceState({reason:'file_protocol'}),{rerender:true});
      setImportButtonState('idle');
      showToast(currentLanguage==='ru'
        ? '⚠️ Автозагрузка workbook отключена в file:// режиме. Запустите локальный сервер или импортируйте .xlsx вручную.'
        : '⚠️ Workbook auto-load is disabled in file:// mode. Start a local server or import .xlsx manually.','error');
      return;
    }
    const filename=project.workbookPath;
    const response=await fetch(`${filename}?t=${Date.now()}`,{cache:'no-store'});
    if(loadToken!==currentProjectLoadToken)return;
    if(!response.ok){
      setModelSourceState(createModelSourceState({reason:response.status===404?'workbook_missing':'fetch_failed'}),{rerender:true});
      setImportButtonState('idle');
      showToast(`❌ Workbook not found: ${project.workbookLabel}`,'error');
      return;
    }
    const buffer=await response.arrayBuffer();
    const wb=XLSX.read(buffer,{type:'array'});
    if(loadToken!==currentProjectLoadToken)return;
    templateCompliance=assessWorkbookTemplate(wb,project.workbookLabel);
    templateComplianceFilter=templateCompliance.summary.fail>0?'fail':'all';
    if(!templateCompliance.ready&&!templateComplianceDetailsExpanded){
      templateComplianceDetailsExpanded=true;
    }
    renderTemplateCompliance();
    if(!templateCompliance.ready){
      console.warn('Default workbook failed template compliance:',templateCompliance);
      setModelSourceState(createModelSourceState({reason:'template_blocked'}),{rerender:true});
      setImportButtonState('idle');
      return;
    }
    parseWorkbook(wb,project.workbookLabel);
    setModelSourceState(createModelSourceState({
      type:'workbook',
      reason:'loaded',
      origin:'auto',
      fileName:project.workbookLabel,
      loadedAt:new Date().toISOString()
    }),{rerender:true});
    setImportButtonState('loaded',project.workbookLabel);
  }catch(err){
    setModelSourceState(createModelSourceState({reason:'fetch_failed'}),{rerender:true});
    setImportButtonState('idle');
    console.warn('Project workbook auto-load skipped:',err);
  }
}

function showToast(msg,type){
  let toast=document.querySelector('.import-toast');
  if(!toast){
    toast=document.createElement('div');
    toast.className='import-toast';
    document.body.appendChild(toast);
  }
  toast.textContent=msg;
  toast.className='import-toast '+(type||'');
  setTimeout(()=>toast.classList.add('show'),10);
  setTimeout(()=>toast.classList.remove('show'),4000);
}

function getSheet(wb,name){
  const ws=wb.Sheets[name];
  return ws?XLSX.utils.sheet_to_json(ws,{header:1,defval:''}):null;
}
function getRow(sheet,rowIdx){
  return sheet&&sheet[rowIdx]?sheet[rowIdx]:[];
}
function getRowNums(sheet,rowIdx,startCol,count){
  const row=getRow(sheet,rowIdx);
  const nums=[];
  for(let c=startCol;c<startCol+count;c++){
    const v=row[c];
    nums.push(typeof v==='number'?v:parseFloat(v)||0);
  }
  return nums;
}

function parseWorkbook(wb,filename){
  const mappedMetricSet=new Set();
  const markMapped=(condition,label)=>{if(condition)mappedMetricSet.add(label);};
  const workbookChecklistRaw={
    wacc:SCENARIOS.base.wacc,
    serviceRevenue:SERVICE_REVENUE_ROWS.map(()=>Array(10).fill(0))
  };
  const pl=getSheet(wb,'P&L');
  const mpl=getSheet(wb,'Monthly P&L');
  const cf=getSheet(wb,'Cash Flow & DCF');
  const cap=getSheet(wb,'CapEx');
  const pa=getSheet(wb,'Patient Acquisition');
  const capacity=getSheet(wb,'Capacity');
  const revenueSheet=getSheet(wb,'Revenue');
  const priceList=getSheet(wb,'Price List');
  const summary=getSheet(wb,'Summary');

  if(!pl){throw new Error('Sheet "P&L" not found');}

  // P&L: Row 3=Revenue, 4=COGS, ..., 14=EBITDA/25=EBITDA, 27=Net Profit
  // Find rows by label
  const normalizeSheetLabel=value=>String(value??'').trim().toLowerCase().replace(/[—–]/g,'-').replace(/\s+/g,' ');
  const findRow=(sheet,label)=>{
    const target=normalizeSheetLabel(label);
    if(!sheet)return -1;
    for(let r=0;r<sheet.length;r++){
      if(normalizeSheetLabel(sheet[r][0]).startsWith(target))return r;
    }
    return -1;
  };
  const findRowContains=(sheet,label)=>{
    const target=normalizeSheetLabel(label);
    if(!sheet)return -1;
    for(let r=0;r<sheet.length;r++){
      if(normalizeSheetLabel(sheet[r][0]).includes(target))return r;
    }
    return -1;
  };

  // P&L data (cols B-K = indices 1-10)
  const revRow=findRow(pl,'Revenue');
  const cogsRow=findRow(pl,'COGS');
  const ebitdaRow=findRow(pl,'EBITDA');
  const npRow=findRow(pl,'NET PROFIT');
  const cashEarnRow=findRow(pl,'Cash Earnings');
  const marketingCostRow=findRow(pl,'Marketing');
  const salaryRow=findRow(pl,'Salary');
  const fixedStaffRow=findRow(pl,'Fixed Staff');
  const adminRow=findRow(pl,'Administrative');
  const rdRow=findRow(pl,'R&D');
  const rentRow=findRow(pl,'Rent');

  if(revRow>=0)BASE.revenue=getRowNums(pl,revRow,1,10);
  if(cogsRow>=0)BASE.cogs=getRowNums(pl,cogsRow,1,10).map(Math.abs);
  if(ebitdaRow>=0)BASE.ebitda=getRowNums(pl,ebitdaRow,1,10);
  if(npRow>=0)BASE.netProfit=getRowNums(pl,npRow,1,10);
  if(revRow>=0)workbookChecklistRaw.revenue=getRowNums(pl,revRow,1,10);
  if(cogsRow>=0)workbookChecklistRaw.cogs=getRowNums(pl,cogsRow,1,10);
  if(ebitdaRow>=0)workbookChecklistRaw.ebitda=getRowNums(pl,ebitdaRow,1,10);
  if(npRow>=0)workbookChecklistRaw.netProfit=getRowNums(pl,npRow,1,10);
  if(marketingCostRow>=0)workbookChecklistRaw.marketingCost=getRowNums(pl,marketingCostRow,1,10);
  if(salaryRow>=0)workbookChecklistRaw.salary=getRowNums(pl,salaryRow,1,10);
  if(fixedStaffRow>=0)workbookChecklistRaw.fixedStaff=getRowNums(pl,fixedStaffRow,1,10);
  if(adminRow>=0)workbookChecklistRaw.admin=getRowNums(pl,adminRow,1,10);
  if(rentRow>=0)workbookChecklistRaw.rent=getRowNums(pl,rentRow,1,10);
  if(rdRow>=0)workbookChecklistRaw.rd=getRowNums(pl,rdRow,1,10);
  if(cashEarnRow>=0)workbookChecklistRaw.cashEarnings=getRowNums(pl,cashEarnRow,1,10);
  if(marketingCostRow>=0)BASE.plBreakdown.marketing=getRowNums(pl,marketingCostRow,1,10).map(Math.abs);
  if(adminRow>=0)BASE.plBreakdown.admin=getRowNums(pl,adminRow,1,10).map(Math.abs);
  if(rentRow>=0)BASE.plBreakdown.rent=getRowNums(pl,rentRow,1,10).map(Math.abs);
  if(rdRow>=0)BASE.plBreakdown.rd=getRowNums(pl,rdRow,1,10).map(Math.abs);
  if(fixedStaffRow>=0)BASE.plBreakdown.staff=getRowNums(pl,fixedStaffRow,1,10).map(Math.abs);
  if(salaryRow>=0&&fixedStaffRow<0)BASE.plBreakdown.staff=getRowNums(pl,salaryRow,1,10).map(Math.abs);
  markMapped(revRow>=0,'Revenue');
  markMapped(cogsRow>=0,'COGS');
  markMapped(ebitdaRow>=0,'EBITDA');
  markMapped(npRow>=0,'Net Profit');
  markMapped(marketingCostRow>=0,'Marketing');
  markMapped(adminRow>=0,'Admin');
  markMapped(rentRow>=0,'Rent');
  markMapped(rdRow>=0,'R&D');
  markMapped(fixedStaffRow>=0||salaryRow>=0,'Staff');

  // EBITDA Margin
  if(revRow>=0&&ebitdaRow>=0){
    BASE.margins=BASE.ebitda.map((e,i)=>e/BASE.revenue[i]);
  }

  // Revenue by department
  if(revenueSheet){
    SERVICE_REVENUE_ROWS.forEach((label,idx)=>{
      const row=findRow(revenueSheet,label);
      if(row>=0){
        const values=getRowNums(revenueSheet,row,1,10);
        BASE.serviceRevenue[idx]=values.slice();
        workbookChecklistRaw.serviceRevenue[idx]=values.slice();
      }
      markMapped(row>=0,`Revenue Dept: ${label}`);
    });
  }
  if(capacity){
    const annualVisitsAfter=label=>{
      const start=findRow(capacity,label);
      if(start<0)return null;
      for(let r=start+1;r<Math.min(start+6,capacity.length);r++){
        if(normalizeSheetLabel(getRow(capacity,r)[0]).includes('annual visits'))return getRowNums(capacity,r,1,10);
      }
      return null;
    };
    const physioVisits=annualVisitsAfter('Physio Clinic');
    const dentalVisits=annualVisitsAfter('Dental Clinic');
    const gpVisits=annualVisitsAfter('GP');
    const gynVisits=annualVisitsAfter('Gynecology');
    const cosmoVisits=annualVisitsAfter('Cosmetology');
    if(physioVisits)BASE.serviceVisits[0]=physioVisits;
    if(dentalVisits)BASE.serviceVisits[1]=dentalVisits;
    if(gpVisits||gynVisits)BASE.serviceVisits[2]=Array.from({length:10},(_,i)=>(gpVisits?.[i]||0)+(gynVisits?.[i]||0));
    if(cosmoVisits)BASE.serviceVisits[3]=cosmoVisits;
    BASE.serviceVisits[4]=Array(10).fill(0);
    markMapped(Boolean(physioVisits),'Dept Visits: Physio');
    markMapped(Boolean(dentalVisits),'Dept Visits: Dental');
    markMapped(Boolean(gpVisits||gynVisits),'Dept Visits: Polyclinic');
    markMapped(Boolean(cosmoVisits),'Dept Visits: Cosmetology');
  }
  if(priceList){
    const parsedProducts=[];
    let currentDepartment='';
    for(let r=0;r<priceList.length;r++){
      const row=getRow(priceList,r);
      const first=String(row[0]||'').trim();
      const second=String(row[1]||'').trim();
      if(BASE.services.includes(first)){
        currentDepartment=first;
        continue;
      }
      if(!currentDepartment||!second)continue;
      const basePrice=Math.max(0,Number(row[2])||0);
      if(basePrice<=0)continue;
      parsedProducts.push({
        department:currentDepartment,
        product:second.replace(/\s*\(lab partner\)\s*/i,'').trim(),
        basePrice,
        monthlyQty:Array.from({length:10},(_,yearIdx)=>Math.max(0,Number(row[3+yearIdx])||0))
      });
    }
    if(parsedProducts.length){
      BASE.productCatalog=parsedProducts;
      syncProductCatalog();
    }
    markMapped(parsedProducts.length>0,'Price List / Product Volumes');
  }

  // Monthly P&L (Y1)
  if(mpl){
    const mRevRow=findRow(mpl,'Revenue');
    const mEbitdaRow=findRow(mpl,'EBITDA');
    if(mRevRow>=0)BASE.mRev=getRowNums(mpl,mRevRow,1,12);
    if(mEbitdaRow>=0)BASE.mEbitda=getRowNums(mpl,mEbitdaRow,1,12);
    if(mRevRow>=0)workbookChecklistRaw.monthlyRevenue=getRowNums(mpl,mRevRow,1,12);
    if(mEbitdaRow>=0)workbookChecklistRaw.monthlyEbitda=getRowNums(mpl,mEbitdaRow,1,12);
    markMapped(mRevRow>=0,'Monthly Revenue');
    markMapped(mEbitdaRow>=0,'Monthly EBITDA');
  }

  // Cash Flow & DCF
  if(cf){
    const capInvRow=findRow(cf,'Capital Investment');
    const ocfRow=findRow(cf,'Operating Cash Flow');
    const totalInvRow=findRow(cf,'Total Investment');
    const npvRow=findRow(cf,'NPV (AED)');
    const stakeRow=findRow(cf,'Investor Equity Stake');
    const payoutRow=findRow(cf,'Dividend Payout Ratio');
    const prefRow=findRow(cf,'Preferred Return');
    const revShareAssRow=findRow(cf,'Revenue Share Rate');
    const targetReturnRow=findRow(cf,'Target Return Multiple');
    const graceRow=findRow(cf,'Grace Period');
    if(capInvRow>=0&&ocfRow>=0){
      const inv=Math.abs(getRowNums(cf,capInvRow,1,1)[0]);
      if(inv>0)BASE.investment=inv;
      const ocfs=getRowNums(cf,ocfRow,1,11); // Y0..Y10
      BASE.annCF=[-BASE.investment,...ocfs.slice(1)]; // first is Y0 OCF=0
      if(ocfs[0]===0)BASE.annCF=[-BASE.investment,...ocfs.slice(1)];
      else BASE.annCF=[-BASE.investment,...ocfs];
      workbookChecklistRaw.capitalInvestment=BASE.investment;
      workbookChecklistRaw.annCF=BASE.annCF.slice();
    }
    if(npvRow>=0){
      const npvCandidates=getRow(cf,npvRow).slice(1).map(v=>Number(v)).filter(v=>Number.isFinite(v));
      if(npvCandidates.length){
        workbookChecklistRaw.npv10=npvCandidates.reduce((best,val)=>Math.abs(val)>Math.abs(best)?val:best,npvCandidates[0]);
      }
    }
    const pctVal=v=>{const num=Number(v)||0;return num<=1?num*100:num;};
    if(totalInvRow>=0)workbookChecklistRaw.totalInvestment=Number(getRow(cf,totalInvRow)[1])||null;
    const investorState=normalizeEquityState({
      // Workbook `Cash Flow & DCF!Total Investment` (6,908,760) double-counts WC — fall back to BASE.investment (6,572,160) which already includes WC.
      totalInvestment:BASE.investment,
      returnMode:equityState.returnMode,
      exitYear:5,
      equityStakePct:stakeRow>=0?pctVal(getRow(cf,stakeRow)[1]):equityState.equityStakePct,
      dividendPayoutPct:payoutRow>=0?pctVal(getRow(cf,payoutRow)[1]):equityState.dividendPayoutPct,
      // Excel Cash Flow & DCF uses pro-rata dividends (no preferred waterfall in practice); ignore prefRow metadata
      preferredReturnPct:0,
      revenueShareRatePct:revShareAssRow>=0?pctVal(getRow(cf,revShareAssRow)[1]):equityState.revenueShareRatePct,
      targetReturnMultiple:targetReturnRow>=0?(Number(getRow(cf,targetReturnRow)[1])||equityState.targetReturnMultiple):equityState.targetReturnMultiple,
      graceMonths:graceRow>=0?(Number(getRow(cf,graceRow)[1])||equityState.graceMonths):equityState.graceMonths
    });
    equityState=investorState;
    writeStoredUI(APP_STORAGE_KEYS.equityModel,equityState,'project');
    markMapped(capInvRow>=0,'Capital Investment');
    markMapped(ocfRow>=0,'Operating Cash Flow');
    markMapped(npvRow>=0,'NPV');
    markMapped(totalInvRow>=0,'Investor Total Investment');
    markMapped(stakeRow>=0,'Investor Equity Stake');
    markMapped(payoutRow>=0,'Dividend Payout Ratio');
    markMapped(prefRow>=0,'Preferred Return');
    markMapped(revShareAssRow>=0,'Revenue Share Investor Rate');
    markMapped(targetReturnRow>=0,'Target Return Multiple');
    markMapped(graceRow>=0,'Grace Period');
  }

  // Patient Acquisition
  if(pa){
    const capRow=findRow(pa,'Annual Visit Capacity');
    const visitRow=findRow(pa,'Total Visits');
    const retainedRow=findRow(pa,'Retained from Prior');
    const referralRow=findRow(pa,'Referral');
    const walkinRow=findRow(pa,'Walk-in');
    const marketingRow=findRow(pa,'Marketing (paid');
    const utilRow=findRow(pa,'Utilization');
    const patientsRow=findRow(pa,'Total Unique Patients');
    const avgRevVisitRow=findRowContains(pa,'Avg Revenue per Visit');
    const avgRevPatientRow=findRowContains(pa,'Avg Revenue per Patient');
    const marketingBudgetRow=findRowContains(pa,'Annual Marketing Budget');
    const pureCacRow=findRowContains(pa,'CAC per Marketing Patient');
    const lifetimeRow=findRowContains(pa,'Patient Lifetime');
    const ltvRow=findRowContains(pa,'Patient LTV');
    const ltvCacRow=findRowContains(pa,'LTV:CAC Ratio');
    const paybackRow=findRowContains(pa,'CAC Payback');
    const blendedCacRow=findRowContains(pa,'Blended CAC');
    const romiRow=findRowContains(pa,'ROMI');
    const retainedVisitsRow=findRowContains(pa,'Visits - Retained Patients');
    const organicVisitsRow=findRowContains(pa,'Visits - Referral + Walk-in');
    const paidVisitsRow=findRowContains(pa,'Visits - Marketing');
    const acquisitionSpendRow=findRowContains(pa,'Acquisition (CAC');
    const retentionSpendRow=findRowContains(pa,'Retention (cost');

    if(capRow>=0)BASE.capacity=getRowNums(pa,capRow,1,10);
    if(visitRow>=0)BASE.visits=getRowNums(pa,visitRow,1,10);
    if(retainedRow>=0)BASE.retained=getRowNums(pa,retainedRow,1,10);
    if(referralRow>=0)BASE.referral=getRowNums(pa,referralRow,1,10);
    if(walkinRow>=0)BASE.walkin=getRowNums(pa,walkinRow,1,10);
    if(marketingRow>=0)BASE.marketing=getRowNums(pa,marketingRow,1,10);
    if(utilRow>=0)BASE.util=getRowNums(pa,utilRow,1,10).map(v=>v>1?v:Math.round(v*100));
    if(patientsRow>=0)BASE.paModel.uniquePatients=getRowNums(pa,patientsRow,1,10);
    if(avgRevVisitRow>=0)BASE.paModel.avgRevenuePerVisit=getRowNums(pa,avgRevVisitRow,1,10);
    if(avgRevPatientRow>=0)BASE.paModel.avgRevenuePerPatient=getRowNums(pa,avgRevPatientRow,1,10);
    if(marketingBudgetRow>=0)BASE.paModel.marketingBudget=getRowNums(pa,marketingBudgetRow,1,10);
    if(pureCacRow>=0)BASE.paModel.pureCAC=getRowNums(pa,pureCacRow,1,10);
    if(lifetimeRow>=0)BASE.paModel.lifetimeYears=getRowNums(pa,lifetimeRow,1,10);
    if(ltvRow>=0)BASE.paModel.ltv=getRowNums(pa,ltvRow,1,10);
    if(ltvCacRow>=0)BASE.paModel.ltvCac=getRowNums(pa,ltvCacRow,1,10);
    if(paybackRow>=0)BASE.paModel.cacPayback=getRowNums(pa,paybackRow,1,10);
    if(blendedCacRow>=0)BASE.paModel.blendedCAC=getRowNums(pa,blendedCacRow,1,10);
    if(romiRow>=0)BASE.paModel.romi=getRowNums(pa,romiRow,1,10);
    if(retainedVisitsRow>=0)BASE.paModel.retainedVisits=getRowNums(pa,retainedVisitsRow,1,10);
    if(organicVisitsRow>=0)BASE.paModel.organicVisits=getRowNums(pa,organicVisitsRow,1,10);
    if(paidVisitsRow>=0)BASE.paModel.paidVisits=getRowNums(pa,paidVisitsRow,1,10);
    if(acquisitionSpendRow>=0)BASE.paModel.acquisitionSpend=getRowNums(pa,acquisitionSpendRow,1,10);
    if(retentionSpendRow>=0)BASE.paModel.retentionSpend=getRowNums(pa,retentionSpendRow,1,10);
    if(visitRow>=0)workbookChecklistRaw.totalVisits=getRowNums(pa,visitRow,1,10);
    if(retainedRow>=0)workbookChecklistRaw.retained=getRowNums(pa,retainedRow,1,10);
    if(referralRow>=0)workbookChecklistRaw.referral=getRowNums(pa,referralRow,1,10);
    if(walkinRow>=0)workbookChecklistRaw.walkin=getRowNums(pa,walkinRow,1,10);
    if(marketingRow>=0)workbookChecklistRaw.marketing=getRowNums(pa,marketingRow,1,10);
    if(patientsRow>=0)workbookChecklistRaw.uniquePatients=getRowNums(pa,patientsRow,1,10);
    if(avgRevPatientRow>=0)workbookChecklistRaw.avgRevenuePerPatient=getRowNums(pa,avgRevPatientRow,1,10);
    if(pureCacRow>=0)workbookChecklistRaw.pureCAC=getRowNums(pa,pureCacRow,1,10);
    if(blendedCacRow>=0)workbookChecklistRaw.blendedCAC=getRowNums(pa,blendedCacRow,1,10);
    if(ltvRow>=0)workbookChecklistRaw.ltv=getRowNums(pa,ltvRow,1,10);
    if(ltvCacRow>=0)workbookChecklistRaw.ltvCac=getRowNums(pa,ltvCacRow,1,10);
    if(paybackRow>=0)workbookChecklistRaw.cacPayback=getRowNums(pa,paybackRow,1,10);
    if(lifetimeRow>=0)workbookChecklistRaw.lifetimeYears=getRowNums(pa,lifetimeRow,1,10);
    if(marketingBudgetRow>=0)workbookChecklistRaw.marketingBudget=getRowNums(pa,marketingBudgetRow,1,10);
    if(retainedVisitsRow>=0)workbookChecklistRaw.retainedVisits=getRowNums(pa,retainedVisitsRow,1,10);
    if(organicVisitsRow>=0)workbookChecklistRaw.organicVisits=getRowNums(pa,organicVisitsRow,1,10);
    if(paidVisitsRow>=0)workbookChecklistRaw.paidVisits=getRowNums(pa,paidVisitsRow,1,10);
    if(acquisitionSpendRow>=0)workbookChecklistRaw.acquisitionSpend=getRowNums(pa,acquisitionSpendRow,1,10);
    markMapped(capRow>=0,'Capacity');
    markMapped(visitRow>=0,'Visits');
    markMapped(retainedRow>=0,'Retained Patients');
    markMapped(referralRow>=0,'Referral Patients');
    markMapped(walkinRow>=0,'Walk-in Patients');
    markMapped(marketingRow>=0,'Marketing Patients');
    markMapped(utilRow>=0,'Utilization');
    markMapped(patientsRow>=0,'Unique Patients');
    markMapped(avgRevVisitRow>=0,'Avg Rev / Visit');
    markMapped(avgRevPatientRow>=0,'Avg Rev / Patient');
    markMapped(marketingBudgetRow>=0,'Annual Marketing Budget');
    markMapped(pureCacRow>=0,'Pure CAC');
    markMapped(lifetimeRow>=0,'Patient Lifetime');
    markMapped(ltvRow>=0,'Patient LTV');
    markMapped(ltvCacRow>=0,'LTV:CAC Ratio');
    markMapped(paybackRow>=0,'CAC Payback');
    markMapped(blendedCacRow>=0,'Blended CAC');
    markMapped(romiRow>=0,'ROMI');
    markMapped(retainedVisitsRow>=0,'Retained Visits');
    markMapped(organicVisitsRow>=0,'Organic Visits');
    markMapped(paidVisitsRow>=0,'Paid Visits');
    markMapped(acquisitionSpendRow>=0,'Acquisition Spend');
    markMapped(retentionSpendRow>=0,'Retention Spend');
  }

  // CapEx breakdown
  if(cap){
    const capexEntries=[
      {label:'Rent Deposit',lookup:'Rent Deposit (3 months)',life:0,color:'#ec4899'},
      {label:'Renovation & Fit-out',lookup:'Renovation & Fit-out',life:7,color:'#6366f1'},
      {label:'Physio Equipment',lookup:'Physio Subtotal',life:7,color:'#22d3ee'},
      {label:'Assessment Equipment',lookup:'Assessment Subtotal',life:7,color:'#3b82f6'},
      {label:'Dental Equipment',lookup:'Dental Subtotal',life:7,color:'#10b981'},
      {label:'Laboratory & Medical Equipment',lookup:'Laboratory & Medical Equipment',life:7,color:'#14b8a6'},
      {label:'Cosmetology & Beauty',lookup:'Cosmetology & Beauty',life:7,color:'#f59e0b'},
      {label:'Office & IT',lookup:'Office & IT Subtotal',life:7,color:'#8b5cf6'},
      {label:'Design / MEP / Engineering',lookup:'Design / MEP / Engineering',life:7,color:'#a855f7'},
      {label:'Licensing / Legal / Hiring / Launch',lookup:'Pre-operational Subtotal',life:0,color:'#64748b',adjustment:-120000},
      {label:'Initial Inventory',lookup:'Initial Inventory',life:0,color:'#f97316'},
      {label:'Pre-opening Salaries',lookup:'Pre-opening Salaries (3 mo)',life:0,color:'#fb7185'},
      {label:'Working Capital',lookup:'Working Capital',life:0,color:'#f43f5e'},
      {label:'Contingency',lookup:'Contingency (5%)',life:0,color:'#94a3b8'}
    ].map(entry=>{
      const row=findRow(cap,entry.lookup);
      let amount=row>=0?(Number(getRow(cap,row)[3])||0):0;
      if(entry.adjustment)amount=Math.max(amount+entry.adjustment,0);
      return{...entry,amount};
    });
    if(capexEntries.some(entry=>entry.amount>0)){
      BASE.capexLabels=capexEntries.map(entry=>entry.label);
      BASE.capexAmounts=capexEntries.map(entry=>entry.amount);
      BASE.capexLife=capexEntries.map(entry=>entry.life);
      BASE.capexColors=capexEntries.map(entry=>entry.color);
    }
    const totalRow=findRow(cap,'TOTAL CAPEX');
    const workingCapitalRow=findRow(cap,'Working Capital');
    if(totalRow>=0){
      const totalVal=getRowNums(cap,totalRow,3,1)[0]; // "Active" column
      if(totalVal>0)BASE.investment=totalVal;
      workbookChecklistRaw.totalCapex=totalVal;
    }
    if(workingCapitalRow>=0){
      workbookChecklistRaw.workingCapital=getRowNums(cap,workingCapitalRow,3,1)[0];
    }
    markMapped(totalRow>=0,'CapEx Total');
    markMapped(workingCapitalRow>=0,'Working Capital');
  }

  // Summary data
  if(summary){
    // Benchmarks update
    const revSumRow=findRow(summary,'Revenue');
    const ebitdaSumRow=findRow(summary,'EBITDA');
    const npSumRow=findRow(summary,'Net Profit');
    // These are same as P&L, just validation
  }

  // Recompute derived ratios
  BASE.npRatio=[];BASE.cfRatio=[];
  for(let i=0;i<10;i++){
    BASE.npRatio[i]=BASE.ebitda[i]!==0?BASE.netProfit[i]/BASE.ebitda[i]:0;
    BASE.cfRatio[i]=BASE.ebitda[i]!==0?(BASE.annCF[i+1]||BASE.ebitda[i])/BASE.ebitda[i]:1;
  }
  BASE.opexRate=BASE.margins.map(m=>1-m-0.40);
  syncPLBreakdown();
  syncServiceRevenue();
  syncPatientAcquisitionModel();
  syncProductCatalog();
  buildExcelChecklistPayload(workbookChecklistRaw,filename);
  const actualSheets=wb.SheetNames.filter(name=>/actual|fact|variance/i.test(String(name)));
  importAudit={
    fileName:filename||'Workbook',
    importedAt:new Date().toISOString(),
    sheets:wb.SheetNames.slice(),
    mappedMetrics:Array.from(mappedMetricSet),
    actualSheets,
    actualsStatus:actualSheets.length
      ? `Detected ${actualSheets.length} actuals-oriented sheet(s): ${actualSheets.join(', ')}`
      : 'No actuals sheet detected in workbook'
  };
  writeStoredUI(APP_STORAGE_KEYS.importAudit,importAudit,'project');
  setModelSourceState(createModelSourceState({
    type:'workbook',
    reason:'loaded',
    origin:'manual',
    fileName:filename||'Workbook',
    loadedAt:importAudit.importedAt
  }));

  // Rebuild dashboard
  monteCarloLastResults=null;
  rebuildDashboard();
}

function rebuildDashboard(){
  // Destroy all existing charts
  Object.keys(charts).forEach(k=>{
    if(charts[k]&&charts[k].destroy){charts[k].destroy();}
    delete charts[k];
  });

  // Recalculate model
  recalculate();

  // Update all UI
  updateKPIs();
  updateLiveMetrics();
  initOverviewCharts();
  initBenchmarks();
  initChecks();

  // Re-init active section
  const sec=document.querySelector('.dashboard-section.active');
  if(sec){
    hydrateSection(sec);
  }

  renderYearSelection(false);
  updateFooterCurrency();
  refreshScenarioLibraryMetrics();
  renderScenarioLibrary();
  renderServiceLineEconomics();
  renderPatientAcquisitionModel();
  renderPatientBaseBridge();
  renderFundingMetrics();
  renderEquityModel();
  renderAssumptionsPanel();
  renderSourcesPanel();
  renderTemplateCompliance();
  renderImportAudit();
  renderExcelChecklist();
  renderActualVsPlan();
  markMonteCarloStale(true);
  console.log('✅ Dashboard rebuilt from Excel data');
}
