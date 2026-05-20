'use strict';

// ========== 1. 表格欄位定義 ==========
// 每個 column: { key, label, type, td_class?, auto? }
//   type: 'number' | 'text' | 'textarea' | 'level'
//   auto: 'dep_app' (= orig + dep_diff) | 'gov_app' (= dep_app + gov_diff)

const COL = {
    // 內部 key 維持 dec112/dec113/bud114/apr114 以維持 autosave 相容；標籤對應新年度
    dec112: { key: 'dec112', label: '113年度決算', type: 'number', td_class: 'sf-num' },
    dec113: { key: 'dec113', label: '114年度決算', type: 'number', td_class: 'sf-num' },
    bud114: { key: 'bud114', label: '115年度預算', type: 'number', td_class: 'sf-num' },
    apr114: { key: 'apr114', label: '115年4月底在職', type: 'number', td_class: 'sf-num' },
    level:  { key: 'level',  label: '層級',         type: 'level',  td_class: 'sf-lv' },
    name:   { key: 'name',   label: '項目/科目名稱', type: 'text',   td_class: 'sf-name' },
    desc:   { key: 'desc',   label: '摘要/說明',     type: 'textarea', td_class: 'sf-desc' },
    orig:    { key: 'orig',    label: '原編數',     type: 'number', td_class: 'sf-num' },
    dep_diff:{ key: 'dep_diff',label: '主管增減(-)', type: 'number', td_class: 'sf-num' },
    dep_app: { key: 'dep_app', label: '主管核列',   type: 'number', td_class: 'sf-num', auto: 'dep_app' },
    gov_diff:{ key: 'gov_diff',label: '院擬增減(-)', type: 'number', td_class: 'sf-num' },
    gov_app: { key: 'gov_app', label: '院擬核列',   type: 'number', td_class: 'sf-num', auto: 'gov_app' }
};

// 業務計畫表共用結構（6 基金 × 甲/乙 = 12 表，共用同一組欄位）
const PLAN_BASE = {
    sheetId: '1-2',
    unit: '新臺幣千元',
    cols: ['level', 'dec113', 'bud114', 'name', 'desc', 'orig', 'dep_diff', 'dep_app', 'gov_diff', 'gov_app'],
    hasLevel: true,
    reviewKey: 'plan',
    isPlan: true
};
const PLAN_FUNDS = [
    { id: 'plan_agri',     name: '農業發展基金',         short: '農發',    rank: 1 },
    { id: 'plan_forest',   name: '林務發展及造林基金',     short: '林務',    rank: 2 },
    { id: 'plan_disaster', name: '農業天然災害救助基金',   short: '天災',    rank: 3 },
    { id: 'plan_fish',     name: '漁業發展基金',         short: '漁業',    rank: 4 },
    { id: 'plan_loss',     name: '農產品受進口損害救助基金', short: '農損',  rank: 5 },
    { id: 'plan_renewal',  name: '農村再生基金',         short: '農再',    rank: 6 }
];
const PLAN_SECTIONS = [
    { suffix: 'src', label: '甲、基金來源' },
    { suffix: 'use', label: '乙、基金用途' }
];

const tableConfigs = {};
PLAN_FUNDS.forEach(f => {
    PLAN_SECTIONS.forEach(sec => {
        const tid = `${f.id}_${sec.suffix}`;
        tableConfigs[tid] = {
            ...PLAN_BASE,
            title: `${f.name} 116年度 主要業務計畫預算表 — ${sec.label}`,
            fundId: f.id,
            fundName: f.name,
            fundShort: f.short,
            section: sec.suffix,
            sectionLabel: sec.label,
            rank: f.rank
        };
    });
});
Object.assign(tableConfigs, {
    headcount: {
        title: '甲、員額',
        sheetId: '1-3 上半部',
        unit: '人',
        cols: ['dec112', 'dec113', 'bud114', 'apr114', 'name', 'orig', 'dep_diff', 'dep_app', 'gov_diff', 'gov_app', 'desc'],
        hasLevel: false,
        reviewKey: null // 員額表自身無審核意見
    },
    personnel_cost: {
        title: '乙、用人費用',
        sheetId: '1-3 下半部',
        unit: '新臺幣千元',
        cols: ['dec112', 'dec113', 'bud114', 'name', 'orig', 'dep_diff', 'dep_app', 'gov_diff', 'gov_app', 'desc'],
        hasLevel: false,
        reviewKey: 'personnel'
    },
    control: {
        title: '116 年度 其他管制性項目及重大事項預算表',
        sheetId: '1-4',
        unit: '新臺幣千元',
        cols: ['dec112', 'dec113', 'bud114', 'name', 'orig', 'dep_diff', 'dep_app', 'gov_diff', 'gov_app', 'desc'],
        hasLevel: false,
        reviewKey: 'control'
    }
});

// 業務計畫表內部層級（甲/乙 已升格為表頭，不再佔層級）
// 註：L2 在實際 Word 中常為「計畫名（無前綴）」，少數情況才以 (一)/(二) 標示
const LEVEL_LABELS = { 0: '─', 1: '一、二、', 2: '計畫名 / (一)', 3: '1./2.', 4: '(1)/(2)' };

// ========== 2. 預設項目（依各分基金 Word 原表，數字留空） ==========
const SAMPLES = {
    // === 業務計畫 — 農發基金（依「農發基金115-212業務計畫」）===
    plan_agri: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一、一般業務計畫' },
        { level: 3, name: '二、國庫撥補款' },
        { level: 2, name: '乙、基金用途：' },
        { level: 3, name: '一、提升農業經營及發展計畫' },
        { level: 4, name: '糧政業務計畫' },
        { level: 4, name: '1.收購糧食' },
        { level: 5, name: '(1)數量：公噸(稻穀)' },
        { level: 5, name: '(2)單位成本：元' },
        { level: 4, name: '2.糧食銷售(主產品)' },
        { level: 5, name: '(1)銷售量：公噸(折糙)' },
        { level: 5, name: '(2)單位成本：元' },
        { level: 4, name: '穩定肥料及相關資材供需計畫' },
        { level: 4, name: '產銷調節計畫' },
        { level: 4, name: '家禽流行性感冒防疫計畫' },
        { level: 3, name: '二、促進農地利用及農業競爭力計畫' },
        { level: 4, name: '老農出租農地獎勵計畫' },
        { level: 4, name: '農地之生產環境整備及維護管理計畫' },
        { level: 4, name: '農業研究、實驗、技術改進計畫' },
        { level: 4, name: '農地對地給付計畫' },
        { level: 3, name: '三、增進農民所得及福利計畫' },
        { level: 4, name: '獎勵農漁民子女就學計畫' },
        { level: 4, name: '農業保險計畫' },
        { level: 4, name: '精進豬隻保險業務計畫' },
        { level: 4, name: '精進乳牛保險業務計畫' },
        { level: 4, name: '家禽禽流感保險計畫' },
        { level: 4, name: '農產業保險計畫' },
        { level: 4, name: '漁產業保險計畫' },
        { level: 4, name: '農業保險推動及輔導計畫' },
        { level: 4, name: '農業貸款利息差額補貼' },
        { level: 4, name: '農機貸款' },
        { level: 4, name: '加速農村建設貸款' },
        { level: 4, name: '擴大家庭農場經營規模協助農民購買耕地貸款' },
        { level: 4, name: '各類專案性農業貸款利息補貼' },
        { level: 4, name: '養豬新式設施(備)導入提供專案政策性農貸利息補貼' },
        { level: 4, name: '委託辦理政策性農業專案貸款業務及印製宣傳摺頁' },
        { level: 4, name: '辦理政策性農業專案貸款行政事務' },
        { level: 3, name: '四、輔導菸農轉型與檳榔廢園轉作計畫' },
        { level: 3, name: '五、處理農會漁會信用部計畫' },
        { level: 3, name: '六、一般行政管理計畫' }
    ],
    // === 業務計畫 — 林務基金（依「林務基金212-115」）===
    plan_forest: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一、一般業務計畫之基金來源合計' },
        { level: 2, name: '乙、基金用途：' },
        { level: 3, name: '一、獎勵輔導造林計畫' },
        { level: 3, name: '二、森林遊樂及林業鐵路經營管理計畫' },
        { level: 3, name: '三、山坡地開發利用回饋金繳交管理計畫' },
        { level: 3, name: '四、原住民保留地竹林更新獎勵計畫' }
    ],
    // === 業務計畫 — 農業天然災害救助基金 ===
    plan_disaster: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一般業務計畫' },
        { level: 3, name: '國庫撥補額' },
        { level: 2, name: '乙、基金用途：' },
        { level: 3, name: '農業天然災害救助計畫' }
    ],
    // === 業務計畫 — 漁業基金（依「漁發基金115-212業務計畫」）===
    plan_fish: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '財產收入' },
        { level: 3, name: '其他收入' },
        { level: 2, name: '乙、基金用途：' },
        { level: 3, name: '漁業發展補助計畫' }
    ],
    // === 業務計畫 — 農產品受進口損害救助基金（依「農損基金-212」）===
    plan_loss: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一、權利金收入' },
        { level: 3, name: '二、利息收入' },
        { level: 3, name: '三、公庫撥款收入' },
        { level: 3, name: '四、其他收入' },
        { level: 2, name: '乙、基金用途：' },
        { level: 3, name: '一、調整產業或防範措施計畫' },
        { level: 3, name: '二、進口損害救助及穩價計畫' },
        { level: 3, name: '三、農糧產業調整與轉型計畫(原「綠色環境給付計畫」)' },
        { level: 3, name: '四、農產品進口管理計畫' }
    ],
    // === 業務計畫 — 農村再生基金（依「農再基金-212」）===
    plan_renewal: [
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一、一般業務計畫' },
        { level: 3, name: '二、國庫撥補款' },
        { level: 2, name: '乙、基金用途：' },
        // 農再有 壹/貳 額外層，所以 壹/貳 對應 L1，下面 一/二/三 對應 L2，1./2. 對應 L3
        { level: 3, name: '壹、農村再生規劃及人力培育計畫' },
        { level: 4, name: '一、農村規劃及培力' },
        { level: 4, name: '1.農村人力及教育推廣' },
        { level: 4, name: '二、農業人才多元培育' },
        { level: 4, name: '三、農村農產業人力活化計畫' },
        { level: 3, name: '貳、農村再生建設及發展計畫' },
        { level: 4, name: '一、農村再生社區發展及環境改善' },
        { level: 4, name: '1.農村再生跨域發展' },
        { level: 4, name: '2.社區農村再生計畫' },
        { level: 4, name: '3.農村社區土地重劃及發展規劃' },
        { level: 4, name: '二、農村發展及活化' },
        { level: 4, name: '1.提升農村農糧產業競爭力' },
        { level: 4, name: '2.發展健康永續的有機產業' },
        { level: 4, name: '3.農村社區畜牧場環境改善及資源利用' },
        { level: 4, name: '4.建設休閒農業優質環境' },
        { level: 4, name: '5.友善漁業生產環境及漁村產業活動推廣' },
        { level: 4, name: '6.山村綠色經濟永續發展計畫' },
        { level: 4, name: '7.農產加工整合服務體系發展' },
        { level: 4, name: '8.運用資通訊技術(ICT)強化農業發展及推廣計畫' },
        { level: 4, name: '9.優化農業推廣教育訓練場域' },
        { level: 4, name: '10.推動化學農藥減量模式促進農村生產環境永續發展' },
        { level: 4, name: '11.農村水資源韌性公共設施建設' },
        { level: 4, name: '12.推動農會經濟事業發展，振興農村產業' },
        { level: 4, name: '13.改善農村金融服務品質' },
        { level: 4, name: '14.農業旅遊創新發展及市場拓展' },
        { level: 4, name: '三、擴大豬場導入新式整合型設施(備)' },
        { level: 4, name: '四、建構農糧產業機械示範體系與營造多元服務價值鏈' },
        { level: 4, name: '五、智能防災設施型農業計畫' }
    ],
    // === 舊版整合用 plan 已移除，向下相容透過 normalizeData() 自動 migrate ===
    _legacy_plan_DELETED: [
        { level: 1, name: '壹、農業發展基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 3, name: '一、一般業務計畫' },
        { level: 3, name: '二、國庫撥補款' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '提升農業經營及發展計畫' },
        { level: 4, name: '糧政業務計畫' },
        { level: 4, name: '收購糧食' },
        { level: 5, name: '(1)數量：公噸(稻穀)' },
        { level: 5, name: '(2)單位成本：元' },
        { level: 4, name: '糧食銷售(主產品)' },
        { level: 5, name: '(1)銷售量：公噸(折糙)' },
        { level: 5, name: '(2)單位成本：元' },
        { level: 4, name: '穩定肥料及相關資材供需計畫' },
        { level: 4, name: '產銷調節計畫' },
        { level: 4, name: '家禽流行性感冒防疫計畫' },
        { level: 4, name: '促進農地利用及農業競爭力計畫' },
        { level: 4, name: '老農出租農地獎勵計畫' },
        { level: 4, name: '農地之生產環境整備及維護管理計畫' },
        { level: 4, name: '農業研究、實驗、技術改進計畫' },
        { level: 4, name: '農地對地給付計畫' },
        { level: 4, name: '增進農民所得及福利計畫' },
        { level: 4, name: '獎勵農漁民子女就學計畫' },
        { level: 4, name: '農業保險計畫' },
        { level: 4, name: '精進豬隻保險業務計畫' },
        { level: 4, name: '精進乳牛保險業務計畫' },
        { level: 4, name: '家禽禽流感保險計畫' },
        { level: 4, name: '農產業保險計畫' },
        { level: 4, name: '漁產業保險計畫' },
        { level: 4, name: '農業保險推動及輔導計畫' },
        { level: 4, name: '農業貸款利息差額補貼' },
        { level: 4, name: '農機貸款' },
        { level: 4, name: '加速農村建設貸款' },
        { level: 4, name: '擴大家庭農場經營規模協助農民購買耕地貸款' },
        { level: 4, name: '各類專案性農業貸款利息補貼' },
        { level: 4, name: '養豬新式設施(備)導入提供專案政策性農貸利息補貼' },
        { level: 4, name: '委託辦理政策性農業專案貸款業務及印製宣傳摺頁' },
        { level: 4, name: '辦理政策性農業專案貸款行政事務' },
        { level: 4, name: '輔導菸農轉型與檳榔廢園轉作計畫' },
        { level: 4, name: '處理農會漁會信用部計畫' },
        { level: 4, name: '一般行政管理計畫' },
        { level: 1, name: '貳、林務發展及造林基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 4, name: '一般業務計畫' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '獎勵輔導造林計畫' },
        { level: 4, name: '森林遊樂及林業鐵路經營管理計畫' },
        { level: 4, name: '山坡地開發利用回饋金繳交管理計畫' },
        { level: 4, name: '原住民保留地竹林更新獎勵計畫' },
        { level: 1, name: '叁、農業天然災害救助基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 4, name: '一般業務計畫' },
        { level: 4, name: '國庫撥補款' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '農業天然災害救助計畫' },
        { level: 1, name: '肆、漁業發展基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 4, name: '財產收入' },
        { level: 4, name: '其他收入' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '漁業發展補助計畫' },
        { level: 1, name: '伍、農產品受進口損害救助基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 4, name: '權利金收入' },
        { level: 4, name: '利息收入' },
        { level: 4, name: '國庫撥補款' },
        { level: 4, name: '其他收入' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '調整產業或防範措施計畫' },
        { level: 4, name: '進口損害救助及穩價計畫' },
        { level: 4, name: '農糧產業調整與轉型計畫(原「綠色環境給付計畫」)' },
        { level: 4, name: '農產品進口管理計畫' },
        { level: 1, name: '陸、農村再生基金' },
        { level: 2, name: '甲、基金來源：' },
        { level: 4, name: '一般業務計畫' },
        { level: 4, name: '國庫撥補款' },
        { level: 2, name: '乙、基金用途：' },
        { level: 4, name: '農村再生規劃及人力培育計畫' },
        { level: 4, name: '農村規劃及培力' },
        { level: 4, name: '農村人力及教育推廣' },
        { level: 4, name: '農業人才多元培育' },
        { level: 4, name: '農村農產業人力活化計畫' },
        { level: 4, name: '農村再生建設及發展計畫' },
        { level: 4, name: '農村再生社區發展及環境改善' },
        { level: 4, name: '農村再生跨域發展' },
        { level: 4, name: '社區農村再生計畫' },
        { level: 4, name: '農村社區土地重劃及發展規劃' },
        { level: 4, name: '農村發展及活化' },
        { level: 4, name: '提升農村農糧產業競爭力' },
        { level: 4, name: '發展健康永續的有機產業' },
        { level: 4, name: '農村社區畜牧場環境改善及資源利用' },
        { level: 4, name: '建設休閒農業優質環境' },
        { level: 4, name: '友善漁業生產環境及漁村產業活動推廣' },
        { level: 4, name: '山村綠色經濟永續發展計畫' },
        { level: 4, name: '農產加工整合服務體系發展' },
        { level: 4, name: '運用資通訊技術(ICT)強化農業發展及推廣計畫' },
        { level: 4, name: '優化農業推廣教育訓練場域' },
        { level: 4, name: '推動化學農藥減量模式促進農村生產環境永續發展' },
        { level: 4, name: '農村水資源韌性公共設施建設' },
        { level: 4, name: '推動農會經濟事業發展，振興農村產業' },
        { level: 4, name: '改善農村金融服務品質' },
        { level: 4, name: '農業旅遊創新發展及市場拓展' },
        { level: 4, name: '擴大豬場導入新式整合型設施(備)' },
        { level: 4, name: '建構農糧產業機械示範體系與營造多元服務價值鏈' },
        { level: 4, name: '智能防災設施型農業計畫' }
    ],
    headcount: [
        { name: '一、編制內預算員額' },
        { name: '　(一)職員' },
        { name: '　(二)警察' },
        { name: '　(三)駐衛警' },
        { name: '　(四)技工' },
        { name: '　(五)工友' },
        { name: '　(六)駕駛' },
        { name: '　(七)聘用' },
        { name: '　(八)約僱' },
        { name: '合　　　　計' },
        { name: '二、管理會委員' },
        { name: '三、顧問人員' },
        { name: '四、兼任人員' },
        { name: '五、資本支出' }
    ],
    personnel_cost: [
        { name: '一、正式員額薪資' },
        { name: '　(一)編制內' },
        { name: '　(二)管理會委員' },
        { name: '　(三)顧問人員' },
        { name: '二、聘僱及兼職人員薪資' },
        { name: '　(一)編制內' },
        { name: '　(二)兼職人員' },
        { name: '　(三)其他' },
        { name: '三、加(夜)班費' },
        { name: '　(一)延長工時加班費' },
        { name: '　(二)其他' },
        { name: '四、津貼' },
        { name: '五、獎金' },
        { name: '六、退休及卹償金' },
        { name: '七、資遣費' },
        { name: '八、福利費' },
        { name: '　(一)分擔員工保險費' },
        { name: '　(二)其他' },
        { name: '九、提繳費' },
        { name: '合　　　　計' },
        { name: '資本支出' }
    ],
    control: [
        { name: '一、水電費' },
        { name: '二、國內旅費' },
        { name: '三、國外旅費' },
        { name: '(一)出國考察、訪問' },
        { name: '(二)參加國際會議、談判' },
        { name: '(三)出國進修、研究及實習計畫' },
        { name: '四、大陸地區旅費' },
        { name: '五、印刷裝訂費' },
        { name: '六、媒體政策及業務宣導費' },
        { name: '七、推展費' },
        { name: '八、一般服務費（不含計時與計件人員酬金）' },
        { name: '(一)一般(不含體育活動費)' },
        { name: '(二)體育活動費' },
        { name: '九、契約勞力（約用人員）' },
        { name: '(一)工程管理費' },
        { name: '(二)一般服務費－計時與計件人員酬金' },
        { name: '(三)專業服務費－專技人員酬金' },
        { name: '十、委託調查研究費' },
        { name: '十一、公共關係費' },
        { name: '十二、員工慰勞費' },
        { name: '十三、用品消耗' },
        { name: '十四、其他費用' },
        { name: '十五、補助與捐助' },
        { name: '(一)以前年度計畫' },
        { name: '(二)新興計畫' },
        { name: '十六、公務車輛' },
        { name: '(一)管理用車輛' },
        { name: '1.新購' },
        { name: '2.汰換' },
        { name: '(二)其他車輛' },
        { name: '1.新購' },
        { name: '2.汰換' },
        { name: '十七、資產變賣' },
        { name: '(一)變賣淨收入' },
        { name: '(二)帳面價值' },
        { name: '(三)資產處分利益' },
        { name: '十八、其他重大事項' },
        { name: '(一)購置無形資產' },
        { name: '(二)補辦預算' }
    ]
};

// ========== 2.5 業務計畫 列拆解輔助（甲/乙 與層級調整）==========
// 業務計畫的層級規則（甲/乙 升格為表頭後）：
//   L1 = 一、二、…    L2 = (一)/(二) 或無編號計畫名
//   L3 = 1./2./…      L4 = (1)/(2)/(一)…
function adjustPlanLevel(name, oldLevel) {
    const nm = (name || '').trim();
    const lv = parseInt(oldLevel) || 0;
    let newLv = lv >= 3 ? lv - 2 : 0;
    // 阿拉伯數字編號 (1./2./3.) 比同層級的「計畫名」深一層
    if (/^\d+\./.test(nm)) newLv++;
    // 阿拉伯數字括弧 ((1)/(2)) 再深一層；中文數字括弧 (一)/(二) 屬同 L2 不再 bump
    else if (/^\(\d+\)/.test(nm)) newLv++;
    return Math.min(Math.max(newLv, 0), 4);
}
function splitPlanRows(combinedRows) {
    const srcRows = [], useRows = [];
    let cur = 'src';
    (combinedRows || []).forEach(row => {
        const nm = (row.name || '').trim();
        const lv = parseInt(row.level) || 0;
        if (lv === 2 && nm.startsWith('甲')) { cur = 'src'; return; }
        if (lv === 2 && nm.startsWith('乙')) { cur = 'use'; return; }
        const newRow = { ...row };
        newRow.level = adjustPlanLevel(nm, lv);
        (cur === 'src' ? srcRows : useRows).push(newRow);
    });
    return { src: srcRows, use: useRows };
}
// 初始化：把 SAMPLES.plan_agri 等拆成 plan_agri_src / plan_agri_use
PLAN_FUNDS.forEach(f => {
    const combined = SAMPLES[f.id];
    if (Array.isArray(combined)) {
        const { src, use } = splitPlanRows(combined);
        SAMPLES[`${f.id}_src`] = src;
        SAMPLES[`${f.id}_use`] = use;
        delete SAMPLES[f.id];
    }
});
delete SAMPLES._legacy_plan_DELETED;

// ========== 3. 行渲染 ==========
// 數值轉「千分號」字串；空值/非數值原樣
function formatNum(v) {
    if (v === '' || v === null || v === undefined) return '';
    const n = parseFloat(String(v).replace(/,/g, '').trim());
    if (isNaN(n)) return String(v);
    return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
}
function parseNum(v) {
    if (v === '' || v === null || v === undefined) return NaN;
    return parseFloat(String(v).replace(/,/g, '').trim());
}

function cellHTML(col, data, tableId) {
    const def = COL[col];
    const val = data?.[col] ?? '';
    if (def.type === 'level') {
        const opts = Object.keys(LEVEL_LABELS).map(n => `<option value="${n}" ${String(val)===String(n)?'selected':''}>${LEVEL_LABELS[n]}</option>`).join('');
        return `<td class="sf-lv"><select class="sf-level-select" data-key="level">${opts}</select></td>`;
    }
    if (def.type === 'textarea') {
        return `<td class="${def.td_class}"><textarea data-key="${col}" placeholder="">${val}</textarea></td>`;
    }
    if (def.type === 'text') {
        return `<td class="${def.td_class}"><input type="text" data-key="${col}" value="${escapeAttr(val)}"></td>`;
    }
    // number → 改用 type=text 以便顯示千分號（HTML number 不允許逗號）
    const isAuto = !!def.auto;
    const displayVal = val !== '' && val !== null ? formatNum(val) : '';
    return `<td class="${def.td_class}"><input type="text" inputmode="numeric" data-key="${col}" data-numeric="1" value="${escapeAttr(displayVal)}" ${isAuto ? 'readonly' : ''}></td>`;
}

function escapeAttr(s) {
    return String(s ?? '').replace(/"/g, '&quot;');
}

function escapeHTML(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function addRow(tableId, data = {}, beforeTr = null) {
    const cfg = tableConfigs[tableId];
    if (!cfg) return null;
    const tbody = document.getElementById('sf-tbody-' + tableId);
    const tr = document.createElement('tr');
    tr.setAttribute('draggable', 'true');
    const lvl = cfg.hasLevel ? (data.level ?? 0) : 0;
    if (cfg.hasLevel) tr.setAttribute('data-level', lvl);
    const actHTML = `<td class="sf-act">
        <div class="sf-insert-zone" title="在此插入一列"><span class="sf-insert-btn">＋</span></div>
        <span class="sf-drag-handle" title="拖曳排序">☰</span>
        <button class="sf-row-btn del" title="刪除">✕</button>
    </td>`;
    tr.innerHTML = actHTML + cfg.cols.map(c => cellHTML(c, data, tableId)).join('');
    if (beforeTr && beforeTr.parentNode === tbody) tbody.insertBefore(tr, beforeTr);
    else tbody.appendChild(tr);
    recalcRow(tr);
    return tr;
}

function recalcRow(tr) {
    const get = k => {
        const v = parseNum(tr.querySelector(`[data-key="${k}"]`)?.value);
        return isNaN(v) ? 0 : v;
    };
    const set = (k, v) => {
        const el = tr.querySelector(`[data-key="${k}"]`);
        if (!el) return;
        el.value = v === 0 ? '' : formatNum(v);
        el.classList.toggle('negative-value', v < 0);
    };
    const depApp = get('orig') + get('dep_diff');
    set('dep_app', depApp);
    const govApp = depApp + get('gov_diff');
    set('gov_app', govApp);
}

// 把列上所有數值欄重新套千分號（focus 出去後）
function reformatNumericCells(tr) {
    tr.querySelectorAll('input[data-numeric]').forEach(el => {
        const raw = parseNum(el.value);
        if (!isNaN(raw)) el.value = formatNum(raw);
    });
}

// === 業務計畫 合計邏輯（父列 = Σ 子列）===
// 例外：「收購糧食」「糧食銷售」之類含子列但子列為數量/單位成本者，不適用合計
const NO_AUTO_SUM_KEYWORDS = ['收購糧食', '糧食銷售'];
function shouldAutoSum(name) {
    const nm = String(name || '');
    return !NO_AUTO_SUM_KEYWORDS.some(k => nm.includes(k));
}

const PLAN_SUM_KEYS = ['dec113', 'bud114', 'orig', 'dep_diff', 'gov_diff'];

function recalcPlanTable(tableId) {
    if (!tableId || !tableId.startsWith('plan_')) return;
    const tbody = document.getElementById('sf-tbody-' + tableId);
    if (!tbody) return;
    const rows = Array.from(tbody.children);
    const nodes = rows.map(tr => ({
        tr,
        level: parseInt(tr.getAttribute('data-level')) || 0,
        name:  tr.querySelector('[data-key="name"]')?.value || '',
        children: [],
        parent: null
    }));
    // 建層級樹
    nodes.forEach((node, i) => {
        for (let j = i - 1; j >= 0; j--) {
            if (nodes[j].level < node.level) {
                nodes[j].children.push(node);
                node.parent = nodes[j];
                break;
            }
        }
    });

    function process(node) {
        node.children.forEach(process);
        const hasKids = node.children.length > 0;
        const autoSum = hasKids && shouldAutoSum(node.name);
        PLAN_SUM_KEYS.forEach(k => {
            const el = node.tr.querySelector(`[data-key="${k}"]`);
            if (!el) return;
            if (autoSum) {
                let sum = 0, has = false;
                node.children.forEach(c => {
                    const v = parseNum(c.tr.querySelector(`[data-key="${k}"]`)?.value);
                    if (!isNaN(v)) { sum += v; has = true; }
                });
                el.value = (has && sum !== 0) ? formatNum(sum) : '';
                el.readOnly = true;
                el.classList.add('sf-auto-sum');
                el.classList.toggle('negative-value', sum < 0);
            } else {
                el.readOnly = false;
                el.classList.remove('sf-auto-sum');
            }
        });
        recalcRow(node.tr);
    }
    nodes.filter(n => !n.parent).forEach(process);
}

// ========== 4. 收集 / 還原資料 ==========
function collectData() {
    const data = {
        meta: {
            fund: val('sf-fund'),
            year: val('sf-year'),
            org:  val('sf-org'),
            user: val('sf-user')
        },
        tables: {},
        reviews: {
            plan:      { org: val('sf-plan-review-org'),      gov: val('sf-plan-review-gov') },
            personnel: { org: val('sf-personnel-review-org'), gov: val('sf-personnel-review-gov') },
            control:   { org: val('sf-control-review-org'),   gov: val('sf-control-review-gov') }
        },
        savedAt: new Date().toISOString()
    };
    Object.keys(tableConfigs).forEach(tid => {
        const cfg = tableConfigs[tid];
        const rows = Array.from(document.querySelectorAll(`#sf-tbody-${tid} tr`));
        data.tables[tid] = rows.map(tr => {
            const item = {};
            if (cfg.hasLevel) item.level = parseInt(tr.getAttribute('data-level')) || 0;
            cfg.cols.forEach(c => {
                const el = tr.querySelector(`[data-key="${c}"]`);
                if (!el) return;
                const def = COL[c];
                let v = el.value;
                if (def && def.type === 'number') {
                    // 數字欄：JSON 內存「無千分號」字串
                    const n = parseNum(v);
                    v = isNaN(n) ? '' : String(n);
                }
                item[c] = v;
            });
            return item;
        });
    });
    return data;
}

function applyData(data) {
    if (!data) return;
    if (data.meta) {
        document.getElementById('sf-fund').value = data.meta.fund || '';
        document.getElementById('sf-year').value = data.meta.year || '';
        document.getElementById('sf-org').value  = data.meta.org  || '';
        document.getElementById('sf-user').value = data.meta.user || '';
    }
    if (data.reviews) {
        ['plan','personnel','control'].forEach(k => {
            const r = data.reviews[k] || {};
            const orgEl = document.getElementById(`sf-${k}-review-org`);
            const govEl = document.getElementById(`sf-${k}-review-gov`);
            if (orgEl) orgEl.value = r.org || '';
            if (govEl) govEl.value = r.gov || '';
        });
    }
    Object.keys(tableConfigs).forEach(tid => {
        const tbody = document.getElementById('sf-tbody-' + tid);
        if (!tbody) return;
        tbody.innerHTML = '';
        const items = data.tables?.[tid] || [];
        if (items.length) items.forEach(it => addRow(tid, it));
        if (tid.startsWith('plan_')) recalcPlanTable(tid);
    });
}

function val(id) { return (document.getElementById(id)?.value || '').trim(); }

// ========== 5. JSON 匯入/匯出 ==========
function exportJSON() {
    const data = collectData();
    const fname = `${data.meta.fund || '特別收入基金'}_${data.meta.year || ''}_細項預算.json`;
    saveAs(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), fname);
}

// 中文欄位別名 → 內部 key（用於 patch 模式）
const FIELD_ALIASES = {
    // 三年表（員額/用人費/管制）
    '113決算': 'dec112', '113年度決算': 'dec112', '113年度決算數': 'dec112',
    '114決算': 'dec113', '114年度決算': 'dec113', '114年度決算數': 'dec113',
    '115預算': 'bud114', '115年度預算': 'bud114', '115年度預算數': 'bud114',
    '115年4月底': 'apr114', '115年4月底在職': 'apr114', '115年4月底在職人員': 'apr114',
    // 共用
    '原編': 'orig', '原編數': 'orig', '原編金額': 'orig', '原編員額': 'orig',
    '主管增減': 'dep_diff', '主管請增減': 'dep_diff', '主管增減數': 'dep_diff',
    '院擬增減': 'gov_diff', '院增減': 'gov_diff', '院請增減': 'gov_diff', '行政院初審擬增減': 'gov_diff', '行政院初審增減': 'gov_diff',
    '說明': 'desc', '摘要': 'desc', '編列說明': 'desc',
    '層級': 'level', 'level': 'level',
    '項目': 'name', '項目名稱': 'name', '科目': 'name', '科目名稱': 'name', '計畫名稱': 'name', '科目或計畫名稱': 'name'
};

function normalizeEntryKeys(entry) {
    const out = {};
    Object.keys(entry).forEach(k => {
        const trimmed = k.trim();
        const mapped = FIELD_ALIASES[trimmed] || trimmed;
        // 自動計算欄位略過（避免覆蓋）
        if (mapped === 'dep_app' || mapped === 'gov_app') return;
        out[mapped] = entry[k];
    });
    return out;
}

function isFullExportFormat(data) {
    return data && typeof data === 'object' && data.tables && typeof data.tables === 'object' && !Array.isArray(data.tables);
}

// 數值 patch 模式：按 name 比對，找到就更新，找不到就新增列
function patchData(patch) {
    const current = collectData();
    let totalUpdates = 0, notFoundAdded = 0;

    Object.keys(patch).forEach(tid => {
        if (!tableConfigs[tid]) {
            console.warn('未知的表格 ID：', tid);
            return;
        }
        const tableData = patch[tid];
        // 兩種輸入格式皆接受：陣列 [{name, ...}] 或 物件 { name1: {...}, name2: {...} }
        const entries = Array.isArray(tableData)
            ? tableData
            : Object.entries(tableData).map(([name, fields]) => ({ name, ...fields }));

        const rows = current.tables[tid] || (current.tables[tid] = []);
        const lookup = new Map();
        rows.forEach((r, i) => {
            const nm = (r.name || '').trim();
            if (nm && !lookup.has(nm)) lookup.set(nm, i);
        });

        entries.forEach(entry => {
            const norm = normalizeEntryKeys(entry);
            const name = (norm.name || '').trim();
            if (!name) return;
            const idx = lookup.get(name);
            if (idx !== undefined) {
                Object.assign(rows[idx], norm);
            } else {
                rows.push(norm);
                lookup.set(name, rows.length - 1);
                notFoundAdded++;
            }
            totalUpdates++;
        });
    });

    applyData(current);
    return { totalUpdates, notFoundAdded };
}

function handleImport(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const raw = JSON.parse(e.target.result);
            if (isFullExportFormat(raw)) {
                // 完整匯出格式 → 整批替換
                const data = normalizeData(raw);
                applyData(data);
                scheduleAutosave();
                flashAutosave('✓ JSON 匯入成功（整批替換）');
            } else {
                // Patch 格式（依項目名稱比對）
                const { totalUpdates, notFoundAdded } = patchData(raw);
                scheduleAutosave();
                flashAutosave(`✓ 數值匯入完成：更新／新增 ${totalUpdates} 項${notFoundAdded ? `（其中 ${notFoundAdded} 項為新增）` : ''}`);
            }
        } catch (err) {
            alert('匯入失敗：' + err.message);
        }
    };
    reader.readAsText(file);
}

// ========== 6. 自動儲存 ==========
const STORE_KEY = 'sf_special_fund_v4'; // v4: 業務計畫每基金拆 甲/乙 上下表
const LEGACY_KEYS = ['sf_special_fund_v3', 'sf_special_fund_v2', 'sf_special_fund_v1'];

// 向下相容：把舊版 tables.plan（一張大表）依 L1 編號拆到 6 個 plan_* 表
function migrateOldPlan(data) {
    if (!data?.tables?.plan) return data;
    const mapping = { '壹': 'plan_agri', '貳': 'plan_forest', '叁': 'plan_disaster', '參': 'plan_disaster', '肆': 'plan_fish', '伍': 'plan_loss', '陸': 'plan_renewal' };
    const newTables = { plan_agri: [], plan_forest: [], plan_disaster: [], plan_fish: [], plan_loss: [], plan_renewal: [] };
    let current = 'plan_agri';
    (data.tables.plan || []).forEach(row => {
        const lv = parseInt(row.level) || 0;
        const nm = (row.name || '').trim();
        if (lv === 1 && nm) {
            const ch = nm[0];
            if (mapping[ch]) current = mapping[ch];
            return; // 不保留 L1 列本身
        }
        const newRow = { ...row };
        // 各 plan_* 的內部最大層級從 2 起算 → 整體 -1
        if (newRow.level && parseInt(newRow.level) > 1) newRow.level = parseInt(newRow.level) - 1;
        newTables[current].push(newRow);
    });
    const result = { ...data, tables: { ...newTables } };
    ['headcount','personnel_cost','control'].forEach(k => {
        if (data.tables[k]) result.tables[k] = data.tables[k];
    });
    return result;
}

// v3 → v4: 把 plan_agri (合併) 拆成 plan_agri_src + plan_agri_use
function migratePlanSplit(data) {
    if (!data?.tables) return data;
    const result = { ...data, tables: { ...data.tables } };
    PLAN_FUNDS.forEach(f => {
        const old = result.tables[f.id];
        if (!Array.isArray(old)) return;
        const { src, use } = splitPlanRows(old);
        result.tables[`${f.id}_src`] = src;
        result.tables[`${f.id}_use`] = use;
        delete result.tables[f.id];
    });
    return result;
}

// 統一資料正規化入口（任何讀進來的資料都先過這層）
function normalizeData(data) {
    if (!data?.tables) return data;
    // v1/v2: tables.plan 存在 → 先拆成 plan_agri/forest/...
    if (Array.isArray(data.tables.plan)) data = migrateOldPlan(data);
    // v3: 若仍有 plan_agri（合併版）→ 再拆成 _src + _use
    const hasV3Plans = PLAN_FUNDS.some(f => Array.isArray(data.tables[f.id]));
    if (hasV3Plans) data = migratePlanSplit(data);
    return data;
}

let autosaveTimer = null;
function scheduleAutosave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => {
        try {
            localStorage.setItem(STORE_KEY, JSON.stringify(collectData()));
            flashAutosave('✓ 已自動儲存');
        } catch (e) { /* quota? */ }
    }, 800);
}
function loadAutosave() {
    try {
        // 先讀 v3
        let raw = localStorage.getItem(STORE_KEY);
        // v3 沒有 → 嘗試遷移舊版
        if (!raw) {
            for (const k of LEGACY_KEYS) {
                const old = localStorage.getItem(k);
                if (old) { raw = old; break; }
            }
        }
        if (!raw) return false;
        const data = normalizeData(JSON.parse(raw));
        applyData(data);
        return true;
    } catch (e) { console.warn('loadAutosave failed', e); return false; }
}
function flashAutosave(msg) {
    const el = document.getElementById('sf-autosave-indicator');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(flashAutosave._t);
    flashAutosave._t = setTimeout(() => el.classList.remove('show'), 1500);
}

// ========== 7. Word 匯出（HTML-as-Word .doc） ==========
// fmtNum / parseNum 已於上方定義（formatNum + parseNum）。Word 匯出沿用 formatNum 與 numColor
const fmtNum = formatNum;
function numColor(v) {
    const n = parseNum(v);
    return (!isNaN(n) && n < 0) ? ' style="color:#c00"' : '';
}

function buildTableDocHTML(tableId) {
    // 基金級 ID（plan_agri…）→ 合併 src+use 為單一表（中間插入「甲、基金來源」「乙、基金用途」分隔列）
    if (PLAN_FUNDS.some(f => f.id === tableId)) {
        return buildPlanFundDocHTML(tableId);
    }
    const cfg = tableConfigs[tableId];
    if (!cfg) return '';
    const data = collectData();
    const meta = data.meta;
    const rows = data.tables[tableId] || [];
    const reviews = cfg.reviewKey ? data.reviews[cfg.reviewKey] : null;

    // header build
    let theadHTML = '';
    if (tableId === 'plan' || tableId.startsWith('plan_')) {
        theadHTML = `
        <thead>
            <tr>
                <th rowspan="2">114年度<br/>決算數</th>
                <th rowspan="2">115年度<br/>預算數</th>
                <th rowspan="2">科目或計畫名稱</th>
                <th rowspan="2">政策目標、計畫實施內容(或工作項目)<br/>及預期效益摘要</th>
                <th rowspan="2">原編數</th>
                <th colspan="2">主管機關</th>
                <th colspan="2">行政院初審</th>
            </tr>
            <tr>
                <th>增減(-)數</th><th>核列數</th>
                <th>擬增減(-)數</th><th>擬核列數</th>
            </tr>
        </thead>`;
    } else if (tableId === 'headcount') {
        theadHTML = `
        <thead>
            <tr>
                <th rowspan="2">113年度<br/>決算</th>
                <th rowspan="2">114年度<br/>決算</th>
                <th rowspan="2">115年度<br/>預算</th>
                <th rowspan="2">115年4月底<br/>在職人員</th>
                <th rowspan="2">項目</th>
                <th rowspan="2">原編數</th>
                <th colspan="2">主管機關</th>
                <th colspan="2">行政院</th>
                <th rowspan="2">說明</th>
            </tr>
            <tr>
                <th>請增減(-)員額</th><th>核列員額</th>
                <th>請增減(-)員額</th><th>擬核列員額</th>
            </tr>
        </thead>`;
    } else if (tableId === 'personnel_cost') {
        theadHTML = `
        <thead>
            <tr>
                <th rowspan="2">113年度<br/>決算</th>
                <th rowspan="2">114年度<br/>決算</th>
                <th rowspan="2">115年度<br/>預算</th>
                <th rowspan="2">項目</th>
                <th rowspan="2">原編金額</th>
                <th colspan="2">主管機關</th>
                <th colspan="2">行政院初審</th>
                <th rowspan="2">編列說明</th>
            </tr>
            <tr>
                <th>增減(-)數</th><th>核列數</th>
                <th>擬增減(-)數</th><th>擬核列數</th>
            </tr>
        </thead>`;
    } else { // control
        theadHTML = `
        <thead>
            <tr>
                <th rowspan="2">113年度<br/>決算數</th>
                <th rowspan="2">114年度<br/>決算數</th>
                <th rowspan="2">115年度<br/>預算數</th>
                <th rowspan="2">項目</th>
                <th rowspan="2">原編數</th>
                <th colspan="2">主管機關</th>
                <th colspan="2">行政院初審</th>
                <th rowspan="2">編列說明</th>
            </tr>
            <tr>
                <th>增減(-)數</th><th>核列數</th>
                <th>擬增減(-)數</th><th>擬核列數</th>
            </tr>
        </thead>`;
    }

    // body rows
    const tbodyRows = rows.map(r => {
        const tds = cfg.cols
            .filter(c => c !== 'level')
            .map(c => {
                const def = COL[c];
                let v = r[c];
                if (def.type === 'number') return `<td class="num"${numColor(v)}>${fmtNum(v) || '-'}</td>`;
                if (c === 'name') {
                    const lv = cfg.hasLevel ? (parseInt(r.level)||0) : 0;
                    const indent = lv > 0 ? '&nbsp;'.repeat((lv-1)*4) : '';
                    const weight = lv === 1 || lv === 2 ? 'font-weight:bold;' : '';
                    return `<td class="name" style="${weight}">${indent}${escapeHTML(v || '')}</td>`;
                }
                return `<td class="desc">${escapeHTML(v || '').replace(/\n/g,'<br/>')}</td>`;
            }).join('');
        return `<tr>${tds}</tr>`;
    }).join('');

    // review block
    const reviewBlock = reviews ? `
        <table class="review">
            <tr><th colspan="2">審 核 意 見</th></tr>
            <tr><td class="lbl">主管機關</td><td>${escapeHTML(reviews.org || '').replace(/\n/g,'<br/>')}</td></tr>
            <tr><td class="lbl">先期審查機關</td><td>${escapeHTML(reviews.gov || '').replace(/\n/g,'<br/>')}</td></tr>
        </table>` : '';

    // single-table page
    return `
    <div class="page">
        <div class="word-header">
            ${escapeHTML(meta.fund || '特別收入基金')} ${escapeHTML(meta.year || '')}年度 ${escapeHTML(cfg.title)}
            <span class="unit">單位：${cfg.unit}</span>
        </div>
        <table class="data-table">${theadHTML}<tbody>${tbodyRows || '<tr><td colspan="20" style="text-align:center;color:#888">（無資料）</td></tr>'}</tbody></table>
        ${reviewBlock}
        <div class="word-footer">${cfg.sheetId}</div>
    </div>`;
}

// 基金級匯出：把該基金的 _src 與 _use 合併成一個完整 1-2 表
function buildPlanFundDocHTML(fundId) {
    const fund = PLAN_FUNDS.find(f => f.id === fundId);
    const data = collectData();
    const meta = data.meta;
    const srcRows = data.tables[`${fundId}_src`] || [];
    const useRows = data.tables[`${fundId}_use`] || [];
    const reviews = data.reviews?.plan;

    const theadHTML = `<thead>
        <tr>
            <th rowspan="2">114年度<br/>決算數</th>
            <th rowspan="2">115年度<br/>預算數</th>
            <th rowspan="2">科目或計畫名稱</th>
            <th rowspan="2">政策目標、計畫實施內容(或工作項目)<br/>及預期效益摘要</th>
            <th rowspan="2">原編數</th>
            <th colspan="2">主管機關</th>
            <th colspan="2">行政院初審</th>
        </tr>
        <tr>
            <th>增減(-)數</th><th>核列數</th>
            <th>擬增減(-)數</th><th>擬核列數</th>
        </tr>
    </thead>`;

    const renderRow = (r) => {
        const lv = parseInt(r.level) || 0;
        const indent = lv > 0 ? '&nbsp;'.repeat((lv-1)*4) : '';
        const weight = lv === 1 ? 'font-weight:bold;' : '';
        const numCell = (v) => `<td class="num"${numColor(v)}>${fmtNum(v) || '-'}</td>`;
        const nameCell = `<td class="name" style="${weight}">${indent}${escapeHTML(r.name || '')}</td>`;
        const descCell = `<td class="desc">${escapeHTML(r.desc || '').replace(/\n/g,'<br/>')}</td>`;
        return `<tr>${numCell(r.dec113)}${numCell(r.bud114)}${nameCell}${descCell}${numCell(r.orig)}${numCell(r.dep_diff)}${numCell(r.dep_app)}${numCell(r.gov_diff)}${numCell(r.gov_app)}</tr>`;
    };

    // 甲 / 乙 分隔列
    const sectionRow = (label) => `<tr><td colspan="9" style="background:#e8e8e8;font-weight:bold;text-align:left;padding:5pt;">${escapeHTML(label)}</td></tr>`;

    const tbodyContent = [
        sectionRow('甲、基金來源：'),
        ...srcRows.map(renderRow),
        sectionRow('乙、基金用途：'),
        ...useRows.map(renderRow)
    ].join('');

    const reviewBlock = reviews ? `
        <table class="review">
            <tr><th colspan="2">審 核 意 見</th></tr>
            <tr><td class="lbl">主管機關</td><td>${escapeHTML(reviews.org || '').replace(/\n/g,'<br/>')}</td></tr>
            <tr><td class="lbl">先期審查機關</td><td>${escapeHTML(reviews.gov || '').replace(/\n/g,'<br/>')}</td></tr>
        </table>` : '';

    return `
    <div class="page">
        <div class="word-header">
            ${escapeHTML(meta.fund || '')} ${escapeHTML(meta.year || '')}年度 ${escapeHTML(fund.name)} 主要業務計畫預算表
            <span class="unit">單位：新臺幣千元</span>
        </div>
        <table class="data-table">${theadHTML}<tbody>${tbodyContent}</tbody></table>
        ${reviewBlock}
        <div class="word-footer">1-2</div>
    </div>`;
}

function exportDoc(scope) {
    const data = collectData();
    const fund = data.meta.fund || '特別收入基金';
    const year = data.meta.year || '';
    // 「基金級」ID：plan_agri 等，由 buildTableDocHTML 自動合併該基金的 _src + _use
    const planFundIds = PLAN_FUNDS.map(f => f.id);

    let tables;
    if (scope === 'all') {
        tables = [...planFundIds, 'headcount', 'personnel_cost', 'control'];
    } else if (scope === 'plan_all') {
        tables = planFundIds;
    } else {
        const main = getActiveTab();
        if (main === 'plan') {
            const active = document.querySelector('.sf-subtab-btn.active')?.dataset.subtab;
            tables = [active || 'plan_agri']; // 該基金的 src+use 合併
        } else if (main === 'personnel') {
            tables = ['headcount', 'personnel_cost'];
        } else {
            tables = [main];
        }
    }

    const uniqueTables = [...new Set(tables)];
    const body = uniqueTables.map(t => buildTableDocHTML(t)).join('');
    let fname;
    if (scope === 'all') {
        fname = `${fund}_${year}_細項預算_全部.doc`;
    } else if (scope === 'plan_all') {
        fname = `${fund}_${year}_業務計畫_六基金合併.doc`;
    } else {
        fname = `${fund}_${year}_${tableConfigs[uniqueTables[0]].title}.doc`;
    }

    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<title>${escapeHTML(fund)} ${escapeHTML(year)}年度 細項預算</title>
<!--[if gte mso 9]>
<xml>
<w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
</w:WordDocument>
</xml>
<![endif]-->
<style>
@page WordSection1 { size: 29.7cm 21cm; mso-page-orientation: landscape; margin: 1.5cm 1.2cm; }
div.WordSection1 { page: WordSection1; }
body { font-family: "標楷體", "DFKai-SB", "Noto Sans TC", sans-serif; font-size: 11pt; color: #000; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }
.word-header { text-align: center; font-weight: bold; font-size: 14pt; margin-bottom: 4pt; position: relative; }
.word-header .unit { position: absolute; right: 0; top: 4pt; font-weight: normal; font-size: 10pt; }
.word-footer { text-align: center; font-size: 9pt; color: #555; margin-top: 8pt; }
table.data-table { border-collapse: collapse; width: 100%; margin-top: 6pt; }
table.data-table th, table.data-table td { border: 1pt solid #000; padding: 3pt 4pt; vertical-align: middle; font-size: 10pt; }
table.data-table th { background: #e8e8e8; text-align: center; font-weight: bold; }
table.data-table td.num { text-align: right; mso-number-format: '\\#\\,\\#\\#0'; }
table.data-table td.name { text-align: left; }
table.data-table td.desc { text-align: left; font-size: 9pt; line-height: 1.35; }
table.review { border-collapse: collapse; width: 100%; margin-top: 10pt; }
table.review th, table.review td { border: 1pt solid #000; padding: 5pt 6pt; font-size: 10pt; }
table.review th { background: #d9d9d9; text-align: center; }
table.review td.lbl { width: 110pt; text-align: center; font-weight: bold; background: #f4f4f4; }
</style>
</head>
<body>
<div class="WordSection1">
${body}
</div>
</body>
</html>`;

    saveAs(new Blob(['﻿', html], { type: 'application/msword' }), fname);
}

// ========== 7.5 彙整模式 ==========
let mergeFunds = []; // [{ fileName, data }]
const MERGE_SUM_FIELDS = ['dec112','dec113','bud114','apr114','orig','dep_diff','dep_app','gov_diff','gov_app'];

function renderMergeList() {
    const list = document.getElementById('sf-merge-list');
    const cnt  = document.getElementById('sf-merge-count');
    if (cnt) cnt.textContent = mergeFunds.length;
    if (!list) return;
    if (!mergeFunds.length) {
        list.innerHTML = '<p class="text-slate-400 text-sm">尚未載入任何基金 JSON</p>';
        return;
    }
    list.innerHTML = mergeFunds.map((f, i) => {
        const meta = f.data.meta || {};
        const t    = f.data.tables || {};
        const cnt  = (tid) => (t[tid] || []).length;
        return `<div class="border border-slate-200 rounded p-3 flex justify-between items-center bg-white">
            <div class="flex-1">
                <div class="font-bold text-slate-800">${escapeHTML(meta.fund || '(未命名基金)')}</div>
                <div class="text-xs text-slate-500 mt-1">
                    📄 ${escapeHTML(f.fileName)} ・ ${escapeHTML(meta.year || '?')}年度 ・
                    計畫 <b>${cnt('plan')}</b> 列 / 員額 <b>${cnt('headcount')}</b> / 用人費 <b>${cnt('personnel_cost')}</b> / 管制 <b>${cnt('control')}</b>
                </div>
            </div>
            <button class="sf-merge-remove text-red-500 text-sm font-bold px-3 py-1 hover:bg-red-50 rounded" data-idx="${i}">移除</button>
        </div>`;
    }).join('');
}

function loadMergeFiles(files) {
    let queued = 0, done = 0, errors = [];
    Array.from(files).forEach(f => {
        if (!f.name.endsWith('.json')) return;
        queued++;
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const data = JSON.parse(e.target.result);
                if (!data.tables) throw new Error('JSON 結構錯誤：缺 tables 欄位（非本系統匯出的格式？）');
                mergeFunds.push({ fileName: f.name, data });
            } catch (err) {
                errors.push(`${f.name}: ${err.message}`);
            }
            done++;
            if (done === queued) {
                renderMergeList();
                document.getElementById('sf-merge-preview')?.classList.add('hidden');
                if (errors.length) alert('部分檔案載入失敗：\n' + errors.join('\n'));
            }
        };
        reader.readAsText(f);
    });
    if (!queued) alert('沒有可用的 JSON 檔案');
}

function computeMerged() {
    const fundName = (document.getElementById('sf-merge-fund')?.value || '').trim();
    const yearVal  = (document.getElementById('sf-merge-year')?.value || '').trim();
    const out = {
        meta: {
            fund: fundName || '彙整基金',
            year: yearVal  || (document.getElementById('sf-year')?.value || ''),
            org:  document.getElementById('sf-org')?.value  || '',
            user: document.getElementById('sf-user')?.value || ''
        },
        tables: (() => {
            const t = { headcount: [], personnel_cost: [], control: [] };
            PLAN_FUNDS.forEach(f => PLAN_SECTIONS.forEach(s => t[`${f.id}_${s.suffix}`] = []));
            return t;
        })(),
        reviews: {
            plan:      { org: '', gov: '' },
            personnel: { org: '', gov: '' },
            control:   { org: '', gov: '' }
        },
        mergedFrom: mergeFunds.map(f => f.data.meta?.fund || f.fileName),
        savedAt: new Date().toISOString()
    };

    mergeFunds.forEach(f => {
        // 自動正規化舊版資料（單一 plan → 6 個 plan_*）
        const normData = normalizeData(f.data);
        const t = normData.tables || {};

        // === 業務計畫 12 表（6 基金 × 甲/乙）：向下追加 ===
        PLAN_FUNDS.forEach(pf => {
            PLAN_SECTIONS.forEach(sec => {
                const tid = `${pf.id}_${sec.suffix}`;
                (t[tid] || []).forEach(row => out.tables[tid].push({ ...row }));
            });
        });

        // === 其他三表：以 name 為 key 加總 ===
        ['headcount','personnel_cost','control'].forEach(tid => {
            (t[tid] || []).forEach(row => {
                const name = (row.name || '').trim();
                if (!name) return;
                let match = out.tables[tid].find(r => (r.name || '').trim() === name);
                if (!match) {
                    match = { name: row.name };
                    out.tables[tid].push(match);
                }
                MERGE_SUM_FIELDS.forEach(k => {
                    const raw = row[k];
                    if (raw === undefined || raw === '' || raw === null) return;
                    const v = parseFloat(raw);
                    if (isNaN(v)) return;
                    const prev = parseFloat(match[k]);
                    match[k] = (isNaN(prev) ? 0 : prev) + v;
                });
                // desc 串接（含來源基金前綴）
                if (row.desc && row.desc.trim()) {
                    const tag = `【${f.data.meta?.fund || f.fileName}】`;
                    match.desc = (match.desc ? match.desc + '\n' : '') + tag + row.desc.trim();
                }
            });
        });

        // 合併 reviews（串接，附來源前綴）
        ['plan','personnel','control'].forEach(rk => {
            const src = f.data.reviews?.[rk];
            if (!src) return;
            const tag = `【${f.data.meta?.fund || f.fileName}】`;
            ['org','gov'].forEach(role => {
                if (src[role] && src[role].trim()) {
                    out.reviews[rk][role] = (out.reviews[rk][role] ? out.reviews[rk][role] + '\n' : '') + tag + src[role].trim();
                }
            });
        });
    });

    // 將加總後的字串格式正規化（轉回字串）
    ['headcount','personnel_cost','control'].forEach(tid => {
        out.tables[tid].forEach(row => {
            MERGE_SUM_FIELDS.forEach(k => {
                if (typeof row[k] === 'number') row[k] = String(row[k]);
            });
        });
    });

    return out;
}

function renderMergePreview() {
    const data = computeMerged();
    const wrap = document.getElementById('sf-merge-preview');
    const cards= document.getElementById('sf-merge-preview-cards');
    if (!wrap || !cards) return;
    const c = (tid) => (data.tables[tid] || []).length;
    const planTotal = PLAN_FUNDS.reduce((a, f) => a + c(`${f.id}_src`) + c(`${f.id}_use`), 0);
    const planBreakdown = PLAN_FUNDS.map(f => `${f.short} ${c(`${f.id}_src`) + c(`${f.id}_use`)}`).join(' / ');
    cards.innerHTML = `
        <div class="sf-merge-card">
            <div class="label">業務計畫（六基金）</div>
            <div class="value">${planTotal}</div>
            <div class="sub">列（追加）<br/><span style="font-size:0.65rem">${planBreakdown}</span></div>
        </div>
        <div class="sf-merge-card">
            <div class="label">員額</div>
            <div class="value">${c('headcount')}</div>
            <div class="sub">項目（加總）</div>
        </div>
        <div class="sf-merge-card">
            <div class="label">用人費用</div>
            <div class="value">${c('personnel_cost')}</div>
            <div class="sub">項目（加總）</div>
        </div>
        <div class="sf-merge-card">
            <div class="label">管制項目</div>
            <div class="value">${c('control')}</div>
            <div class="sub">項目（加總）</div>
        </div>
    `;
    wrap.classList.remove('hidden');
}

function downloadMergedJSON() {
    if (!mergeFunds.length) { alert('請先載入至少一個基金 JSON'); return; }
    const data = computeMerged();
    const fname = `${data.meta.fund}_${data.meta.year}_彙整.json`;
    saveAs(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), fname);
}

function applyMergedToTables() {
    if (!mergeFunds.length) { alert('請先載入至少一個基金 JSON'); return; }
    if (!confirm('將以彙整結果覆蓋上方四張表的當前資料，確定要繼續？\n（操作不可復原，建議先「下載彙整 JSON」備份）')) return;
    const data = computeMerged();
    applyData(data);
    scheduleAutosave();
    flashAutosave('✓ 已套用彙整結果至表格');
    switchTab('plan');
}

// ========== 7.6 業務計畫子 tab（6 基金，每基金 甲/乙 兩表）==========
function planThead() {
    return `<thead>
        <tr>
            <th rowspan="2" class="sf-col-act">操作</th>
            <th rowspan="2" class="sf-col-lv">層級</th>
            <th rowspan="2">114年度<br>決算數</th>
            <th rowspan="2">115年度<br>預算數</th>
            <th rowspan="2" class="sf-col-name">科目或計畫名稱</th>
            <th rowspan="2" class="sf-col-desc">政策目標、計畫實施內容(或工作項目)<br>及預期效益摘要</th>
            <th rowspan="2">原編數</th>
            <th colspan="2" class="sf-col-group">主管機關</th>
            <th colspan="2" class="sf-col-group">行政院初審</th>
        </tr>
        <tr>
            <th>增減(-)數</th>
            <th class="sf-col-auto">核列數</th>
            <th>擬增減(-)數</th>
            <th class="sf-col-auto">擬核列數</th>
        </tr>
    </thead>`;
}
function renderPlanSubpanels() {
    const container = document.getElementById('sf-plan-subpanels');
    if (!container) return;
    container.innerHTML = PLAN_FUNDS.map((f, idx) => {
        const sectionHTML = PLAN_SECTIONS.map(sec => {
            const tid = `${f.id}_${sec.suffix}`;
            return `
            <div class="section-card sf-plan-section" data-section="${sec.suffix}">
                <div class="flex justify-between items-center mb-3">
                    <h4 class="text-lg font-bold text-purple-700">${escapeHTML(sec.label)}</h4>
                    <div class="flex gap-2">
                        <button class="sf-add-row bg-purple-600 text-white px-3 py-1 rounded text-sm" data-table="${tid}">＋ 新增列</button>
                        <button class="sf-load-sample bg-slate-200 text-slate-700 px-3 py-1 rounded text-sm" data-table="${tid}">重置為 Word 預設項目</button>
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table class="sf-table">
                        ${planThead()}
                        <tbody id="sf-tbody-${tid}"></tbody>
                    </table>
                </div>
            </div>`;
        }).join('');
        return `<div class="sf-plan-subpanel" data-fund="${f.id}" ${idx > 0 ? 'style="display:none"' : ''}>
            <div class="bg-purple-50 border border-purple-200 rounded-lg p-3 mb-3">
                <h3 class="text-lg font-bold text-purple-700">${escapeHTML(f.name)}</h3>
                <p class="text-xs text-slate-600 mt-1">116年度 主要業務計畫預算表 · 表號 1-2 · 單位：新臺幣千元</p>
            </div>
            ${sectionHTML}
        </div>`;
    }).join('');
}

function switchPlanSubtab(fundId) {
    document.querySelectorAll('.sf-subtab-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.subtab === fundId);
    });
    document.querySelectorAll('.sf-plan-subpanel').forEach(p => {
        p.style.display = (p.dataset.fund === fundId) ? '' : 'none';
    });
}

// ========== 8. 分頁切換 ==========
function getActiveTab() {
    return document.querySelector('.sf-tab-btn.active')?.dataset.tab || 'plan';
}
function switchTab(tab) {
    document.querySelectorAll('.sf-tab-btn').forEach(b => {
        const on = b.dataset.tab === tab;
        b.classList.toggle('active', on);
        b.classList.toggle('bg-white', on);
        b.classList.toggle('shadow-sm', on);
        b.classList.toggle('text-slate-500', !on);
    });
    document.querySelectorAll('.sf-tab-panel').forEach(p => p.classList.add('hidden'));
    document.getElementById('sf-tab-' + tab)?.classList.remove('hidden');
}

// ========== 9. 事件綁定 ==========
function bindEvents() {
    // tab switch（主分頁）
    document.querySelectorAll('.sf-tab-btn').forEach(b => {
        b.onclick = () => switchTab(b.dataset.tab);
    });

    // 業務計畫子 tab（六基金）
    document.querySelectorAll('.sf-subtab-btn').forEach(b => {
        b.onclick = () => switchPlanSubtab(b.dataset.subtab);
    });

    // 工具列
    document.getElementById('sf-btn-export-json').onclick = exportJSON;
    document.getElementById('sf-btn-import').onclick = () => document.getElementById('sf-import-file').click();
    document.getElementById('sf-import-file').onchange = e => { if (e.target.files[0]) handleImport(e.target.files[0]); e.target.value = ''; };
    document.getElementById('sf-btn-export-doc').onclick     = () => exportDoc('current');
    document.getElementById('sf-btn-export-doc-all').onclick = () => exportDoc('all');
    document.getElementById('sf-btn-clear').onclick = () => {
        if (!confirm('確定要清空所有欄位與資料列？此動作無法復原。')) return;
        localStorage.removeItem(STORE_KEY);
        Object.keys(tableConfigs).forEach(tid => {
            document.getElementById('sf-tbody-' + tid).innerHTML = '';
            addRow(tid);
        });
        ['sf-fund','sf-year','sf-org','sf-user'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = (id === 'sf-year') ? '115' : '';
        });
        document.querySelectorAll('textarea[id$="-review-org"], textarea[id$="-review-gov"]').forEach(t => t.value = '');
    };

    // 新增列
    document.querySelectorAll('.sf-add-row').forEach(btn => {
        btn.onclick = () => { addRow(btn.dataset.table); scheduleAutosave(); };
    });

    // 套用範例骨架
    document.querySelectorAll('.sf-load-sample').forEach(btn => {
        btn.onclick = () => {
            const tid = btn.dataset.table;
            if (!confirm(`將清除「${tableConfigs[tid].title}」目前所有列並重新載入 Word 預設項目？`)) return;
            document.getElementById('sf-tbody-' + tid).innerHTML = '';
            (SAMPLES[tid] || []).forEach(r => addRow(tid, r));
            if (tid.startsWith('plan_')) recalcPlanTable(tid);
            scheduleAutosave();
        };
    });

    // 行內事件委派：插入鈕、刪除、層級變更、輸入計算、拖曳排序
    Object.keys(tableConfigs).forEach(tid => {
        const tbody = document.getElementById('sf-tbody-' + tid);

        const isPlan = tid.startsWith('plan_');
        const triggerRecalc = (tr) => {
            if (isPlan) recalcPlanTable(tid);
            else recalcRow(tr);
        };

        tbody.addEventListener('click', e => {
            // 列上方插入鈕（在這列前插入一列新空白列）
            const insertZone = e.target.closest('.sf-insert-zone');
            if (insertZone) {
                const tr = insertZone.closest('tr');
                addRow(tid, {}, tr);
                if (isPlan) recalcPlanTable(tid);
                scheduleAutosave();
                return;
            }
            // 刪除
            const btn = e.target.closest('.sf-row-btn');
            if (!btn) return;
            const tr = btn.closest('tr');
            if (btn.classList.contains('del')) {
                if (tbody.children.length <= 1) {
                    tr.querySelectorAll('input,textarea,select').forEach(el => el.value = '');
                    tr.setAttribute('data-level', 0);
                    recalcRow(tr);
                } else tr.remove();
                if (isPlan) recalcPlanTable(tid);
                scheduleAutosave();
            }
        });

        tbody.addEventListener('change', e => {
            if (e.target.dataset.key === 'level') {
                e.target.closest('tr').setAttribute('data-level', e.target.value);
                if (isPlan) recalcPlanTable(tid);
            }
        });

        tbody.addEventListener('input', e => {
            const tr = e.target.closest('tr');
            if (!tr) return;
            const key = e.target.dataset.key;
            if (['orig','dep_diff','gov_diff','dec112','dec113','bud114','apr114'].includes(key)) {
                triggerRecalc(tr);
            }
            scheduleAutosave();
        });

        // 數字欄 blur 重新套用千分號
        tbody.addEventListener('focusout', e => {
            const el = e.target;
            if (el.matches && el.matches('input[data-numeric]:not([readonly])')) {
                const n = parseNum(el.value);
                if (!isNaN(n) && el.value.trim() !== '') el.value = formatNum(n);
            }
        });

        // ===== 拖曳排序 =====
        let draggedTr = null;
        tbody.addEventListener('dragstart', e => {
            // 只允許從 .sf-drag-handle 或 .sf-act td 起始（避免 input 文字選取被誤觸發）
            const t = e.target;
            const tr = t.closest('tr');
            if (!tr) return;
            const fromHandle = t.closest('.sf-drag-handle') || (t.closest('td')?.classList.contains('sf-act') && !t.closest('input,textarea,select,button'));
            if (!fromHandle) { e.preventDefault(); return; }
            draggedTr = tr;
            tr.classList.add('sf-dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', tid);
        });
        tbody.addEventListener('dragover', e => {
            if (!draggedTr) return;
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            const tr = e.target.closest('tr');
            if (!tr || tr === draggedTr) return;
            tbody.querySelectorAll('.sf-drop-above, .sf-drop-below').forEach(el =>
                el.classList.remove('sf-drop-above', 'sf-drop-below'));
            const rect = tr.getBoundingClientRect();
            const above = e.clientY < rect.top + rect.height / 2;
            tr.classList.add(above ? 'sf-drop-above' : 'sf-drop-below');
        });
        tbody.addEventListener('drop', e => {
            if (!draggedTr) return;
            e.preventDefault();
            const tr = e.target.closest('tr');
            if (!tr || tr === draggedTr) return;
            const rect = tr.getBoundingClientRect();
            const above = e.clientY < rect.top + rect.height / 2;
            if (above) tbody.insertBefore(draggedTr, tr);
            else tbody.insertBefore(draggedTr, tr.nextSibling);
            if (isPlan) recalcPlanTable(tid);
            scheduleAutosave();
        });
        tbody.addEventListener('dragend', () => {
            if (draggedTr) draggedTr.classList.remove('sf-dragging');
            tbody.querySelectorAll('.sf-drop-above, .sf-drop-below').forEach(el =>
                el.classList.remove('sf-drop-above', 'sf-drop-below'));
            draggedTr = null;
        });
    });

    // metadata 輸入也自動儲存
    ['sf-fund','sf-year','sf-org','sf-user'].forEach(id => {
        document.getElementById(id).addEventListener('input', scheduleAutosave);
    });
    document.querySelectorAll('textarea[id$="-review-org"], textarea[id$="-review-gov"]').forEach(t => {
        t.addEventListener('input', scheduleAutosave);
    });

    // ===== 彙整模式事件 =====
    const mdz = document.getElementById('sf-merge-dropzone');
    if (mdz) {
        mdz.addEventListener('click', () => document.getElementById('sf-merge-file').click());
        mdz.addEventListener('dragover', e => { e.preventDefault(); mdz.classList.add('dragover'); });
        mdz.addEventListener('dragleave', () => mdz.classList.remove('dragover'));
        mdz.addEventListener('drop', e => {
            e.preventDefault(); mdz.classList.remove('dragover');
            loadMergeFiles(e.dataTransfer.files);
        });
    }
    const mfile = document.getElementById('sf-merge-file');
    if (mfile) mfile.addEventListener('change', e => {
        if (e.target.files.length) loadMergeFiles(e.target.files);
        e.target.value = '';
    });
    document.getElementById('sf-merge-list')?.addEventListener('click', e => {
        const btn = e.target.closest('.sf-merge-remove');
        if (!btn) return;
        mergeFunds.splice(parseInt(btn.dataset.idx), 1);
        renderMergeList();
        document.getElementById('sf-merge-preview')?.classList.add('hidden');
    });
    document.getElementById('sf-merge-preview-btn')?.addEventListener('click', () => {
        if (!mergeFunds.length) { alert('請先載入至少一個基金 JSON'); return; }
        renderMergePreview();
    });
    document.getElementById('sf-merge-download')?.addEventListener('click', downloadMergedJSON);
    document.getElementById('sf-merge-apply')?.addEventListener('click', applyMergedToTables);
    document.getElementById('sf-merge-clear')?.addEventListener('click', () => {
        if (mergeFunds.length && !confirm('清除所有已載入的基金 JSON？')) return;
        mergeFunds = [];
        renderMergeList();
        document.getElementById('sf-merge-preview')?.classList.add('hidden');
    });
    renderMergeList(); // 初始

    // 鍵盤導航：Enter → 下一列；表格貼上
    document.addEventListener('keydown', e => {
        if (e.key !== 'Enter') return;
        const el = document.activeElement;
        if (!el || (el.tagName !== 'INPUT' && el.tagName !== 'SELECT')) return;
        if (el.readOnly) return;
        if (!el.closest('.sf-table')) return;
        e.preventDefault();
        const tr = el.closest('tr');
        const tbody = tr.parentNode;
        const tid = tbody.id.replace('sf-tbody-', '');
        const tds = Array.from(tr.children);
        const colIdx = tds.indexOf(el.closest('td'));
        let nextTr = tr.nextElementSibling;
        if (!nextTr) { addRow(tid); nextTr = tbody.lastElementChild; }
        let target = nextTr.children[colIdx]?.querySelector('input,select,textarea');
        if (target?.readOnly) {
            target = nextTr.querySelector('input:not([readonly]),select,textarea');
        }
        target?.focus();
    });

    document.addEventListener('paste', e => {
        const el = document.activeElement;
        if (!el || el.tagName !== 'INPUT' || el.readOnly) return;
        if (!el.closest('.sf-table')) return;

        const cd = e.clipboardData || window.clipboardData;
        const html = cd.getData('text/html');
        const text = cd.getData('text/plain');

        // 1) 優先解析 text/html 的 <table>（Word 表格複製保留結構，多行儲存格不會錯位）
        let rowsData = parseClipboardTable(html);
        // 2) 退回純文字（Excel 表格複製或單欄列表）
        if (!rowsData) {
            if (!text || (!text.includes('\t') && !text.includes('\n'))) return;
            rowsData = text.split(/\r\n|\r|\n/).filter(r => r.length).map(r => r.split('\t'));
        }
        if (!rowsData.length) return;

        e.preventDefault();
        const startTr = el.closest('tr');
        const tbody = startTr.parentNode;
        const tid = tbody.id.replace('sf-tbody-', '');
        const startCol = Array.from(startTr.children).indexOf(el.closest('td'));
        const startRowIdx = Array.from(tbody.children).indexOf(startTr);

        rowsData.forEach((cells, i) => {
            let targetTr = tbody.children[startRowIdx + i];
            if (!targetTr) { addRow(tid); targetTr = tbody.lastElementChild; }
            cells.forEach((txt, j) => {
                const td = targetTr.children[startCol + j];
                if (!td) return;
                const inp = td.querySelector('input:not([readonly]),textarea');
                if (!inp) return;
                if (inp.tagName === 'TEXTAREA') {
                    // 摘要欄保留換行
                    inp.value = txt.replace(/\r\n?/g, '\n').trim();
                } else if (inp.type === 'number') {
                    // 數字欄：取第一行、去千分位逗號、去除非數字尾巴
                    const firstLine = txt.split(/\r?\n/)[0].trim().replace(/,/g, '');
                    const m = firstLine.match(/-?\d+(\.\d+)?/);
                    inp.value = m ? m[0] : '';
                } else {
                    // 文字欄（項目名稱）：取第一行、去逗號
                    inp.value = txt.split(/\r?\n/)[0].trim().replace(/,/g, '');
                }
            });
            recalcRow(targetTr);
        });
        scheduleAutosave();
    });
}

// ===== 解析剪貼簿中的 HTML 表格 =====
// 回傳 rowsData：cell 內的 <br>、<p>、<div> 邊界轉成 \n，其餘標籤剝除
function parseClipboardTable(html) {
    if (!html || !/<table[\s>]/i.test(html)) return null;
    try {
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        const table = tmp.querySelector('table');
        if (!table) return null;
        const rows = Array.from(table.querySelectorAll('tr')).map(tr =>
            Array.from(tr.children).map(cell => {
                // 把段落、區塊、換行 → \n；其餘標籤靠 textContent 自動剝除
                const wrap = document.createElement('div');
                wrap.innerHTML = cell.innerHTML
                    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
                    .replace(/<\/\s*(p|div|h[1-6]|li)\s*>/gi, '\n');
                return wrap.textContent
                    .replace(/ /g, ' ')                  // &nbsp; → 一般空格
                    .replace(/[ \t]+/g, ' ')                // 多空白/Tab → 1 空白
                    .replace(/[ \t]*\n[ \t]*/g, '\n')     // 換行前後空白清掉
                    .replace(/\n{2,}/g, '\n')              // 多重換行縮成 1 個
                    .trim();
            })
        ).filter(r => r.length); // 跳過完全空列
        return rows.length ? rows : null;
    } catch (e) {
        console.warn('parseClipboardTable failed', e);
        return null;
    }
}

// ========== 10. 啟動 ==========
function loadAllSamples() {
    Object.keys(tableConfigs).forEach(tid => {
        const tbody = document.getElementById('sf-tbody-' + tid);
        if (!tbody) return;
        tbody.innerHTML = '';
        (SAMPLES[tid] || []).forEach(r => addRow(tid, r));
        if (tid.startsWith('plan_')) recalcPlanTable(tid);
    });
}

document.addEventListener('DOMContentLoaded', () => {
    // 先渲染 6 個業務計畫子面板（建立 tbody 與按鈕），再綁事件
    renderPlanSubpanels();
    bindEvents();
    if (loadAutosave()) {
        flashAutosave('✓ 已還原上次資料');
    } else {
        loadAllSamples();
        flashAutosave('✓ 已載入 Word 預設項目');
    }
    switchTab('plan');
    switchPlanSubtab('plan_agri');
});
