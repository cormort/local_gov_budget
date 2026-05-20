'use strict';

// ========== 1. 表格欄位定義 ==========
// 每個 column: { key, label, type, td_class?, auto? }
//   type: 'number' | 'text' | 'textarea' | 'level'
//   auto: 'dep_app' (= orig + dep_diff) | 'gov_app' (= dep_app + gov_diff)

const COL = {
    // 內部 key 維持 dec112/dec113/bud114/apr114 以維持 autosave 相容；標籤對應新年度
    dec112: { key: 'dec112', label: '114年度決算', type: 'number', td_class: 'sf-num' },
    dec113: { key: 'dec113', label: '115年度決算', type: 'number', td_class: 'sf-num' },
    bud114: { key: 'bud114', label: '116年度預算', type: 'number', td_class: 'sf-num' },
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

// 業務計畫表共用結構（6 個基金子 tab 共用同一組欄位）
const PLAN_BASE = {
    sheetId: '1-2',
    unit: '新臺幣千元',
    cols: ['level', 'dec113', 'bud114', 'name', 'desc', 'orig', 'dep_diff', 'dep_app', 'gov_diff', 'gov_app'],
    hasLevel: true,
    reviewKey: 'plan',
    isPlan: true
};
// 6 個基金代號 + 顯示名稱
const PLAN_FUNDS = [
    { id: 'plan_agri',     name: '農業發展基金',         short: '農發',    rank: 1 },
    { id: 'plan_forest',   name: '林務發展及造林基金',     short: '林務',    rank: 2 },
    { id: 'plan_disaster', name: '農業天然災害救助基金',   short: '天災',    rank: 3 },
    { id: 'plan_fish',     name: '漁業發展基金',         short: '漁業',    rank: 4 },
    { id: 'plan_loss',     name: '農產品受進口損害救助基金', short: '農損',  rank: 5 },
    { id: 'plan_renewal',  name: '農村再生基金',         short: '農再',    rank: 6 }
];

const tableConfigs = {};
PLAN_FUNDS.forEach(f => {
    tableConfigs[f.id] = {
        ...PLAN_BASE,
        title: `${f.name} 116年度 主要業務計畫預算表`,
        fundShort: f.short,
        fundName: f.name,
        rank: f.rank
    };
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

const LEVEL_LABELS = { 0: '─', 1: '壹/貳', 2: '甲/乙', 3: '一/二', 4: '(一)/(二)', 5: '項目' };

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
        { level: 3, name: '壹、農村再生規劃及人力培育計畫' },
        { level: 4, name: '一、農村規劃及培力' },
        { level: 5, name: '1.農村人力及教育推廣' },
        { level: 4, name: '二、農業人才多元培育' },
        { level: 4, name: '三、農村農產業人力活化計畫' },
        { level: 3, name: '貳、農村再生建設及發展計畫' },
        { level: 4, name: '一、農村再生社區發展及環境改善' },
        { level: 5, name: '1.農村再生跨域發展' },
        { level: 5, name: '2.社區農村再生計畫' },
        { level: 5, name: '3.農村社區土地重劃及發展規劃' },
        { level: 4, name: '二、農村發展及活化' },
        { level: 5, name: '1.提升農村農糧產業競爭力' },
        { level: 5, name: '2.發展健康永續的有機產業' },
        { level: 5, name: '3.農村社區畜牧場環境改善及資源利用' },
        { level: 5, name: '4.建設休閒農業優質環境' },
        { level: 5, name: '5.友善漁業生產環境及漁村產業活動推廣' },
        { level: 5, name: '6.山村綠色經濟永續發展計畫' },
        { level: 5, name: '7.農產加工整合服務體系發展' },
        { level: 5, name: '8.運用資通訊技術(ICT)強化農業發展及推廣計畫' },
        { level: 5, name: '9.優化農業推廣教育訓練場域' },
        { level: 5, name: '10.推動化學農藥減量模式促進農村生產環境永續發展' },
        { level: 5, name: '11.農村水資源韌性公共設施建設' },
        { level: 5, name: '12.推動農會經濟事業發展，振興農村產業' },
        { level: 5, name: '13.改善農村金融服務品質' },
        { level: 5, name: '14.農業旅遊創新發展及市場拓展' },
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
        { name: '(一)購置無形資產' }
    ]
};

// ========== 3. 行渲染 ==========
function cellHTML(col, data, tableId) {
    const def = COL[col];
    const val = data?.[col] ?? '';
    if (def.type === 'level') {
        const opts = [0,1,2,3,4,5].map(n => `<option value="${n}" ${String(val)===String(n)?'selected':''}>${LEVEL_LABELS[n]}</option>`).join('');
        return `<td class="sf-lv"><select class="sf-level-select" data-key="level">${opts}</select></td>`;
    }
    if (def.type === 'textarea') {
        return `<td class="${def.td_class}"><textarea data-key="${col}" placeholder="">${val}</textarea></td>`;
    }
    if (def.type === 'text') {
        return `<td class="${def.td_class}"><input type="text" data-key="${col}" value="${escapeAttr(val)}"></td>`;
    }
    // number
    const isAuto = !!def.auto;
    return `<td class="${def.td_class}"><input type="number" data-key="${col}" value="${val !== '' && val !== null ? val : ''}" ${isAuto ? 'readonly' : ''}></td>`;
}

function escapeAttr(s) {
    return String(s ?? '').replace(/"/g, '&quot;');
}

function escapeHTML(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function addRow(tableId, data = {}) {
    const cfg = tableConfigs[tableId];
    if (!cfg) return;
    const tbody = document.getElementById('sf-tbody-' + tableId);
    const tr = document.createElement('tr');
    const lvl = cfg.hasLevel ? (data.level ?? 0) : 0;
    if (cfg.hasLevel) tr.setAttribute('data-level', lvl);
    const actHTML = `<td class="sf-act">
        <button class="sf-row-btn up" title="上移">▲</button><button class="sf-row-btn down" title="下移">▼</button><button class="sf-row-btn del" title="刪除">✕</button>
    </td>`;
    tr.innerHTML = actHTML + cfg.cols.map(c => cellHTML(c, data, tableId)).join('');
    tbody.appendChild(tr);
    recalcRow(tr);
}

function recalcRow(tr) {
    const get = k => parseFloat(tr.querySelector(`[data-key="${k}"]`)?.value) || 0;
    const set = (k, v) => {
        const el = tr.querySelector(`[data-key="${k}"]`);
        if (!el) return;
        el.value = (v === 0 ? '' : v);
        el.classList.toggle('negative-value', v < 0);
    };
    const depApp = get('orig') + get('dep_diff');
    set('dep_app', depApp);
    const govApp = depApp + get('gov_diff');
    set('gov_app', govApp);
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
                if (el) item[c] = el.value;
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
    });
}

function val(id) { return (document.getElementById(id)?.value || '').trim(); }

// ========== 5. JSON 匯入/匯出 ==========
function exportJSON() {
    const data = collectData();
    const fname = `${data.meta.fund || '特別收入基金'}_${data.meta.year || ''}_細項預算.json`;
    saveAs(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), fname);
}

function handleImport(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const data = normalizeData(JSON.parse(e.target.result));
            applyData(data);
            scheduleAutosave();
            flashAutosave('✓ 匯入成功');
        } catch (err) {
            alert('匯入失敗：' + err.message);
        }
    };
    reader.readAsText(file);
}

// ========== 6. 自動儲存 ==========
const STORE_KEY = 'sf_special_fund_v3'; // v3: 業務計畫拆成 6 個基金子 tab
const LEGACY_KEYS = ['sf_special_fund_v2', 'sf_special_fund_v1'];

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

// 統一資料正規化入口（任何讀進來的資料都先過這層）
function normalizeData(data) {
    if (!data?.tables) return data;
    const hasOldPlan = Array.isArray(data.tables.plan);
    const hasSubPlans = PLAN_FUNDS.some(f => Array.isArray(data.tables[f.id]));
    if (hasOldPlan && !hasSubPlans) return migrateOldPlan(data);
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
function fmtNum(v) {
    if (v === '' || v === null || v === undefined) return '';
    const n = parseFloat(v);
    if (isNaN(n)) return escapeHTML(String(v));
    return n.toLocaleString();
}
function numColor(v) {
    const n = parseFloat(v);
    return (!isNaN(n) && n < 0) ? ' style="color:#c00"' : '';
}

function buildTableDocHTML(tableId) {
    const cfg = tableConfigs[tableId];
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
                <th rowspan="2">114年度<br/>決算</th>
                <th rowspan="2">115年度<br/>決算</th>
                <th rowspan="2">116年度<br/>預算</th>
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
                <th rowspan="2">114年度<br/>決算</th>
                <th rowspan="2">115年度<br/>決算</th>
                <th rowspan="2">116年度<br/>預算</th>
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
                <th rowspan="2">114年度<br/>決算數</th>
                <th rowspan="2">115年度<br/>決算數</th>
                <th rowspan="2">116年度<br/>預算數</th>
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

function exportDoc(scope) {
    const data = collectData();
    const fund = data.meta.fund || '特別收入基金';
    const year = data.meta.year || '';
    const planIds = PLAN_FUNDS.map(f => f.id);

    let tables;
    if (scope === 'all') {
        tables = [...planIds, 'headcount', 'personnel_cost', 'control'];
    } else if (scope === 'plan_all') {
        // 6 個基金的業務計畫合併匯出
        tables = planIds;
    } else {
        const main = getActiveTab();
        if (main === 'plan') {
            const active = document.querySelector('.sf-subtab-btn.active')?.dataset.subtab;
            tables = [active || 'plan_agri'];
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
            PLAN_FUNDS.forEach(f => t[f.id] = []);
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

        // === 業務計畫 6 個基金表：向下追加（保留每基金所有列）===
        PLAN_FUNDS.forEach(pf => {
            (t[pf.id] || []).forEach(row => out.tables[pf.id].push({ ...row }));
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
    const planTotal = PLAN_FUNDS.reduce((a, f) => a + c(f.id), 0);
    const planBreakdown = PLAN_FUNDS.map(f => `${f.short} ${c(f.id)}`).join(' / ');
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

// ========== 7.6 業務計畫子 tab（6 個基金）==========
function renderPlanSubpanels() {
    const container = document.getElementById('sf-plan-subpanels');
    if (!container) return;
    container.innerHTML = PLAN_FUNDS.map((f, idx) => `
        <div class="sf-plan-subpanel section-card" data-fund="${f.id}" ${idx > 0 ? 'style="display:none"' : ''}>
            <div class="flex justify-between items-center mb-3">
                <div>
                    <h3 class="text-xl font-bold text-purple-700">${escapeHTML(f.name)}</h3>
                    <p class="text-xs text-slate-500 mt-1">116年度 主要業務計畫預算表 · 表號 1-2 · 單位：新臺幣千元</p>
                </div>
                <div class="flex gap-2">
                    <button class="sf-add-row bg-purple-600 text-white px-3 py-1 rounded text-sm" data-table="${f.id}">＋ 新增列</button>
                    <button class="sf-load-sample bg-slate-200 text-slate-700 px-3 py-1 rounded text-sm" data-table="${f.id}">重置為 Word 預設項目</button>
                </div>
            </div>
            <div class="overflow-x-auto">
                <table class="sf-table">
                    <thead>
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
                    </thead>
                    <tbody id="sf-tbody-${f.id}"></tbody>
                </table>
            </div>
        </div>
    `).join('');
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
            SAMPLES[tid].forEach(r => addRow(tid, r));
            scheduleAutosave();
        };
    });

    // 行內事件委派：刪除、上下移、層級變更、輸入計算
    Object.keys(tableConfigs).forEach(tid => {
        const tbody = document.getElementById('sf-tbody-' + tid);
        tbody.addEventListener('click', e => {
            const btn = e.target.closest('.sf-row-btn');
            if (!btn) return;
            const tr = btn.closest('tr');
            if (btn.classList.contains('del')) {
                if (tbody.children.length <= 1) {
                    // 至少留一列，但清空
                    tr.querySelectorAll('input,textarea,select').forEach(el => el.value = '');
                    tr.setAttribute('data-level', 0);
                    recalcRow(tr);
                } else tr.remove();
            } else if (btn.classList.contains('up')) {
                if (tr.previousElementSibling) tr.parentNode.insertBefore(tr, tr.previousElementSibling);
            } else if (btn.classList.contains('down')) {
                if (tr.nextElementSibling) tr.parentNode.insertBefore(tr.nextElementSibling, tr);
            }
            scheduleAutosave();
        });

        tbody.addEventListener('change', e => {
            if (e.target.dataset.key === 'level') {
                e.target.closest('tr').setAttribute('data-level', e.target.value);
            }
        });

        tbody.addEventListener('input', e => {
            const tr = e.target.closest('tr');
            if (!tr) return;
            const key = e.target.dataset.key;
            if (['orig','dep_diff','gov_diff'].includes(key)) recalcRow(tr);
            scheduleAutosave();
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
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text || (!text.includes('\t') && !text.includes('\n'))) return;
        e.preventDefault();
        const rows = text.split(/\r\n|\n|\r/).filter(r => r.length);
        const startTr = el.closest('tr');
        const tbody = startTr.parentNode;
        const tid = tbody.id.replace('sf-tbody-', '');
        const startCol = Array.from(startTr.children).indexOf(el.closest('td'));
        const startRowIdx = Array.from(tbody.children).indexOf(startTr);
        rows.forEach((rowText, i) => {
            let targetTr = tbody.children[startRowIdx + i];
            if (!targetTr) { addRow(tid); targetTr = tbody.lastElementChild; }
            const cells = rowText.split('\t');
            cells.forEach((txt, j) => {
                const td = targetTr.children[startCol + j];
                if (!td) return;
                const inp = td.querySelector('input:not([readonly]),textarea');
                if (inp) inp.value = txt.trim().replace(/,/g, '');
            });
            recalcRow(targetTr);
        });
        scheduleAutosave();
    });
}

// ========== 10. 啟動 ==========
function loadAllSamples() {
    Object.keys(tableConfigs).forEach(tid => {
        const tbody = document.getElementById('sf-tbody-' + tid);
        if (!tbody) return;
        tbody.innerHTML = '';
        (SAMPLES[tid] || []).forEach(r => addRow(tid, r));
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
