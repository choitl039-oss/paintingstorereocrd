import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.2/package/xlsx.mjs';
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import { adminPin, supabaseConfig } from './supabase-config.js';

const state = {
  role: 'staff',
  adminUnlocked: false,
  language: localStorage.getItem('store-language') || 'zh-Hans',
  inventory: [],
  sales: [],
  staffSearchTerm: '',
  busyAction: '',
  statusMessage: '',
  errorMessage: '',
  selectedImageFile: null,
  selectedImagePreviewUrl: '',
};

const dom = {
  adminCardGrid: document.getElementById('admin-card-grid'),
  adminLayout: document.getElementById('admin-layout'),
  adminLockCard: document.getElementById('admin-lock-card'),
  adminPinInput: document.getElementById('admin-pin-input'),
  adminTools: document.getElementById('admin-tools'),
  adminViewButton: document.getElementById('admin-view-button'),
  errorBanner: document.getElementById('error-banner'),
  exportButton: document.getElementById('export-button'),
  languageLabel: document.getElementById('language-label'),
  languageSelect: document.getElementById('language-select'),
  inventoryForm: document.getElementById('inventory-form'),
  salesLogPanel: document.getElementById('sales-log-panel'),
  paintingCount: document.getElementById('painting-count'),
  paintingCost: document.getElementById('painting-cost'),
  paintingHeightInch: document.getElementById('painting-height-inch'),
  paintingImage: document.getElementById('painting-image'),
  paintingImageName: document.getElementById('upload-file-name'),
  paintingName: document.getElementById('painting-name'),
  paintingPreviewImage: document.getElementById('upload-preview-image'),
  paintingPrice: document.getElementById('painting-price'),
  paintingQuantity: document.getElementById('painting-quantity'),
  paintingSerialNumber: document.getElementById('painting-serial-number'),
  paintingWidthInch: document.getElementById('painting-width-inch'),
  salesTableBody: document.getElementById('sales-table-body'),
  statsGrid: document.getElementById('stats-grid'),
  setupCard: document.getElementById('setup-card'),
  soldUnits: document.getElementById('sold-units'),
  staffCardGrid: document.getElementById('staff-card-grid'),
  staffSearch: document.getElementById('staff-search'),
  staffName: document.getElementById('staff-name'),
  staffLayout: document.getElementById('staff-layout'),
  staffViewButton: document.getElementById('staff-view-button'),
  statusBanner: document.getElementById('status-banner'),
  totalCash: document.getElementById('total-cash'),
  totalInventoryValue: document.getElementById('inventory-value'),
  unlockAdminButton: document.getElementById('unlock-admin-button'),
  uploadDropzone: document.getElementById('upload-dropzone'),
  brandLabel: document.getElementById('brand-label'),
  pageTitle: document.getElementById('page-title'),
  heroCopy: document.getElementById('hero-copy'),
  setupTitle: document.getElementById('setup-title'),
  setupCopy: document.getElementById('setup-copy'),
  totalCashLabel: document.getElementById('total-cash-label'),
  totalUnitsLabel: document.getElementById('total-units-label'),
  inventoryValueLabel: document.getElementById('inventory-value-label'),
  paintingsListedLabel: document.getElementById('paintings-listed-label'),
  unlockTitle: document.getElementById('unlock-title'),
  unlockCopy: document.getElementById('unlock-copy'),
  adminEyebrow: document.getElementById('admin-eyebrow'),
  addTitle: document.getElementById('add-title'),
  uploadTitle: document.getElementById('upload-title'),
  uploadSubtitle: document.getElementById('upload-subtitle'),
  paintingNameLabel: document.getElementById('painting-name-label'),
  serialNumberLabel: document.getElementById('serial-number-label'),
  costLabel: document.getElementById('cost-label'),
  priceLabel: document.getElementById('price-label'),
  sizeLabel: document.getElementById('size-label'),
  heightLabel: document.getElementById('height-label'),
  widthLabel: document.getElementById('width-label'),
  quantityLabel: document.getElementById('quantity-label'),
  savePaintingButton: document.getElementById('save-painting-button'),
  inventoryEyebrow: document.getElementById('inventory-eyebrow'),
  updateTitle: document.getElementById('update-title'),
  staffEyebrow: document.getElementById('staff-eyebrow'),
  staffTitle: document.getElementById('staff-title'),
  staffSearchLabel: document.getElementById('staff-search-label'),
  staffNameLabel: document.getElementById('staff-name-label'),
  recordsEyebrow: document.getElementById('records-eyebrow'),
  recentSalesTitle: document.getElementById('recent-sales-title'),
  tablePainting: document.getElementById('table-painting'),
  tableQuantitySold: document.getElementById('table-quantity-sold'),
  tableCashReceived: document.getElementById('table-cash-received'),
  tableActions: document.getElementById('table-actions'),
  tableStaff: document.getElementById('table-staff'),
  tableSoldAt: document.getElementById('table-sold-at'),
};

const COPY = {
  en: {
    documentTitle: 'Jennifer Store Keeping',
    brandLabel: 'Jennifer Store Keeping',
    pageTitle: 'Painting inventory and sales tracker',
    heroCopy:
      'Add each painting with its image, quantity, and price. Staff can record sales, and the system keeps live stock records plus an Excel export for reporting.',
    languageLabel: 'Language',
    staffView: 'Staff view',
    adminView: 'Admin view',
    exportReport: 'Export Excel report',
    setupTitle: 'Supabase setup is still required',
    setupCopy:
      'Open supabase-config.js and add your Supabase URL and anon key. GitHub Pages can host the website, while Supabase keeps the shared stock and sales data online for you and your staff.',
    totalCash: 'Total cash received',
    totalUnits: 'Total units sold',
    inventoryValue: 'Current inventory value',
    paintingsListed: 'Paintings listed',
    unlockTitle: 'Unlock admin tools',
    unlockCopy: 'Use your admin PIN before adding stock or changing quantities.',
    adminPinPlaceholder: 'Enter admin PIN',
    unlockButton: 'Unlock',
    adminEyebrow: 'Admin',
    addTitle: 'Add a new painting',
    uploadTitle: 'Drag and drop the painting photo here',
    uploadSubtitle: 'or click to choose the image file',
    noImageSelected: 'No image selected yet.',
    paintingNameLabel: 'Painting name',
    paintingNamePlaceholder: 'Sunset over the lake',
    serialNumberLabel: 'Serial number',
    serialNumberPlaceholder: 'JEN-2026-001',
    costLabel: 'Cost',
    costPlaceholder: '800',
    priceLabel: 'Price',
    pricePlaceholder: '1200',
    sizeLabel: 'Size (inches)',
    heightLabel: 'Height',
    heightPlaceholder: '24',
    widthLabel: 'Width',
    widthPlaceholder: '36',
    quantityLabel: 'Quantity',
    quantityPlaceholder: '3',
    savePainting: 'Save painting',
    inventoryEyebrow: 'Inventory',
    updateTitle: 'Update painting details',
    staffEyebrow: 'Staff',
    staffTitle: 'Record a sale',
    staffSearchLabel: 'Search painting',
    staffSearchPlaceholder: 'Type painting name or serial number',
    staffNameLabel: 'Staff name',
    staffNamePlaceholder: 'Optional',
    recordsEyebrow: 'Records',
    recentSalesTitle: 'Recent sales',
    tablePainting: 'Painting',
    tableQuantitySold: 'Quantity sold',
    tableCashReceived: 'Cash received',
    tableActions: 'Actions',
    tableStaff: 'Staff',
    tableSoldAt: 'Sold at',
    noPaintings: 'No paintings added yet.',
    noSearchResults: 'No paintings match your search.',
    noSales: 'No sales have been recorded yet.',
    saveChanges: 'Save changes',
    removeItem: 'Remove item',
    removeSale: 'Remove sale',
    saving: 'Saving...',
    removing: 'Removing...',
    recordSale: 'Record sale',
    adminPinIncorrect: 'Admin PIN is incorrect.',
    adminUnlocked: 'Admin tools unlocked.',
    paintingAdded: 'Painting added to inventory.',
    paintingUpdated: 'Painting details updated.',
    paintingRemoved: 'Painting removed.',
    saleRecorded: 'Sale recorded.',
    saleRemoved: 'Sale record removed and stock restored.',
    saleRemovedNoRestore: 'Sale record removed.',
    providedImageNamePriceQuantity: 'Please provide image, name, serial number, cost, price, and quantity.',
    validNamePriceQuantity: 'Please enter a valid name, serial number, cost, price, and quantity.',
    validSoldQuantity: 'Please enter a valid sold quantity.',
    paintingNotFound: 'Painting not found.',
    notEnoughStock: 'Not enough stock available.',
    imageFileType: 'Please drop or choose an image file.',
    selectLanguage: 'Language',
    deleteConfirm: 'Delete "{name}" from inventory? This cannot be undone.',
    saleDeleteConfirm: 'Delete sale record for "{name}"? Stock will be restored.',
    configuredMissing: 'Supabase is not configured yet. Please check URL and anon key in supabase-config.js.',
    totalSold: 'Total sold',
    availableNow: 'Available now',
    savePaintingBusy: 'Saving...',
    savePaintingIdle: 'Save painting',
    promptSoldQuantity: 'How many pieces were sold?',
    couldNotReadImageFile: 'Could not read image file.',
    inventoryLoadFailed: 'Inventory load failed: {message}',
    salesLoadFailed: 'Sales load failed: {message}',
    saveFailed: 'Save failed: {message}',
    updateFailed: 'Update failed: {message}',
    deleteFailed: 'Delete failed: {message}',
    inventoryUpdateFailed: 'Inventory update failed: {message}',
    saleRecordFailed: 'Sale record failed: {message}',
    sheetSummary: 'Summary',
    sheetInventory: 'Inventory',
    sheetSales: 'Sales',
    metricTotalUnits: 'Total units sold',
    metricTotalCash: 'Total cash received',
    metricInventoryValue: 'Current inventory value',
    metricActivePaintings: 'Active paintings',
    soldOut: 'Sold out',
  },
  'zh-Hans': {
    documentTitle: 'Jennifer 店铺管理',
    brandLabel: 'Jennifer 店铺管理',
    pageTitle: '绘画库存与销售记录',
    heroCopy: '录入每一幅画的图片、数量和价格。员工可以记录销售，系统会实时保存库存记录，并可导出 Excel 报表。',
    languageLabel: '语言',
    staffView: '员工界面',
    adminView: '管理员界面',
    exportReport: '导出 Excel 报表',
    setupTitle: '仍需要 Supabase 设置',
    setupCopy: '打开 supabase-config.js，填写你的 Supabase 地址和 anon key。GitHub Pages 用来托管网站，Supabase 负责保存库存和销售数据。',
    totalCash: '收到总现金',
    totalUnits: '售出总数量',
    inventoryValue: '当前库存总值',
    paintingsListed: '画作数量',
    unlockTitle: '解锁管理员工具',
    unlockCopy: '在添加库存或修改数量前，请先输入管理员 PIN。',
    adminPinPlaceholder: '输入管理员 PIN',
    unlockButton: '解锁',
    adminEyebrow: '管理员',
    addTitle: '新增画作',
    uploadTitle: '把画作照片拖到这里',
    uploadSubtitle: '或者点击选择图片文件',
    noImageSelected: '尚未选择图片。',
    paintingNameLabel: '画作名称',
    paintingNamePlaceholder: '湖边日落',
    serialNumberLabel: '序列号',
    serialNumberPlaceholder: 'JEN-2026-001',
    costLabel: '成本',
    costPlaceholder: '800',
    priceLabel: '价格',
    pricePlaceholder: '1200',
    sizeLabel: '尺寸（英寸）',
    heightLabel: '高',
    heightPlaceholder: '24',
    widthLabel: '宽',
    widthPlaceholder: '36',
    quantityLabel: '数量',
    quantityPlaceholder: '3',
    savePainting: '保存画作',
    inventoryEyebrow: '库存',
    updateTitle: '修改画作资料',
    staffEyebrow: '员工',
    staffTitle: '记录销售',
    staffSearchLabel: '搜索画作',
    staffSearchPlaceholder: '输入画作名称或序列号',
    staffNameLabel: '员工姓名',
    staffNamePlaceholder: '可选',
    recordsEyebrow: '记录',
    recentSalesTitle: '最近销售',
    tablePainting: '画作',
    tableQuantitySold: '售出数量',
    tableCashReceived: '收到现金',
    tableActions: '操作',
    tableStaff: '员工',
    tableSoldAt: '销售时间',
    noPaintings: '还没有添加画作。',
    noSearchResults: '没有符合搜索条件的画作。',
    noSales: '还没有记录任何销售。',
    saveChanges: '保存修改',
    removeItem: '删除项目',
    removeSale: '删除销售',
    saving: '保存中...',
    removing: '删除中...',
    recordSale: '记录销售',
    adminPinIncorrect: '管理员 PIN 不正确。',
    adminUnlocked: '管理员工具已解锁。',
    paintingAdded: '画作已加入库存。',
    paintingUpdated: '画作资料已更新。',
    paintingRemoved: '画作已删除。',
    saleRecorded: '销售已记录。',
    saleRemoved: '销售记录已删除，库存已恢复。',
    saleRemovedNoRestore: '销售记录已删除。',
    providedImageNamePriceQuantity: '请填写图片、名称、序列号、成本、价格和数量。',
    validNamePriceQuantity: '请输入有效的名称、序列号、成本、价格和数量。',
    validSoldQuantity: '请输入有效的售出数量。',
    paintingNotFound: '找不到该画作。',
    notEnoughStock: '库存不足。',
    imageFileType: '请选择图片文件。',
    selectLanguage: '语言',
    deleteConfirm: '确定要删除“{name}”吗？此操作无法撤销。',
    saleDeleteConfirm: '确定要删除“{name}”的销售记录吗？库存将会恢复。',
    configuredMissing: 'Supabase 尚未配置。请检查 supabase-config.js 里的 URL 和 anon key。',
    totalSold: '总售出',
    availableNow: '现有库存',
    savePaintingBusy: '保存中...',
    savePaintingIdle: '保存画作',
    promptSoldQuantity: '卖出了多少件？',
    couldNotReadImageFile: '无法读取图片文件。',
    inventoryLoadFailed: '库存加载失败：{message}',
    salesLoadFailed: '销售加载失败：{message}',
    saveFailed: '保存失败：{message}',
    updateFailed: '更新失败：{message}',
    deleteFailed: '删除失败：{message}',
    inventoryUpdateFailed: '库存更新失败：{message}',
    saleRecordFailed: '销售记录失败：{message}',
    sheetSummary: '摘要',
    sheetInventory: '库存',
    sheetSales: '销售',
    metricTotalUnits: '售出总数量',
    metricTotalCash: '收到总现金',
    metricInventoryValue: '当前库存总值',
    metricActivePaintings: '现有画作',
    soldOut: '已售罄',
  },
  'zh-Hant': {
    documentTitle: 'Jennifer 店鋪管理',
    brandLabel: 'Jennifer 店鋪管理',
    pageTitle: '繪畫庫存與銷售記錄',
    heroCopy: '輸入每幅畫的圖片、數量和價格。員工可以記錄銷售，系統會即時保存庫存資料，並可匯出 Excel 報表。',
    languageLabel: '語言',
    staffView: '員工介面',
    adminView: '管理員介面',
    exportReport: '匯出 Excel 報表',
    setupTitle: '仍需要 Supabase 設定',
    setupCopy: '打開 supabase-config.js，填入你的 Supabase URL 和 anon key。GitHub Pages 用來託管網站，Supabase 負責保存庫存與銷售資料。',
    totalCash: '收到總現金',
    totalUnits: '售出總數量',
    inventoryValue: '目前庫存總值',
    paintingsListed: '畫作數量',
    unlockTitle: '解鎖管理工具',
    unlockCopy: '在新增庫存或修改數量前，請先輸入管理員 PIN。',
    adminPinPlaceholder: '輸入管理員 PIN',
    unlockButton: '解鎖',
    adminEyebrow: '管理員',
    addTitle: '新增畫作',
    uploadTitle: '把畫作照片拖到這裡',
    uploadSubtitle: '或者點擊選擇圖片檔案',
    noImageSelected: '尚未選擇圖片。',
    paintingNameLabel: '畫作名稱',
    paintingNamePlaceholder: '湖邊日落',
    serialNumberLabel: '序列號',
    serialNumberPlaceholder: 'JEN-2026-001',
    costLabel: '成本',
    costPlaceholder: '800',
    priceLabel: '價格',
    pricePlaceholder: '1200',
    sizeLabel: '尺寸（英吋）',
    heightLabel: '高',
    heightPlaceholder: '24',
    widthLabel: '寬',
    widthPlaceholder: '36',
    quantityLabel: '數量',
    quantityPlaceholder: '3',
    savePainting: '儲存畫作',
    inventoryEyebrow: '庫存',
    updateTitle: '修改畫作資料',
    staffEyebrow: '員工',
    staffTitle: '記錄銷售',
    staffSearchLabel: '搜尋畫作',
    staffSearchPlaceholder: '輸入畫作名稱或序列號',
    staffNameLabel: '員工姓名',
    staffNamePlaceholder: '可選',
    recordsEyebrow: '記錄',
    recentSalesTitle: '最近銷售',
    tablePainting: '畫作',
    tableQuantitySold: '售出數量',
    tableCashReceived: '收到現金',
    tableActions: '操作',
    tableStaff: '員工',
    tableSoldAt: '銷售時間',
    noPaintings: '還沒有新增畫作。',
    noSearchResults: '沒有符合搜尋條件的畫作。',
    noSales: '還沒有記錄任何銷售。',
    saveChanges: '儲存修改',
    removeItem: '刪除項目',
    removeSale: '刪除銷售',
    saving: '儲存中...',
    removing: '刪除中...',
    recordSale: '記錄銷售',
    adminPinIncorrect: '管理員 PIN 不正確。',
    adminUnlocked: '管理工具已解鎖。',
    paintingAdded: '畫作已加入庫存。',
    paintingUpdated: '畫作資料已更新。',
    paintingRemoved: '畫作已刪除。',
    saleRecorded: '銷售已記錄。',
    saleRemoved: '銷售記錄已刪除，庫存已恢復。',
    saleRemovedNoRestore: '銷售記錄已刪除。',
    providedImageNamePriceQuantity: '請填寫圖片、名稱、序列號、成本、價格和數量。',
    validNamePriceQuantity: '請輸入有效的名稱、序列號、成本、價格和數量。',
    validSoldQuantity: '請輸入有效的售出數量。',
    paintingNotFound: '找不到該畫作。',
    notEnoughStock: '庫存不足。',
    imageFileType: '請選擇圖片檔案。',
    selectLanguage: '語言',
    deleteConfirm: '確定要刪除「{name}」嗎？此操作無法復原。',
    saleDeleteConfirm: '確定要刪除「{name}」的銷售記錄嗎？庫存將會恢復。',
    configuredMissing: 'Supabase 尚未設定。請檢查 supabase-config.js 的 URL 和 anon key。',
    totalSold: '總售出',
    availableNow: '現有庫存',
    savePaintingBusy: '儲存中...',
    savePaintingIdle: '儲存畫作',
    promptSoldQuantity: '賣出了多少件？',
    couldNotReadImageFile: '無法讀取圖片檔案。',
    inventoryLoadFailed: '庫存載入失敗：{message}',
    salesLoadFailed: '銷售載入失敗：{message}',
    saveFailed: '儲存失敗：{message}',
    updateFailed: '更新失敗：{message}',
    deleteFailed: '刪除失敗：{message}',
    inventoryUpdateFailed: '庫存更新失敗：{message}',
    saleRecordFailed: '銷售記錄失敗：{message}',
    sheetSummary: '摘要',
    sheetInventory: '庫存',
    sheetSales: '銷售',
    metricTotalUnits: '售出總數量',
    metricTotalCash: '收到總現金',
    metricInventoryValue: '目前庫存總值',
    metricActivePaintings: '現有畫作',
    soldOut: '已售罄',
  },
};

function getCopy(key) {
  return COPY[state.language]?.[key] ?? COPY.en[key] ?? key;
}

function renderLanguage() {
  const copy = COPY[state.language] || COPY.en;

  document.documentElement.lang = state.language;
  document.title = copy.documentTitle;
  dom.languageSelect.value = state.language;

  dom.brandLabel.textContent = copy.brandLabel;
  dom.pageTitle.textContent = copy.pageTitle;
  dom.heroCopy.textContent = copy.heroCopy;
  dom.languageLabel.textContent = copy.languageLabel;
  dom.staffViewButton.textContent = copy.staffView;
  dom.adminViewButton.textContent = copy.adminView;
  dom.exportButton.textContent = copy.exportReport;

  dom.setupTitle.textContent = copy.setupTitle;
  dom.setupCopy.textContent = copy.setupCopy;

  dom.totalCashLabel.textContent = copy.totalCash;
  dom.totalUnitsLabel.textContent = copy.totalUnits;
  dom.inventoryValueLabel.textContent = copy.inventoryValue;
  dom.paintingsListedLabel.textContent = copy.paintingsListed;

  dom.unlockTitle.textContent = copy.unlockTitle;
  dom.unlockCopy.textContent = copy.unlockCopy;
  dom.adminPinInput.placeholder = copy.adminPinPlaceholder;
  dom.unlockAdminButton.textContent = copy.unlockButton;

  dom.adminEyebrow.textContent = copy.adminEyebrow;
  dom.addTitle.textContent = copy.addTitle;
  dom.uploadTitle.textContent = copy.uploadTitle;
  dom.uploadSubtitle.textContent = copy.uploadSubtitle;
  dom.paintingImageName.textContent = state.selectedImageFile ? state.selectedImageFile.name : copy.noImageSelected;
  dom.paintingNameLabel.textContent = copy.paintingNameLabel;
  dom.paintingName.placeholder = copy.paintingNamePlaceholder;
  dom.serialNumberLabel.textContent = copy.serialNumberLabel;
  dom.paintingSerialNumber.placeholder = copy.serialNumberPlaceholder;
  dom.costLabel.textContent = copy.costLabel;
  dom.paintingCost.placeholder = copy.costPlaceholder;
  dom.priceLabel.textContent = copy.priceLabel;
  dom.paintingPrice.placeholder = copy.pricePlaceholder;
  dom.sizeLabel.textContent = copy.sizeLabel;
  dom.heightLabel.textContent = copy.heightLabel;
  dom.paintingHeightInch.placeholder = copy.heightPlaceholder;
  dom.widthLabel.textContent = copy.widthLabel;
  dom.paintingWidthInch.placeholder = copy.widthPlaceholder;
  dom.quantityLabel.textContent = copy.quantityLabel;
  dom.paintingQuantity.placeholder = copy.quantityPlaceholder;
  dom.savePaintingButton.textContent = state.busyAction === 'add-painting' ? copy.savePaintingBusy : copy.savePaintingIdle;

  dom.inventoryEyebrow.textContent = copy.inventoryEyebrow;
  dom.updateTitle.textContent = copy.updateTitle;

  dom.staffEyebrow.textContent = copy.staffEyebrow;
  dom.staffTitle.textContent = copy.staffTitle;
  dom.staffSearchLabel.textContent = copy.staffSearchLabel;
  dom.staffSearch.placeholder = copy.staffSearchPlaceholder;
  dom.staffNameLabel.textContent = copy.staffNameLabel;
  dom.staffName.placeholder = copy.staffNamePlaceholder;

  dom.recordsEyebrow.textContent = copy.recordsEyebrow;
  dom.recentSalesTitle.textContent = copy.recentSalesTitle;
  dom.tablePainting.textContent = copy.tablePainting;
  dom.tableQuantitySold.textContent = copy.tableQuantitySold;
  dom.tableCashReceived.textContent = copy.tableCashReceived;
  dom.tableActions.textContent = copy.tableActions;
  dom.tableStaff.textContent = copy.tableStaff;
  dom.tableSoldAt.textContent = copy.tableSoldAt;
}

const supabaseUrl = String(supabaseConfig?.url ?? '').trim();
const supabaseAnonKey = String(supabaseConfig?.anonKey ?? '').trim();
const isPlaceholderUrl = supabaseUrl.startsWith('REPLACE_WITH_');
const isPlaceholderKey = supabaseAnonKey.startsWith('REPLACE_WITH_');
const hasHttpUrl = /^https?:\/\//i.test(supabaseUrl);
const hasAnonKey = supabaseAnonKey.length > 20;

const configReady =
  Boolean(supabaseUrl) &&
  Boolean(supabaseAnonKey) &&
  !isPlaceholderUrl &&
  !isPlaceholderKey &&
  hasHttpUrl &&
  hasAnonKey;

let supabase = null;

if (configReady) {
  supabase = createClient(supabaseUrl, supabaseAnonKey);
}

if (!configReady) {
  dom.setupCard.hidden = false;
}

function ensureSupabaseConfigured() {
  if (supabase) {
    return true;
  }

  setMessage('error', getCopy('configuredMissing'));
  return false;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 2,
  }).format(Number(value || 0));
}

function formatDate(value) {
  if (!value) {
    return '-';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '-';
  }

  return date.toLocaleString();
}

function setMessage(type, message) {
  state.statusMessage = type === 'status' ? message : '';
  state.errorMessage = type === 'error' ? message : '';
  renderBanners();

  window.clearTimeout(setMessage.timerId);
  if (message) {
    setMessage.timerId = window.setTimeout(() => {
      state.statusMessage = '';
      state.errorMessage = '';
      renderBanners();
    }, 3500);
  }
}

function getSummary() {
  const inventoryValue = state.inventory.reduce(
    (sum, item) => sum + Number(item.price || 0) * Number(item.quantity || 0),
    0,
  );
  const soldUnits = state.sales.reduce((sum, sale) => sum + Number(sale.quantity_sold || 0), 0);
  const totalCash = state.sales.reduce((sum, sale) => sum + Number(sale.total_amount || 0), 0);

  return {
    inventoryValue,
    soldUnits,
    totalCash,
  };
}

function renderStats() {
  const summary = getSummary();
  dom.totalCash.textContent = formatCurrency(summary.totalCash);
  dom.soldUnits.textContent = String(summary.soldUnits);
  dom.totalInventoryValue.textContent = formatCurrency(summary.inventoryValue);
  dom.paintingCount.textContent = String(state.inventory.length);
}

function renderBanners() {
  dom.statusBanner.hidden = !state.statusMessage;
  dom.errorBanner.hidden = !state.errorMessage;
  dom.statusBanner.textContent = state.statusMessage;
  dom.errorBanner.textContent = state.errorMessage;
}

function renderRole() {
  const isAdmin = state.role === 'admin';
  const canViewAdminDetails = isAdmin && state.adminUnlocked;
  dom.adminLayout.hidden = !isAdmin;
  dom.staffLayout.hidden = isAdmin;
  dom.statsGrid.hidden = !canViewAdminDetails;
  dom.salesLogPanel.hidden = !isAdmin;
  dom.staffViewButton.classList.toggle('active', !isAdmin);
  dom.adminViewButton.classList.toggle('active', isAdmin);
  dom.adminTools.hidden = !canViewAdminDetails;
  dom.adminLockCard.hidden = state.adminUnlocked;
  dom.salesLogPanel.hidden = !canViewAdminDetails;
}

function clearSelectedImage() {
  if (state.selectedImagePreviewUrl) {
    URL.revokeObjectURL(state.selectedImagePreviewUrl);
  }

  state.selectedImageFile = null;
  state.selectedImagePreviewUrl = '';
  dom.paintingImage.value = '';
  dom.paintingPreviewImage.hidden = true;
  dom.paintingPreviewImage.removeAttribute('src');
  dom.paintingImageName.textContent = getCopy('noImageSelected');
}

function setSelectedImage(file) {
  if (!file) {
    clearSelectedImage();
    return;
  }

  if (!file.type.startsWith('image/')) {
    setMessage('error', getCopy('imageFileType'));
    return;
  }

  if (state.selectedImagePreviewUrl) {
    URL.revokeObjectURL(state.selectedImagePreviewUrl);
  }

  state.selectedImageFile = file;
  state.selectedImagePreviewUrl = URL.createObjectURL(file);
  dom.paintingPreviewImage.src = state.selectedImagePreviewUrl;
  dom.paintingPreviewImage.hidden = false;
  dom.paintingImageName.textContent = file.name;
}

function adminCardMarkup(item) {
  const busy = state.busyAction === `update-${item.id}`;
  const deleting = state.busyAction === `delete-${item.id}`;
  const sizeText = `${Number(item.height_inch || 0)} x ${Number(item.width_inch || 0)} in`;
  const soldOut = Number(item.quantity || 0) <= 0;
  return `
    <article class="painting-card${soldOut ? ' sold-out' : ''}">
      <img alt="${escapeHtml(item.name)}" class="painting-image" src="${escapeHtml(item.image_url || '')}" />
      ${soldOut ? `<span class="sold-out-badge">${escapeHtml(getCopy('soldOut'))}</span>` : ''}
      <div class="painting-details">
        <h3>${escapeHtml(item.name)}</h3>
        <p>#${escapeHtml(item.serial_number || '')}</p>
        <p>${escapeHtml(getCopy('costLabel'))}: ${formatCurrency(item.cost || 0)}</p>
        <p>${formatCurrency(item.price)}</p>
        <div class="detail-row">
          <span>${escapeHtml(getCopy('sizeLabel'))}</span>
          <strong>${escapeHtml(sizeText)}</strong>
        </div>
        <div class="detail-row">
          <span>${escapeHtml(getCopy('totalSold'))}</span>
          <strong>${Number(item.total_sold || 0)}</strong>
        </div>
        <div class="admin-editor">
          <label>
            ${escapeHtml(getCopy('paintingNameLabel'))}
            <input data-edit-name-id="${item.id}" type="text" value="${escapeHtml(item.name)}" />
          </label>
          <label>
            ${escapeHtml(getCopy('serialNumberLabel'))}
            <input data-edit-serial-id="${item.id}" type="text" value="${escapeHtml(item.serial_number || '')}" />
          </label>
          <label>
            ${escapeHtml(getCopy('costLabel'))}
            <input data-edit-cost-id="${item.id}" type="number" min="0" step="0.01" value="${Number(item.cost || 0)}" />
          </label>
          <label>
            ${escapeHtml(getCopy('priceLabel'))}
            <input data-edit-price-id="${item.id}" type="number" min="0" step="0.01" value="${Number(item.price || 0)}" />
          </label>
          <label>
            ${escapeHtml(getCopy('heightLabel'))}
            <input data-edit-height-id="${item.id}" type="number" min="0" step="0.01" value="${Number(item.height_inch || 0)}" />
          </label>
          <label>
            ${escapeHtml(getCopy('widthLabel'))}
            <input data-edit-width-id="${item.id}" type="number" min="0" step="0.01" value="${Number(item.width_inch || 0)}" />
          </label>
          <label>
            ${escapeHtml(getCopy('quantityLabel'))}
            <input data-edit-quantity-id="${item.id}" type="number" min="0" step="1" value="${Number(item.quantity || 0)}" />
          </label>
          <div class="admin-editor-actions">
            <button data-update-id="${item.id}" type="button" ${busy ? 'disabled' : ''}>${busy ? getCopy('saving') : getCopy('saveChanges')}</button>
            <button data-delete-id="${item.id}" type="button" ${deleting ? 'disabled' : ''}>${deleting ? getCopy('removing') : getCopy('removeItem')}</button>
            <p class="admin-editor-note">${escapeHtml(getCopy('updateTitle'))}</p>
          </div>
        </div>
      </div>
    </article>
  `;
}

function staffCardMarkup(item) {
  const busy = state.busyAction === `sale-${item.id}`;
  const disabled = !Number(item.quantity || 0) || busy;
  const sizeText = `${Number(item.height_inch || 0)} x ${Number(item.width_inch || 0)} in`;
  const soldOut = Number(item.quantity || 0) <= 0;
  return `
    <article class="painting-card${soldOut ? ' sold-out' : ''}">
      <img alt="${escapeHtml(item.name)}" class="painting-image" src="${escapeHtml(item.image_url || '')}" />
      ${soldOut ? `<span class="sold-out-badge">${escapeHtml(getCopy('soldOut'))}</span>` : ''}
      <div class="painting-details">
        <div class="detail-row">
          <span>${escapeHtml(getCopy('availableNow'))}</span>
          <strong>${Number(item.quantity || 0)}</strong>
        </div>
        <div class="detail-row">
          <span>${escapeHtml(getCopy('sizeLabel'))}</span>
          <strong>${escapeHtml(sizeText)}</strong>
        </div>
        <h3>${escapeHtml(item.name)}</h3>
        <p>#${escapeHtml(item.serial_number || '')}</p>
        <p>${formatCurrency(item.price)}</p>
        <button data-sale-id="${item.id}" type="button" ${disabled ? 'disabled' : ''}>${busy ? getCopy('saving') : getCopy('recordSale')}</button>
      </div>
    </article>
  `;
}

function renderInventoryCards() {
  if (!state.inventory.length) {
    const empty = `<div class="empty-state">${escapeHtml(getCopy('noPaintings'))}</div>`;
    dom.adminCardGrid.innerHTML = empty;
    dom.staffCardGrid.innerHTML = empty;
    return;
  }

  dom.adminCardGrid.innerHTML = state.inventory.map(adminCardMarkup).join('');

  const searchTerm = state.staffSearchTerm.trim().toLowerCase();
  const filteredInventory = searchTerm
    ? state.inventory.filter((item) => {
        const name = String(item.name || '').toLowerCase();
        const serial = String(item.serial_number || '').toLowerCase();
        return name.includes(searchTerm) || serial.includes(searchTerm);
      })
    : state.inventory;

  if (!filteredInventory.length) {
    dom.staffCardGrid.innerHTML = `<div class="empty-state">${escapeHtml(getCopy('noSearchResults'))}</div>`;
  } else {
    dom.staffCardGrid.innerHTML = filteredInventory.map(staffCardMarkup).join('');
  }

  dom.adminCardGrid.querySelectorAll('[data-update-id]').forEach((button) => {
    button.addEventListener('click', () => handleUpdatePainting(button.dataset.updateId));
  });

  dom.adminCardGrid.querySelectorAll('[data-delete-id]').forEach((button) => {
    button.addEventListener('click', () => handleDeletePainting(button.dataset.deleteId));
  });

  dom.staffCardGrid.querySelectorAll('[data-sale-id]').forEach((button) => {
    button.addEventListener('click', () => handleSale(button.dataset.saleId));
  });
}

function renderSalesTable() {
  const canManageSales = state.role === 'admin' && state.adminUnlocked;
  dom.tableActions.hidden = !canManageSales;

  if (!state.sales.length) {
    dom.salesTableBody.innerHTML = `<tr><td colspan="${canManageSales ? 6 : 5}">${escapeHtml(getCopy('noSales'))}</td></tr>`;
    return;
  }

  dom.salesTableBody.innerHTML = state.sales
    .map(
      (sale) => `
        <tr>
          <td>${escapeHtml(sale.item_name)}</td>
          <td>${Number(sale.quantity_sold || 0)}</td>
          <td>${formatCurrency(sale.total_amount)}</td>
          ${canManageSales ? `<td><button class="danger-button" data-delete-sale-id="${sale.id}" type="button">${escapeHtml(getCopy('removeSale'))}</button></td>` : ''}
          <td>${escapeHtml(sale.staff_name || 'Staff')}</td>
          <td>${escapeHtml(formatDate(sale.sold_at))}</td>
        </tr>
      `,
    )
    .join('');

  if (canManageSales) {
    dom.salesTableBody.querySelectorAll('[data-delete-sale-id]').forEach((button) => {
      button.addEventListener('click', () => handleDeleteSale(button.dataset.deleteSaleId));
    });
  }
}

function renderAll() {
  renderLanguage();
  renderRole();
  renderStats();
  renderBanners();
  renderInventoryCards();
  renderSalesTable();
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error(getCopy('couldNotReadImageFile')));
    reader.readAsDataURL(file);
  });
}

async function fetchInventory() {
  if (!supabase) {
    return;
  }

  const { data, error } = await supabase
    .from('paintings')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) {
    throw new Error(getCopy('inventoryLoadFailed').replace('{message}', error.message));
  }

  state.inventory = data ?? [];
}

async function fetchSales() {
  if (!supabase) {
    return;
  }

  const { data, error } = await supabase
    .from('sales')
    .select('*')
    .order('sold_at', { ascending: false });

  if (error) {
    throw new Error(getCopy('salesLoadFailed').replace('{message}', error.message));
  }

  state.sales = data ?? [];
}

async function refreshData() {
  if (!supabase) {
    renderAll();
    return;
  }

  try {
    await Promise.all([fetchInventory(), fetchSales()]);
    renderAll();
  } catch (error) {
    setMessage('error', error.message);
  }
}

async function handleAddPainting(event) {
  event.preventDefault();

  if (!ensureSupabaseConfigured()) {
    return;
  }

  const name = dom.paintingName.value.trim();
  const serialNumber = dom.paintingSerialNumber.value.trim();
  const cost = Number(dom.paintingCost.value);
  const price = Number(dom.paintingPrice.value);
  const quantity = Number(dom.paintingQuantity.value);
  const heightInch = Number(dom.paintingHeightInch.value);
  const widthInch = Number(dom.paintingWidthInch.value);
  const imageFile = state.selectedImageFile;

  if (
    !name ||
    !serialNumber ||
    !Number.isFinite(cost) ||
    !Number.isFinite(price) ||
    !Number.isFinite(quantity) ||
    !Number.isFinite(heightInch) ||
    !Number.isFinite(widthInch) ||
    cost < 0 ||
    price < 0 ||
    quantity < 0 ||
    heightInch <= 0 ||
    widthInch <= 0 ||
    !imageFile
  ) {
    setMessage('error', getCopy('providedImageNamePriceQuantity'));
    return;
  }

  state.busyAction = 'add-painting';
  document.getElementById('save-painting-button').disabled = true;
  document.getElementById('save-painting-button').textContent = getCopy('savePaintingBusy');

  try {
    // Store image as data URL for display only to avoid storage upload setup.
    const imageUrl = await fileToDataUrl(imageFile);

    const { error } = await supabase.from('paintings').insert({
      image_url: imageUrl,
      name,
      serial_number: serialNumber,
      cost,
      price,
      quantity,
      height_inch: heightInch,
      width_inch: widthInch,
      total_sold: 0,
    });

    if (error) {
      throw new Error(getCopy('saveFailed').replace('{message}', error.message));
    }

    dom.inventoryForm.reset();
    clearSelectedImage();
    await refreshData();
    setMessage('status', getCopy('paintingAdded'));
  } catch (error) {
    setMessage('error', error.message);
  } finally {
    state.busyAction = '';
    document.getElementById('save-painting-button').disabled = false;
    document.getElementById('save-painting-button').textContent = getCopy('savePaintingIdle');
  }
}

async function handleUpdatePainting(itemId) {
  if (!ensureSupabaseConfigured()) {
    return;
  }

  const item = state.inventory.find((entry) => entry.id === itemId);
  if (!item) {
    return;
  }

  const nameInput = dom.adminCardGrid.querySelector(`[data-edit-name-id="${itemId}"]`);
  const serialInput = dom.adminCardGrid.querySelector(`[data-edit-serial-id="${itemId}"]`);
  const costInput = dom.adminCardGrid.querySelector(`[data-edit-cost-id="${itemId}"]`);
  const priceInput = dom.adminCardGrid.querySelector(`[data-edit-price-id="${itemId}"]`);
  const heightInput = dom.adminCardGrid.querySelector(`[data-edit-height-id="${itemId}"]`);
  const widthInput = dom.adminCardGrid.querySelector(`[data-edit-width-id="${itemId}"]`);
  const quantityInput = dom.adminCardGrid.querySelector(`[data-edit-quantity-id="${itemId}"]`);

  const name = nameInput?.value.trim() || '';
  const serialNumber = serialInput?.value.trim() || '';
  const cost = Number(costInput?.value);
  const price = Number(priceInput?.value);
  const heightInch = Number(heightInput?.value);
  const widthInch = Number(widthInput?.value);
  const quantity = Number(quantityInput?.value);

  if (
    !name ||
    !serialNumber ||
    !Number.isFinite(cost) ||
    !Number.isFinite(price) ||
    !Number.isFinite(quantity) ||
    !Number.isFinite(heightInch) ||
    !Number.isFinite(widthInch) ||
    cost < 0 ||
    price < 0 ||
    quantity < 0 ||
    heightInch <= 0 ||
    widthInch <= 0
  ) {
    setMessage('error', getCopy('validNamePriceQuantity'));
    return;
  }

  state.busyAction = `update-${itemId}`;
  renderInventoryCards();

  try {
    const { error } = await supabase
      .from('paintings')
      .update({
        name,
        serial_number: serialNumber,
        cost,
        price,
        quantity,
        height_inch: heightInch,
        width_inch: widthInch,
        updated_at: new Date().toISOString(),
      })
      .eq('id', itemId);

    if (error) {
      throw new Error(getCopy('updateFailed').replace('{message}', error.message));
    }

    await refreshData();
    setMessage('status', getCopy('paintingUpdated'));
  } catch (error) {
    setMessage('error', error.message);
  } finally {
    state.busyAction = '';
    renderInventoryCards();
  }
}

async function handleDeletePainting(itemId) {
  if (!ensureSupabaseConfigured()) {
    return;
  }

  const item = state.inventory.find((entry) => entry.id === itemId);
  if (!item) {
    return;
  }

  const confirmed = window.confirm(getCopy('deleteConfirm').replace('{name}', item.name));
  if (!confirmed) {
    return;
  }

  state.busyAction = `delete-${itemId}`;
  renderInventoryCards();

  try {
    const { error } = await supabase.from('paintings').delete().eq('id', itemId);

    if (error) {
      throw new Error(getCopy('deleteFailed').replace('{message}', error.message));
    }

    await refreshData();
    setMessage('status', getCopy('paintingRemoved'));
  } catch (error) {
    setMessage('error', error.message);
  } finally {
    state.busyAction = '';
    renderInventoryCards();
  }
}

async function handleDeleteSale(saleId) {
  if (!ensureSupabaseConfigured()) {
    return;
  }

  const sale = state.sales.find((entry) => entry.id === saleId);
  if (!sale) {
    return;
  }

  const confirmed = window.confirm(getCopy('saleDeleteConfirm').replace('{name}', sale.item_name));
  if (!confirmed) {
    return;
  }

  const paintingId = sale.painting_id;
  const quantitySold = Number(sale.quantity_sold || 0);
  const painting = state.inventory.find((entry) => entry.id === paintingId);

  state.busyAction = `delete-sale-${saleId}`;
  renderSalesTable();

  const hasPainting = Boolean(painting);
  const originalQuantity = hasPainting ? Number(painting.quantity || 0) : 0;
  const originalTotalSold = hasPainting ? Number(painting.total_sold || 0) : 0;
  const restoredQuantity = hasPainting ? originalQuantity + quantitySold : 0;
  const restoredTotalSold = hasPainting ? Math.max(0, originalTotalSold - quantitySold) : 0;

  try {
    if (hasPainting) {
      const { error: restoreError } = await supabase
        .from('paintings')
        .update({
          quantity: restoredQuantity,
          total_sold: restoredTotalSold,
          updated_at: new Date().toISOString(),
        })
        .eq('id', paintingId);

      if (restoreError) {
        throw new Error(getCopy('inventoryUpdateFailed').replace('{message}', restoreError.message));
      }
    }

    const { error: deleteError } = await supabase.from('sales').delete().eq('id', saleId);

    if (deleteError) {
      const { error: rollbackError } = await supabase
        .from('paintings')
        .update({
          quantity: originalQuantity,
          total_sold: originalTotalSold,
          updated_at: new Date().toISOString(),
        })
        .eq('id', paintingId);

      if (rollbackError) {
        throw new Error(
          `${getCopy('deleteFailed').replace('{message}', deleteError.message)}; ${getCopy('inventoryUpdateFailed').replace('{message}', rollbackError.message)}`,
        );
      }

      throw new Error(getCopy('deleteFailed').replace('{message}', deleteError.message));
    }

    await refreshData();
    setMessage('status', hasPainting ? getCopy('saleRemoved') : getCopy('saleRemovedNoRestore'));
  } catch (error) {
    setMessage('error', error.message);
  } finally {
    state.busyAction = '';
    renderSalesTable();
  }
}

async function handleSale(itemId) {
  if (!ensureSupabaseConfigured()) {
    return;
  }

  const enteredQuantity = window.prompt(getCopy('promptSoldQuantity'), '1');
  if (!enteredQuantity) {
    return;
  }

  const quantitySold = Number(enteredQuantity);
  if (!Number.isFinite(quantitySold) || quantitySold <= 0) {
    setMessage('error', getCopy('validSoldQuantity'));
    return;
  }

  state.busyAction = `sale-${itemId}`;
  renderInventoryCards();

  try {
    const item = state.inventory.find((entry) => entry.id === itemId);
    if (!item) {
      throw new Error(getCopy('paintingNotFound'));
    }

    const availableQuantity = Number(item.quantity || 0);
    if (availableQuantity < quantitySold) {
      throw new Error(getCopy('notEnoughStock'));
    }

    const nextQuantity = availableQuantity - quantitySold;
    const nextTotalSold = Number(item.total_sold || 0) + quantitySold;

    const { error: updateError } = await supabase
      .from('paintings')
      .update({
        quantity: nextQuantity,
        total_sold: nextTotalSold,
        updated_at: new Date().toISOString(),
      })
      .eq('id', itemId);

    if (updateError) {
      throw new Error(getCopy('inventoryUpdateFailed').replace('{message}', updateError.message));
    }

    const unitPrice = Number(item.price || 0);
    const { error: saleError } = await supabase.from('sales').insert({
      painting_id: itemId,
      item_name: item.name,
      unit_price: unitPrice,
      quantity_sold: quantitySold,
      total_amount: unitPrice * quantitySold,
      staff_name: dom.staffName.value.trim() || 'Staff',
    });

    if (saleError) {
      throw new Error(getCopy('saleRecordFailed').replace('{message}', saleError.message));
    }

    await refreshData();
    setMessage('status', getCopy('saleRecorded'));
  } catch (error) {
    setMessage('error', error.message);
  } finally {
    state.busyAction = '';
    renderInventoryCards();
  }
}

function unlockAdmin() {
  if (!adminPin || dom.adminPinInput.value === adminPin) {
    state.adminUnlocked = true;
    dom.adminPinInput.value = '';
    setMessage('status', getCopy('adminUnlocked'));
    renderAll();
    return;
  }

  setMessage('error', getCopy('adminPinIncorrect'));
}

function exportWorkbook() {
  const summary = getSummary();
  const inventoryRows = state.inventory.map((item) => ({
    Name: item.name,
    SerialNumber: item.serial_number || '',
    Cost: Number(item.cost || 0),
    Price: Number(item.price || 0),
    HeightInch: Number(item.height_inch || 0),
    WidthInch: Number(item.width_inch || 0),
    QuantityInStock: Number(item.quantity || 0),
    TotalSold: Number(item.total_sold || 0),
    ProfitPerUnit: Number(item.price || 0) - Number(item.cost || 0),
    StockValue: Number(item.price || 0) * Number(item.quantity || 0),
    UpdatedAt: formatDate(item.updated_at),
  }));

  const salesRows = state.sales.map((sale) => ({
    Painting: sale.item_name,
    QuantitySold: Number(sale.quantity_sold || 0),
    UnitPrice: Number(sale.unit_price || 0),
    TotalAmount: Number(sale.total_amount || 0),
    Staff: sale.staff_name || 'Staff',
    SoldAt: formatDate(sale.sold_at),
  }));

  const summaryRows = [
    { Metric: getCopy('metricTotalUnits'), Value: summary.soldUnits },
    { Metric: getCopy('metricTotalCash'), Value: summary.totalCash },
    { Metric: getCopy('metricInventoryValue'), Value: summary.inventoryValue },
    { Metric: getCopy('metricActivePaintings'), Value: state.inventory.length },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), getCopy('sheetSummary'));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(inventoryRows), getCopy('sheetInventory'));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(salesRows), getCopy('sheetSales'));
  XLSX.writeFile(workbook, `painting-inventory-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function subscribeToData() {
  if (!supabase) {
    renderAll();
    return;
  }

  supabase
    .channel('paintings-changes')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'paintings' }, refreshData)
    .subscribe();

  supabase
    .channel('sales-changes')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'sales' }, refreshData)
    .subscribe();

  refreshData();
}

function bindEvents() {
  dom.staffViewButton.addEventListener('click', () => {
    state.role = 'staff';
    renderAll();
  });

  dom.adminViewButton.addEventListener('click', () => {
    state.role = 'admin';
    renderAll();
  });

  dom.unlockAdminButton.addEventListener('click', unlockAdmin);
  dom.inventoryForm.addEventListener('submit', handleAddPainting);
  dom.exportButton.addEventListener('click', exportWorkbook);
  dom.paintingImage.addEventListener('change', (event) => {
    setSelectedImage(event.target.files?.[0] ?? null);
  });
  dom.uploadDropzone.addEventListener('click', () => {
    dom.paintingImage.click();
  });
  dom.uploadDropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      dom.paintingImage.click();
    }
  });
  dom.uploadDropzone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dom.uploadDropzone.classList.add('drag-active');
  });
  dom.uploadDropzone.addEventListener('dragleave', () => {
    dom.uploadDropzone.classList.remove('drag-active');
  });
  dom.uploadDropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    dom.uploadDropzone.classList.remove('drag-active');
    setSelectedImage(event.dataTransfer?.files?.[0] ?? null);
  });

  dom.languageSelect.addEventListener('change', (event) => {
    state.language = event.target.value;
    localStorage.setItem('store-language', state.language);
    renderAll();
  });

  dom.staffSearch.addEventListener('input', (event) => {
    state.staffSearchTerm = event.target.value || '';
    renderInventoryCards();
  });
}

bindEvents();
renderAll();
subscribeToData();
