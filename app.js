// ============================================
// Ad Creative Analyzer - Main Application
// ============================================

let rawData = [];
let filteredData = [];
let charts = {};
let currentSection = 'overview';
let creativePage = 0;
const PAGE_SIZE = 20;

// Column mapping - matches actual CSV headers
const COL_MAP = {
  date: '日',
  cost: '費用',
  imps: 'Imps',
  clicks: 'Clicks',
  cv: 'CV (CVs)',
  cpa: 'CV (CPA)',
  lp: 'LP名',
  targeting: 'ターゲティング',
  appeal: 'バナー訴求',
  bannerType: 'バナータイプ',
  adName: '広告名',
  adText: 'テキスト 1',
  adText2: 'テキスト 2',
  operation: '運用',
  ebisCV: 'EBiS遷移数 (CVs)',
  interest: '興味関心カテゴリ',
  campaign: 'キャンペーン名',
  adGroup: '広告グループ名',
  cpc: 'CPC',
  cpm: 'CPM',
  lpCategory: 'LPカテゴリ',
  bannerTarget: 'バナーターゲット',
};

const COLORS = ['#e88c7a','#d4956a','#c97b8a','#e8b86a','#b8a060','#7ab8d4','#7dbd8a','#e07a7a','#c4a888','#a89078','#d4a07a','#8db88a','#8aafb8','#c9887a','#e0c07a'];

// ============================================
// Initialization
// ============================================
document.addEventListener('DOMContentLoaded', () => {
  setupSidebar();
  setupUploadZone();
  setupFilters();
  // 保存済みデータがあれば自動復帰
  autoRestoreData();
});

function setupSidebar() {
  document.getElementById('sidebar-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
  });
  document.getElementById('mobile-menu-btn').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('mobile-open');
  });
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      const section = item.dataset.section;
      if (!rawData.length && section !== 'overview') return;
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      item.classList.add('active');
      showSection(section);
      document.getElementById('sidebar').classList.remove('mobile-open');
    });
  });
}

function setupUploadZone() {
  const zone = document.getElementById('upload-zone');
  const input = document.getElementById('csv-input');
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
  });
  input.addEventListener('change', e => handleFiles(e.target.files));
}

function setupFilters() {
  document.getElementById('filter-month').addEventListener('change', function() {
    // 月フィルタ変更時に日付範囲をリセット
    document.getElementById('filter-date-start').value = '';
    document.getElementById('filter-date-end').value = '';
    applyFilters();
  });
  document.getElementById('filter-operation').addEventListener('change', applyFilters);
  document.getElementById('filter-date-start').addEventListener('change', function() {
    // 日付指定時は月フィルタを「全期間」に戻す
    document.getElementById('filter-month').value = 'all';
    applyFilters();
  });
  document.getElementById('filter-date-end').addEventListener('change', function() {
    document.getElementById('filter-month').value = 'all';
    applyFilters();
  });
}

// ============================================
// Section Navigation
// ============================================
function showSection(id) {
  currentSection = id;
  // Hide welcome
  const welcomeEl = document.getElementById('welcome-section');
  if (welcomeEl) welcomeEl.style.display = 'none';
  // Hide all sections
  document.querySelectorAll('.dashboard-section').forEach(s => {
    s.classList.add('hidden');
    s.style.display = 'none';
  });
  // Show target section
  const section = document.getElementById('section-' + id);
  if (section) {
    section.classList.remove('hidden');
    section.style.display = 'block';
  }
  const titles = { overview:'概要ダッシュボード', creative:'クリエイティブ分析', lp:'LP分析', targeting:'ターゲティング分析', cross:'クロス分析', trend:'トレンド分析', text:'テキスト分析', combo:'掛け合わせ効果分析' };
  document.getElementById('page-title').textContent = titles[id] || '';
  if (rawData.length) refreshSection(id);
}

function refreshSection(id) {
  try {
    switch(id) {
      case 'overview': renderOverview(); break;
      case 'creative': renderCreative(); break;
      case 'lp': renderLP(); break;
      case 'targeting': renderTargeting(); break;
      case 'cross': updateHeatmap(); break;
      case 'trend': renderTrend(); break;
      case 'text': renderText(); break;
      case 'combo': renderCombo(); break;
    }
  } catch(err) {
    console.error('Section render error:', id, err);
  }
}

// ============================================
// Data Modal
// ============================================
function showDataModal() { document.getElementById('data-modal').classList.remove('hidden'); }
function hideDataModal() { document.getElementById('data-modal').classList.add('hidden'); }

function switchModalTab(tab) {
  document.querySelectorAll('.modal-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.add('hidden'));
  document.querySelector(`[data-tab="${tab}"]`).classList.add('active');
  document.getElementById('tab-' + tab).classList.remove('hidden');
}

// ============================================
// CSV Parsing - Robust version
// ============================================
function handleFiles(files) {
  if (!files.length) return;
  showLoading('CSVファイルを読み込んでいます...');
  
  const file = files[0];
  console.log('📁 File:', file.name, 'Size:', file.size, 'Type:', file.type);
  
  const reader = new FileReader();
  reader.onload = e => {
    try {
      let text = e.target.result;
      console.log('📄 Text length:', text.length);
      console.log('📄 First 500 chars:', text.substring(0, 500));
      
      rawData = parseCSV(text);
      console.log('✅ Parsed rows:', rawData.length);
      
      if (rawData.length > 0) {
        console.log('📊 Columns:', Object.keys(rawData[0]));
        console.log('📊 First row:', rawData[0]);
        // IndexedDB にCSVテキストを保存（永続化）
        saveDataToIDB(text, file.name);
        hideDataModal();
        processData();
      } else {
        alert('CSVのパースに失敗しました。データが0行です。');
        hideLoading();
      }
    } catch(err) {
      console.error('CSV parse error:', err);
      alert('CSV読み込みエラー: ' + err.message);
      hideLoading();
    }
  };
  reader.onerror = err => {
    console.error('FileReader error:', err);
    alert('ファイル読み込みエラー');
    hideLoading();
  };
  
  // Try UTF-8 first, fallback to Shift_JIS
  reader.readAsText(file, 'UTF-8');
}

function parseCSV(text) {
  // Normalize CRLF but keep LF inside quotes
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  // Detect delimiter from first line (up to first newline)
  const firstNewline = text.indexOf('\n');
  const firstLine = firstNewline >= 0 ? text.substring(0, firstNewline) : text;
  const tabCount = (firstLine.match(/\t/g) || []).length;
  const commaCount = (firstLine.match(/,/g) || []).length;
  const delim = tabCount > commaCount ? '\t' : ',';
  console.log(`🔍 Delimiter: "${delim === '\t' ? 'TAB' : 'COMMA'}" (tabs:${tabCount}, commas:${commaCount})`);
  
  // Character-by-character full-text parser (handles embedded newlines in quoted fields)
  const allRows = [];
  let current = '';
  let inQuotes = false;
  let currentRow = [];
  
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inQuotes && i + 1 < text.length && text[i + 1] === '"') {
        current += '"'; i++; // escaped quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delim && !inQuotes) {
      currentRow.push(current.trim());
      current = '';
    } else if (ch === '\n' && !inQuotes) {
      currentRow.push(current.trim());
      current = '';
      if (currentRow.some(c => c !== '')) allRows.push(currentRow);
      currentRow = [];
    } else {
      current += ch;
    }
  }
  // last field/row
  if (current.trim() || currentRow.length > 0) {
    currentRow.push(current.trim());
    if (currentRow.some(c => c !== '')) allRows.push(currentRow);
  }
  
  if (allRows.length < 2) { console.warn('No data rows found'); return []; }
  
  const headers = allRows[0].map(h => h.trim());
  console.log('📋 Headers:', headers.length, headers.slice(0, 10));
  
  const data = [];
  let errorCount = 0;
  const minCols = Math.floor(headers.length * 0.5); // allow up to 50% missing cols
  
  for (let i = 1; i < allRows.length; i++) {
    const vals = allRows[i];
    if (vals.length < minCols) { errorCount++; continue; }
    const row = {};
    headers.forEach((h, idx) => { row[h] = (vals[idx] || '').trim(); });
    data.push(row);
  }
  
  if (errorCount > 0) console.warn(`⚠️ Skipped ${errorCount} short rows`);
  return data;
}

function parseCSVLine(line, delim) {
  // Keep for backward compat (pasted data / single line use)
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
        current += '"'; i++;
      } else { inQuotes = !inQuotes; }
    } else if (ch === delim && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

function processPastedData() {
  const text = document.getElementById('paste-data').value;
  if (!text.trim()) return;
  showLoading('ペーストデータを処理中...');
  rawData = parseCSV(text);
  if (rawData.length) {
    saveDataToIDB(text, 'pasted_data');
    hideDataModal();
    processData();
  }
  else { alert('データのパースに失敗しました'); hideLoading(); }
}

// ============================================
// IndexedDB データ永続化
// ============================================
const IDB_NAME = 'AdAnalyzerDB';
const IDB_STORE = 'csvData';
const IDB_VERSION = 1;

function openIDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}

async function saveDataToIDB(csvText, fileName) {
  try {
    const db = await openIDB();
    const tx = db.transaction(IDB_STORE, 'readwrite');
    const store = tx.objectStore(IDB_STORE);
    store.put(csvText, 'csv');
    store.put(fileName, 'fileName');
    store.put(new Date().toLocaleString('ja-JP'), 'savedAt');
    tx.oncomplete = () => {
      console.log('Data saved to IndexedDB (' + (csvText.length / 1024 / 1024).toFixed(1) + 'MB)');
      localStorage.setItem('data_connected', 'true');
      updateDataConnectionUI();
    };
  } catch(e) {
    console.warn('IndexedDB save failed:', e);
  }
}

async function loadDataFromIDB() {
  try {
    const db = await openIDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const store = tx.objectStore(IDB_STORE);
      const csvReq = store.get('csv');
      const nameReq = store.get('fileName');
      const dateReq = store.get('savedAt');
      tx.oncomplete = () => {
        resolve({
          csv: csvReq.result || null,
          fileName: nameReq.result || '',
          savedAt: dateReq.result || ''
        });
      };
      tx.onerror = () => reject(tx.error);
    });
  } catch(e) {
    console.warn('IndexedDB load failed:', e);
    return null;
  }
}

async function clearIDB() {
  try {
    const db = await openIDB();
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).clear();
    tx.oncomplete = () => console.log('IndexedDB cleared');
  } catch(e) {
    console.warn('IndexedDB clear failed:', e);
  }
}

function disconnectData() {
  clearIDB();
  localStorage.removeItem('data_connected');
  localStorage.removeItem('sheets_config');
  localStorage.removeItem('sheets_last_sync');
  if (_sheetsRefreshTimer) {
    clearInterval(_sheetsRefreshTimer);
    _sheetsRefreshTimer = null;
  }
  rawData = [];
  filteredData = [];
  // ヘッダーボタンを元に戻す
  const dataBtn = document.querySelector('[onclick="showDataModal()"]');
  if (dataBtn) {
    dataBtn.style.borderColor = '';
    dataBtn.querySelector('span').textContent = 'データ接続';
  }
  const dcBtn = document.getElementById('disconnect-btn');
  if (dcBtn) dcBtn.style.display = 'none';
  // ステータス更新
  const sb = document.getElementById('status-bar');
  if (sb) sb.textContent = '連携解除しました';
  // モーダル内の連携中表示をリセット
  const connInfo = document.getElementById('sheets-connected-info');
  const connForm = document.getElementById('sheets-connect-form');
  if (connInfo) connInfo.classList.add('hidden');
  if (connForm) connForm.style.display = '';
  console.log('Data disconnected');
}

function updateDataConnectionUI() {
  const dataBtn = document.querySelector('[onclick="showDataModal()"]');
  if (dataBtn) {
    dataBtn.style.borderColor = '#34d399';
    dataBtn.querySelector('span').textContent = '連携中';
  }
  const dcBtn = document.getElementById('disconnect-btn');
  if (dcBtn) dcBtn.style.display = '';
}

async function autoRestoreData() {
  // Sheets連携を先に試行
  const sheetsConfig = localStorage.getItem('sheets_config');
  if (sheetsConfig) {
    try {
      const config = JSON.parse(sheetsConfig);
      console.log('Auto-reconnecting to Sheets...');
      await fetchSheetsData(config);
      return;
    } catch(e) {
      console.warn('Sheets auto-reconnect failed, trying IndexedDB...', e);
    }
  }
  // IndexedDB からCSVデータを復帰
  if (localStorage.getItem('data_connected') === 'true') {
    try {
      const stored = await loadDataFromIDB();
      if (stored && stored.csv) {
        console.log('Restoring data from IndexedDB... (' + stored.fileName + ', saved ' + stored.savedAt + ')');
        showLoading('保存済みデータを復帰中...');
        rawData = parseCSV(stored.csv);
        if (rawData.length > 0) {
          updateDataConnectionUI();
          processData();
          console.log('Data restored: ' + rawData.length + ' rows from ' + stored.fileName);
        } else {
          hideLoading();
        }
      }
    } catch(e) {
      console.warn('Auto-restore failed:', e);
    }
  }
}

let _sheetsRefreshTimer = null;

async function connectGoogleSheets() {
  const url = document.getElementById('sheets-url').value.trim();
  if (!url) return;
  const tabName = (document.getElementById('sheets-tab-name')?.value || '').trim();
  const apiKey = (document.getElementById('sheets-api-key')?.value || '').trim();
  const autoRefresh = document.getElementById('sheets-auto-refresh')?.checked !== false;

  // Save to localStorage
  const config = { url, tabName, apiKey, autoRefresh };
  localStorage.setItem('sheets_config', JSON.stringify(config));

  await fetchSheetsData(config);
}

function fetchSheetsData(config) {
  const { url, tabName, apiKey } = config;
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = match ? match[1] : url;

  showLoading('Google Sheetsからデータを取得中...');

  // JSONP方式: file://プロトコルでもCORS制限を回避
  // Google Sheets の gviz エンドポイントに responseHandler を指定して
  // <script>タグで読み込むことでCORSをバイパス
  return new Promise((resolve, reject) => {
    const callbackName = '_sheetsCb_' + Date.now();
    const sheet = tabName ? '&sheet=' + encodeURIComponent(tabName) : '';
    const scriptUrl = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:json;responseHandler:${callbackName}${sheet}`;

    // タイムアウト処理
    const timeout = setTimeout(() => {
      cleanup();
      alert('Google Sheets接続エラー: タイムアウト。スプシが共有設定「リンクを知っている全員に閲覧権限」になっているか確認してください');
      hideLoading();
      reject(new Error('timeout'));
    }, 15000);

    function cleanup() {
      clearTimeout(timeout);
      delete window[callbackName];
      const s = document.getElementById('sheets-jsonp-script');
      if (s) s.remove();
    }

    // コールバック関数を登録
    window[callbackName] = function(response) {
      cleanup();

      if (response.status !== 'ok') {
        const errMsg = response.errors?.[0]?.detailed_message || response.errors?.[0]?.message || '不明なエラー';
        alert('Google Sheets接続エラー: ' + errMsg);
        hideLoading();
        reject(new Error(errMsg));
        return;
      }

      const table = response.table;
      if (!table || !table.cols || !table.rows || table.rows.length === 0) {
        alert('Google Sheets接続エラー: データが空です');
        hideLoading();
        reject(new Error('empty'));
        return;
      }

      // ヘッダー名を取得（cols[i].label が空の場合は cols[i].id を使用）
      const headers = table.cols.map((c, i) => c.label || c.id || ('col' + i));

      // 行データをパース
      rawData = table.rows.map(row => {
        const obj = {};
        if (!row.c) return obj;
        row.c.forEach((cell, i) => {
          if (i >= headers.length) return;
          if (!cell || cell.v === null || cell.v === undefined) {
            obj[headers[i]] = '';
          } else {
            // gviz は日付を "Date(y,m,d)" 形式で返す → 文字列に変換
            let val = cell.v;
            if (typeof val === 'string' && val.startsWith('Date(')) {
              // Date(2025,0,15) → 2025/01/15
              const dm = val.match(/Date\((\d+),(\d+),(\d+)/);
              if (dm) {
                val = dm[1] + '/' + String(Number(dm[2]) + 1).padStart(2, '0') + '/' + dm[3].padStart(2, '0');
              }
            }
            // formatted value があればそちらを優先（日付や通貨の表示用）
            obj[headers[i]] = cell.f || String(val);
          }
        });
        return obj;
      }).filter(r => Object.values(r).some(v => v !== ''));  // 空行を除外

      if (!rawData.length) {
        alert('Google Sheets接続エラー: データが空です');
        hideLoading();
        reject(new Error('empty'));
        return;
      }

      // 同期時刻を記録
      const now = new Date().toLocaleString('ja-JP');
      localStorage.setItem('sheets_last_sync', now);

      // UI更新
      updateSheetsUI(config, now);
      hideDataModal();
      processData();

      // 自動更新
      if (config.autoRefresh) {
        startSheetsAutoRefresh(config);
      }

      console.log('Sheets connected: ' + rawData.length + ' rows loaded');
      resolve();
    };

    // scriptタグを作成して読み込み
    const script = document.createElement('script');
    script.id = 'sheets-jsonp-script';
    script.src = scriptUrl;
    script.onerror = function() {
      cleanup();
      alert('Google Sheets接続エラー: スクリプトの読み込みに失敗しました。URLが正しいか、スプシが共有されているか確認してください');
      hideLoading();
      reject(new Error('script error'));
    };
    document.head.appendChild(script);
  });
}

function updateSheetsUI(config, lastSync) {
  const connInfo = document.getElementById('sheets-connected-info');
  const connForm = document.getElementById('sheets-connect-form');
  const connUrl = document.getElementById('sheets-connected-url');
  const syncEl = document.getElementById('sheets-last-sync');

  if (connInfo && connForm) {
    connInfo.classList.remove('hidden');
    connForm.style.display = 'none';
    if (connUrl) connUrl.textContent = config.url + (config.tabName ? ' [' + config.tabName + ']' : '');
    if (syncEl) syncEl.textContent = '最終同期: ' + (lastSync || '-');
  }

  // Update header button to show connected state
  const dataBtn = document.querySelector('[onclick="showDataModal()"]');
  if (dataBtn) {
    dataBtn.style.borderColor = '#34d399';
    dataBtn.querySelector('span').textContent = '連携中';
  }
}

function disconnectSheets() {
  disconnectData();
}

async function refreshSheetsData() {
  const stored = localStorage.getItem('sheets_config');
  if (!stored) return;
  const config = JSON.parse(stored);
  await fetchSheetsData(config);
}

function startSheetsAutoRefresh(config) {
  if (_sheetsRefreshTimer) clearInterval(_sheetsRefreshTimer);
  _sheetsRefreshTimer = setInterval(async () => {
    console.log('Auto-refreshing Sheets data...');
    try {
      await fetchSheetsData(config);
    } catch(e) {
      console.warn('Auto-refresh failed:', e);
    }
  }, 5 * 60 * 1000); // 5分ごと
}

// autoReconnectSheets は autoRestoreData に統合済み

// ============================================
// Data Processing
// ============================================
function processData() {
  showLoading(`${rawData.length.toLocaleString()}行のデータを処理中...`);
  
  const cols = Object.keys(rawData[0]);
  console.log('🔧 Processing with columns:', cols);
  
  // Auto-detect column mapping by fuzzy match
  function findCol(candidates) {
    for (const c of candidates) {
      const exact = cols.find(col => col === c);
      if (exact) return exact;
    }
    for (const c of candidates) {
      const partial = cols.find(col => col.includes(c));
      if (partial) return partial;
    }
    return null;
  }
  
  // Override mappings with auto-detected columns
  const dateCol   = findCol(['日']);
  const costCol   = findCol(['費用']);
  const impsCol   = findCol(['Imps']);
  const clicksCol = findCol(['Clicks']);
  // CV = EBiSCV (CVs) が正 (EBiS計測値を使用)
  const cvCol     = findCol(['EBiSCV (CVs)']);
  const lpCol     = findCol(['LP名']);
  const tgCol     = findCol(['ターゲティング']);
  const appealCol = findCol(['バナー訴求']);
  const btCol     = findCol(['バナータイプ']);
  const adNameCol = findCol(['広告名']);
  const textCol   = findCol(['テキスト 1', 'テキスト1']);
  const text2Col  = findCol(['テキスト 2', 'テキスト2']);
  const opCol     = findCol(['運用']);
  const ebisCol   = findCol(['EBiS遷移数 (CVs)']);
  const ebisCvColAlt = findCol(['EBiSCV (CVs)']); // same as cvCol, used for redundancy check
  const intCol    = findCol(['興味関心カテゴリ', '興味関心']);
  const campCol   = findCol(['キャンペーン名', 'キャンペーン']);
  
  console.log('🗂️ Column mapping:', { dateCol, costCol, impsCol, clicksCol, cvCol, lpCol, tgCol, appealCol, btCol, adNameCol, textCol, opCol, ebisCol, intCol, campCol });
  
  // Parse data
  rawData.forEach(row => {
    row._cost = numParse(row[costCol]);
    row._imps = numParse(row[impsCol]);
    row._clicks = numParse(row[clicksCol]);
    row._cv = numParse(row[cvCol]);
    row._ebisCV = numParse(row[ebisCol]);
    row._ctr = row._imps > 0 ? (row._clicks / row._imps * 100) : 0;
    row._cpa = row._cv > 0 ? (row._cost / row._cv) : null;
    
    // LP名: まず「LP名」列を参照、空ならキャンペーン名から抽出
    let lp = (lpCol ? row[lpCol] : '') || '';
    if (!lp && campCol && row[campCol]) {
      // キャンペーン名から「〇〇LP」「LP〇〇」パターンを抽出
      // 例: "記事LP012(疑いLPver2)" や "本サイトsp548(カータンさんLP)" 等
      const campName = row[campCol];
      const lpMatch = campName.match(/((?:記事|本サイト|)[A-Za-z]*LP\d*[^\s_)]*(?:\([^)]+\))?)/i)
                   || campName.match(/(LP\([^)]+\))/i)
                   || campName.match(/(LP[^\s_]*)/i);
      if (lpMatch) {
        lp = lpMatch[1];
      }
    }
    row._lp = lp || '不明';
    
    row._targeting = (tgCol ? row[tgCol] : '') || '不明';
    row._appeal = (appealCol ? row[appealCol] : '') || '不明';
    row._bannerType = (btCol ? row[btCol] : '') || '不明';
    row._adName = (adNameCol ? row[adNameCol] : '') || '';
    row._adText = (textCol ? row[textCol] : '') || '';
    if (text2Col && row[text2Col]) row._adText += ' ' + row[text2Col];
    row._operation = (opCol ? row[opCol] : '') || '';
    row._interest = (intCol ? row[intCol] : '') || '';
    
    // Extract month and date from date column
    row._month = '';
    row._date = '';
    if (dateCol && row[dateCol]) {
      const d = String(row[dateCol]);
      // yyyy/mm/dd or yyyy-mm-dd
      const m = d.match(/(\d{4})[\/\-](\d{1,2})[\/\-]?(\d{1,2})?/);
      if (m) {
        row._month = m[1] + '-' + m[2].padStart(2, '0');
        if (m[3]) {
          row._date = m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
        }
      }
    }
    
    // Extract creator
    row._creator = extractCreator(row._adName);
  });
  
  // Verify parsing worked
  const sampleCosts = rawData.slice(0, 10).map(r => r._cost);
  console.log('💰 Sample costs:', sampleCosts);
  const totalCost = rawData.reduce((s,r) => s + r._cost, 0);
  const totalCV = rawData.reduce((s,r) => s + r._cv, 0);
  console.log(`💰 Total cost: ¥${totalCost.toLocaleString()}, Total CV: ${totalCV}`);
  
  // Populate month filter
  const months = [...new Set(rawData.map(r => r._month).filter(Boolean))].sort();
  const sel = document.getElementById('filter-month');
  sel.innerHTML = '<option value="all">全期間</option>';
  months.forEach(m => { sel.innerHTML += `<option value="${m}">${m}</option>`; });

  // 日付レンジのmin/maxを設定
  const dates = rawData.map(r => r._date).filter(Boolean).sort();
  if (dates.length) {
    const startInput = document.getElementById('filter-date-start');
    const endInput = document.getElementById('filter-date-end');
    if (startInput) { startInput.min = dates[0]; startInput.max = dates[dates.length - 1]; }
    if (endInput) { endInput.min = dates[0]; endInput.max = dates[dates.length - 1]; }
  }
  
  // Populate operation filter
  const ops = [...new Set(rawData.map(r => r._operation).filter(Boolean))];
  const opSel = document.getElementById('filter-operation');
  opSel.innerHTML = '<option value="all">全運用</option>';
  ops.forEach(o => { opSel.innerHTML += `<option value="${o}">${o}</option>`; });
  
  // Update status
  document.getElementById('data-status').innerHTML = `<div class="status-dot online"></div><span>${rawData.length.toLocaleString()}行</span>`;
  
  // Hide welcome, show overview
  const welcomeEl = document.getElementById('welcome-section');
  if (welcomeEl) welcomeEl.style.display = 'none';
  
  // Force show overview section
  currentSection = 'overview';
  document.querySelectorAll('.dashboard-section').forEach(s => {
    s.classList.add('hidden');
    s.style.display = '';
  });
  const overviewEl = document.getElementById('section-overview');
  if (overviewEl) {
    overviewEl.classList.remove('hidden');
    overviewEl.style.display = 'block';
  }
  document.getElementById('page-title').textContent = '概要ダッシュボード';
  
  // Apply filters (this triggers refreshSection)
  applyFilters();
  
  // Hide loading after a small delay to ensure DOM is ready
  setTimeout(() => {
    hideLoading();
    console.log('✅ Dashboard ready');
  }, 100);
}

function numParse(v) {
  if (v === null || v === undefined || v === '') return 0;
  // Strip everything except digits, decimal point, and minus sign
  // Handles ¥, ￥, full-width yen variants, commas, spaces, %, etc.
  const cleaned = String(v).replace(/[^\d.\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '.') return 0;
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function extractCreator(name) {
  if (!name) return 'other';
  const n = String(name).toLowerCase();
  const creators = [
    ['おっちゃん','おっちゃん'],['まりまり','まりまり'],['マミーマウス子ビッツ','ﾏﾐｰﾏｳｽ子'],
    ['新庄アキラ','新庄アキラ'],['ミロチ','mirochi'],['ミニマリストますみ','ますみ'],
    ['のなか海','のなか海'],['みたん','みたん'],['しまゆう','しまゆう'],['中山少年','中山少年'],
    ['inui','inui'],['omochico','omochico'],['kanako','kanako'],['yuri','yuri'],['MizukiOkami','mizukiokami']
  ];
  for (const [label, key] of creators) {
    if (n.includes(key.toLowerCase())) return label;
  }
  if (n.includes('manga_carousel') || (n.includes('manga') && n.includes('carousel'))) return 'manga_carousel';
  if (/_mnm_/.test(n) || n.includes('_mnm')) return 'mnm';
  if (/_mro_/.test(n) || n.includes('_mro')) return 'mro';
  if (/_d_/.test(n)) return 'd';
  return 'other';
}

function applyFilters() {
  const month = document.getElementById('filter-month').value;
  const op = document.getElementById('filter-operation').value;
  const dateStart = document.getElementById('filter-date-start').value;  // YYYY-MM-DD or ''
  const dateEnd = document.getElementById('filter-date-end').value;

  filteredData = rawData.filter(r => {
    // 月フィルタ
    if (month !== 'all' && r._month !== month) return false;
    // 運用フィルタ
    if (op !== 'all' && r._operation !== op) return false;
    // 日付範囲フィルタ
    if (dateStart && r._date && r._date < dateStart) return false;
    if (dateEnd && r._date && r._date > dateEnd) return false;
    return true;
  });
  console.log(`🔽 Filtered: ${filteredData.length} / ${rawData.length} rows`);
  // ステータスバーにフィルタ情報を表示
  const sb = document.getElementById('status-bar');
  if (sb) {
    let label = '';
    if (month !== 'all') label += month + ' ';
    if (dateStart || dateEnd) label += (dateStart || '...') + ' 〜 ' + (dateEnd || '...') + ' ';
    if (op !== 'all') label += op + ' ';
    sb.textContent = label ? label + `(${filteredData.length.toLocaleString()}行)` : `全データ (${filteredData.length.toLocaleString()}行)`;
  }
  if (rawData.length) refreshSection(currentSection);
}

// ============================================
// Aggregation Helpers
// ============================================
function aggregate(data, groupKey) {
  const groups = {};
  data.forEach(r => {
    const key = typeof groupKey === 'function' ? groupKey(r) : r[groupKey];
    if (!key || key === '不明' || key === '#REF!' || key === '') return;
    if (!groups[key]) groups[key] = { cost:0, imps:0, clicks:0, cv:0, ebisCV:0, count:0 };
    const g = groups[key];
    g.cost += r._cost; g.imps += r._imps; g.clicks += r._clicks;
    g.cv += r._cv; g.ebisCV += r._ebisCV; g.count++;
  });
  return Object.entries(groups).map(([key, g]) => ({
    label: key, ...g,
    ctr: g.imps > 0 ? g.clicks / g.imps * 100 : 0,
    cpa: g.cv > 0 ? g.cost / g.cv : null,
    ebCpa: g.ebisCV > 0 ? g.cost / g.ebisCV : null
  }));
}

function crossAggregate(data, rowKey, colKey) {
  const result = {};
  data.forEach(r => {
    const rk = r[rowKey] || '不明';
    const ck = r[colKey] || '不明';
    if (rk === '不明' || rk === '#REF!' || ck === '不明' || ck === '#REF!' || rk === '' || ck === '') return;
    if (!result[rk]) result[rk] = {};
    if (!result[rk][ck]) result[rk][ck] = { cost:0, cv:0, clicks:0, imps:0 };
    const g = result[rk][ck];
    g.cost += r._cost; g.cv += r._cv; g.clicks += r._clicks; g.imps += r._imps;
  });
  return result;
}

function fmt(n) { return n == null ? 'N/A' : '¥' + Math.round(n).toLocaleString(); }
function fmtN(n) { return n == null ? 'N/A' : Math.round(n).toLocaleString(); }
function fmtP(n) { return n == null ? 'N/A' : n.toFixed(2) + '%'; }

// ============================================
// Chart Helpers
// ============================================
function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

function createBarChart(id, labels, datasets, opts = {}) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      indexAxis: opts.horizontal ? 'y' : 'x',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: datasets.length > 1, labels: { color: '#a0a0b8', font: { size: 11 } } }, datalabels: { display: false } },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80', font: { size: 10 }, maxRotation: 45 } },
        y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80', font: { size: 10 } } }
      }
    }
  });
}

function createLineChart(id, labels, datasets) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  charts[id] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { intersect: false, mode: 'index' },
      plugins: { legend: { labels: { color: '#a0a0b8', font: { size: 11 }, usePointStyle: true } }, datalabels: { display: false } },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80' } },
        y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80', callback: v => '¥' + (v/1000).toFixed(0) + 'K' } }
      }
    }
  });
}

function createDoughnutChart(id, labels, data, colors) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 0 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: { legend: { position: 'right', labels: { color: '#a0a0b8', font: { size: 10 }, padding: 8, usePointStyle: true } }, datalabels: { display: false } }
    }
  });
}

// ============================================
// Overview Section
// ============================================
function renderOverview() {
  const d = filteredData;
  const totalCost = d.reduce((s,r) => s + r._cost, 0);
  const totalCV = d.reduce((s,r) => s + r._cv, 0);
  const totalImps = d.reduce((s,r) => s + r._imps, 0);
  const totalClicks = d.reduce((s,r) => s + r._clicks, 0);
  const cpa = totalCV > 0 ? totalCost / totalCV : null;
  const ctr = totalImps > 0 ? totalClicks / totalImps * 100 : 0;

  setKPI('kpi-cost', fmt(totalCost));
  setKPI('kpi-cv', fmtN(totalCV));
  setKPI('kpi-cpa', fmt(cpa));
  setKPI('kpi-ctr', fmtP(ctr));

  const monthly = aggregate(d, '_month');
  monthly.sort((a,b) => a.label.localeCompare(b.label));
  if (monthly.length) {
    createLineChart('chart-monthly-trend', monthly.map(m=>m.label), [{
      label: 'CPA', data: monthly.map(m=>m.cpa), borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.1)', fill: true, tension: 0.3, pointRadius: 4
    }]);
  }

  const lpData = aggregate(d, '_lp').sort((a,b) => b.cost - a.cost).slice(0, 10);
  if (lpData.length) createDoughnutChart('chart-lp-share', lpData.map(l => shortLabel(l.label)), lpData.map(l=>l.cost), COLORS);

  const btData = aggregate(d, '_bannerType').filter(b => b.cv >= 5).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (btData.length) createBarChart('chart-banner-type-cpa', btData.map(b=>b.label), [{ label:'CPA', data:btData.map(b=>b.cpa), backgroundColor:COLORS.slice(0,btData.length) }]);

  const tgData = aggregate(d, '_targeting').filter(t => t.cv >= 5).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (tgData.length) createBarChart('chart-targeting-efficiency', tgData.map(t=>t.label), [{ label:'CPA', data:tgData.map(t=>t.cpa), backgroundColor:'#6366f1' }]);

  renderInsights();
}

function setKPI(id, value) {
  const el = document.getElementById(id);
  if (el) el.querySelector('.kpi-value').textContent = value;
}

function switchOverviewChart(metric) {
  document.querySelectorAll('#section-overview .chart-btn').forEach(b => b.classList.toggle('active', b.dataset.metric === metric));
  const monthly = aggregate(filteredData, '_month').sort((a,b) => a.label.localeCompare(b.label));
  const config = { cpa:{data:monthly.map(m=>m.cpa),label:'CPA',color:'#6366f1'}, cv:{data:monthly.map(m=>m.cv),label:'CV',color:'#34d399'}, cost:{data:monthly.map(m=>m.cost),label:'費用',color:'#f97316'} };
  const c = config[metric];
  createLineChart('chart-monthly-trend', monthly.map(m=>m.label), [{ label:c.label, data:c.data, borderColor:c.color, backgroundColor:c.color+'1a', fill:true, tension:0.3, pointRadius:4 }]);
}

function renderInsights() {
  const container = document.getElementById('auto-insights');
  if (!container) return;
  const insights = [];
  
  const lpData = aggregate(filteredData, '_lp').filter(l => l.cv >= 10).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (lpData.length >= 2) {
    insights.push({ icon:'🏆', type:'success', title:'Best LP', text:`${shortLabel(lpData[0].label)} がCPA ${fmt(lpData[0].cpa)} で最効率（CV ${lpData[0].cv}件）` });
    const worst = lpData[lpData.length-1];
    insights.push({ icon:'⚠️', type:'warning', title:'要改善LP', text:`${shortLabel(worst.label)} がCPA ${fmt(worst.cpa)} で最も高い` });
  }
  
  const creatorData = aggregate(filteredData, '_creator').filter(c => c.cv >= 5 && c.label !== 'other').sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (creatorData.length) insights.push({ icon:'🎨', type:'', title:'Best クリエイター', text:`${creatorData[0].label} がCPA ${fmt(creatorData[0].cpa)} で最効率` });

  const btData = aggregate(filteredData, '_bannerType').filter(b => b.cv >= 10);
  if (btData.length >= 2) {
    btData.sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
    insights.push({ icon:'📐', type:'', title:'Best バナータイプ', text:`${btData[0].label} がCPA ${fmt(btData[0].cpa)} で最効率` });
  }

  container.innerHTML = insights.map(i => `
    <div class="insight-item ${i.type}">
      <div class="insight-icon">${i.icon}</div>
      <div class="insight-content"><h4>${i.title}</h4><p>${i.text}</p></div>
    </div>
  `).join('') || '<p style="color:var(--text-tertiary)">データが不足しています</p>';
}

// ============================================
// Creative Section
// ============================================
function renderCreative() {
  updateCreatorChart();
  
  const appealData = aggregate(filteredData, '_appeal').filter(a => a.cv >= 5).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (appealData.length) createBarChart('chart-appeal-cpa', appealData.map(a=>a.label), [{ label:'CPA', data:appealData.map(a=>a.cpa), backgroundColor:COLORS }], { horizontal:true });

  const fmtData = aggregate(filteredData, r => {
    const n = (r._adName || '').toLowerCase();
    if (n.includes('carousel')) return 'カルーセル';
    if (n.includes('1080_1920') || n.includes('1080x1920')) return 'ストーリーズ';
    if (n.includes('1620_1620')) return 'フィード(1:1大)';
    if (n.includes('1080_1080') || n.includes('1080x1080')) return 'フィード(1:1)';
    return 'その他';
  }).filter(f => f.cv >= 3).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (fmtData.length) createBarChart('chart-format-perf', fmtData.map(f=>f.label), [{ label:'CPA', data:fmtData.map(f=>f.cpa), backgroundColor:'#a855f7' }]);

  renderCreativeTable();
}

function updateCreatorChart() {
  const sort = document.getElementById('creator-sort')?.value || 'cost';
  let data = aggregate(filteredData, '_creator').filter(c => c.cv >= 2 && c.label !== 'other');
  const sortFn = { cost:(a,b)=>b.cost-a.cost, cpa:(a,b)=>(a.cpa||99999)-(b.cpa||99999), cv:(a,b)=>b.cv-a.cv, ctr:(a,b)=>b.ctr-a.ctr };
  data.sort(sortFn[sort] || sortFn.cost);
  data = data.slice(0, 20);
  if (data.length) createBarChart('chart-creator-perf', data.map(d=>d.label), [{ label:'CPA', data:data.map(d=>d.cpa), backgroundColor:COLORS.slice(0,data.length) }], { horizontal:true });
}

function renderCreativeTable() {
  const tbody = document.getElementById('creative-table-body');
  if (!tbody) return;
  
  const textAgg = {};
  filteredData.forEach(r => {
    const key = (r._adText || '').substring(0, 80) || r._adName || '(不明)';
    if (!textAgg[key]) textAgg[key] = { text:r._adText, lp:r._lp, tg:r._targeting, cost:0, cv:0, clicks:0, imps:0 };
    const g = textAgg[key];
    g.cost += r._cost; g.cv += r._cv; g.clicks += r._clicks; g.imps += r._imps;
  });

  let rows = Object.values(textAgg).filter(r => r.cv >= 1);
  rows.sort((a,b) => (a.cv > 0 ? a.cost/a.cv : 99999) - (b.cv > 0 ? b.cost/b.cv : 99999));
  
  const start = creativePage * PAGE_SIZE;
  const pageRows = rows.slice(start, start + PAGE_SIZE);
  
  tbody.innerHTML = pageRows.map((r, i) => {
    const cpa = r.cv > 0 ? r.cost / r.cv : null;
    const ctr = r.imps > 0 ? r.clicks / r.imps * 100 : 0;
    const cls = cpa ? (cpa < 20000 ? 'cpa-good' : cpa > 30000 ? 'cpa-bad' : 'cpa-ok') : '';
    return `<tr>
      <td class="rank-col">${start + i + 1}</td>
      <td title="${escHtml(r.text)}">${escHtml((r.text||'').substring(0, 60))}</td>
      <td>${shortLabel(r.lp)}</td><td>${r.tg}</td>
      <td class="num-col">${fmt(r.cost)}</td><td class="num-col">${r.cv}</td>
      <td class="num-col ${cls}">${fmt(cpa)}</td><td class="num-col">${fmtP(ctr)}</td>
    </tr>`;
  }).join('');

  const totalPages = Math.ceil(rows.length / PAGE_SIZE);
  const pgEl = document.getElementById('creative-pagination');
  if (pgEl && totalPages > 1) {
    pgEl.innerHTML = Array.from({length: Math.min(totalPages, 10)}, (_, i) => 
      `<button class="page-btn ${i === creativePage ? 'active' : ''}" onclick="creativePage=${i};renderCreativeTable()">${i+1}</button>`
    ).join('');
  }
}

function filterCreativeTable() { creativePage = 0; renderCreativeTable(); }
function sortCreativeTable() { creativePage = 0; renderCreativeTable(); }

// ============================================
// LP Section
// ============================================
function renderLP() {
  const lpData = aggregate(filteredData, '_lp').filter(l => l.cv >= 5).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
  if (lpData.length) createBarChart('chart-lp-comparison', lpData.map(l=>shortLabel(l.label)), [{ label:'CPA', data:lpData.map(l=>l.cpa), backgroundColor:COLORS }], { horizontal:true });

  const lps = lpData.slice(0, 8).map(l => l.label);
  const months = [...new Set(filteredData.map(r => r._month).filter(Boolean))].sort();
  if (months.length && lps.length) {
    const datasets = lps.map((lp, i) => {
      const monthly = months.map(m => {
        const rows = filteredData.filter(r => r._lp === lp && r._month === m);
        const cost = rows.reduce((s,r) => s + r._cost, 0);
        const cv = rows.reduce((s,r) => s + r._cv, 0);
        return cv > 0 ? cost / cv : null;
      });
      return { label: shortLabel(lp), data: monthly, borderColor: COLORS[i], backgroundColor: 'transparent', tension: 0.3, pointRadius: 3 };
    });
    createLineChart('chart-lp-monthly', months, datasets);
  }

  renderFatigueAlerts();

  const tbody = document.getElementById('lp-table-body');
  if (tbody) {
    tbody.innerHTML = lpData.map(l => `<tr>
      <td>${shortLabel(l.label)}</td>
      <td class="num-col">${fmt(l.cost)}</td><td class="num-col">${fmtN(l.imps)}</td>
      <td class="num-col">${fmtN(l.clicks)}</td><td class="num-col">${fmtP(l.ctr)}</td>
      <td class="num-col">${l.cv}</td>
      <td class="num-col ${l.cpa < 22000 ? 'cpa-good' : l.cpa > 28000 ? 'cpa-bad' : ''}">${fmt(l.cpa)}</td>
      <td class="num-col">${fmt(l.ebCpa)}</td>
    </tr>`).join('');
  }
}

function renderFatigueAlerts() {
  const el = document.getElementById('fatigue-list');
  if (!el) return;
  const months = [...new Set(filteredData.map(r => r._month).filter(Boolean))].sort();
  const lps = [...new Set(filteredData.map(r => r._lp).filter(l => l && l !== '不明'))];
  const alerts = [];
  lps.forEach(lp => {
    months.forEach((m, i) => {
      if (i === 0) return;
      const prev = months[i-1];
      const prevRows = filteredData.filter(r => r._lp === lp && r._month === prev);
      const currRows = filteredData.filter(r => r._lp === lp && r._month === m);
      const prevCost = prevRows.reduce((s,r) => s+r._cost, 0);
      const prevCV = prevRows.reduce((s,r) => s+r._cv, 0);
      const currCost = currRows.reduce((s,r) => s+r._cost, 0);
      const currCV = currRows.reduce((s,r) => s+r._cv, 0);
      if (prevCV < 10 || currCV < 10) return;
      const prevCPA = prevCost/prevCV, currCPA = currCost/currCV;
      const change = (currCPA - prevCPA) / prevCPA * 100;
      if (change > 20) alerts.push({ lp: shortLabel(lp), period: `${prev}→${m}`, change, prevCPA, currCPA });
    });
  });
  alerts.sort((a,b) => b.change - a.change);
  el.innerHTML = alerts.slice(0, 10).map(a => `
    <div class="alert-item"><span class="alert-badge danger">+${a.change.toFixed(0)}%</span>
    <span><strong>${a.lp}</strong> ${a.period} CPA ${fmt(a.prevCPA)} → ${fmt(a.currCPA)}</span></div>
  `).join('') || '<p style="color:var(--text-tertiary)">重大な疲弊パターンは検出されませんでした</p>';
}

function updateLpChart() { renderLP(); }

// ============================================
// Targeting Section
// ============================================
function renderTargeting() {
  const tgData = aggregate(filteredData, '_targeting').filter(t => t.cv >= 5).sort((a,b) => b.cost - a.cost);
  if (tgData.length) createBarChart('chart-tg-overview', tgData.map(t=>t.label), [{ label:'CPA', data:tgData.map(t=>t.cpa), backgroundColor:'#6366f1' }]);

  const intData = aggregate(filteredData, '_interest').filter(i => i.cv >= 5 && i.label.length > 2).sort((a,b) => (a.cpa||99999) - (b.cpa||99999)).slice(0, 15);
  if (intData.length) createBarChart('chart-interest-cpa', intData.map(i=>i.label.substring(0,25)), [{ label:'CPA', data:intData.map(i=>i.cpa), backgroundColor:COLORS }], { horizontal:true });

  const tbody = document.getElementById('tg-table-body');
  if (tbody) {
    tbody.innerHTML = tgData.map(t => `<tr>
      <td>${t.label}</td><td class="num-col">${fmt(t.cost)}</td><td class="num-col">${t.cv}</td>
      <td class="num-col">${fmt(t.cpa)}</td><td class="num-col">${fmtP(t.ctr)}</td>
      <td class="num-col">${fmtN(t.ebisCV)}</td><td class="num-col">${fmt(t.ebCpa)}</td>
    </tr>`).join('');
  }
}

// ============================================
// Cross Analysis (Heatmap)
// ============================================
function updateHeatmap() {
  const rowKey = '_' + document.getElementById('cross-row').value;
  const colKey = '_' + document.getElementById('cross-col').value;
  const metric = document.getElementById('cross-metric').value;
  const minCV = parseInt(document.getElementById('cross-min-cv').value) || 3;

  const cross = crossAggregate(filteredData, rowKey, colKey);
  const allCols = [...new Set(Object.values(cross).flatMap(c => Object.keys(c)))].sort();
  const rows = Object.entries(cross).filter(([_, cols]) => {
    const totalCV = Object.values(cols).reduce((s,c) => s + c.cv, 0);
    return totalCV >= minCV;
  }).sort((a,b) => {
    const aTotal = Object.values(a[1]).reduce((s,c) => s+c.cost, 0);
    const bTotal = Object.values(b[1]).reduce((s,c) => s+c.cost, 0);
    return bTotal - aTotal;
  }).slice(0, 20);

  if (!rows.length) {
    document.getElementById('heatmap-container').innerHTML = '<div class="heatmap-placeholder">該当データがありません</div>';
    return;
  }

  let allVals = [];
  rows.forEach(([_, cols]) => {
    allCols.forEach(c => {
      if (cols[c] && cols[c].cv >= minCV) {
        const v = getMetricValue(cols[c], metric);
        if (v !== null) allVals.push(v);
      }
    });
  });
  const minVal = Math.min(...allVals);
  const maxVal = Math.max(...allVals);

  let html = '<table class="heatmap-table"><thead><tr><th></th>';
  allCols.forEach(c => { html += `<th>${shortLabel(c)}</th>`; });
  html += '</tr></thead><tbody>';
  rows.forEach(([rowLabel, cols]) => {
    html += `<tr><td class="row-header" title="${rowLabel}">${shortLabel(rowLabel)}</td>`;
    allCols.forEach(c => {
      if (cols[c] && cols[c].cv >= minCV) {
        const v = getMetricValue(cols[c], metric);
        const color = heatmapColor(v, minVal, maxVal, metric);
        const display = metric === 'cpa' ? fmt(v) : metric === 'ctr' ? fmtP(v) : fmtN(v);
        html += `<td class="heatmap-cell" style="background:${color}">${display}<span class="heatmap-sub">CV ${cols[c].cv}</span></td>`;
      } else { html += '<td style="color:var(--text-tertiary)">-</td>'; }
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  document.getElementById('heatmap-container').innerHTML = html;
}

function getMetricValue(g, metric) {
  if (metric === 'cpa') return g.cv > 0 ? g.cost / g.cv : null;
  if (metric === 'cv') return g.cv;
  if (metric === 'cost') return g.cost;
  if (metric === 'ctr') return g.imps > 0 ? g.clicks / g.imps * 100 : 0;
  return null;
}

function heatmapColor(val, min, max, metric) {
  if (val === null) return 'transparent';
  const ratio = max !== min ? (val - min) / (max - min) : 0.5;
  const inverted = metric === 'cpa';
  const r = inverted ? ratio : 1 - ratio;
  if (r < 0.5) { const t = r * 2; return `rgba(52,211,153,${0.1 + t * 0.25})`; }
  else { const t = (r - 0.5) * 2; return `rgba(248,113,113,${0.1 + t * 0.25})`; }
}

// ============================================
// Trend Section
// ============================================
function renderTrend() {
  const months = [...new Set(filteredData.map(r => r._month).filter(Boolean))].sort();
  const winnerData = months.map(m => {
    const lpAgg = aggregate(filteredData.filter(r => r._month === m), '_lp').filter(l => l.cv >= 10);
    lpAgg.sort((a,b) => (a.cpa||99999) - (b.cpa||99999));
    return { month: m, best: lpAgg[0], worst: lpAgg[lpAgg.length - 1] };
  }).filter(w => w.best);

  if (winnerData.length) {
    createBarChart('chart-monthly-winner', winnerData.map(w=>w.month), [
      { label:'Best CPA', data:winnerData.map(w=>w.best?.cpa), backgroundColor:'#34d399' },
      { label:'Worst CPA', data:winnerData.map(w=>w.worst?.cpa), backgroundColor:'#f87171' }
    ]);
  }

  renderConcentrationChart(months);

  const el = document.getElementById('trend-insights');
  if (el) {
    el.innerHTML = winnerData.map(w => `<div class="alert-item"><span>${w.month}: 🏆 ${shortLabel(w.best.label)} (CPA ${fmt(w.best.cpa)})</span></div>`).join('');
  }
}

function renderConcentrationChart(months) {
  const points = [];
  const lps = [...new Set(filteredData.map(r => r._lp).filter(l => l && l !== '不明'))];
  lps.forEach(lp => {
    for (let i = 1; i < months.length; i++) {
      const prevMonth = months[i-1], currMonth = months[i];
      const allPrev = filteredData.filter(r => r._month === prevMonth);
      const lpPrev = allPrev.filter(r => r._lp === lp);
      const lpCurr = filteredData.filter(r => r._lp === lp && r._month === currMonth);
      const totalPrevCost = allPrev.reduce((s,r) => s+r._cost, 0);
      const lpPrevCost = lpPrev.reduce((s,r) => s+r._cost, 0);
      const prevCV = lpPrev.reduce((s,r) => s+r._cv, 0);
      const currCV = lpCurr.reduce((s,r) => s+r._cv, 0);
      if (prevCV < 10 || currCV < 10 || totalPrevCost === 0) continue;
      const share = lpPrevCost / totalPrevCost * 100;
      const prevCPA = lpPrevCost / prevCV;
      const currCost = lpCurr.reduce((s,r) => s+r._cost, 0);
      const currCPA = currCost / currCV;
      points.push({ x: share, y: (currCPA - prevCPA) / prevCPA * 100 });
    }
  });

  destroyChart('chart-concentration-impact');
  const ctx = document.getElementById('chart-concentration-impact');
  if (!ctx || !points.length) return;
  charts['chart-concentration-impact'] = new Chart(ctx, {
    type: 'scatter',
    data: { datasets: [{ label:'LP月次', data: points, backgroundColor: points.map(p => p.y > 0 ? 'rgba(248,113,113,0.6)' : 'rgba(52,211,153,0.6)'), pointRadius: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, datalabels: { display: false } },
      scales: {
        x: { title: { display: true, text: '前月構成比 (%)', color: '#a0a0b8' }, grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80' } },
        y: { title: { display: true, text: '翌月CPA変動 (%)', color: '#a0a0b8' }, grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#6a6a80' } }
      }
    }
  });
}

// ============================================
// Text Analysis
// ============================================
function renderText() {
  const keywords = ['今さら','バタバタ','パワー','電池切れ','半信半疑','プレゼント','ミニ','2,700万袋','すっぽんってそんなに','限界ママ','元気','ロングセラー','朝','本音','ぐったり','調子','キャンペーン','疑'];
  const kwData = keywords.map(kw => {
    const rows = filteredData.filter(r => (r._adText||'').includes(kw));
    const cost = rows.reduce((s,r) => s+r._cost, 0);
    const cv = rows.reduce((s,r) => s+r._cv, 0);
    return { label: kw, cost, cv, cpa: cv > 0 ? cost/cv : null };
  }).filter(k => k.cv >= 3).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));

  if (kwData.length) createBarChart('chart-keyword-cpa', kwData.map(k=>k.label), [{
    label: 'CPA', data: kwData.map(k=>k.cpa),
    backgroundColor: kwData.map(k => k.cpa < 22000 ? '#34d399' : k.cpa > 26000 ? '#f87171' : '#fbbf24')
  }]);

  const goodEl = document.getElementById('good-keywords');
  const badEl = document.getElementById('bad-keywords');
  if (goodEl) goodEl.innerHTML = kwData.filter(k => k.cpa && k.cpa < 22000).map(k => `<div class="keyword-item"><span class="keyword-name">「${k.label}」</span><span class="keyword-badge good">${fmt(k.cpa)}</span></div>`).join('') || '<p style="color:var(--text-tertiary)">該当なし</p>';
  if (badEl) badEl.innerHTML = kwData.filter(k => k.cpa && k.cpa > 26000).map(k => `<div class="keyword-item"><span class="keyword-name">「${k.label}」</span><span class="keyword-badge bad">${fmt(k.cpa)}</span></div>`).join('') || '<p style="color:var(--text-tertiary)">該当なし</p>';
}

function filterTextTable() {}

// ============================================
// Demo Data
// ============================================
function loadDemoData() {
  showLoading('デモデータを生成中...');
  const lps = ['記事LP012(疑いLPver2)','記事LP001(InstagramLP)','記事LP035(Season3LP)','記事LP038(矢部LP)','記事LP032(朝InstagramLP)','記事LP027(オトナチャージLP)','記事LP026(ママ漫画LP)'];
  const targets = ['興味関心','ASC','ブロード','ATG','RTG'];
  const appeals = ['疑い','元気','悩み（元気がない）','レビュー・実体験','栄養素','キャンペーン'];
  const types = ['マンガ','実写','イラスト'];
  const months = ['2025-07','2025-08','2025-09','2025-10','2025-11','2025-12','2026-01','2026-02'];
  const creators = ['inui','omochico','manga_carousel','みたん','新庄アキラ','まりまり','ミロチ','おっちゃん'];
  const texts = ['選ばれ続けて19年のロングセラーサプリ。ココロとカラダの元気補給。バタバタの日々は変わらなくとも、私にパワーをくれるもの。','【PR】今さらなんて言わないで！！すっぽん小町始めました✨','全国の限界ママが大絶賛✨ 2,700万袋売れているサプリ！','「実際どうなの？」本音体験レポ 半信半疑で始めてみたら…','電池切れ寸前のママへ。毎日がんばるあなたに、ミニサイズプレゼント中！'];

  rawData = [];
  for (let i = 0; i < 5000; i++) {
    const cost = Math.random() * 50000 + 5000;
    const imps = Math.floor(cost / (Math.random()*2+1));
    const clicks = Math.floor(imps * (Math.random()*0.012+0.003));
    const cv = Math.random() > 0.6 ? Math.floor(cost / (Math.random()*20000+15000)) : 0;
    const mo = months[Math.floor(Math.random()*months.length)];
    const cr = creators[Math.floor(Math.random()*creators.length)];
    rawData.push({
      '日': mo + '-' + String(Math.floor(Math.random()*28)+1).padStart(2,'0'),
      '費用': cost, 'Imps': imps, 'Clicks': clicks,
      'CV (CVs)': cv,
      'EBiSCV (CVs)': cv,  // EBiSCV = 正のCV値（実際データに合わせる）
      'LP名': lps[Math.floor(Math.random()*lps.length)],
      'ターゲティング': targets[Math.floor(Math.random()*targets.length)],
      'バナー訴求': appeals[Math.floor(Math.random()*appeals.length)],
      'バナータイプ': types[Math.floor(Math.random()*types.length)],
      '広告名': `${mo}_k_${cr}_ad${i}`,
      'テキスト 1': texts[Math.floor(Math.random()*texts.length)],
      '運用': '自社運用',
      'EBiS遷移数 (CVs)': Math.floor(cv * (Math.random()*5+2)),
      '興味関心カテゴリ': ['子育てAND楽天','百貨店ANDアンチエイジング','子育てANDレゴORトミカ',''][Math.floor(Math.random()*4)]
    });
  }
  processData();
}

// ============================================
// Utility
// ============================================
function shortLabel(s) { return (!s) ? '' : s.length > 18 ? s.substring(0, 18) + '…' : s; }
function escHtml(s) { return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function showLoading(msg) {
  const el = document.getElementById('loading-overlay');
  el.classList.remove('hidden');
  const sub = document.getElementById('loading-status');
  if (sub) sub.textContent = msg || '';
}
function hideLoading() { document.getElementById('loading-overlay').classList.add('hidden'); }

// ============================================
// CSV Export - 集計過程の可視化
// 広告主がダッシュボードの数値を裏取りできるよう、
// 「どう絞り込み、どう集計したか」を明記してCSV出力
// ============================================

// CSVセーフな値をエスケープ
function csvVal(v) {
  if (v === null || v === undefined) return '';
  v = String(v);
  if (v.includes(',') || v.includes('"') || v.includes('\n')) {
    return '"' + v.replace(/"/g, '""') + '"';
  }
  return v;
}

// メタデータ付きTSVを生成してクリップボードにコピー
function downloadCSVWithMeta(meta, dataRows, filename) {
  const lines = [];

  // === セクション1: 集計条件 ===
  if (meta.length > 0) {
    meta.forEach(m => {
      lines.push((m[0] || '') + '\t' + (m[1] || ''));
    });
    lines.push('');
  }

  // === セクション2: データ ===
  if (dataRows.length > 0) {
    const headers = Object.keys(dataRows[0]);
    lines.push(headers.join('\t'));
    dataRows.forEach(row => {
      lines.push(headers.map(h => {
        let v = row[h];
        if (v === null || v === undefined) return '';
        return String(v);
      }).join('\t'));
    });
  }

  const tsvContent = lines.join('\n');
  copyToClipboard(tsvContent, filename);
}

// クリップボードにコピー（全セクション共通）
async function copyToClipboard(text, label) {
  try {
    await navigator.clipboard.writeText(text);
    showClipStatus(label + ' をコピーしました。スプシに貼り付けてください。');
    console.log('Copied to clipboard: ' + label);
  } catch(e) {
    // フォールバック: execCommand
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try {
      document.execCommand('copy');
      showClipStatus(label + ' をコピーしました。スプシに貼り付けてください。');
      console.log('Copied to clipboard (execCommand): ' + label);
    } catch(e2) {
      // 最終手段: 新しいウィンドウで表示
      const w = window.open('', '_blank');
      w.document.write('<pre>' + text.replace(/</g, '&lt;') + '</pre>');
      w.document.title = label;
      showClipStatus('クリップボードに失敗。新しいタブにデータを表示しました。');
    }
    document.body.removeChild(ta);
  }
}

function showClipStatus(msg) {
  // combo section status
  const el = document.getElementById('clip-status');
  if (el) { el.textContent = msg; setTimeout(() => { el.textContent = ''; }, 5000); }
  // also show a toast-like notification at the top
  const toast = document.createElement('div');
  toast.textContent = msg;
  toast.style.cssText = 'position:fixed;top:16px;left:50%;transform:translateX(-50%);background:#6366f1;color:white;padding:10px 24px;border-radius:8px;z-index:9999;font-size:14px;box-shadow:0 4px 20px rgba(0,0,0,0.3);white-space:nowrap';
  document.body.appendChild(toast);
  setTimeout(() => { toast.style.opacity = '0'; toast.style.transition = 'opacity 0.5s'; }, 2500);
  setTimeout(() => { document.body.removeChild(toast); }, 3200);
}

// 旧CSVダウンロードのラッパー（既存セクションもクリップボードに変更）
function downloadRawCSV(csvContent, filename) {
  // CSVをTSVに変換してクリップボードにコピー
  const tsvContent = csvContent.replace(/^\uFEFF/, '').split('\n').map(line => {
    // Simple CSV to TSV: replace commas not inside quotes
    return line.replace(/,/g, '\t');
  }).join('\n');
  copyToClipboard(tsvContent, filename);
}

// 現在のフィルタ条件をメタデータ配列で返す
function getFilterMeta() {
  const m = document.getElementById('filter-month').value;
  const o = document.getElementById('filter-operation').value;
  const meta = [
    ['出力日時', new Date().toLocaleString('ja-JP')],
    ['元データ行数(全体)', rawData.length.toLocaleString() + '行'],
    ['フィルタ後行数', filteredData.length.toLocaleString() + '行'],
    ['期間フィルタ', m === 'all' ? '全期間' : m],
    ['運用フィルタ', o === 'all' ? '全運用' : o],
    ['CVソース', 'EBiSCV (CVs) 列を使用（EBiS計測値）'],
    ['費用ソース', '費用 列を使用'],
  ];
  return meta;
}

function getFilterLabel() {
  const m = document.getElementById('filter-month').value;
  const o = document.getElementById('filter-operation').value;
  let label = '';
  if (m !== 'all') label += m + '_';
  if (o !== 'all') label += o + '_';
  return label;
}

// --------------------------------------------------
// セクション別エクスポート
// --------------------------------------------------

function exportOverview() {
  const d = filteredData;
  const totalCost = d.reduce((s,r) => s + r._cost, 0);
  const totalCV = d.reduce((s,r) => s + r._cv, 0);
  const totalImps = d.reduce((s,r) => s + r._imps, 0);
  const totalClicks = d.reduce((s,r) => s + r._clicks, 0);

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- KPI計算ロジック ---', ''],
    ['広告費用', 'SUM(費用列) = ¥' + Math.round(totalCost).toLocaleString()],
    ['CV数', 'SUM(EBiSCV(CVs)列) = ' + totalCV.toLocaleString()],
    ['CPA', '広告費用 ÷ CV数 = ¥' + (totalCV > 0 ? Math.round(totalCost/totalCV).toLocaleString() : 'N/A')],
    ['CTR', 'SUM(Clicks) ÷ SUM(Imps) × 100 = ' + (totalImps > 0 ? (totalClicks/totalImps*100).toFixed(2) + '%' : 'N/A')],
    ['', ''],
    ['--- 以下はフィルタ後の全行ローデータ ---', ''],
    ['用途', 'スプレッドシートでSUMIF等を使い上記KPIを再計算して検算してください'],
  ];

  const rows = d.map(r => ({
    '日付': r[findColCached('日')] || '',
    'LP名': r._lp, 'ターゲティング': r._targeting, 'バナー訴求': r._appeal,
    'バナータイプ': r._bannerType, '運用': r._operation,
    '費用': r._cost, 'Imps': r._imps, 'Clicks': r._clicks,
    'EBiSCV': r._cv, 'EBiS遷移数': r._ebisCV,
    'CTR(%)': r._ctr.toFixed(4), 'CPA': r._cpa != null ? Math.round(r._cpa) : ''
  }));
  downloadCSVWithMeta(meta, rows, `overview_verify_${getFilterLabel()}${d.length}rows.csv`);
}

function exportLP() {
  const minCV = 5;
  const lpAgg = aggregate(filteredData, '_lp').filter(l => l.cv >= minCV).sort((a,b) => (a.cpa||99999) - (b.cpa||99999));

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['グルーピングキー', 'LP名 列'],
    ['費用', 'グループ内の「費用」列をSUM'],
    ['Imps', 'グループ内の「Imps」列をSUM'],
    ['Clicks', 'グループ内の「Clicks」列をSUM'],
    ['EBiSCV', 'グループ内の「EBiSCV (CVs)」列をSUM'],
    ['CPA計算式', '費用 ÷ EBiSCV'],
    ['CTR計算式', 'Clicks ÷ Imps × 100'],
    ['遷移CPA計算式', '費用 ÷ EBiS遷移数(CVs)'],
    ['足切り条件', 'EBiSCV >= ' + minCV + '件 のグループのみ表示'],
    ['ソート順', 'CPA 昇順（低い=効率良い順）'],
    ['データ行数', '各グループに何行のローデータが含まれるかをCOUNT'],
    ['', ''],
    ['--- 検算方法 ---', ''],
    ['手順', '元CSVを開き、LP名列でフィルタ → 費用列をSUM → EBiSCV(CVs)列をSUM → 割り算でCPAを検算'],
  ];

  const rows = lpAgg.map(l => ({
    'LP名': l.label,
    '費用_SUM': Math.round(l.cost),
    'Imps_SUM': l.imps,
    'Clicks_SUM': l.clicks,
    'CTR(Clicks÷Imps×100)': l.ctr.toFixed(4),
    'EBiSCV_SUM': l.cv,
    'CPA(費用÷EBiSCV)': l.cpa ? Math.round(l.cpa) : '',
    'EBiS遷移数_SUM': l.ebisCV,
    '遷移CPA(費用÷遷移数)': l.ebCpa ? Math.round(l.ebCpa) : '',
    'ローデータ行数_COUNT': l.count
  }));
  downloadCSVWithMeta(meta, rows, `lp_analysis_verify_${getFilterLabel()}.csv`);
}

function exportCreative() {
  const minCV = 2;
  const creatorAgg = aggregate(filteredData, '_creator').filter(c => c.cv >= minCV && c.label !== 'other').sort((a,b) => b.cost - a.cost);

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['グルーピングキー', '広告名からクリエイター名を正規表現で抽出（推定）'],
    ['抽出ルール例', '広告名に「おっちゃん」→おっちゃん、「mirochi」→ミロチ、等'],
    ['費用', 'グループ内の「費用」列をSUM'],
    ['EBiSCV', 'グループ内の「EBiSCV (CVs)」列をSUM'],
    ['CPA計算式', '費用 ÷ EBiSCV'],
    ['足切り条件', 'EBiSCV >= ' + minCV + '件、かつ「other」カテゴリは除外'],
    ['ソート順', '費用 降順'],
  ];

  const rows = creatorAgg.map(c => ({
    'クリエイター(推定)': c.label,
    '費用_SUM': Math.round(c.cost),
    'Imps_SUM': c.imps,
    'Clicks_SUM': c.clicks,
    'CTR(Clicks÷Imps×100)': c.ctr.toFixed(4),
    'EBiSCV_SUM': c.cv,
    'CPA(費用÷EBiSCV)': c.cpa ? Math.round(c.cpa) : '',
    'ローデータ行数_COUNT': c.count
  }));
  downloadCSVWithMeta(meta, rows, `creative_verify_${getFilterLabel()}.csv`);
}

function exportTargeting() {
  const minCV = 5;
  const tgAgg = aggregate(filteredData, '_targeting').filter(t => t.cv >= minCV).sort((a,b) => b.cost - a.cost);

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['グルーピングキー', 'ターゲティング 列'],
    ['費用', 'グループ内の「費用」列をSUM'],
    ['EBiSCV', 'グループ内の「EBiSCV (CVs)」列をSUM'],
    ['CPA計算式', '費用 ÷ EBiSCV'],
    ['足切り条件', 'EBiSCV >= ' + minCV + '件'],
    ['ソート順', '費用 降順'],
    ['', ''],
    ['--- 検算方法 ---', ''],
    ['手順', '元CSVの「ターゲティング」列でフィルタ → 費用/EBiSCV各列をSUM → CPA計算'],
  ];

  const rows = tgAgg.map(t => ({
    'ターゲティング': t.label,
    '費用_SUM': Math.round(t.cost),
    'Imps_SUM': t.imps,
    'Clicks_SUM': t.clicks,
    'CTR(Clicks÷Imps×100)': t.ctr.toFixed(4),
    'EBiSCV_SUM': t.cv,
    'CPA(費用÷EBiSCV)': t.cpa ? Math.round(t.cpa) : '',
    'EBiS遷移数_SUM': t.ebisCV,
    '遷移CPA(費用÷遷移数)': t.ebCpa ? Math.round(t.ebCpa) : '',
    'ローデータ行数_COUNT': t.count
  }));
  downloadCSVWithMeta(meta, rows, `targeting_verify_${getFilterLabel()}.csv`);
}

function exportCross() {
  const rowKey = '_' + document.getElementById('cross-row').value;
  const colKey = '_' + document.getElementById('cross-col').value;
  const rowLabel = document.getElementById('cross-row').selectedOptions[0].text;
  const colLabel = document.getElementById('cross-col').selectedOptions[0].text;
  const minCV = parseInt(document.getElementById('cross-min-cv').value) || 5;

  const cross = crossAggregate(filteredData, rowKey, colKey);
  const summaryRows = [];
  Object.entries(cross).forEach(([rk, cols]) => {
    Object.entries(cols).forEach(([ck, g]) => {
      summaryRows.push({
        [rowLabel]: rk, [colLabel]: ck,
        '費用_SUM': Math.round(g.cost), 'Imps_SUM': g.imps, 'Clicks_SUM': g.clicks,
        'EBiSCV_SUM': g.cv,
        'CPA(費用÷EBiSCV)': g.cv > 0 ? Math.round(g.cost / g.cv) : '',
        'CTR(Clicks÷Imps×100)': g.imps > 0 ? (g.clicks / g.imps * 100).toFixed(4) : ''
      });
    });
  });
  summaryRows.sort((a,b) => (b['費用_SUM'] || 0) - (a['費用_SUM'] || 0));

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['行軸（グルーピングキー1）', rowLabel + ' 列'],
    ['列軸（グルーピングキー2）', colLabel + ' 列'],
    ['集計方法', '行軸×列軸の組み合わせごとに、費用・Imps・Clicks・EBiSCVをそれぞれSUM'],
    ['CPA計算式', '費用_SUM ÷ EBiSCV_SUM'],
    ['最低CV足切り', 'ヒートマップ表示時は EBiSCV >= ' + minCV + '件 のセルのみ色表示'],
    ['ソート順', '費用 降順'],
    ['', ''],
    ['--- 検算方法 ---', ''],
    ['手順', '元CSVを「' + rowLabel + '」列と「' + colLabel + '」列の2軸でピボット → 費用/EBiSCVをSUM → CPA計算'],
  ];
  downloadCSVWithMeta(meta, summaryRows, `cross_analysis_verify_${getFilterLabel()}.csv`);
}

function exportTrend() {
  const months = [...new Set(filteredData.map(r => r._month).filter(Boolean))].sort();
  const lps = [...new Set(filteredData.map(r => r._lp).filter(l => l && l !== '不明'))];

  const rows = [];
  lps.forEach(lp => {
    months.forEach(m => {
      const mRows = filteredData.filter(r => r._lp === lp && r._month === m);
      if (!mRows.length) return;
      const cost = mRows.reduce((s,r) => s + r._cost, 0);
      const cv = mRows.reduce((s,r) => s + r._cv, 0);
      const imps = mRows.reduce((s,r) => s + r._imps, 0);
      const clicks = mRows.reduce((s,r) => s + r._clicks, 0);
      rows.push({
        '月': m, 'LP名': lp,
        '費用_SUM': Math.round(cost), 'Imps_SUM': imps, 'Clicks_SUM': clicks,
        'EBiSCV_SUM': cv,
        'CPA(費用÷EBiSCV)': cv > 0 ? Math.round(cost / cv) : '',
        'CTR(Clicks÷Imps×100)': imps > 0 ? (clicks / imps * 100).toFixed(4) : '',
        'ローデータ行数_COUNT': mRows.length
      });
    });
  });
  rows.sort((a,b) => a['月'].localeCompare(b['月']) || (b['費用_SUM'] || 0) - (a['費用_SUM'] || 0));

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['グルーピングキー', '月（日付列からyyyy-MM抽出） × LP名'],
    ['月の抽出方法', '「日」列をyyyy/MM/ddまたはyyyy-MM-ddとしてパースし、yyyy-MMに変換'],
    ['費用', 'グループ内の「費用」列をSUM'],
    ['EBiSCV', 'グループ内の「EBiSCV (CVs)」列をSUM'],
    ['CPA計算式', '費用 ÷ EBiSCV'],
    ['ソート順', '月 昇順 → 費用 降順'],
    ['', ''],
    ['--- 検算方法 ---', ''],
    ['手順', '元CSVの日付列から月を抽出（YEAR&MONTH関数等） → LP名×月でピボット → 費用/EBiSCV SUM → CPA'],
  ];
  downloadCSVWithMeta(meta, rows, `trend_lp_monthly_verify_${getFilterLabel()}.csv`);
}

function exportText() {
  const keywords = ['今さら','バタバタ','パワー','電池切れ','半信半疑','プレゼント','ミニ','2,700万袋','すっぽんってそんなに','限界ママ','元気','ロングセラー','朝','本音','ぐったり','調子','キャンペーン','疑'];

  const summaryRows = keywords.map(kw => {
    const matched = filteredData.filter(r => (r._adText || '').includes(kw));
    const cost = matched.reduce((s,r) => s + r._cost, 0);
    const cv = matched.reduce((s,r) => s + r._cv, 0);
    return {
      'キーワード': kw,
      '該当行数_COUNT': matched.length,
      '費用_SUM': Math.round(cost),
      'EBiSCV_SUM': cv,
      'CPA(費用÷EBiSCV)': cv > 0 ? Math.round(cost / cv) : ''
    };
  }).filter(r => r['該当行数_COUNT'] > 0);

  const meta = [
    ...getFilterMeta(),
    ['', ''],
    ['--- 集計ロジック ---', ''],
    ['対象列', '「テキスト 1」列（広告のプライマリテキスト）'],
    ['マッチ方法', '各キーワードが「テキスト 1」列に部分一致(includes)する行を抽出'],
    ['費用', 'マッチした行の「費用」列をSUM'],
    ['EBiSCV', 'マッチした行の「EBiSCV (CVs)」列をSUM'],
    ['CPA計算式', '費用 ÷ EBiSCV'],
    ['注意', '1行が複数キーワードにマッチする場合、それぞれのキーワード集計に重複加算されます'],
    ['', ''],
    ['--- 検算方法 ---', ''],
    ['手順', '元CSVの「テキスト 1」列で当該キーワードをCOUNTIF → 該当行の費用/EBiSCVをSUMIFS → CPA'],
  ];
  downloadCSVWithMeta(meta, summaryRows, `text_keyword_verify_${getFilterLabel()}.csv`);
}

// カラム名キャッシュ
let _colCache = {};
function findColCached(name) {
  if (_colCache[name]) return _colCache[name];
  if (rawData.length > 0) {
    const cols = Object.keys(rawData[0]);
    const found = cols.find(c => c === name) || cols.find(c => c.includes(name));
    if (found) { _colCache[name] = found; return found; }
  }
  return name;
}

// セクション別エクスポート dispatcher
function exportCurrentSection() {
  if (!filteredData.length) { alert('データがありません'); return; }
  switch(currentSection) {
    case 'overview': exportOverview(); break;
    case 'creative': exportCreative(); break;
    case 'lp': exportLP(); break;
    case 'targeting': exportTargeting(); break;
    case 'cross': exportCross(); break;
    case 'trend': exportTrend(); break;
    case 'text': exportText(); break;
    case 'combo': clipComboMatrix(); break;
  }
}

// ============================================
// Combo Analysis (Creative × LP Effect Separation)
// ============================================

function renderCombo() {
  const minCV = parseInt(document.getElementById('combo-min-cv')?.value) || 5;
  const d = filteredData;

  // Step 1: Build combo aggregation: adName × LP × TG
  const comboMap = {};
  d.forEach(r => {
    const key = r._adName + '|||' + r._lp;
    if (!comboMap[key]) comboMap[key] = { adName: r._adName, lp: r._lp, tg: r._targeting, appeal: r._appeal, bType: r._bannerType, cost:0, cv:0, clicks:0, imps:0 };
    const g = comboMap[key];
    g.cost += r._cost; g.cv += r._cv; g.clicks += r._clicks; g.imps += r._imps;
  });
  const combos = Object.values(comboMap).filter(c => c.cv >= 2);

  // Step 2: LP baseline CPA
  const lpMap = {};
  combos.forEach(c => {
    if (!lpMap[c.lp]) lpMap[c.lp] = { cost:0, cv:0, count:0 };
    lpMap[c.lp].cost += c.cost; lpMap[c.lp].cv += c.cv; lpMap[c.lp].count++;
  });
  const lpBaseline = Object.entries(lpMap)
    .filter(([_,g]) => g.cv >= minCV)
    .map(([lp, g]) => ({ lp, cost: g.cost, cv: g.cv, cpa: g.cost/g.cv, count: g.count }))
    .sort((a,b) => b.cost - a.cost);

  const lpCpaMap = {};
  lpBaseline.forEach(l => { lpCpaMap[l.lp] = l.cpa; });

  // Render LP baseline table
  const blBody = document.getElementById('combo-lp-baseline-body');
  if (blBody) {
    blBody.innerHTML = lpBaseline.slice(0, 15).map(l => `<tr>
      <td>${shortLabel(l.lp)}</td>
      <td class="num-col">${fmt(l.cost)}</td><td class="num-col">${l.cv}</td>
      <td class="num-col">${fmt(l.cpa)}</td><td class="num-col">${l.count}</td>
    </tr>`).join('');
  }

  // Step 3: Appeal × LP compatibility matrix
  renderComboMatrix('combo-appeal-matrix', combos, '_appeal', lpCpaMap, minCV);
  renderComboMatrix('combo-type-matrix', combos, '_bType', lpCpaMap, minCV);

  // Step 4: Creative scoring (LP effect removed)
  // Group by adName, compute deviation from each LP's average
  const crMap = {};
  combos.forEach(c => {
    const lpBase = lpCpaMap[c.lp];
    if (!lpBase || c.cv < 2) return;
    const cpa = c.cost / c.cv;
    const dev = (cpa / lpBase - 1) * 100;
    if (!crMap[c.adName]) crMap[c.adName] = { adName: c.adName, appeal: c.appeal, bType: c.bType, cost:0, cv:0, lps: new Set(), devs: [], devWeights: [] };
    const cr = crMap[c.adName];
    cr.cost += c.cost; cr.cv += c.cv; cr.lps.add(c.lp);
    cr.devs.push(dev); cr.devWeights.push(c.cv);
  });

  const crScores = Object.values(crMap)
    .filter(c => c.lps.size >= 2)
    .map(c => {
      const totalW = c.devWeights.reduce((s,w)=>s+w, 0);
      const wDev = c.devWeights.reduce((s,w,i) => s + w * c.devs[i], 0) / (totalW || 1);
      return { ...c, lpCount: c.lps.size, cpa: c.cost/c.cv, wDev };
    })
    .sort((a,b) => a.wDev - b.wDev);

  const scoreBody = document.getElementById('combo-score-body');
  if (scoreBody) {
    const show = [...crScores.slice(0, 15), ...crScores.slice(-10).reverse()];
    scoreBody.innerHTML = show.map((c, i) => {
      const rank = i < 15 ? (i+1) : '—';
      const cls = c.wDev < -10 ? 'cpa-good' : c.wDev > 10 ? 'cpa-bad' : '';
      const badge = c.wDev < -10 ? '🟢' : c.wDev > 10 ? '🔴' : '';
      const adShort = c.adName.length > 40 ? '…' + c.adName.slice(-38) : c.adName;
      return `<tr>
        <td class="rank-col">${rank}</td>
        <td title="${escHtml(c.adName)}">${escHtml(adShort)}</td>
        <td>${c.appeal || '-'}</td><td>${c.bType || '-'}</td>
        <td class="num-col">${fmt(c.cost)}</td><td class="num-col">${c.cv}</td>
        <td class="num-col">${fmt(c.cpa)}</td>
        <td class="num-col">${c.lpCount}</td>
        <td class="num-col ${cls}">${badge} ${c.wDev > 0 ? '+' : ''}${c.wDev.toFixed(0)}%</td>
      </tr>`;
    }).join('');
    if (crScores.length > 25) {
      scoreBody.innerHTML = scoreBody.innerHTML.replace('</tr><tr>\n        <td class="rank-col">—', '</tr><tr style="background:rgba(255,255,255,0.02)"><td colspan="9" style="text-align:center;color:var(--text-tertiary);padding:8px">… 中間 ' + (crScores.length - 25) + '件省略 …</td></tr><tr>\n        <td class="rank-col">—');
    }
  }

  // Store for export
  window._comboData = { lpBaseline, combos, crScores, lpCpaMap, minCV };
}

function renderComboMatrix(containerId, combos, attrKey, lpCpaMap, minCV) {
  const el = document.getElementById(containerId);
  if (!el) return;

  // Remap attrKey from internal names
  const getAttr = c => attrKey === '_appeal' ? c.appeal : attrKey === '_bType' ? c.bType : c.appeal;

  // Group by LP × attr
  const matrix = {};
  combos.forEach(c => {
    const attr = getAttr(c);
    const lp = c.lp;
    if (!attr || attr === '不明' || !lpCpaMap[lp]) return;
    if (!matrix[lp]) matrix[lp] = {};
    if (!matrix[lp][attr]) matrix[lp][attr] = { cost:0, cv:0 };
    matrix[lp][attr].cost += c.cost; matrix[lp][attr].cv += c.cv;
  });

  const allAttrs = [...new Set(combos.map(c => getAttr(c)).filter(a => a && a !== '不明'))]
    .filter(attr => {
      const totalCV = combos.filter(c => getAttr(c) === attr).reduce((s,c)=>s+c.cv,0);
      return totalCV >= minCV;
    });
  const lps = Object.keys(matrix).filter(lp => lpCpaMap[lp]).sort((a,b) => (lpCpaMap[a]||0) - (lpCpaMap[b]||0));

  if (!lps.length || !allAttrs.length) {
    el.innerHTML = '<div class="heatmap-placeholder">該当データがありません</div>';
    return;
  }

  let html = '<table class="heatmap-table"><thead><tr><th></th>';
  allAttrs.forEach(a => { html += `<th>${a.length > 8 ? a.substring(0,8)+'…' : a}</th>`; });
  html += '</tr></thead><tbody>';

  lps.forEach(lp => {
    const lpBase = lpCpaMap[lp];
    html += `<tr><td class="row-header" title="${lp}">${shortLabel(lp)}</td>`;
    allAttrs.forEach(attr => {
      const cell = matrix[lp]?.[attr];
      if (cell && cell.cv >= minCV) {
        const cellCpa = cell.cost / cell.cv;
        const dev = (cellCpa / lpBase - 1) * 100;
        const color = dev < -10 ? `rgba(52,211,153,${Math.min(0.35, 0.1 + Math.abs(dev)/100)})` :
                     dev > 10 ? `rgba(248,113,113,${Math.min(0.35, 0.1 + Math.abs(dev)/100)})` :
                     'rgba(255,255,255,0.03)';
        const emoji = dev < -10 ? '🟢' : dev > 10 ? '🔴' : '';
        html += `<td class="heatmap-cell" style="background:${color}">${emoji}${dev > 0 ? '+' : ''}${dev.toFixed(0)}%<span class="heatmap-sub">CV ${cell.cv}</span></td>`;
      } else {
        html += '<td style="color:var(--text-tertiary)">-</td>';
      }
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  el.innerHTML = html;
}

// ============================================
// Combo Exports - クリップボードにTSVコピー
// Tab1: 計算方法  Tab2: 分析表  Tab3: rawデータ
// ============================================

// --- Tab1: 計算方法 ---
function clipComboMethodology() {
  if (!window._comboData) { renderCombo(); }
  const { minCV } = window._comboData;

  const rows = [
    { '項目': '分析名', '内容': '掛け合わせ効果分析（クリエイティブ x LP 相性評価）' },
    { '項目': '分析目的', '内容': '同じクリエイティブでもLPによってCPAが異なる。LP自体の平均CPAをベースラインとし、そこからの乖離でクリエイティブの純粋な実力を推定する' },
    { '項目': '', '内容': '' },
    { '項目': '--- 使用データ ---', '内容': '' },
    { '項目': '対象期間', '内容': document.getElementById('filter-month').value === 'all' ? '全期間' : document.getElementById('filter-month').value },
    { '項目': '対象行数', '内容': filteredData.length + '行' },
    { '項目': 'CVソース', '内容': 'EBiSCV (CVs) 列' },
    { '項目': '費用ソース', '内容': '費用 列' },
    { '項目': '最低CVフィルタ', '内容': minCV + '件以上' },
    { '項目': '', '内容': '' },
    { '項目': '--- 計算手順 ---', '内容': '' },
    { '項目': '', '内容': '' },
    { '項目': 'Step 1: LP平均CPA（ベースライン）の算出', '内容': '' },
    { '項目': '概要', '内容': 'LP単位で費用合計とCV合計を集計し、費用合計 / CV合計 = LP平均CPA を算出' },
    { '項目': '関数', '内容': '=SUMIFS(H:H, B:B, "対象LP名") / SUMIFS(K:K, B:B, "対象LP名")  ※H列=費用, B列=LP名, K列=EBiSCV' },
    { '項目': '具体例', '内容': '=SUMIFS(rawデータ!H:H, rawデータ!B:B, "記事LP001(InstagramLP)") / SUMIFS(rawデータ!K:K, rawデータ!B:B, "記事LP001(InstagramLP)")' },
    { '項目': '', '内容': '' },
    { '項目': 'Step 2: 訴求 x LP ごとのCPA算出', '内容': '' },
    { '項目': '概要', '内容': 'LP名 と バナー訴求 の2条件でSUMIFSし、その組み合わせのCPAを算出' },
    { '項目': '関数', '内容': '=SUMIFS(H:H, B:B, "対象LP", D:D, "対象訴求") / SUMIFS(K:K, B:B, "対象LP", D:D, "対象訴求")  ※D列=バナー訴求' },
    { '項目': '', '内容': '' },
    { '項目': 'Step 3: 乖離率の算出', '内容': '' },
    { '項目': '概要', '内容': '各組み合わせのCPAが、LP平均CPAと比べて何%良い/悪いかを算出' },
    { '項目': '計算式', '内容': '乖離率(%) = (組み合わせCPA / LP平均CPA - 1) x 100' },
    { '項目': '関数', '内容': '=(組み合わせCPAのセル / LP平均CPAのセル - 1) * 100' },
    { '項目': '', '内容': '' },
    { '項目': 'Step 4: 加重乖離率（クリエイティブスコア）', '内容': '' },
    { '項目': '概要', '内容': '同一クリエイティブが複数LPで配信されている場合、各LPでの乖離率をCV件数で加重平均する' },
    { '項目': '計算式', '内容': '加重乖離率 = (各LP乖離率 x 各LP上CV) の合計 / 全LP上CV合計' },
    { '項目': '関数', '内容': '=SUMPRODUCT(乖離率範囲, CV範囲) / SUM(CV範囲)' },
    { '項目': '対象条件', '内容': '2LP以上で展開されたクリエイティブのみ対象' },
    { '項目': '', '内容': '' },
    { '項目': '--- 判定基準 ---', '内容': '' },
    { '項目': '乖離率 < -10%', '内容': 'LP平均より10%以上低いCPA = 相性が良い / クリエイティブ実力が高い' },
    { '項目': '乖離率 > +10%', '内容': 'LP平均より10%以上高いCPA = 相性が悪い / 改善余地あり' },
    { '項目': '', '内容': '' },
    { '項目': '--- スプレッドシート構成 ---', '内容': '' },
    { '項目': 'Tab1（本シート）', '内容': '計算方法の記載' },
    { '項目': 'Tab2', '内容': '分析結果の表（LP平均CPA、相性マトリクス、スコアリング）' },
    { '項目': 'Tab3', '内容': 'rawデータ（元CSVのフィルタ済みデータ）' },
    { '項目': '', '内容': '' },
    { '項目': '--- 関数を使った再現方法 ---', '内容': '' },
    { '項目': '--- rawデータ（Tab3）の列構成 ---', '内容': '' },
    { '項目': 'A列', '内容': '日付' },
    { '項目': 'B列', '内容': 'LP名' },
    { '項目': 'C列', '内容': 'ターゲティング' },
    { '項目': 'D列', '内容': 'バナー訴求' },
    { '項目': 'E列', '内容': 'バナータイプ' },
    { '項目': 'F列', '内容': '広告名' },
    { '項目': 'G列', '内容': '運用' },
    { '項目': 'H列', '内容': '費用' },
    { '項目': 'I列', '内容': 'Imps' },
    { '項目': 'J列', '内容': 'Clicks' },
    { '項目': 'K列', '内容': 'EBiSCV' },
    { '項目': 'L列', '内容': 'EBiS遷移数' },
    { '項目': 'M列', '内容': 'CTR(%)' },
    { '項目': 'N列', '内容': 'CPA' },
    { '項目': '', '内容': '' },
    { '項目': '--- 関数例（Tab3のrawデータを参照する場合） ---', '内容': '' },
    { '項目': 'LP平均CPA', '内容': '=SUMIFS(rawデータ!H:H, rawデータ!B:B, LP名) / SUMIFS(rawデータ!K:K, rawデータ!B:B, LP名)' },
    { '項目': '訴求xLP CPA', '内容': '=SUMIFS(rawデータ!H:H, rawデータ!B:B, LP名, rawデータ!D:D, 訴求名) / SUMIFS(rawデータ!K:K, rawデータ!B:B, LP名, rawデータ!D:D, 訴求名)' },
    { '項目': '乖離率', '内容': '=(訴求xLP CPA / LP平均CPA - 1) * 100' },
    { '項目': '条件付き書式', '内容': '乖離率セルに適用: < -10 で緑背景, > 10 で赤背景' },
  ];

  downloadCSVWithMeta([], rows, '計算方法');
}

// --- Tab2: 分析表（相性マトリクス + スコアリング） ---
// 全セルをSUMIFS関数で出力。rawデータシートの列構成:
// A=日付, B=LP名, C=ターゲティング, D=バナー訴求, E=バナータイプ,
// F=広告名, G=運用, H=費用, I=Imps, J=Clicks, K=EBiSCV,
// L=EBiS遷移数, M=CTR(%), N=CPA
function clipComboMatrix() {
  if (!window._comboData) { renderCombo(); }
  const { lpBaseline, combos, crScores, lpCpaMap, minCV } = window._comboData;

  const appeals = [...new Set(combos.map(c => c.appeal).filter(a => a && a !== '不明'))];
  const bTypes = [...new Set(combos.map(c => c.bType).filter(a => a && a !== '不明'))];
  const lps = lpBaseline.map(l => l.lp);

  const lines = [];

  // ===== Section 1: LP別ベースラインCPA（全セル関数） =====
  lines.push('--- LP別ベースラインCPA ---');
  lines.push(['LP名', '費用合計', 'CV合計', '平均CPA', '組み合わせ数'].join('\t'));
  lpBaseline.forEach((l, i) => {
    const r = i + 3;  // ヘッダー行=1, カラム名行=2 なので3から
    lines.push([
      l.lp,
      `=SUMIFS(rawデータ!H:H, rawデータ!B:B, A${r})`,
      `=SUMIFS(rawデータ!K:K, rawデータ!B:B, A${r})`,
      `=IF(C${r}>0, B${r}/C${r}, "")`,
      l.count
    ].join('\t'));
  });

  // ===== Section 2: 訴求 x LP マトリクス（全セル関数） =====
  lines.push('');
  lines.push('--- バナー訴求 x LP 相性マトリクス ---');
  // ヘッダー行: LP名 | LP平均CPA | 訴求A_費用 | 訴求A_CV | 訴求A_CPA | 訴求A_乖離率(%) | 訴求B_費用 | ...
  const appealHeader = ['LP名', 'LP平均CPA', ...appeals.flatMap(a => [a+'_費用', a+'_CV', a+'_CPA', a+'_乖離率(%)'])];
  lines.push(appealHeader.join('\t'));

  // セクション2のデータ開始行（スプシ上での行番号）
  // セクション1: タイトル1行 + ヘッダー1行 + データN行 + 空行1行 + タイトル1行 + ヘッダー1行 = lpBaseline.length + 5
  const sec2DataStart = lpBaseline.length + 5 + 1; // +1 for 1-indexed

  lps.forEach((lp, lpIdx) => {
    const r = sec2DataStart + lpIdx;
    const vals = [lp];
    // LP平均CPA = rawデータ上のH列合計/K列合計 where B列=LP名
    vals.push(`=SUMIFS(rawデータ!H:H, rawデータ!B:B, A${r}) / SUMIFS(rawデータ!K:K, rawデータ!B:B, A${r})`);

    appeals.forEach((appeal, aIdx) => {
      // 各訴求の列位置（1訴求あたり4列: 費用, CV, CPA, 乖離率）
      // 費用列 = C + aIdx*4, CV列 = D + aIdx*4, CPA列 = E + aIdx*4, 乖離率列 = F + aIdx*4
      const costColLetter = getColLetter(2 + aIdx * 4);  // 0-indexed: A=0, B=1, C=2...
      const cvColLetter = getColLetter(3 + aIdx * 4);
      const cpaColLetter = getColLetter(4 + aIdx * 4);

      // 費用: =SUMIFS(rawデータ!H:H, rawデータ!B:B, $A行, rawデータ!D:D, "訴求名")
      vals.push(`=SUMIFS(rawデータ!H:H, rawデータ!B:B, $A${r}, rawデータ!D:D, "${appeal}")`);
      // CV: =SUMIFS(rawデータ!K:K, rawデータ!B:B, $A行, rawデータ!D:D, "訴求名")
      vals.push(`=SUMIFS(rawデータ!K:K, rawデータ!B:B, $A${r}, rawデータ!D:D, "${appeal}")`);
      // CPA: =IF(CV>=minCV, 費用/CV, "")
      vals.push(`=IF(${cvColLetter}${r}>=${minCV}, ${costColLetter}${r}/${cvColLetter}${r}, "")`);
      // 乖離率: =IF(CPA<>"", (CPA/LP平均CPA-1)*100, "")
      vals.push(`=IF(${cpaColLetter}${r}<>"", (${cpaColLetter}${r}/B${r}-1)*100, "")`);
    });
    lines.push(vals.join('\t'));
  });

  // ===== Section 3: バナータイプ x LP マトリクス（全セル関数） =====
  lines.push('');
  lines.push('--- バナータイプ x LP 相性マトリクス ---');
  const typeHeader = ['LP名', 'LP平均CPA', ...bTypes.flatMap(t => [t+'_費用', t+'_CV', t+'_CPA', t+'_乖離率(%)'])];
  lines.push(typeHeader.join('\t'));

  // セクション3のデータ開始行
  const sec3DataStart = sec2DataStart + lps.length + 3; // 空行1 + タイトル1 + ヘッダー1

  lps.forEach((lp, lpIdx) => {
    const r = sec3DataStart + lpIdx;
    const vals = [lp];
    vals.push(`=SUMIFS(rawデータ!H:H, rawデータ!B:B, A${r}) / SUMIFS(rawデータ!K:K, rawデータ!B:B, A${r})`);

    bTypes.forEach((bType, tIdx) => {
      const costColLetter = getColLetter(2 + tIdx * 4);
      const cvColLetter = getColLetter(3 + tIdx * 4);
      const cpaColLetter = getColLetter(4 + tIdx * 4);

      // 費用: =SUMIFS(rawデータ!H:H, rawデータ!B:B, $A行, rawデータ!E:E, "タイプ名")  ※E列=バナータイプ
      vals.push(`=SUMIFS(rawデータ!H:H, rawデータ!B:B, $A${r}, rawデータ!E:E, "${bType}")`);
      // CV
      vals.push(`=SUMIFS(rawデータ!K:K, rawデータ!B:B, $A${r}, rawデータ!E:E, "${bType}")`);
      // CPA
      vals.push(`=IF(${cvColLetter}${r}>=${minCV}, ${costColLetter}${r}/${cvColLetter}${r}, "")`);
      // 乖離率
      vals.push(`=IF(${cpaColLetter}${r}<>"", (${cpaColLetter}${r}/B${r}-1)*100, "")`);
    });
    lines.push(vals.join('\t'));
  });

  // ===== Section 4: クリエイティブ実力スコア =====
  lines.push('');
  lines.push('--- クリエイティブ実力スコア（LP効果除去） ---');
  lines.push(['順位', '広告名', 'バナー訴求', 'バナータイプ', '費用合計', 'CV合計', '実CPA', '展開LP数', '加重乖離率(%)', '判定'].join('\t'));
  crScores.forEach((c, i) => {
    const judge = c.wDev < -10 ? '実力○' : c.wDev > 10 ? '要改善' : '標準';
    lines.push([i+1, c.adName, c.appeal||'', c.bType||'', Math.round(c.cost), c.cv, Math.round(c.cpa), c.lpCount, c.wDev.toFixed(1), judge].join('\t'));
  });

  // ===== Section 5: 関数の説明 =====
  lines.push('');
  lines.push('--- 関数の説明 ---');
  lines.push('rawデータシートの列構成: A=日付, B=LP名, C=ターゲティング, D=バナー訴求, E=バナータイプ, F=広告名, G=運用, H=費用, I=Imps, J=Clicks, K=EBiSCV');
  lines.push('LP平均CPA: =SUMIFS(rawデータ!H:H, rawデータ!B:B, LP名) / SUMIFS(rawデータ!K:K, rawデータ!B:B, LP名)');
  lines.push('訴求別CPA: =SUMIFS(rawデータ!H:H, rawデータ!B:B, LP名, rawデータ!D:D, 訴求名) / SUMIFS(rawデータ!K:K, rawデータ!B:B, LP名, rawデータ!D:D, 訴求名)');
  lines.push('タイプ別CPA: =SUMIFS(rawデータ!H:H, rawデータ!B:B, LP名, rawデータ!E:E, タイプ名) / SUMIFS(rawデータ!K:K, rawデータ!B:B, LP名, rawデータ!E:E, タイプ名)');
  lines.push('乖離率: =(訴求別CPA / LP平均CPA - 1) * 100');
  lines.push('条件付き書式: 乖離率 < -10 で緑背景（相性良）, > 10 で赤背景（相性悪）');

  copyToClipboard(lines.join('\n'), '分析表');
}

// 列番号(0-indexed) → スプシの列文字変換（A, B, ... Z, AA, AB, ...）
function getColLetter(n) {
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// --- Tab3: rawデータ ---
function clipComboRaw() {
  const rows = filteredData.map(r => ({
    '日付': r[findColCached('日')] || '',
    'LP名': r._lp,
    'ターゲティング': r._targeting,
    'バナー訴求': r._appeal,
    'バナータイプ': r._bannerType,
    '広告名': r._adName,
    '運用': r._operation,
    '費用': r._cost,
    'Imps': r._imps,
    'Clicks': r._clicks,
    'EBiSCV': r._cv,
    'EBiS遷移数': r._ebisCV,
    'CTR(%)': r._ctr.toFixed(4),
    'CPA': r._cpa != null ? Math.round(r._cpa) : ''
  }));

  downloadCSVWithMeta([], rows, 'rawデータ');
}

// トップバーの「CSV出力」ボタンのdispatcher
// comboセクションの場合は分析表をクリップボードにコピー


