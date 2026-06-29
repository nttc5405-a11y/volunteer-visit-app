/**
 * 志工居家訪視系統 — 共用 JavaScript 模組 (app.js)
 *
 * ╔══════════════════════════════════════════════════════════╗
 * ║  部署後請將 GAS 網址填入 CONFIG.GAS_API_URL              ║
 * ║  步驟：GAS 部署 → 複製 Web App 網址 → 貼到下方           ║
 * ╚══════════════════════════════════════════════════════════╝
 */

// ============================================================
// 系統設定
// ============================================================
const CONFIG = {
  // ↓↓↓ 部署 GAS 後，將 Web App 網址貼入此處 ↓↓↓
  GAS_API_URL: 'https://script.google.com/macros/s/AKfycbyq7qvcLI7GbsXnRjZOY-3i0RbZHeibURHer9Mkws45hl15qqdc5pq1ja1n64JN3lRh8Q/exec',
  // ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

  SESSION_KEY: 'vas_session',
  FORM_KEY:    'vas_form',
};

// ============================================================
// API 模組（與 GAS 後端溝通）
// ============================================================
const API = {
  /**
   * 發送 GET 請求
   * @param {string} action
   * @param {Object} params
   */
  async get(action, params = {}) {
    if (!CONFIG.GAS_API_URL || CONFIG.GAS_API_URL === 'YOUR_GAS_WEB_APP_URL_HERE') {
      throw new Error('尚未設定 GAS_API_URL，請參閱 README.md 完成部署設定。');
    }

    const url = new URL(CONFIG.GAS_API_URL);
    url.searchParams.set('action', action);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));

    const res = await fetch(url.toString(), {
      method: 'GET',
      cache: 'no-cache',
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    return json;
  },

  /**
   * 發送 POST 請求（使用 text/plain 避免 CORS Preflight）
   * @param {Object} data
   */
  async post(data) {
    if (!CONFIG.GAS_API_URL || CONFIG.GAS_API_URL === 'YOUR_GAS_WEB_APP_URL_HERE') {
      throw new Error('尚未設定 GAS_API_URL，請參閱 README.md 完成部署設定。');
    }

    const res = await fetch(CONFIG.GAS_API_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain' }, // 避免 CORS Preflight
      body:    JSON.stringify(data),
      cache:   'no-cache',
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  },

  /** 取得訪視題庫 */
  getQuestions:      ()             => API.get('getQuestions'),
  /** 取得所有人員（不含手機） */
  getAllMembers:      ()             => API.get('getAllMembers'),
  /** 依分隊取得人員 */
  getMembersByBranch: (branch)      => API.get('getMembersByBranch', { branch }),
  /** 取得分隊列表 */
  getBranches:       ()             => API.get('getBranches'),
  /** 登入驗證 */
  verifyLogin:       (name, idCard, phone)  => API.get('verifyLogin', { name, idCard, phone }),
  /** 提交訪視表單 */
  submitForm:        (record)       => API.post({ action: 'submitForm', record }),
};

// ============================================================
// Session 管理（登入狀態）
// ============================================================
const Session = {
  set(data) {
    try { sessionStorage.setItem(CONFIG.SESSION_KEY, JSON.stringify(data)); } catch (_) {}
  },
  get() {
    try {
      const raw = sessionStorage.getItem(CONFIG.SESSION_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (_) { return null; }
  },
  clear() {
    try { sessionStorage.removeItem(CONFIG.SESSION_KEY); } catch (_) {}
  },
  /** 若未登入則重導至登入頁，並回傳 null */
  require() {
    const s = Session.get();
    if (!s) { window.location.href = 'index.html'; return null; }
    return s;
  },
};

// ============================================================
// 表單資料管理（跨頁面保存）
// ============================================================
const FormStore = {
  set(data) {
    try { sessionStorage.setItem(CONFIG.FORM_KEY, JSON.stringify(data)); } catch (_) {}
  },
  get() {
    try {
      const raw = sessionStorage.getItem(CONFIG.FORM_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (_) { return {}; }
  },
  update(partial) {
    FormStore.set({ ...FormStore.get(), ...partial });
  },
  clear() {
    try { sessionStorage.removeItem(CONFIG.FORM_KEY); } catch (_) {}
  },
};

// ============================================================
// Toast 通知系統
// ============================================================
const Toast = {
  _container: null,

  _getContainer() {
    if (!this._container) {
      this._container = document.createElement('div');
      this._container.className = 'toast-container';
      document.body.appendChild(this._container);
    }
    return this._container;
  },

  show(message, type = 'info', duration = 3500) {
    const icons = { success: '✅', error: '❌', warning: '⚠️', info: 'ℹ️' };
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.innerHTML = `<span>${icons[type] || icons.info}</span><span>${message}</span>`;
    this._getContainer().appendChild(el);

    setTimeout(() => {
      el.style.animation = 'toastOut .3s ease forwards';
      el.addEventListener('animationend', () => el.remove());
    }, duration);
  },

  success: (msg)          => Toast.show(msg, 'success'),
  error:   (msg)          => Toast.show(msg, 'error',   5000),
  warning: (msg)          => Toast.show(msg, 'warning'),
  info:    (msg)          => Toast.show(msg, 'info'),
};

// ============================================================
// Loading Overlay（全螢幕載入遮罩）
// ============================================================
const Loading = {
  _el: null,

  _create() {
    const el = document.createElement('div');
    el.className = 'loading-overlay';
    el.innerHTML = `
      <div class="spinner"></div>
      <p class="loading-text">載入中...</p>
    `;
    document.body.appendChild(el);
    this._el = el;
  },

  show(text = '載入中...') {
    if (!this._el) this._create();
    this._el.querySelector('.loading-text').textContent = text;
    this._el.style.display = 'flex';
  },

  hide() {
    if (this._el) this._el.style.display = 'none';
  },
};

// ============================================================
// GPS 定位模組
// ============================================================
const GPS = {
  /**
   * 取得目前位置
   * @returns {Promise<{lat, lng, formatted, accuracy}>}
   */
  getPosition() {
    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) {
        reject(new Error('您的瀏覽器不支援 GPS 定位'));
        return;
      }
      navigator.geolocation.getCurrentPosition(
        (pos) => resolve({
          lat:       pos.coords.latitude.toFixed(6),
          lng:       pos.coords.longitude.toFixed(6),
          accuracy:  Math.round(pos.coords.accuracy),
          formatted: `${pos.coords.latitude.toFixed(6)}, ${pos.coords.longitude.toFixed(6)}`,
        }),
        (err) => {
          const msgs = {
            1: '位置存取被拒絕，請在瀏覽器設定中允許定位',
            2: '無法取得位置，請確認 GPS 是否開啟',
            3: '定位逾時，請稍後再試',
          };
          reject(new Error(msgs[err.code] || '定位失敗'));
        },
        { enableHighAccuracy: true, timeout: 12000, maximumAge: 60000 }
      );
    });
  },

  /**
   * 反向地理編碼（座標 → 台灣格式地址）
   * 使用 OpenStreetMap Nominatim，解析 address 物件組合台灣慣用寫法：
   * 縣市 + 鄉鎮區 + 路段 + 門牌
   */
  async reverseGeocode(lat, lng) {
    try {
      const res = await fetch(
        `https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json`,
        { headers: { 'Accept-Language': 'zh-TW,zh;q=0.9', 'User-Agent': 'VolunteerVisitApp/1.0' } }
      );
      const data = await res.json();
      if (!data || !data.address) return '';

      const a = data.address;
      
      // 1. 縣市 (County / City / State)
      let county = a.county || a.city || a.state || '';
      if (county.includes('臺灣') || county.includes('Taiwan')) county = '';

      // 2. 鄉鎮市區 (Town / Township / District / Suburb)
      let district = a.town || a.township || a.district || a.suburb || a.city_district || '';

      // 3. 村里 (Village / Neighbourhood)
      let village = a.village || a.neighbourhood || '';
      if (village === district) village = '';

      // 4. 路街巷弄 (Road)
      let road = a.road || '';

      // 5. 門牌 (House number)
      let num = a.house_number || '';
      if (num && !num.endsWith('號')) {
        num = num + '號';
      }

      // 重組地址，避免重複與遺漏
      let addressParts = [];
      if (county && !addressParts.includes(county)) addressParts.push(county);
      if (district && !addressParts.includes(district)) addressParts.push(district);
      if (village && !addressParts.includes(village) && !district.includes(village)) addressParts.push(village);
      if (road && !addressParts.includes(road)) addressParts.push(road);
      if (num && !addressParts.includes(num)) addressParts.push(num);

      return addressParts.join('');
    } catch (_) { return ''; }
  },
};

// ============================================================
// 工具函式
// ============================================================

/** 格式化日期時間 → "YYYY/MM/DD HH:mm" */
function formatDateTime(date = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}/${p(date.getMonth()+1)}/${p(date.getDate())} ` +
         `${p(date.getHours())}:${p(date.getMinutes())}`;
}

/** 檢查值是否非空（含陣列） */
function isNotEmpty(val) {
  if (Array.isArray(val)) return val.length > 0;
  return val !== null && val !== undefined && String(val).trim() !== '';
}

/** 防抖 */
function debounce(fn, ms = 300) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

/** 捲動至指定元素 */
function scrollToEl(el, block = 'center') {
  if (el) el.scrollIntoView({ behavior: 'smooth', block });
}
