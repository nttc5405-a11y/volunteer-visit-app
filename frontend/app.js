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

    // Google Apps Script 在冷啟動或忙碌時會間歇性回 404／5xx（實測確有發生），
    // 因此失敗時自動重試。此處所有 action 皆為唯讀查詢，重試不會造成重複寫入；
    // 送出訪視紀錄走的是下方的 post()，不適用重試。
    const MAX_RETRY = 2;
    let lastErr;

    for (let attempt = 0; attempt <= MAX_RETRY; attempt++) {
      if (attempt > 0) {
        await new Promise(r => setTimeout(r, 800 * attempt));
        console.warn(`「${action}」連線失敗，重試第 ${attempt} 次…`);
      }
      try {
        const res = await fetch(url.toString(), { method: 'GET', cache: 'no-cache' });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return await res.json();
      } catch (err) {
        lastErr = err;
      }
    }

    throw lastErr;
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
  /** 取得跑馬燈公告（page：填報頁 / 儀表板） */
  getAnnouncements:  (page)         => API.get('getAnnouncements', { page }),
  /** 登入驗證 */
  verifyLogin:       (idCard, phone, name)  => API.get('verifyLogin', { idCard, phone, name }),
  /** 取得儀表板資料（含個資，需帶登入身分，後端會再驗證一次並依角色過濾） */
  getDashboardData() {
    const s = Session.get() || {};
    return API.get('getDashboardData', { idCard: s.idCard, phone: s.phone, name: s.name });
  },
  
  /** 提交訪視表單 (含離線暫存邏輯) */
  async submitForm(record) {
    if (!navigator.onLine) {
      OfflineQueue.add(record);
      return { success: true, offline: true, message: '目前為離線狀態，已將此筆紀錄儲存於本機暫存！' };
    }

    try {
      const res = await API.post({ action: 'submitForm', record });
      return res;
    } catch (err) {
      OfflineQueue.add(record);
      return { success: true, offline: true, message: '網路連線失敗，已將此筆紀錄儲存於本機暫存！' };
    }
  },
};

// ============================================================
// 跑馬燈公告
//
// 內容維護於 Google 試算表的「系統公告」工作表。
// 沒有任何公告時整條隱藏；公告載入失敗也只是不顯示，不影響其他功能。
// ============================================================
const Marquee = {
  CACHE_TTL: 5 * 60 * 1000,   // 公告不常變，5 分鐘內重複開頁直接用快取，不再打 API

  /**
   * @param {string} page   登入頁 / 填報頁 / 儀表板
   * @param {number} delay  延後幾毫秒才向後端查詢。登入頁會設定延遲，
   *                        讓使用者按下登入時，驗證請求不必排在公告請求後面。
   */
  async load(page, delay = 0) {
    const bar = document.getElementById('marquee');
    if (!bar) return;

    // 快取仍在有效期內就直接用，完全不連線。
    // 這能明顯減少 Apps Script 的請求量（登入頁尤其重要），
    // 代價是公告改動後最多 5 分鐘才會反映，對公告而言可以接受。
    const cached = Marquee.readCache(page);
    if (cached) {
      Marquee.render(cached);
      return;
    }

    if (delay) await new Promise(r => setTimeout(r, delay));

    try {
      const res = await API.getAnnouncements(page);
      const items = (res && res.success && Array.isArray(res.data)) ? res.data : [];
      Marquee.writeCache(page, items);
      Marquee.render(items);
    } catch (err) {
      // 公告失敗不影響主要功能，靜默隱藏即可
      console.warn('公告載入失敗：', err.message);
      bar.classList.add('hidden');
      document.body.classList.remove('has-marquee');
    }
  },

  readCache(page) {
    try {
      const raw = sessionStorage.getItem('vas_ann_' + page);
      if (!raw) return null;
      const box = JSON.parse(raw);
      if (Date.now() - box.t > Marquee.CACHE_TTL) return null;
      return Array.isArray(box.d) ? box.d : null;
    } catch (_) { return null; }
  },

  writeCache(page, items) {
    try {
      sessionStorage.setItem('vas_ann_' + page, JSON.stringify({ t: Date.now(), d: items }));
    } catch (_) { /* 無痕模式等情況存不了，略過即可 */ }
  },

  render(items) {
    const bar   = document.getElementById('marquee');
    const track = document.getElementById('marqueeTrack');
    if (!bar || !track) return;

    const texts = (items || []).map(t => String(t).trim()).filter(Boolean);
    if (texts.length === 0) {          // 沒有公告 → 整條隱藏
      bar.classList.add('hidden');
      document.body.classList.remove('has-marquee');
      return;
    }
    // 供版面調整用：登入頁需要扣掉跑馬燈高度才不會多出捲軸
    document.body.classList.add('has-marquee');

    const content = texts.join('　◆　');
    track.innerHTML = '';
    const first = document.createElement('span');
    first.className = 'marquee-text';
    first.textContent = content;
    track.appendChild(first);
    bar.classList.remove('hidden');

    Marquee.fit(track);

    // 可用寬度改變時要重新判斷（轉動手機方向、視窗縮放、分割畫面）。
    // 用 ResizeObserver 直接監看容器本身，比 window resize 事件可靠
    // ——實測某些情況下視窗尺寸變了卻不會觸發 resize，導致停在錯誤的模式。
    if (!Marquee._bound) {
      Marquee._bound = true;
      const refit = () => Marquee.fit(document.getElementById('marqueeTrack'));

      // 必須保留 observer 的參考：沒有變數指向它時，
      // 瀏覽器可能將其回收，導致尺寸改變後不再觸發（實測發生過）
      if (typeof ResizeObserver !== 'undefined' && track.parentElement) {
        Marquee._observer = new ResizeObserver(refit);
        Marquee._observer.observe(track.parentElement);
      }
      window.addEventListener('resize', refit);             // 舊瀏覽器的後備
      window.addEventListener('orientationchange', refit);  // 手機轉向
      document.addEventListener('visibilitychange', refit); // 背景分頁載入時量不到寬度

      // 字型載入完成後文字寬度會變，且某些情況下初次量測會偏早，
      // 因此稍後再量一次（不依賴畫面更新迴圈）
      setTimeout(refit, 1200);
    }
  },

  /**
   * 依內容長度自動決定顯示方式：
   *   • 一行放得下 → 靜止顯示，使用者隨時看得到完整內容
   *   • 放不下     → 捲動顯示，自畫面右緣進場、跑完離場再重來
   * 公告只保留一份，同一時間不會出現重複的文字。
   * 量不到寬度（頁面尚未排版）時直接返回，等下次事件再試。
   */
  fit(track) {
    if (!track) return;
    const text = track.querySelector('.marquee-text');
    const bar  = track.closest('.marquee');
    if (!text || !bar) return;

    const unit = text.getBoundingClientRect().width;
    const view = track.parentElement ? track.parentElement.getBoundingClientRect().width : 0;
    if (!unit || !view) return;

    // 保險：確保只有一份內容
    while (track.children.length > 1) track.removeChild(track.lastChild);

    const fits = unit <= view - 4;          // 留幾像素邊距，避免貼邊看起來被切到
    bar.classList.toggle('marquee-static', fits);

    if (fits) {
      track.style.removeProperty('--marquee-start');
      track.style.removeProperty('--marquee-shift');
      track.style.animationDuration = '';
      return;
    }

    track.style.setProperty('--marquee-start', view + 'px');   // 起點：畫面右緣外
    track.style.setProperty('--marquee-shift', unit + 'px');   // 終點：完全離開左緣
    // 速度固定約每秒 60px，內容越長跑越久，閱讀節奏一致
    track.style.animationDuration = Math.max(8, Math.round((view + unit) / 60)) + 's';
  },
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

// ============================================================
// 離線填報佇列管理 (OfflineQueue)
// ============================================================
const OfflineQueue = {
  get() {
    try {
      return JSON.parse(localStorage.getItem('offline_submissions') || '[]');
    } catch (_) {
      return [];
    }
  },

  add(record) {
    const queue = OfflineQueue.get();
    record.offlineId = 'off_' + Date.now();
    record.queuedTime = new Date().toISOString();
    queue.push(record);
    localStorage.setItem('offline_submissions', JSON.stringify(queue));
    console.log('✓ 離線紀錄已安全儲存本機：', record.offlineId);
  },

  remove(offlineId) {
    let queue = OfflineQueue.get();
    queue = queue.filter(r => r.offlineId !== offlineId);
    localStorage.setItem('offline_submissions', JSON.stringify(queue));
  },

  async sync() {
    const queue = OfflineQueue.get();
    if (queue.length === 0) return;

    if (!navigator.onLine) return; // 依然處於離線狀態

    console.log(`🔄 網路已回復！開始上傳同步 ${queue.length} 筆離線紀錄...`);
    Toast.info(`連線已回復，正在自動同步 ${queue.length} 筆離線暫存紀錄...`);

    for (let i = 0; i < queue.length; i++) {
      const record = queue[i];
      try {
        const res = await API.post({ action: 'submitForm', record });
        if (res.success) {
          OfflineQueue.remove(record.offlineId);
          Toast.success(`案家「${record.clientName}」的離線暫存已上傳成功！`);
        } else {
          throw new Error(res.error || '後端錯誤');
        }
      } catch (err) {
        console.error(`❌ 離線紀錄同步失敗: ${record.offlineId}，Error: ${err.message}`);
        Toast.error(`同步「${record.clientName}」的暫存紀錄失敗，將於稍後重試：${err.message}`);
        break; // 發生錯誤時先中斷，待下次重試以維持時間序
      }
    }
  }
};

// ─── 監聽連線恢復與頁面加載完成 ──────────────────────────────────────
window.addEventListener('online', () => OfflineQueue.sync());
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    OfflineQueue.sync();
  }, 1500);
});
