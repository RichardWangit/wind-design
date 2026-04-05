/**
 * shared/hub-state.js
 * Hub 層全域狀態物件
 *
 * 統一管理跨區（section）資料流，提供：
 *   - O(1) 查詢（以 type 為 key 的 Map）
 *   - Pub/Sub 訂閱通知（支援 wildcard '*'）
 *   - localStorage 持久化（hub_state_v1）
 *   - 歷史記錄（最近 50 筆）
 *
 * API：
 *   HubState.set(type, data)       儲存並通知訂閱者
 *   HubState.get(type)             O(1) 取得資料
 *   HubState.getAll()              回傳全部 { type: data } map
 *   HubState.remove(type)          移除特定 type
 *   HubState.clear()               清空全部
 *   HubState.subscribe(type, fn)   訂閱（type 可為 '*'）
 *   HubState.unsubscribe(type, fn) 取消訂閱
 *   HubState.restore()             從 localStorage 還原（不觸發通知）
 *   HubState.getHistory()          取得歷史記錄陣列
 */

window.HubState = (function () {
  const STORAGE_KEY = 'hub_state_v1';
  const MAX_HISTORY = 50;

  // { [type]: data }
  const _store = Object.create(null);

  // { [type | '*']: Set<fn> }
  const _subs = Object.create(null);

  // [{ type, prev, next, timestamp }, ...]
  const _history = [];

  // ── Internal helpers ──────────────────────────────────────────────

  function _notify(type, data, prev) {
    // Type-specific subscribers
    if (_subs[type]) {
      _subs[type].forEach(function (fn) {
        try { fn(type, data, prev); } catch (e) { console.error('[HubState] subscriber error', e); }
      });
    }
    // Wildcard subscribers
    if (_subs['*']) {
      _subs['*'].forEach(function (fn) {
        try { fn(type, data, prev); } catch (e) { console.error('[HubState] wildcard subscriber error', e); }
      });
    }
  }

  function _persist() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(_store));
    } catch (e) {
      // Quota exceeded or private mode — silently ignore
    }
  }

  function _pushHistory(type, prev, next) {
    _history.push({ type: type, prev: prev, next: next, timestamp: new Date().toISOString() });
    if (_history.length > MAX_HISTORY) _history.shift();
  }

  // ── Public API ────────────────────────────────────────────────────

  /**
   * 儲存資料、寫入 localStorage、通知訂閱者。
   * @param {string} type  資料類型，例如 'wind' / 'terr' / 'occ'
   * @param {object} data  任意可序列化物件
   */
  function set(type, data) {
    const prev = _store[type] !== undefined ? _store[type] : null;
    _store[type] = data;
    _pushHistory(type, prev, data);
    _persist();
    _notify(type, data, prev);
  }

  /**
   * O(1) 查詢。
   * @param  {string} type
   * @returns {object|null}
   */
  function get(type) {
    return _store[type] !== undefined ? _store[type] : null;
  }

  /**
   * 取得全部資料的淺複製 map（{ type: data }）。
   * @returns {object}
   */
  function getAll() {
    var out = Object.create(null);
    var keys = Object.keys(_store);
    for (var i = 0; i < keys.length; i++) {
      out[keys[i]] = _store[keys[i]];
    }
    return out;
  }

  /**
   * 移除特定 type，並通知訂閱者（data 為 null）。
   * @param {string} type
   */
  function remove(type) {
    if (_store[type] === undefined) return;
    var prev = _store[type];
    delete _store[type];
    _pushHistory(type, prev, null);
    _persist();
    _notify(type, null, prev);
  }

  /**
   * 清空全部資料、清除 localStorage、通知 wildcard 訂閱者。
   */
  function clear() {
    var keys = Object.keys(_store);
    keys.forEach(function (k) { delete _store[k]; });
    _history.length = 0;
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) {}
    if (_subs['*']) {
      _subs['*'].forEach(function (fn) {
        try { fn('*', null, null); } catch (e) {}
      });
    }
  }

  /**
   * 訂閱特定 type 或 '*'（所有變更）。
   * callback 簽名：fn(type, data, prev)
   * @param {string}   type
   * @param {function} fn
   */
  function subscribe(type, fn) {
    if (!_subs[type]) _subs[type] = new Set();
    _subs[type].add(fn);
  }

  /**
   * 取消訂閱。
   * @param {string}   type
   * @param {function} fn
   */
  function unsubscribe(type, fn) {
    if (_subs[type]) _subs[type].delete(fn);
  }

  /**
   * 從 localStorage 還原先前的狀態（不觸發通知）。
   * 應在頁面啟動時呼叫一次。
   */
  function restore() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      var saved = JSON.parse(raw);
      if (saved && typeof saved === 'object') {
        var keys = Object.keys(saved);
        for (var i = 0; i < keys.length; i++) {
          _store[keys[i]] = saved[keys[i]];
        }
      }
    } catch (e) {
      console.warn('[HubState] restore failed', e);
    }
  }

  /**
   * 取得歷史記錄（最近 MAX_HISTORY 筆）。
   * @returns {Array}
   */
  function getHistory() {
    return _history.slice();
  }

  return {
    set: set,
    get: get,
    getAll: getAll,
    remove: remove,
    clear: clear,
    subscribe: subscribe,
    unsubscribe: unsubscribe,
    restore: restore,
    getHistory: getHistory
  };
})();
