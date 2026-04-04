/**
 * shared/bridge.js
 * postMessage 通訊封裝
 *
 * 各子頁面（sections/sec-*.html）引用此檔案，
 * 透過 Bridge API 與 hub（index.html）進行雙向通訊。
 */

const Bridge = (function () {
  /**
   * 從子頁面送訊息至 hub
   * @param {object} data  含 source 欄位的訊息物件
   */
  function sendToHub(data) {
    window.parent.postMessage(data, '*');
  }

  /**
   * 在子頁面監聽來自 hub 的訊息
   * @param {function} handler  callback(data)，data 為訊息物件
   * @returns {function}        呼叫即可移除監聽
   */
  function onHubMessage(handler) {
    function listener(event) {
      if (!event.data || typeof event.data !== 'object') return;
      handler(event.data);
    }
    window.addEventListener('message', listener);
    return function () { window.removeEventListener('message', listener); };
  }

  return { sendToHub, onHubMessage };
})();
