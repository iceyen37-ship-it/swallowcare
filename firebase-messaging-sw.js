/* 安心呷：手機通知（iPhone 加到主畫面後、電腦瀏覽器）。只處理通知，不快取任何網頁內容。 */
importScripts(
  "https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js",
  "https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js"
);
// 跟 index.html 的 FIREBASE_CONFIG 一樣
const FIREBASE_CONFIG = {apiKey:"AIzaSyAVRhh8bDu41w7Op4DqjcM6-w2zcFFC4tM", authDomain:"ansimjia.firebaseapp.com", projectId:"ansimjia", storageBucket:"ansimjia.firebasestorage.app", messagingSenderId:"306016418002", appId:"1:306016418002:web:ddc1cc919fd4f23beb950b"};
if (FIREBASE_CONFIG) {
  firebase.initializeApp(FIREBASE_CONFIG);
  firebase.messaging();   // 背景收到通知時，Firebase 會自動顯示；點通知會打開訊息裡的連結
}
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
