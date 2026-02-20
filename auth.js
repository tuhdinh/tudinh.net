// auth.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import {
  getAuth,
  onAuthStateChanged,
  signOut
} from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

// --- Firebase config ---
const firebaseConfig = {
  apiKey: "AIzaSyBXqrLMnNtRnQz7rNqf5eKf_oPd80zcuPI",
  authDomain: "tudinhnet.firebaseapp.com",
  projectId: "tudinhnet",
  storageBucket: "tudinhnet.firebasestorage.app",
  messagingSenderId: "563689283361",
  appId: "1:563689283361:web:d8e0f48889bcc0f68a6860",
  measurementId: "G-56RGXDFED8"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// ---- SESSION TIMEOUT ----
const SESSION_TIMEOUT = 30 * 60 * 1000; // 30 minutes
const LAST_ACTIVE_KEY = "lastActiveTime";
let timeoutInterval = null;

// ---- PAGES ----
const protectedPages = [
"abaqusfea.html",
"contacts.html",
"conversion.html",
"esprit.html",
"externalflush.html",
"externalflush.png",
"externalwedge.html",
"fea.html",
"index.html",
"internalflush.html",
"internalwedge.html",
"learn.html",
"login.html",
"noaccess.html",
"projects.html",
"resume.html",
"wedge_thread_coupling.html"


];

const currentPage = location.pathname.split("/").pop() || "index.html";

// Hide page until auth resolves
document.body.style.display = "none";

// ---- ACTIVITY TRACKING ----
function updateLastActive() {
  localStorage.setItem(LAST_ACTIVE_KEY, Date.now());
}

["click", "keydown", "scroll", "touchstart"].forEach(evt => {
  window.addEventListener(evt, updateLastActive, true);
});

// ---- FORCE TIMEOUT CHECK (ON LOAD / RETURN) ----
function checkSessionTimeout() {
  const lastActive = Number(localStorage.getItem(LAST_ACTIVE_KEY));
  if (lastActive && Date.now() - lastActive > SESSION_TIMEOUT) {
    logout();
    return true;
  }
  return false;
}

// ---- AUTH GUARD ----
onAuthStateChanged(auth, (user) => {
  const isProtected = protectedPages.includes(currentPage);

  // Not logged in but page is protected
  if (!user && isProtected) {
    window.location.replace("login.html");
    return;
  }

  // Logged in but on login page
  if (user && currentPage === "login.html") {
    window.location.replace("index.html");
    return;
  }

  if (user) {
    // 🚨 Check timeout immediately (mobile-safe)
    if (checkSessionTimeout()) return;

    // Initialize activity time
    updateLastActive();

    // Clear old interval
    if (timeoutInterval) clearInterval(timeoutInterval);

    // Check inactivity every minute
    timeoutInterval = setInterval(() => {
      if (checkSessionTimeout()) {
        clearInterval(timeoutInterval);
      }
    }, 60 * 1000);
  }

  document.body.style.display = "block";

  // ---- NAVBAR BUTTON ----
  const navButton = document.querySelector("#navcol-1 .btn");
  if (navButton) {
    if (user) {
      navButton.textContent = "Logout";
      navButton.className = "btn btn-outline-danger shadow";
      navButton.onclick = logout;
    } else {
      navButton.textContent = "Login";
      navButton.className = "btn btn-primary shadow";
      navButton.onclick = () => window.location.replace("login.html");
    }
  }
});

// ---- LOGOUT ----
export async function logout() {
  localStorage.removeItem(LAST_ACTIVE_KEY);
  if (timeoutInterval) clearInterval(timeoutInterval);
  await signOut(auth);
  window.location.replace("login.html");
}

window.logout = logout;
