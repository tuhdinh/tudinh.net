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

// ---- CONFIG ----

// Pages that REQUIRE login
const protectedPages = [
  "index.html",
  "projects.html",
  "internal_nps.html",
  "external_nps.html",
  "contacts.html"
];

// Pages that should be accessible without login
const publicPages = [
  "login.html",
  "noaccess.html"
];

// ---- AUTH GUARD ----

const currentPage = location.pathname.split("/").pop() || "index.html";

// Hide content until auth resolves
document.body.style.display = "none";

onAuthStateChanged(auth, (user) => {
  const isProtected = protectedPages.includes(currentPage);
  const isPublic = publicPages.includes(currentPage);

  if (!user && isProtected) {
    // Not logged in → redirect to login
    window.location.replace("login.html");
    return;
  }

  if (user && currentPage === "login.html") {
    // Logged in user visiting login page → redirect home
    window.location.replace("index.html");
    return;
  }

  // Show page once auth is resolved
  document.body.style.display = "block";

  // Optional: Update navbar button dynamically
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
  await signOut(auth);
  window.location.replace("login.html");
}

// Make logout globally available (for inline onclick)
window.logout = logout;
