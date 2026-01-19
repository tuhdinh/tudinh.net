// auth.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import {
  getAuth,
  onAuthStateChanged,
  signOut
} from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyBXqrLMnNtRnQz7rNqf5eKf_oPd80zcuPI",
  authDomain: "tudinhnet.firebaseapp.com",
  projectId: "tudinhnet",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// 🔓 Pages that do NOT require login
const publicPages = [
  "login.html",
  "resume.html",

];

const currentPage =
  location.pathname.split("/").pop() || "index.html";

// Hide body by default
document.documentElement.style.visibility = "hidden";

onAuthStateChanged(auth, (user) => {
  const isPublic = publicPages.includes(currentPage);

  if (!user && !isPublic) {
    window.location.replace("login.html");
    return;
  }

  if (user && currentPage === "login.html") {
    window.location.replace("index.html");
    return;
  }

  // Auth resolved → show page
  document.documentElement.style.visibility = "visible";
});

// Logout available everywhere
window.logout = async () => {
  await signOut(auth);
  window.location.replace("login.html");
};
