// logout.js
import { getAuth, signOut } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

const auth = getAuth(); // assumes Firebase app is already initialized in auth.js

// Find the logout button
const btn = document.getElementById("logoutBtn");

if (btn) {
    btn.addEventListener("click", async () => {
        try {
            await signOut(auth);
            window.location.href = "login.html";
        } catch (err) {
            console.error("Logout failed:", err);
            alert("Logout failed, check console");
        }
    });
}
