// auth.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { 
  getAuth, 
  sendSignInLinkToEmail, 
  isSignInWithEmailLink, 
  signInWithEmailLink, 
  signOut,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

// --- Your Firebase config ---
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

/**
 * Send Email Link login
 * @param {string} email
 */
export async function login(email) {
  const actionCodeSettings = {
    url: window.location.origin + '/index.html', // redirect after login
    handleCodeInApp: true
  };
  
  await sendSignInLinkToEmail(auth, email, actionCodeSettings);
  window.localStorage.setItem('emailForSignIn', email);
  alert("Check your email for the login link!");
}

/**
 * Complete login if user clicked Email Link
 */
export async function completeLogin() {
  if (isSignInWithEmailLink(auth, window.location.href)) {
    let email = window.localStorage.getItem('emailForSignIn');
    if (!email) {
      email = prompt("Please enter your email for verification");
    }
    try {
      await signInWithEmailLink(auth, email, window.location.href);
      win
