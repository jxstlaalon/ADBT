import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  deleteDoc,
  getDocs,
  collection,
  query,
} from 'firebase/firestore';
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from 'firebase/auth';

const firebaseConfig = {
  apiKey: "AIzaSyAPLqv8UlLSv-I_Gr5Ar_Gvm6hmrmUsIMk",
  authDomain: "ad-bakery-tracker.firebaseapp.com",
  projectId: "ad-bakery-tracker",
  storageBucket: "ad-bakery-tracker.firebasestorage.app",
  messagingSenderId: "828184249303",
  appId: "1:828184249303:web:39d3fe575a411cc51d4192"
};

const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
const db = getFirestore(app);
const auth = getAuth(app);

export { app, db, auth, signInWithEmailAndPassword, signOut, onAuthStateChanged };

export const firebaseStorage = {
  async get(key) {
    try {
      const snap = await getDoc(doc(db, key));
      return snap.exists() ? snap.data() : null;
    } catch (e) {
      console.error('Firebase get error:', e);
      return null;
    }
  },
  async set(key, value) {
    try {
      await setDoc(doc(db, key), value);
    } catch (e) {
      console.error('Firebase set error:', e);
    }
  },
  async delete(key) {
    try {
      await deleteDoc(doc(db, key));
    } catch (e) {
      console.error('Firebase delete error:', e);
    }
  },
  async list(prefix) {
    try {
      const coll = collection(db, prefix.replace(/\/$/, ''));
      const snap = await getDocs(coll);
      return snap.docs.map(d => prefix + d.id);
    } catch (e) {
      console.error('Firebase list error:', e);
      return [];
    }
  }
};

export const initFirebase = async () => {
  console.log('Firebase initialized');
};