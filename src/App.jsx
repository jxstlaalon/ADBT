import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Plus, Minus, Calendar, RotateCcw, ChevronRight, Trash2, AlertCircle, FileDown,
  X, Check, CheckCircle2, History, CalendarRange, Info, ClipboardList, Save,
  Coffee, Croissant, Cake, Cookie, Wine, Search, Printer, DollarSign, User, ArrowLeft,
  Wallet, FileText
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import { firebaseStorage, initFirebase, signInWithEmailAndPassword, signOut, onAuthStateChanged, auth } from './firebase';

const PRINT_STYLES = null;

initFirebase();

window.storage = firebaseStorage;

const CATEGORIES = [
  {
    id: 'savoury',
    label: 'Savoury',
    icon: Coffee,
    items: [
      { id: 'pie', label: 'Pie', price: 10 },
      { id: 'sausage_roll', label: 'Sausage Roll', price: 10 },
      { id: 'mini_pizza', label: 'Mini Pizza', price: 10 },
      { id: 'cheesy_meat_roll', label: 'Cheesy Meat Roll', price: 10 },
      { id: 'puff', label: 'Puff', price: 7 },
    ]
  },
  {
    id: 'specialty',
    label: 'Specialty',
    icon: Croissant,
    items: [
      { id: 'strawberry_cheesecake', label: 'Strawberry Cheese Cake', price: 25 },
      { id: 'cinnamon_rolls', label: 'Cinnamon Rolls', price: 12 },
      { id: 'muffins', label: 'Muffins', price: 12 },
      { id: 'donuts', label: 'Donuts', price: 12 },
    ]
  },
  {
    id: 'pastries',
    label: 'Pastries & Dessert',
    icon: Cake,
    items: [
      { id: 'sweet_bread', label: 'Sweet Bread', price: 8 },
      { id: 'coconut_drops', label: 'Coconut Drops', price: 5 },
      { id: 'currants_roll', label: 'Currants Roll', price: 11 },
      { id: 'coconut_turnover', label: 'Coconut Turnover', price: 8 },
      { id: 'chocolate_cake', label: 'Chocolate Cake', price: 10 },
      { id: 'sponge_cake', label: 'Sponge Cake', price: 10 },
      { id: 'marble_cake', label: 'Marble Cake', price: 10 },
      { id: 'carrot_cake', label: 'Carrot Cake', price: 10 },
      { id: 'banana_bread', label: 'Banana Bread', price: 7 },
      { id: 'oatmeal_raisin_cookies', label: 'Oatmeal Raisin Cookies', price: 7 },
      { id: 'chocolate_chip_cookies', label: 'Chocolate Chip Cookies', price: 9 },
      { id: 'cinnamon_swirl_bread', label: 'Cinnamon Swirl Bread', price: 5 },
      { id: 'chocolate_swirl_bread', label: 'Chocolate Swirl Bread', price: 7 },
    ]
  },
  {
    id: 'beverages',
    label: 'Beverages',
    icon: Wine,
    items: [
      { id: 'water', label: 'Water', price: 5 },
      { id: 'classic_cola', label: 'Classic Cola', price: 3 },
      { id: 'solo', label: 'Solo', price: 4 },
      { id: 'coca_cola', label: 'Coca Cola', price: 8 },
      { id: 'cole_cole', label: 'Cole Cole', price: 3 },
      { id: 'fruitopia', label: 'Fruitopia', price: 6 },
      { id: 'natural_juice', label: 'Natural Juice', price: 10 },
    ]
  },
  {
    id: 'breads',
    label: 'Breads',
    icon: Coffee,
    items: [
      { id: 'coconut_bake_sm', label: 'Coconut Bake (Small)', price: 11 },
      { id: 'coconut_bake_lg', label: 'Coconut Bake (Large)', price: 13 },
      { id: 'hops_sm', label: 'Hops (Small)', price: 6 },
      { id: 'hops_lg', label: 'Hops (Large)', price: 13 },
      { id: 'loaf_sm', label: 'Loaf (Small)', price: 8 },
      { id: 'loaf_lg', label: 'Loaf (Large)', price: 11 },
      { id: 'hotdog_bread_sm', label: 'Hotdog Bread (Small)', price: 'TBA' },
      { id: 'hotdog_bread_lg', label: 'Hotdog Bread (Large)', price: 'TBA' },
    ]
  },
];

const NAV_ITEMS = [
  { id: 'editor', label: 'Daily Tally', icon: ClipboardList },
  { id: 'past',   label: 'Past Days',   icon: History },
  { id: 'export', label: 'Export',      icon: FileDown },
];

const ALL_ITEMS = CATEGORIES.flatMap(cat =>
  cat.items.map(item => ({
    ...item,
    categoryLabel: cat.label,
    categoryId: cat.id,
  }))
);

const C = {
  ink:       '#2B2319',
  mute:      '#6B6558',
  faint:     '#B8AE9B',
  bg:        '#FAF8F3',
  card:      '#FFFFFF',
  line:      '#E8E2D5',
  lineStrong:'#D4C7B0',
  cardinal:  '#5eb76a',
  cardinalDark: '#4a9454',
  gold:      '#C9A34A',
  softBg:    '#F5F1E8',
  success:   '#3F6B3A',
  successBg: '#EEF3EC',
  danger:    '#A02F2A',
  dangerBg:  '#F7ECEB',
};

const FONT_DISPLAY = "'Fraunces', 'Libre Caslon Text', Georgia, serif";
const FONT_BODY    = "'Instrument Sans', system-ui, -apple-system, sans-serif";
const FONT_MONO    = "'JetBrains Mono', 'SF Mono', ui-monospace, monospace";

const todayISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
};

const firstOfMonth = (iso) => iso.slice(0, 8) + '01';

const formatDateLong = (iso) => {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m-1, d).toLocaleDateString('en-US',
    { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
};

const formatDateShort = (iso) => {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m-1, d).toLocaleDateString('en-US',
    { weekday: 'short', month: 'short', day: 'numeric' });
};

const formatMoney = (n) => {
  if (n === 'TBA') return 'TBA';
  return `$${Number(n || 0).toFixed(2)}`;
};

const computeTotal = (qs) =>
  ALL_ITEMS.reduce((s, i) => {
    if (i.price === 'TBA') return s;
    return s + (qs[i.id] || 0) * i.price;
  }, 0);

const totalItemsSold = (qs) =>
  ALL_ITEMS.reduce((s, i) => s + (qs[i.id] || 0), 0);

const emptyQuantities = () =>
  Object.fromEntries(ALL_ITEMS.map(i => [i.id, 0]));

const STORAGE_PREFIX = 'adbt_daily_/';

export default function App() {
  const [view, setView]             = useState('editor');
  const [workingDate, setWorkingDate] = useState(todayISO());
  const [quantities, setQuantities] = useState(emptyQuantities());
  const [sales, setSales] = useState([]);
  const [existingRecord, setExistingRecord] = useState(null);
  const [isDirty, setIsDirty]       = useState(false);
  const [loaded, setLoaded]         = useState(false);
  const [pastDays, setPastDays]     = useState([]);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmReset, setConfirmReset]   = useState(false);
  const [undoSnapshot, setUndoSnapshot]   = useState(null);
  const [exportStart, setExportStart] = useState(firstOfMonth(todayISO()));
  const [exportEnd,   setExportEnd]   = useState(todayISO());
  const [exporting, setExporting]     = useState(false);
  const [toast, setToast]           = useState(null);
  const [sigOpen, setSigOpen]       = useState(false);
  const [sigName, setSigName]       = useState('');
  const [receiptView, setReceiptView] = useState(null);
  const [user, setUser]             = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginError, setLoginError]   = useState('');
  const [startDayOpen, setStartDayOpen] = useState(false);
  const [startName, setStartName]   = useState('');
  const [startFloat, setStartFloat] = useState('');
  const [notes, setNotes] = useState({ payouts: [], clears: [], custom: [] });
  const [noteModalOpen, setNoteModalOpen] = useState(null);
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthLoading(false);
    });
    return unsubscribe;
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    const password = e.target.password.value;
    setLoginError('');
    try {
      await signInWithEmailAndPassword(auth, 'admin@adbt.org', password);
    } catch (err) {
      setLoginError(err.message);
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
  };

  useEffect(() => {
    if (!user) return;
    let cancelled = false;
    setLoaded(false);
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_PREFIX + workingDate);
        if (cancelled) return;
        if (res) {
          const rec = res;
          setExistingRecord(rec);
          setQuantities(rec.quantities || emptyQuantities());
          setSales(rec.sales || []);
          setStartName(rec.startName || '');
          setStartFloat(rec.startFloat !== undefined ? String(rec.startFloat) : '');
          setNotes(rec.notes || { payouts: [], clears: [], custom: [] });
        } else {
          setExistingRecord(null);
          setQuantities(emptyQuantities());
          setSales([]);
          setStartName('');
          setStartFloat('');
          setNotes({ payouts: [], clears: [], custom: [] });
        }
      } catch {
        if (cancelled) return;
        setExistingRecord(null);
        setQuantities(emptyQuantities());
        setSales([]);
      } finally {
        if (!cancelled) {
          setIsDirty(false);
          setUndoSnapshot(null);
          setLoaded(true);
        }
      }
    })();
    return () => { cancelled = true; };
  }, [workingDate, refreshKey, user]);

  useEffect(() => {
    if (!user) return;
    (async () => {
      try {
        const keys = await window.storage.list(STORAGE_PREFIX, false);
        const days = [];
        for (const key of keys) {
          const data = await window.storage.get(key);
          if (data) {
            days.push(data);
          }
        }
        days.sort((a,b) => b.date.localeCompare(a.date));
        setPastDays(days);
      } catch (err) {
        console.error(err);
        setPastDays([]);
      }
    })();
  }, [user]);

  useEffect(() => {
    if (view !== 'past') return;
    (async () => {
      try {
        const keys = await window.storage.list(STORAGE_PREFIX);
        const days = [];
        for (const key of keys) {
          const data = await window.storage.get(key);
          if (data) {
            days.push(data);
          }
        }
        days.sort((a,b) => b.date.localeCompare(a.date));
        setPastDays(days);
      } catch (err) {
        console.error(err);
        setPastDays([]);
      }
    })();
  }, [view]);

  const dailyTotal  = useMemo(() => computeTotal(quantities), [quantities]);
  const itemsCount  = useMemo(() => totalItemsSold(quantities), [quantities]);
  const isToday     = workingDate === todayISO();
  const isPastDate  = workingDate < todayISO();
  const isFutureDate = workingDate > todayISO();
  const hasStartedDay = !!(existingRecord?.startName || (startName && startName.trim()));

  const bump = (id, delta) => {
    if (delta === 0) return;
    const timestamp = new Date().toISOString();
    setQuantities(q => {
      const next = { ...q, [id]: Math.max(0, (q[id] || 0) + delta) };
      return next;
    });
    setSales(prev => [...prev, {
      itemId: id,
      delta,
      timestamp,
    }]);
    setIsDirty(true);
  };

  const resetCounters = () => {
    if (itemsCount === 0) return;
    setConfirmReset(true);
  };

  const doReset = () => {
    setUndoSnapshot({ quantities: { ...quantities }, sales: [...sales] });
    setQuantities(emptyQuantities());
    setSales([]);
    setIsDirty(true);
    setConfirmReset(false);
    showToast('success', 'Counters reset. Click "Undo reset" to restore.');
  };

  const undoReset = () => {
    if (!undoSnapshot) return;
    setQuantities(undoSnapshot.quantities);
    setSales(undoSnapshot.sales || []);
    setUndoSnapshot(null);
    if (existingRecord && JSON.stringify(undoSnapshot.quantities) === JSON.stringify(existingRecord.quantities)) {
      setIsDirty(false);
    } else {
      setIsDirty(true);
    }
    showToast('success', 'Reset undone.');
  };

  const openEndDay = () => {
    if (itemsCount === 0) {
      showToast('error', 'No items tallied. Add at least one sale before ending the day.');
      return;
    }
    setSigName(existingRecord?.signature || '');
    setSigOpen(true);
  };

  const openStartDay = () => {
    setStartName(existingRecord?.startName || '');
    setStartFloat(existingRecord?.startFloat !== undefined ? String(existingRecord.startFloat) : '');
    setStartDayOpen(true);
  };

  const quickSave = async () => {
    const record = {
      date: workingDate,
      quantities: { ...quantities },
      sales: [...sales],
      total: dailyTotal,
      startName: existingRecord?.startName || startName.trim() || null,
      startFloat: existingRecord?.startFloat ?? (startFloat ? parseFloat(startFloat) : null),
      signature: existingRecord?.signature || null,
      savedAt: existingRecord?.savedAt || new Date().toISOString(),
      lastEdited: new Date().toISOString(),
      editCount: (existingRecord?.editCount || 0) + (existingRecord ? 1 : 0),
      notes,
    };
    try {
      await window.storage.set(STORAGE_PREFIX + workingDate, record);
      setExistingRecord(record);
      setIsDirty(false);
      setUndoSnapshot(null);
      showToast('success', `Saved at ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`);
    } catch (err) {
      showToast('error', 'Unable to save: ' + err.message);
    }
  };

  const saveStartDay = async () => {
    const float = parseFloat(startFloat) || 0;
    const record = {
      date: workingDate,
      quantities: { ...quantities },
      sales: [...sales],
      total: dailyTotal,
      startName: startName.trim(),
      startFloat: float,
      signature: existingRecord?.signature || null,
      savedAt: existingRecord?.savedAt || new Date().toISOString(),
      lastEdited: new Date().toISOString(),
      editCount: (existingRecord?.editCount || 0) + (existingRecord ? 1 : 0),
      notes,
    };
    try {
      await window.storage.set(STORAGE_PREFIX + workingDate, record);
      setExistingRecord(record);
      setIsDirty(true);
      setStartDayOpen(false);
      showToast('success', `Day started with float of ${formatMoney(float)}.`);
    } catch (err) {
      showToast('error', 'Unable to save: ' + err.message);
    }
  };

  const addNote = (type, value) => {
    const note = { id: Date.now(), value, createdAt: new Date().toISOString() };
    setNotes(prev => ({ ...prev, [type]: [...prev[type], note] }));
    setNoteModalOpen(null);
  };

  const removeNote = (type, id) => {
    setNotes(prev => ({ ...prev, [type]: prev[type].filter(n => n.id !== id) }));
  };

  const submitEndDay = async () => {
    const name = sigName.trim();
    if (!name) {
      showToast('error', 'Please enter your name as a signature.');
      return;
    }
    const record = {
      date: workingDate,
      quantities: { ...quantities },
      sales: [...sales],
      total: dailyTotal,
      startName: existingRecord?.startName || startName.trim() || null,
      startFloat: existingRecord?.startFloat ?? (startFloat ? parseFloat(startFloat) : null),
      signature: name,
      savedAt: existingRecord?.savedAt || new Date().toISOString(),
      lastEdited: new Date().toISOString(),
      editCount: (existingRecord?.editCount || 0) + (existingRecord ? 1 : 0),
      notes,
    };
    try {
      await window.storage.set(STORAGE_PREFIX + workingDate, record);
      setExistingRecord(record);
      setIsDirty(false);
      setUndoSnapshot(null);
      setSigOpen(false);
      setSigName('');
      showToast('success', existingRecord
        ? `Record for ${formatDateShort(workingDate)} updated.`
        : `Day closed and saved for ${formatDateShort(workingDate)}.`);
    } catch (err) {
      showToast('error', 'Unable to save: ' + err.message);
    }
  };

  const openPastDay = (dateISO) => {
    setWorkingDate(dateISO);
    setRefreshKey(k => k + 1);
    setView('editor');
  };

  const confirmDeletePastDay = (dateISO) => setConfirmDelete(dateISO);

  const doDelete = async () => {
    const d = confirmDelete;
    if (!d) return;
    try {
      await window.storage.delete(STORAGE_PREFIX + d, false);
      setConfirmDelete(null);
      if (workingDate === d) {
        setExistingRecord(null);
        setQuantities(emptyQuantities());
        setIsDirty(false);
      }
      setPastDays(prev => prev.filter(day => day.date !== d));
      showToast('success', `Record for ${formatDateShort(d)} deleted.`);
    } catch (err) {
      showToast('error', 'Unable to delete: ' + err.message);
    }
  };

  const jumpToToday = () => setWorkingDate(todayISO());
  const goBackToPastDays = () => {
    setWorkingDate(todayISO());
    setView('past');
  };

  const showToast = (kind, msg) => {
    setToast({ kind, msg });
    setTimeout(() => setToast(null), 4000);
  };

  const presetRange = (preset) => {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();
    const fmt = (d) => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    if (preset === 'thisMonth') {
      setExportStart(fmt(new Date(y, m, 1)));
      setExportEnd(fmt(now));
    } else if (preset === 'lastMonth') {
      const first = new Date(y, m-1, 1);
      const last  = new Date(y, m, 0);
      setExportStart(fmt(first));
      setExportEnd(fmt(last));
    } else if (preset === 'last7') {
      const start = new Date(y, m, now.getDate() - 6);
      setExportStart(fmt(start));
      setExportEnd(fmt(now));
    } else if (preset === 'thisYear') {
      setExportStart(`${y}-01-01`);
      setExportEnd(fmt(now));
    }
  };

  const runExport = async () => {
    if (exportStart > exportEnd) {
      showToast('error', 'Start date must be on or before end date.');
      return;
    }
    setExporting(true);
    try {
      const inRange = pastDays
        .filter(d => d.date >= exportStart && d.date <= exportEnd)
        .sort((a,b) => a.date.localeCompare(b.date));

      if (inRange.length === 0) {
        showToast('error', 'No saved records in that date range.');
        setExporting(false);
        return;
      }

      const grandTotal = inRange.reduce((s,d) => s + (d.total || 0), 0);

      const summary = [
        ['ALON AND DEKLON BAKERY TRACKER — MONTHLY SUMMARY'],
        [`Report period: ${exportStart} to ${exportEnd}`],
        [`Days with entries: ${inRange.length}`],
        [],
        ['Date', 'Day', 'Daily Total'],
        ...inRange.map(d => {
          const [y,m,dd] = d.date.split('-').map(Number);
          const weekday = new Date(y, m-1, dd).toLocaleDateString('en-US', { weekday: 'long' });
          return [d.date, weekday, d.total || 0];
        }),
        [],
        ['', 'GRAND TOTAL', grandTotal],
      ];
      const wsSummary = XLSX.utils.aoa_to_sheet(summary);
      wsSummary['!cols'] = [{ wch: 16 }, { wch: 16 }, { wch: 18 }];
      for (let r = 5; r < 5 + inRange.length; r++) {
        const cellRef = XLSX.utils.encode_cell({ r, c: 2 });
        if (wsSummary[cellRef]) wsSummary[cellRef].z = '"$"#,##0.00';
      }
      const grandCellRef = XLSX.utils.encode_cell({ r: 5 + inRange.length + 1, c: 2 });
      if (wsSummary[grandCellRef]) wsSummary[grandCellRef].z = '"$"#,##0.00';

      const itemized = [
        ['ALON AND DEKLON BAKERY TRACKER — ITEMIZED DAILY BREAKDOWN'],
        [`Report period: ${exportStart} to ${exportEnd}`],
        [],
        ['Date', 'Category', 'Item', 'Unit Price', 'Quantity', 'Subtotal'],
      ];
      inRange.forEach((day, idx) => {
        let hasAny = false;
        ALL_ITEMS.forEach(it => {
          const qty = (day.quantities || {})[it.id] || 0;
          if (qty > 0) {
            hasAny = true;
            const price = it.price === 'TBA' ? 0 : it.price;
            itemized.push([day.date, it.categoryLabel, it.label, it.price, qty, qty * price]);
          }
        });
        if (!hasAny) {
          itemized.push([day.date, '(no items recorded)', '', '', '', 0]);
        }
        itemized.push(['', '', '', '', `Daily total — ${day.date}`, day.total || 0]);
        if (idx < inRange.length - 1) itemized.push([]);
      });
      itemized.push([]);
      itemized.push(['', '', '', '', 'GRAND TOTAL', grandTotal]);
      const wsItem = XLSX.utils.aoa_to_sheet(itemized);
      wsItem['!cols'] = [{wch:14},{wch:26},{wch:22},{wch:14},{wch:12},{wch:18}];
      const range = XLSX.utils.decode_range(wsItem['!ref']);
      for (let r = 4; r <= range.e.r; r++) {
        for (const c of [3, 5]) {
          const ref = XLSX.utils.encode_cell({ r, c });
          if (wsItem[ref] && typeof wsItem[ref].v === 'number') {
            wsItem[ref].z = '"$"#,##0.00';
          }
        }
      }

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');
      XLSX.utils.book_append_sheet(wb, wsItem, 'Itemized');

      const filename = `Bakery_Report_${exportStart}_to_${exportEnd}.xlsx`;
      XLSX.writeFile(wb, filename);
      showToast('success', `Exported ${inRange.length} day${inRange.length===1?'':'s'} — total ${formatMoney(grandTotal)}.`);
    } catch (err) {
      showToast('error', 'Export error: ' + err.message);
    } finally {
      setExporting(false);
    }
  };

  if (authLoading) {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#FAF8F3' }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: 18, color: '#2B2319' }}>Loading...</div>
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div style={{
        minHeight: '100vh',
        background: `linear-gradient(135deg, ${C.ink} 0%, #1e1813 100%)`,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 20,
        fontFamily: FONT_BODY,
      }}>
        <div style={{
          background: C.card,
          borderRadius: 16,
          padding: '40px 36px',
          width: '100%',
          maxWidth: 400,
          boxShadow: '0 24px 80px rgba(0,0,0,0.4)',
        }}>
          <div style={{ textAlign: 'center', marginBottom: 32 }}>
            <img src="logo.png" alt="Aalon & Deklon Bakery" style={{ width: 100, height: 100, objectFit: 'contain', marginBottom: 16 }} />
            <div style={{ fontFamily: FONT_DISPLAY, fontSize: 24, fontWeight: 500, color: C.ink, marginBottom: 4 }}>
A & D's Bakery
            </div>
            <div style={{ fontSize: 13, color: C.mute }}>
              Enter your password to continue
            </div>
          </div>

          <form onSubmit={handleLogin}>
            <div style={{ marginBottom: 16 }}>
              <label style={{ display: 'block', fontSize: 11, color: C.mute, fontWeight: 500, marginBottom: 6, textTransform: 'uppercase', letterSpacing: '0.15em' }}>
                Account
              </label>
              <div style={{
                padding: '12px 14px',
                border: `1px solid ${C.lineStrong}`,
                borderRadius: 8,
                fontSize: 14,
                color: C.ink,
                background: C.softBg,
              }}>
                Admin
              </div>
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={{ display: 'block', fontSize: 11, color: C.mute, fontWeight: 500, marginBottom: 6, textTransform: 'uppercase', letterSpacing: '0.15em' }}>
                Password
              </label>
              <input name="password" type="password" required
                style={{ width: '100%', padding: '12px 14px', border: `1px solid ${C.lineStrong}`, borderRadius: 8, fontSize: 14, boxSizing: 'border-box' }} />
            </div>
            {loginError && <div style={{ color: C.danger, fontSize: 13, marginBottom: 16 }}>{loginError}</div>}
            <button type="submit"
              style={{ width: '100%', padding: '14px', background: C.cardinal, color: '#FFFFFF', border: 'none', borderRadius: 8, fontSize: 14, fontWeight: 500, cursor: 'pointer' }}>
              Sign In
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', backgroundColor: C.bg, color: C.ink, fontFamily: FONT_BODY }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Instrument+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; }
        input[type="date"]::-webkit-calendar-picker-indicator { opacity: 0.55; cursor: pointer; }
        input:focus, button:focus, select:focus { outline: none; }
        input:focus-visible, button:focus-visible { box-shadow: 0 0 0 3px ${C.softBg}, 0 0 0 4px ${C.cardinal}; }
        .fdsa-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
        .fdsa-scroll::-webkit-scrollbar-thumb { background: ${C.lineStrong}; border-radius: 4px; }
        @keyframes fdsaFade { from { opacity: 0; transform: translateY(-4px); } to { opacity: 1; transform: none; } }
        .fdsa-fade { animation: fdsaFade .25s ease; }

        @media (max-width: 1024px) {
          .fdsa-header { flex-direction: column !important; gap: 16px !important; padding: 16px 20px !important; }
          .fdsa-logo-img { width: 140px !important; height: auto !important; }
          .fdsa-title { font-size: 22px !important; }
          .fdsa-main { padding: 20px 16px 16px !important; }
          .fdsa-nav { padding: 0 16px !important; overflow-x: auto; }
          .fdsa-nav button { padding: 12px 14px !important; font-size: 12px !important; }
          .fdsa-grid { grid-template-columns: 1fr !important; }
          .fdsa-footer { padding: 24px 20px !important; }
        }

        @media (max-width: 640px) {
          .fdsa-header { padding: 12px 14px !important; gap: 12px !important; flex-direction: column !important; align-items: center !important; text-align: center !important; }
          .fdsa-title { font-size: 18px !important; }
          .fdsa-main { padding: 16px 12px 12px !important; }
          .fdsa-nav { padding: 0 12px !important; overflow-x: auto; -webkit-overflow-scrolling: touch; }
          .fdsa-nav button { padding: 10px 12px !important; font-size: 11px !important; white-space: nowrap; }
          .fdsa-editor-bar { flex-direction: column !important; align-items: stretch !important; gap: 12px !important; }
          .fdsa-total-bar { flex-direction: column !important; gap: 16px !important; padding: 20px !important; text-align: center; }
          .fdsa-footer { padding: 20px 14px !important; }
          .fdsa-export-grid { grid-template-columns: 1fr !important; }
          .fdsa-footer-buttons { flex-wrap: wrap !important; justify-content: center !important; }
          .fdsa-footer-buttons button { flex: 1 1 calc(50% - 8px) !important; min-width: 100px !important; font-size: 12px !important; padding: 10px 12px !important; }
          .receipt-modal { width: 95% !important; max-width: none !important; padding: 16px !important; }
          .receipt-modal h2 { font-size: 18px !important; }
          .receipt-modal .item-row { font-size: 12px !important; padding: 8px 12px !important; }
          .receipt-modal .item-qty { min-width: 20px !important; font-size: 10px !important; }
          .past-day-row { flex-direction: column !important; gap: 12px !important; align-items: flex-start !important; }
          .past-day-row .day-info { width: 100% !important; }
          .past-day-row .day-actions { width: 100% !important; justify-content: flex-start !important; }
          .category-card { padding: 12px !important; }
          .category-card .item-row { padding: 8px 0 !important; }
          .category-card .item-btn { width: 32px !important; height: 32px !important; }
          .category-card .item-btn svg { width: 14px !important; height: 14px !important; }
        }
      `}</style>

      <header style={{ backgroundColor: C.card, borderBottom: `1px solid ${C.line}` }}>
        <div className="fdsa-header" style={{ maxWidth: 1200, margin: '0 auto', padding: '18px 28px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 24 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
            <img src="logo.png" alt="Aalon & Deklon Bakery" style={{ width: 120, height: 120, objectFit: 'contain', borderRadius: 8 }} />
            <div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.28em', color: C.mute, fontWeight: 500 }}>
                Daily Operations
              </div>
              <h1 className="fdsa-title" style={{ margin: '2px 0 0', fontFamily: FONT_DISPLAY, fontSize: 28, fontWeight: 500, letterSpacing: '-0.015em', color: C.ink }}>
                A & D's Bakery Tracker
              </h1>
            </div>
          </div>
          <div style={{ textAlign: 'right', display: 'flex', alignItems: 'center', gap: 16 }}>
            <div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.28em', color: C.mute, fontWeight: 500 }}>Today</div>
              <div style={{ marginTop: 3, fontSize: 14, color: C.ink, fontFamily: FONT_MONO }}>{formatDateShort(todayISO())}</div>
            </div>
            <button onClick={handleLogout}
              style={{
                padding: '8px 14px', background: 'transparent', color: C.mute,
                border: `1px solid ${C.lineStrong}`, borderRadius: 6, cursor: 'pointer',
                fontSize: 12, fontFamily: FONT_BODY,
              }}>
              Sign Out
            </button>
          </div>
        </div>

        <nav className="fdsa-nav" style={{ maxWidth: 1200, margin: '0 auto', padding: '0 28px', display: 'flex', gap: 2, borderTop: `1px solid ${C.line}` }}>
          {NAV_ITEMS.map(t => {
            const active = view === t.id;
            const Ico = t.icon;
            return (
              <button
                key={t.id}
                onClick={() => setView(t.id)}
                style={{
                  background: 'none', border: 'none', cursor: 'pointer',
                  padding: '14px 18px', fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  color: active ? C.cardinal : C.mute,
                  borderBottom: `2px solid ${active ? C.cardinal : 'transparent'}`,
                  marginBottom: -1,
                  display: 'flex', alignItems: 'center', gap: 8,
                  letterSpacing: '0.02em',
                  transition: 'color .15s',
                }}
              >
                <Ico size={15} strokeWidth={active ? 2.2 : 1.8} />
                {t.label}
              </button>
            );
          })}
        </nav>
      </header>

      <main className="fdsa-main" style={{ maxWidth: 1200, margin: '0 auto', padding: '32px 28px 24px' }}>
        {view === 'editor' && (
          <EditorView
            workingDate={workingDate} setWorkingDate={setWorkingDate}
            existingRecord={existingRecord} quantities={quantities}
            bump={bump} resetCounters={resetCounters}
            isDirty={isDirty} isToday={isToday} isPastDate={isPastDate} isFutureDate={isFutureDate}
            dailyTotal={dailyTotal} itemsCount={itemsCount}
            openEndDay={openEndDay} jumpToToday={jumpToToday} goBackToPastDays={goBackToPastDays}
            loaded={loaded}
            undoSnapshot={undoSnapshot} undoReset={undoReset}
            quickSave={quickSave}
            sales={sales}
            setReceiptView={setReceiptView}
            startName={startName} startFloat={startFloat}
            openStartDay={openStartDay}
            setStartName={setStartName} setStartFloat={setStartFloat}
            saveStartDay={saveStartDay}
            hasStartedDay={hasStartedDay}
            notes={notes} addNote={addNote} removeNote={removeNote}
            setNoteModalOpen={setNoteModalOpen}
          />
        )}
        {view === 'past' && (
          <PastView
            days={pastDays}
            onOpen={openPastDay} onDelete={confirmDeletePastDay}
            onGoEditor={() => setView('editor')}
          />
        )}
        {view === 'export' && (
          <ExportView
            start={exportStart} end={exportEnd}
            setStart={setExportStart} setEnd={setExportEnd}
            presetRange={presetRange}
            pastDays={pastDays} exporting={exporting}
            runExport={runExport}
          />
        )}
        {receiptView && (
          <ModalShell onCancel={() => setReceiptView(null)}>
            <ReceiptView
              day={receiptView}
              onClose={() => setReceiptView(null)}
            />
          </ModalShell>
        )}
      </main>

      <footer className="fdsa-footer" style={{ borderTop: `1px solid ${C.line}`, backgroundColor: C.card, marginTop: 40 }}>
        <div style={{ maxWidth: 1200, margin: '0 auto', padding: '32px 28px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
          <img src="logo.png" alt="Aalon & Deklon Bakery" style={{ width: 112, height: 112, objectFit: 'contain', marginBottom: 8 }} />
          <div style={{ fontSize: 13, color: C.mute, fontStyle: 'italic', textAlign: 'center', maxWidth: 600, lineHeight: 1.6 }}>
            &ldquo;Wealth from get-rich-quick schemes quickly disappears; wealth from hard work grows over time.&rdquo;
          </div>
          <div style={{ fontSize: 12, color: C.mute, textAlign: 'center' }}>
            Proverbs 13:21
          </div>
          <div style={{ fontSize: 13, color: C.mute, fontStyle: 'italic', textAlign: 'center', maxWidth: 600, lineHeight: 1.6 }}>
            &ldquo;For I know the plans I have for you,&rdquo; declares the LORD, &ldquo;plans to prosper you and not to harm you, plans to give you hope and a future.&rdquo;
          </div>
          <div style={{ fontSize: 12, color: C.mute, textAlign: 'center' }}>
            Jeremiah 29:11
          </div>
          <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.28em', color: C.faint, fontWeight: 500 }}>
            Made by Aalon Peters
          </div>
        </div>
      </footer>

      {sigOpen && (
        <SignatureModal
          workingDate={workingDate}
          dailyTotal={dailyTotal}
          itemsCount={itemsCount}
          existing={existingRecord}
          value={sigName}
          setValue={setSigName}
          onCancel={() => { setSigOpen(false); setSigName(''); }}
          onSubmit={submitEndDay}
        />
      )}

      {startDayOpen && (
        <StartDayModal
          workingDate={workingDate}
          startName={startName}
          setStartName={setStartName}
          startFloat={startFloat}
          setStartFloat={setStartFloat}
          onCancel={() => { 
            setStartDayOpen(false); 
            setStartName(existingRecord?.startName || '');
            setStartFloat(existingRecord?.startFloat !== undefined ? String(existingRecord.startFloat) : '');
          }}
          onSave={saveStartDay}
        />
      )}

      {noteModalOpen && (
        <NoteModal
          type={noteModalOpen}
          onCancel={() => setNoteModalOpen(null)}
          onSave={(val) => addNote(noteModalOpen === 'payout' ? 'payouts' : noteModalOpen === 'clear' ? 'clears' : 'custom', val)}
        />
      )}

      {confirmDelete && (
        <ConfirmModal
          title="Delete day record?"
          body={`This will permanently remove the saved record for ${formatDateLong(confirmDelete)}. This cannot be undone.`}
          confirmLabel="Delete record"
          danger
          onConfirm={doDelete}
          onCancel={() => setConfirmDelete(null)}
        />
      )}

      {confirmReset && (
        <ConfirmModal
          title="Reset counters?"
          body="This clears the current entries on screen. Saved records are not affected until you save again."
          confirmLabel="Reset counters"
          onConfirm={doReset}
          onCancel={() => setConfirmReset(false)}
        />
      )}

      {toast && <Toast kind={toast.kind} msg={toast.msg} />}
    </div>
  );
}

function EditorView({
  workingDate, setWorkingDate,
  existingRecord, quantities, bump, resetCounters,
  isDirty, isToday, isPastDate, isFutureDate,
  dailyTotal, itemsCount,
  openEndDay, jumpToToday, goBackToPastDays, loaded,
  undoSnapshot, undoReset,
  quickSave, sales, setReceiptView,
  startName, startFloat, openStartDay,
  setStartName, setStartFloat, saveStartDay,
  hasStartedDay,
  notes, addNote, removeNote, setNoteModalOpen,
}) {
  return (
    <div className="fdsa-fade">
      <div className="fdsa-editor-bar" style={{
        display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap',
        padding: '16px 20px',
        backgroundColor: isToday ? C.card : (isPastDate ? '#FBF4E8' : '#FAF0F0'),
        border: `1px solid ${isToday ? C.line : (isPastDate ? '#E8D9B6' : '#E6C8C5')}`,
        borderRadius: 10, marginBottom: 24,
      }}>
        {isPastDate && goBackToPastDays && (
          <button onClick={goBackToPastDays}
            style={{
              padding: '9px 14px', border: `1px solid ${C.lineStrong}`, borderRadius: 6,
              background: C.card, cursor: 'pointer', fontSize: 12, fontWeight: 500,
              color: C.ink, display: 'flex', alignItems: 'center', gap: 6, fontFamily: FONT_BODY,
            }}>
            <ArrowLeft size={14} /> Back
          </button>
        )}
        <Calendar size={18} style={{ color: isToday ? C.cardinal : (isPastDate ? '#946C20' : C.danger) }} />
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.22em', color: C.mute, fontWeight: 600 }}>
            {isToday ? 'Current day' : isPastDate ? 'Editing past day' : 'Future date'}
          </div>
          <div style={{ fontFamily: FONT_DISPLAY, fontSize: 22, fontWeight: 500, color: C.ink, marginTop: 2 }}>
            {formatDateLong(workingDate)}
          </div>
        </div>
        <input
          type="date"
          value={workingDate}
          onChange={e => e.target.value && setWorkingDate(e.target.value)}
          style={{
            padding: '9px 12px', border: `1px solid ${C.lineStrong}`, borderRadius: 6,
            background: C.card, fontFamily: FONT_MONO, fontSize: 13, color: C.ink,
          }}
        />
        {!isToday && (
          <button onClick={jumpToToday}
            style={{
              padding: '9px 14px', border: `1px solid ${C.lineStrong}`, borderRadius: 6,
              background: C.card, cursor: 'pointer', fontSize: 12, fontWeight: 500,
              color: C.ink, display: 'flex', alignItems: 'center', gap: 6, fontFamily: FONT_BODY,
            }}>
            Jump to today
          </button>
        )}
        <button onClick={() => setRefreshKey(k => k + 1)}
          title="Refresh from database"
          style={{
            padding: '9px 14px', border: `1px solid ${C.lineStrong}`, borderRadius: 6,
            background: C.card, cursor: 'pointer', fontSize: 12, fontWeight: 500,
            color: C.ink, display: 'flex', alignItems: 'center', gap: 6, fontFamily: FONT_BODY,
          }}>
          <RotateCcw size={14} /> Refresh
        </button>
        {existingRecord && (
          <div style={{
            display: 'flex', alignItems: 'center', gap: 6,
            padding: '6px 10px', borderRadius: 4,
            backgroundColor: C.successBg, color: C.success,
            fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.12em',
          }}>
            <CheckCircle2 size={13} /> Saved
            {existingRecord.editCount > 0 && <span style={{ color: C.mute, fontWeight: 500, textTransform: 'none', letterSpacing: 0 }}>· edited {existingRecord.editCount}×</span>}
          </div>
        )}
      </div>

      <>
        {hasStartedDay && (
        <div style={{
          padding: '10px 16px', backgroundColor: C.softBg, borderRadius: 8, marginBottom: 20,
          fontSize: 12, color: C.mute, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
        }}>
          <Info size={14} />
          <span>
            Started by <span style={{ color: C.ink, fontWeight: 600 }}>{existingRecord?.startName || startName}</span>
            {' · '}Float: <span style={{ color: C.ink, fontWeight: 600, fontFamily: FONT_MONO }}>{formatMoney(existingRecord?.startFloat ?? (startFloat ? parseFloat(startFloat) : 0))}</span>
            {isPastDate && (
              <button onClick={openStartDay}
                title="Edit float"
                style={{
                  background: 'none', border: 'none', padding: '2px 6px', marginLeft: 4,
                  color: C.cardinal, cursor: 'pointer', fontSize: 11, fontWeight: 600,
                  fontFamily: FONT_BODY, borderRadius: 4,
                }}>
                Edit
              </button>
            )}
            {existingRecord?.signature && <><span> · </span>Signed by <span style={{ color: C.ink, fontWeight: 600 }}>{existingRecord.signature}</span></>}
            <span> · </span>Total on file: <span style={{ color: C.ink, fontWeight: 600, fontFamily: FONT_MONO }}>{formatMoney(existingRecord?.total || dailyTotal)}</span>
          </span>
          {isDirty && <span style={{ color: C.cardinal, fontWeight: 600 }}>· Unsaved changes</span>}
          {undoSnapshot && (
            <>
              <span style={{ color: C.faint, userSelect: 'none' }}>|</span>
              <button onClick={undoReset}
                style={{
                  background: 'none', border: 'none', padding: 0, margin: 0,
                  color: C.cardinal, cursor: 'pointer',
                  fontSize: 12, fontWeight: 600, fontFamily: FONT_BODY,
                  textDecoration: 'underline', textUnderlineOffset: 3,
                  display: 'inline-flex', alignItems: 'center', gap: 4,
                }}>
                <RotateCcw size={12} strokeWidth={2.2} />
                Undo reset
              </button>
            </>
          )}
        </div>
      )}

      {!existingRecord && undoSnapshot && (
        <div style={{
          padding: '10px 16px', backgroundColor: '#FBF4E8', border: `1px solid #E8D9B6`, borderRadius: 8, marginBottom: 20,
          fontSize: 12, color: '#946C20', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
        }}>
          <Info size={14} />
          <span>Counters were reset.</span>
          <span style={{ color: '#C9A34A', userSelect: 'none' }}>|</span>
          <button onClick={undoReset}
            style={{
              background: 'none', border: 'none', padding: 0, margin: 0,
              color: '#946C20', cursor: 'pointer',
              fontSize: 12, fontWeight: 600, fontFamily: FONT_BODY,
              textDecoration: 'underline', textUnderlineOffset: 3,
              display: 'inline-flex', alignItems: 'center', gap: 4,
            }}>
            <RotateCcw size={12} strokeWidth={2.2} />
            Undo reset
          </button>
        </div>
)}

      {!loaded && (
        <div style={{ padding: 60, textAlign: 'center', color: C.mute }}>Loading…</div>
      )}

      {loaded && !hasStartedDay && (
        <div style={{
          padding: '60px 40px', textAlign: 'center',
          background: C.card, border: `1px solid ${C.line}`, borderRadius: 16,
          maxWidth: 500, margin: '40px auto',
        }}>
          <div style={{
            width: 64, height: 64, borderRadius: 16,
            background: `linear-gradient(135deg, ${C.cardinal} 0%, ${C.cardinalDark} 100%)`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 24px',
          }}>
            <DollarSign size={28} color="#FAF8F3" />
          </div>
          <div style={{ fontFamily: FONT_DISPLAY, fontSize: 26, fontWeight: 500, color: C.ink, marginBottom: 8 }}>
            Start Your Day
          </div>
          <p style={{ color: C.mute, fontSize: 14, maxWidth: 340, margin: '0 auto 24px', lineHeight: 1.6 }}>
            Enter your name and starting float to begin tallying sales for the day.
          </p>
          <button onClick={openStartDay}
            style={{
              padding: '14px 28px', background: C.cardinal, color: '#FAF8F3',
              border: 'none', borderRadius: 8, cursor: 'pointer',
              fontSize: 15, fontWeight: 600, fontFamily: FONT_BODY,
              display: 'inline-flex', alignItems: 'center', gap: 10,
              boxShadow: '0 4px 12px rgba(139,29,47,0.25)',
            }}>
            <DollarSign size={18} />
            Start Day
          </button>
        </div>
      )}
      {loaded && hasStartedDay && (
        <>
          <div className="fdsa-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(460px, 1fr))', gap: 20 }}>
            {CATEGORIES.map(cat => (
              <CategoryCard key={cat.id} category={cat} quantities={quantities} bump={bump} />
            ))}
          </div>

          <div style={{
            marginTop: 20, padding: '20px 24px',
            background: C.card, border: `1px solid ${C.line}`, borderRadius: 12,
          }}>
            <div style={{
              fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.25em',
              color: C.mute, fontWeight: 600, marginBottom: 14,
            }}>
              Notes
            </div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <button onClick={() => setNoteModalOpen('payout')}
                style={{
                  padding: '12px 18px', background: C.softBg, color: C.ink,
                  border: `1px solid ${C.line}`, borderRadius: 8, cursor: 'pointer',
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                <Wallet size={16} /> Payout
              </button>
              <button onClick={() => setNoteModalOpen('clear')}
                style={{
                  padding: '12px 18px', background: C.softBg, color: C.ink,
                  border: `1px solid ${C.line}`, borderRadius: 8, cursor: 'pointer',
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                <Wallet size={16} /> Clear
              </button>
              <button onClick={() => setNoteModalOpen('custom')}
                style={{
                  padding: '12px 18px', background: C.softBg, color: C.ink,
                  border: `1px solid ${C.line}`, borderRadius: 8, cursor: 'pointer',
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                <FileText size={16} /> Custom
              </button>
            </div>
            {(notes.payouts.length > 0 || notes.clears.length > 0 || notes.custom.length > 0) && (
              <div style={{ marginTop: 16, display: 'flex', flexDirection: 'column', gap: 8 }}>
                {notes.payouts.map(n => (
                  <div key={n.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 12px', background: C.softBg, borderRadius: 6, fontSize: 13 }}>
                    <span style={{ color: C.danger }}>Payout: {formatMoney(n.value)}</span>
                    <button onClick={() => removeNote('payouts', n.id)} style={{ background: 'none', border: 'none', color: C.mute, cursor: 'pointer', padding: 4 }}>×</button>
                  </div>
                ))}
                {notes.clears.map(n => (
                  <div key={n.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 12px', background: C.softBg, borderRadius: 6, fontSize: 13 }}>
                    <span style={{ color: C.success }}>Clear: {formatMoney(n.value)}</span>
                    <button onClick={() => removeNote('clears', n.id)} style={{ background: 'none', border: 'none', color: C.mute, cursor: 'pointer', padding: 4 }}>×</button>
                  </div>
                ))}
                {notes.custom.map(n => (
                  <div key={n.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 12px', background: C.softBg, borderRadius: 6, fontSize: 13 }}>
                    <span>Note: {n.value}</span>
                    <button onClick={() => removeNote('custom', n.id)} style={{ background: 'none', border: 'none', color: C.mute, cursor: 'pointer', padding: 4 }}>×</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="fdsa-total-bar" style={{
            marginTop: 28, padding: '24px 28px',
            background: `linear-gradient(135deg, ${C.ink} 0%, #1e1813 100%)`,
            color: C.bg, borderRadius: 12,
            display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 24, flexWrap: 'wrap',
          }}>
            <div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.3em', color: C.gold, fontWeight: 600 }}>
                Daily Total · {itemsCount} item{itemsCount===1?'':'s'} sold
              </div>
              <div style={{ fontFamily: FONT_DISPLAY, fontSize: 44, fontWeight: 500, letterSpacing: '-0.02em', marginTop: 4, lineHeight: 1 }}>
                {formatMoney(dailyTotal)}
              </div>
            </div>
            <div className="fdsa-footer-buttons" style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <button onClick={resetCounters}
                disabled={itemsCount === 0}
                style={{
                  padding: '12px 18px',
                  background: 'transparent', color: C.bg,
                  border: `1px solid rgba(250,248,243,0.3)`, borderRadius: 8,
                  cursor: itemsCount === 0 ? 'not-allowed' : 'pointer',
                  opacity: itemsCount === 0 ? 0.4 : 1,
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                <RotateCcw size={14} /> Reset counters
              </button>
              <button onClick={quickSave}
                style={{
                  padding: '12px 18px',
                  background: 'transparent', color: C.bg,
                  border: `1px solid rgba(250,248,243,0.3)`, borderRadius: 8,
                  cursor: 'pointer',
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                <Save size={14} /> Save
              </button>
              {(existingRecord || isPastDate) && (
                <button onClick={() => setReceiptView({
                  date: workingDate,
                  quantities,
                  sales,
                  total: dailyTotal,
                  signature: existingRecord?.signature || null,
                  notes,
                })}
                style={{
                  padding: '12px 18px',
                  background: 'transparent', color: C.bg,
                  border: `1px solid rgba(250,248,243,0.3)`, borderRadius: 8,
                  cursor: 'pointer',
                  fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                }}>
                  <Printer size={14} /> Receipt
                </button>
              )}
              <button onClick={openEndDay}
                disabled={itemsCount === 0}
                style={{
                  padding: '12px 22px',
                  background: itemsCount === 0 ? 'rgba(201,163,74,0.3)' : C.gold,
                  color: itemsCount === 0 ? 'rgba(250,248,243,0.5)' : C.ink,
                  border: 'none', borderRadius: 8,
                  cursor: itemsCount === 0 ? 'not-allowed' : 'pointer',
                  fontSize: 14, fontWeight: 600, fontFamily: FONT_BODY,
                  display: 'flex', alignItems: 'center', gap: 8,
                  letterSpacing: '0.01em',
                }}>
                End Day & Save
              </button>
            </div>
          </div>
        </>
        )}
      </>
    </div>
  );
}

function CategoryCard({ category, quantities, bump }) {
  const Icon = category.icon;
  const catTotal = category.items.reduce((s, it) => {
    if (it.price === 'TBA') return s;
    return s + quantities[it.id] * it.price;
  }, 0);
  const catCount = category.items.reduce((s, it) => s + quantities[it.id], 0);

  return (
    <section className="category-card" style={{
      background: C.card, border: `1px solid ${C.line}`, borderRadius: 12,
      overflow: 'hidden',
      boxShadow: catCount > 0 ? '0 2px 16px rgba(43,35,25,0.04)' : 'none',
      transition: 'box-shadow .2s',
    }}>
      <header style={{
        padding: '16px 20px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        borderBottom: `1px solid ${C.line}`,
        background: C.softBg,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 34, height: 34, borderRadius: 8,
            background: C.card, border: `1px solid ${C.lineStrong}`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: C.cardinal,
          }}>
            <Icon size={17} strokeWidth={1.8} />
          </div>
          <h2 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 19, fontWeight: 500, color: C.ink, letterSpacing: '-0.01em' }}>
            {category.label}
          </h2>
        </div>
        <div style={{ textAlign: 'right' }}>
          <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: '0.22em', color: C.mute, fontWeight: 600 }}>Subtotal</div>
          <div style={{ fontFamily: FONT_MONO, fontSize: 15, fontWeight: 600, color: catCount > 0 ? C.cardinal : C.ink, marginTop: 1 }}>
            {formatMoney(catTotal)}
          </div>
        </div>
      </header>
      <div>
        {category.items.map((item, i) => (
          <ItemRow
            key={item.id} item={item}
            quantity={quantities[item.id]}
            onInc={() => bump(item.id, +1)}
            onDec={() => bump(item.id, -1)}
            isLast={i === category.items.length - 1}
          />
        ))}
      </div>
    </section>
  );
}

function ItemRow({ item, quantity, onInc, onDec, isLast }) {
  const subtotal = item.price === 'TBA' ? 'TBA' : quantity * item.price;
  const canDec = quantity > 0;
  return (
    <div style={{
      display: 'flex', alignItems: 'center',
      padding: '16px 20px',
      borderBottom: isLast ? 'none' : `1px solid ${C.line}`,
      gap: 16,
    }}>
      <div style={{ flex: '1 1 140px', minWidth: 0 }}>
        <div style={{ fontSize: 15, fontWeight: 500, color: C.ink }}>{item.label}</div>
        <div style={{ marginTop: 3, fontSize: 11, color: C.mute, fontFamily: FONT_MONO }}>
          {formatMoney(item.price)}
        </div>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
        <button
          onClick={onDec}
          disabled={!canDec}
          aria-label={`Decrease ${item.label}`}
          style={{
            width: 40, height: 40, borderRadius: '50%',
            border: `1px solid ${canDec ? C.lineStrong : C.line}`,
            background: canDec ? C.card : C.softBg,
            color: canDec ? C.cardinal : C.faint,
            cursor: canDec ? 'pointer' : 'not-allowed',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            transition: 'all .12s',
          }}
        >
          <Minus size={16} strokeWidth={2.5} />
        </button>
        <div style={{
          minWidth: 44, textAlign: 'center',
          fontFamily: FONT_DISPLAY, fontSize: 26, fontWeight: 500,
          fontVariantNumeric: 'tabular-nums',
          color: quantity > 0 ? C.ink : C.faint,
          lineHeight: 1,
        }}>
          {quantity}
        </div>
        <button
          onClick={onInc}
          aria-label={`Increase ${item.label}`}
          style={{
            width: 40, height: 40, borderRadius: '50%',
            border: 'none',
            background: C.cardinal, color: '#FAF8F3',
            cursor: 'pointer',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            boxShadow: '0 2px 6px rgba(139,29,47,0.25)',
            transition: 'transform .1s',
          }}
          onMouseDown={e => e.currentTarget.style.transform = 'scale(0.94)'}
          onMouseUp={e => e.currentTarget.style.transform = 'scale(1)'}
          onMouseLeave={e => e.currentTarget.style.transform = 'scale(1)'}
        >
          <Plus size={16} strokeWidth={2.5} />
        </button>
      </div>

      <div style={{ width: 94, textAlign: 'right', flexShrink: 0 }}>
        <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: '0.2em', color: C.mute, fontWeight: 600 }}>Current</div>
        <div style={{
          marginTop: 2,
          fontFamily: FONT_MONO, fontSize: 16, fontWeight: 600,
          color: subtotal === 'TBA' ? C.gold : (subtotal > 0 ? C.cardinal : C.ink),
          fontVariantNumeric: 'tabular-nums',
        }}>
          {formatMoney(subtotal)}
        </div>
      </div>
    </div>
  );
}

function PastView({ days, onOpen, onDelete, onGoEditor }) {
  const grandTotal = days.reduce((s,d) => s + (d.total || 0), 0);
  return (
    <div className="fdsa-fade">
      <div style={{ display: 'flex', alignItems: 'end', justifyContent: 'space-between', marginBottom: 22, gap: 20, flexWrap: 'wrap' }}>
        <div>
          <h2 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 30, fontWeight: 500, letterSpacing: '-0.015em' }}>Past Days</h2>
          <p style={{ margin: '6px 0 0', color: C.mute, fontSize: 14, maxWidth: 620 }}>
            Every saved day lives here. Open a record to edit, or export to Excel.
          </p>
        </div>
        {days.length > 0 && (
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.28em', color: C.mute, fontWeight: 600 }}>
              All-time recorded · {days.length} day{days.length===1?'':'s'}
            </div>
            <div style={{ fontFamily: FONT_DISPLAY, fontSize: 26, fontWeight: 500, color: C.ink }}>
              {formatMoney(grandTotal)}
            </div>
          </div>
        )}
      </div>

      {days.length === 0 ? (
        <div style={{
          padding: '64px 32px', textAlign: 'center',
          background: C.card, border: `1px dashed ${C.lineStrong}`, borderRadius: 12,
        }}>
          <div style={{ fontFamily: FONT_DISPLAY, fontSize: 20, color: C.ink, marginBottom: 6 }}>No saved days yet</div>
          <p style={{ color: C.mute, fontSize: 14, maxWidth: 380, margin: '0 auto 18px' }}>
            Once you close out your first day, records will appear here.
          </p>
          <button onClick={onGoEditor}
            style={{
              padding: '10px 18px', background: C.cardinal, color: '#FAF8F3',
              border: 'none', borderRadius: 6, cursor: 'pointer',
              fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
              display: 'inline-flex', alignItems: 'center', gap: 8,
            }}>
            Start a daily tally <ChevronRight size={14} />
          </button>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {days.map(d => <PastDayRow key={d.date} day={d} onOpen={() => onOpen(d.date)} onDelete={() => onDelete(d.date)} />)}
        </div>
      )}
    </div>
  );
}

function PastDayRow({ day, onOpen, onDelete }) {
  const itemCount = totalItemsSold(day.quantities || {});
  return (
    <div className="past-day-row" style={{
      display: 'flex', alignItems: 'center', gap: 20,
      padding: '16px 20px', background: C.card,
      border: `1px solid ${C.line}`, borderRadius: 10,
      transition: 'border-color .15s, box-shadow .15s',
    }}
    onMouseEnter={e => {
      e.currentTarget.style.borderColor = C.lineStrong;
      e.currentTarget.style.boxShadow = '0 2px 8px rgba(43,35,25,0.04)';
    }}
    onMouseLeave={e => {
      e.currentTarget.style.borderColor = C.line;
      e.currentTarget.style.boxShadow = 'none';
    }}
    >
      <div style={{ width: 80, flexShrink: 0 }}>
        <div style={{ fontFamily: FONT_DISPLAY, fontSize: 28, fontWeight: 500, color: C.ink, lineHeight: 1 }}>
          {day.date.slice(8,10)}
        </div>
        <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.2em', color: C.mute, fontWeight: 600, marginTop: 4 }}>
          {(() => {
            const [y,m,d] = day.date.split('-').map(Number);
            return new Date(y,m-1,d).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          })()}
        </div>
      </div>
      <div className="day-info" style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 15, fontWeight: 500, color: C.ink }}>
          {formatDateLong(day.date).split(',').slice(0,1)[0]}
        </div>
        <div style={{ fontSize: 12, color: C.mute, marginTop: 3, display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <span>Signed by <span style={{ color: C.ink, fontWeight: 500 }}>{day.signature || '—'}</span></span>
          <span>·</span>
          <span>{itemCount} item{itemCount===1?'':'s'}</span>
          {day.editCount > 0 && <><span>·</span><span style={{ color: C.gold, fontWeight: 500 }}>Edited {day.editCount}×</span></>}
        </div>
      </div>
      <div style={{ textAlign: 'right', minWidth: 110 }}>
        <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: '0.22em', color: C.mute, fontWeight: 600 }}>Total</div>
        <div style={{ fontFamily: FONT_MONO, fontSize: 18, fontWeight: 600, color: C.ink, marginTop: 1 }}>
          {formatMoney(day.total)}
        </div>
      </div>
      <div className="day-actions" style={{ display: 'flex', gap: 6 }}>
        <button onClick={onOpen}
          style={{
            padding: '9px 14px', background: C.card, color: C.ink,
            border: `1px solid ${C.lineStrong}`, borderRadius: 6, cursor: 'pointer',
            fontSize: 12, fontWeight: 500, fontFamily: FONT_BODY,
            display: 'flex', alignItems: 'center', gap: 6,
          }}>
          Open <ChevronRight size={13} />
        </button>
        <button onClick={onDelete}
          aria-label="Delete record"
          style={{
            padding: 9, background: C.card, color: C.danger,
            border: `1px solid ${C.line}`, borderRadius: 6, cursor: 'pointer',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
          <Trash2 size={14} />
        </button>
      </div>
    </div>
  );
}

function ExportView({ start, end, setStart, setEnd, presetRange, pastDays, exporting, runExport }) {
  const inRange = pastDays.filter(d => d.date >= start && d.date <= end);
  const rangeTotal = inRange.reduce((s,d) => s + (d.total || 0), 0);

  return (
    <div className="fdsa-fade">
      <div style={{ marginBottom: 26 }}>
        <h2 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 30, fontWeight: 500, letterSpacing: '-0.015em' }}>Export Report</h2>
        <p style={{ margin: '6px 0 0', color: C.mute, fontSize: 14, maxWidth: 620 }}>
          Generate an Excel workbook for any date range.
        </p>
      </div>

      <div className="fdsa-export-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
        <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 24 }}>
          <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.25em', color: C.mute, fontWeight: 600, marginBottom: 14 }}>
            Date range
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
            <DateField label="Start" value={start} onChange={setStart} />
            <DateField label="End"   value={end}   onChange={setEnd} />
          </div>

          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 16 }}>
            {[
              { id: 'thisMonth',  label: 'This month' },
              { id: 'lastMonth',  label: 'Last month' },
              { id: 'last7',      label: 'Last 7 days' },
              { id: 'thisYear',   label: 'This year' },
            ].map(p => (
              <button key={p.id} onClick={() => presetRange(p.id)}
                style={{
                  padding: '7px 12px', background: C.softBg, color: C.ink,
                  border: `1px solid ${C.line}`, borderRadius: 999, cursor: 'pointer',
                  fontSize: 11, fontWeight: 500, fontFamily: FONT_BODY,
                }}>
                {p.label}
              </button>
            ))}
          </div>

          <div style={{ marginTop: 22, padding: '14px 16px', background: C.softBg, borderRadius: 8, fontSize: 12, color: C.mute }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
              <CalendarRange size={14} />
              <span style={{ color: C.ink, fontWeight: 600, fontSize: 13 }}>Preview</span>
            </div>
            <div>{inRange.length} saved day{inRange.length===1?'':'s'} in range</div>
            <div>Range total: <span style={{ color: C.ink, fontWeight: 600, fontFamily: FONT_MONO }}>{formatMoney(rangeTotal)}</span></div>
          </div>

          <button onClick={runExport} disabled={exporting || inRange.length === 0}
            style={{
              width: '100%', marginTop: 20, padding: '14px 18px',
              background: inRange.length === 0 ? C.softBg : C.cardinal,
              color: inRange.length === 0 ? C.faint : '#FAF8F3',
              border: 'none', borderRadius: 8,
              cursor: (exporting || inRange.length === 0) ? 'not-allowed' : 'pointer',
              fontSize: 14, fontWeight: 600, fontFamily: FONT_BODY,
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
              letterSpacing: '0.01em',
              boxShadow: inRange.length > 0 ? '0 2px 8px rgba(139,29,47,0.2)' : 'none',
            }}>
            <FileDown size={16} />
            {exporting ? 'Exporting…' : 'Download Excel workbook'}
          </button>
        </div>

        <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 24 }}>
          <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.25em', color: C.mute, fontWeight: 600, marginBottom: 14 }}>
            Workbook contents
          </div>
          <SheetDescription
            index="1" name="Summary"
            desc="Each day + daily total, followed by the grand total."
          />
          <SheetDescription
            index="2" name="Itemized"
            desc="Line-by-line: date, category, item, unit price, quantity, subtotal."
            last
          />
        </div>
      </div>
    </div>
  );
}

function DateField({ label, value, onChange }) {
  return (
    <label style={{ display: 'block' }}>
      <div style={{ fontSize: 11, color: C.mute, fontWeight: 500, marginBottom: 6 }}>{label}</div>
      <input type="date" value={value} onChange={e => onChange(e.target.value)}
        style={{
          width: '100%', padding: '10px 12px',
          border: `1px solid ${C.lineStrong}`, borderRadius: 6,
          background: C.card, fontSize: 14, fontFamily: FONT_MONO, color: C.ink,
        }}
      />
    </label>
  );
}

function SheetDescription({ index, name, desc, last }) {
  return (
    <div style={{
      display: 'flex', gap: 14, padding: '12px 0',
      borderBottom: last ? 'none' : `1px solid ${C.line}`,
    }}>
      <div style={{
        width: 28, height: 28, borderRadius: 6,
        background: C.softBg, color: C.cardinal,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        fontFamily: FONT_MONO, fontSize: 12, fontWeight: 600, flexShrink: 0,
      }}>{index}</div>
      <div>
        <div style={{ fontSize: 14, fontWeight: 600, color: C.ink }}>{name}</div>
        <div style={{ fontSize: 12, color: C.mute, marginTop: 2, lineHeight: 1.5 }}>{desc}</div>
      </div>
    </div>
  );
}

function SignatureModal({ workingDate, dailyTotal, itemsCount, existing, value, setValue, onCancel, onSubmit }) {
  return (
    <ModalShell onCancel={onCancel}>
      <div style={{ padding: 28, maxWidth: 480 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 18 }}>
          <div style={{
            width: 40, height: 40, borderRadius: 10,
            background: C.softBg, color: C.cardinal,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <Save size={20} />
          </div>
          <div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.25em', color: C.mute, fontWeight: 600 }}>
              Staff signature required
            </div>
            <h3 style={{ margin: '2px 0 0', fontFamily: FONT_DISPLAY, fontSize: 22, fontWeight: 500 }}>
              {existing ? 'Update day record' : 'Close out the day'}
            </h3>
          </div>
        </div>

        <div style={{
          padding: 14, background: C.softBg, borderRadius: 8, marginBottom: 18,
          display: 'grid', gridTemplateColumns: 'auto 1fr', gap: '6px 12px', fontSize: 13,
        }}>
          <div style={{ color: C.mute }}>Date</div>
          <div style={{ color: C.ink, fontWeight: 500 }}>{formatDateLong(workingDate)}</div>
          <div style={{ color: C.mute }}>Items sold</div>
          <div style={{ color: C.ink, fontWeight: 500 }}>{itemsCount}</div>
          <div style={{ color: C.mute }}>Total</div>
          <div style={{ color: C.cardinal, fontWeight: 600, fontFamily: FONT_MONO }}>{formatMoney(dailyTotal)}</div>
        </div>

        <label style={{ display: 'block' }}>
          <div style={{ fontSize: 12, color: C.mute, fontWeight: 500, marginBottom: 6 }}>
            Your name (signature)
          </div>
          <input
            type="text" autoFocus value={value}
            onChange={e => setValue(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') onSubmit(); }}
            placeholder="Enter your name"
            style={{
              width: '100%', padding: '12px 14px',
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              fontSize: 15, fontFamily: FONT_BODY, color: C.ink,
              background: C.card,
            }}
          />
        </label>

        <div style={{ display: 'flex', gap: 10, marginTop: 22 }}>
          <button onClick={onCancel}
            style={{
              flex: '0 0 auto', padding: '11px 16px',
              background: C.card, color: C.ink,
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              cursor: 'pointer', fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
            }}>
            Cancel
          </button>
          <button onClick={onSubmit}
            disabled={!value.trim()}
            style={{
              flex: 1, padding: '11px 16px',
              background: value.trim() ? C.cardinal : C.softBg,
              color: value.trim() ? '#FAF8F3' : C.faint,
              border: 'none', borderRadius: 8,
              cursor: value.trim() ? 'pointer' : 'not-allowed',
              fontSize: 14, fontWeight: 600, fontFamily: FONT_BODY,
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
            }}>
            <Save size={14} />
            {existing ? 'Save changes' : 'Sign & save day'}
          </button>
        </div>
      </div>
    </ModalShell>
  );
}

function StartDayModal({ workingDate, startName, setStartName, startFloat, setStartFloat, onCancel, onSave }) {
  const isValid = startName.trim() && startFloat && !isNaN(parseFloat(startFloat)) && parseFloat(startFloat) >= 0;
  
  return (
    <ModalShell onCancel={onCancel}>
      <div style={{ padding: 28, maxWidth: 480 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 18 }}>
          <div style={{
            width: 40, height: 40, borderRadius: 10,
            background: C.softBg, color: C.cardinal,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <DollarSign size={20} />
          </div>
          <div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.25em', color: C.mute, fontWeight: 600 }}>
              Start the day
            </div>
            <h3 style={{ margin: '2px 0 0', fontFamily: FONT_DISPLAY, fontSize: 22, fontWeight: 500 }}>
              Enter your details
            </h3>
          </div>
        </div>

        <div style={{
          padding: 14, background: C.softBg, borderRadius: 8, marginBottom: 18,
          fontSize: 12, color: C.mute,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <Calendar size={14} />
            <span style={{ color: C.ink, fontWeight: 500 }}>{formatDateLong(workingDate)}</span>
          </div>
        </div>

        <label style={{ display: 'block', marginBottom: 16 }}>
          <div style={{ fontSize: 12, color: C.mute, fontWeight: 500, marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
            <User size={14} /> Your name
          </div>
          <input
            type="text" autoFocus value={startName}
            onChange={e => setStartName(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter' && isValid) onSave(); }}
            placeholder="Enter your name"
            style={{
              width: '100%', padding: '12px 14px',
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              fontSize: 15, fontFamily: FONT_BODY, color: C.ink,
              background: C.card,
            }}
          />
        </label>

        <label style={{ display: 'block' }}>
          <div style={{ fontSize: 12, color: C.mute, fontWeight: 500, marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
            <DollarSign size={14} /> Starting float (change money)
          </div>
          <div style={{ position: 'relative' }}>
            <span style={{
              position: 'absolute', left: 14, top: '50%', transform: 'translateY(-50%)',
              color: C.mute, fontSize: 15, fontFamily: FONT_MONO,
            }}>$</span>
            <input
              type="number" min="0" step="0.01" value={startFloat}
              onChange={e => setStartFloat(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter' && isValid) onSave(); }}
              placeholder="0.00"
              style={{
                width: '100%', padding: '12px 14px 12px 28px',
                border: `1px solid ${C.lineStrong}`, borderRadius: 8,
                fontSize: 15, fontFamily: FONT_MONO, color: C.ink,
                background: C.card, boxSizing: 'border-box',
              }}
            />
          </div>
        </label>

        <div style={{ display: 'flex', gap: 10, marginTop: 22 }}>
          <button onClick={onCancel}
            style={{
              flex: '0 0 auto', padding: '11px 16px',
              background: C.card, color: C.ink,
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              cursor: 'pointer', fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
            }}>
            Cancel
          </button>
          <button onClick={onSave}
            disabled={!isValid}
            style={{
              flex: 1, padding: '11px 16px',
              background: isValid ? C.cardinal : C.softBg,
              color: isValid ? '#FAF8F3' : C.faint,
              border: 'none', borderRadius: 8,
              cursor: isValid ? 'pointer' : 'not-allowed',
              fontSize: 14, fontWeight: 600, fontFamily: FONT_BODY,
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
            }}>
            <Save size={14} />
            Start Day
          </button>
        </div>
      </div>
    </ModalShell>
  );
}

function ConfirmModal({ title, body, confirmLabel, danger, onConfirm, onCancel }) {
  return (
    <ModalShell onCancel={onCancel}>
      <div style={{ padding: 28, maxWidth: 420 }}>
        <div style={{ display: 'flex', gap: 14, marginBottom: 14 }}>
          <div style={{
            width: 40, height: 40, borderRadius: 10,
            background: danger ? C.dangerBg : C.softBg,
            color: danger ? C.danger : C.ink,
            display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
          }}>
            <AlertCircle size={20} />
          </div>
          <div>
            <h3 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 20, fontWeight: 500 }}>{title}</h3>
            <p style={{ margin: '6px 0 0', color: C.mute, fontSize: 13, lineHeight: 1.5 }}>{body}</p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
          <button onClick={onCancel}
            style={{
              padding: '10px 16px', background: C.card, color: C.ink,
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              cursor: 'pointer', fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
            }}>
            Cancel
          </button>
          <button onClick={onConfirm}
            style={{
              padding: '10px 16px',
              background: danger ? C.danger : C.cardinal, color: '#FAF8F3',
              border: 'none', borderRadius: 8,
              cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: FONT_BODY,
            }}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </ModalShell>
  );
}

function NoteModal({ type, onCancel, onSave }) {
  const [value, setValue] = useState('');
  const isAmount = type === 'payout' || type === 'clear';
  const isValid = isAmount 
    ? value && !isNaN(parseFloat(value)) && parseFloat(value) > 0
    : value.trim().length > 0;

  const handleSave = () => {
    if (!isValid) return;
    if (isAmount) {
      onSave(parseFloat(value));
    } else {
      onSave(value.trim());
    }
  };

  const labels = { payout: 'Payout', clear: 'Clear', custom: 'Custom Note' };
  const placeholders = { payout: '0.00', clear: '0.00', custom: 'Enter note...' };

  return (
    <ModalShell onCancel={onCancel}>
      <div style={{ padding: 28, maxWidth: 420 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 18 }}>
          <div style={{
            width: 40, height: 40, borderRadius: 10,
            background: C.softBg, color: C.ink,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            {type === 'custom' ? <FileText size={20} /> : <Wallet size={20} />}
          </div>
          <div>
            <h3 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 20, fontWeight: 500 }}>
              Add {labels[type]}
            </h3>
          </div>
        </div>

        {isAmount ? (
          <div style={{ position: 'relative' }}>
            <span style={{
              position: 'absolute', left: 14, top: '50%', transform: 'translateY(-50%)',
              color: C.mute, fontSize: 16, fontFamily: FONT_MONO,
            }}>$</span>
            <input
              type="number" min="0" step="0.01" autoFocus
              value={value} onChange={e => setValue(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter' && isValid) handleSave(); }}
              placeholder="0.00"
              style={{
                width: '100%', padding: '14px 14px 14px 32px',
                border: `1px solid ${C.lineStrong}`, borderRadius: 8,
                fontSize: 16, fontFamily: FONT_MONO, color: C.ink,
                background: C.card, boxSizing: 'border-box',
              }}
            />
          </div>
        ) : (
          <input
            type="text" autoFocus
            value={value} onChange={e => setValue(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter' && isValid) handleSave(); }}
            placeholder={placeholders[type]}
            style={{
              width: '100%', padding: '14px',
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              fontSize: 15, fontFamily: FONT_BODY, color: C.ink,
              background: C.card, boxSizing: 'border-box',
            }}
          />
        )}

        <div style={{ display: 'flex', gap: 10, marginTop: 22 }}>
          <button onClick={onCancel}
            style={{
              flex: '0 0 auto', padding: '11px 16px',
              background: C.card, color: C.ink,
              border: `1px solid ${C.lineStrong}`, borderRadius: 8,
              cursor: 'pointer', fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
            }}>
            Cancel
          </button>
          <button onClick={handleSave} disabled={!isValid}
            style={{
              flex: 1, padding: '11px 16px',
              background: isValid ? C.cardinal : C.softBg,
              color: isValid ? '#FAF8F3' : C.faint,
              border: 'none', borderRadius: 8,
              cursor: isValid ? 'pointer' : 'not-allowed',
              fontSize: 14, fontWeight: 600, fontFamily: FONT_BODY,
            }}>
            Add
          </button>
        </div>
      </div>
    </ModalShell>
  );
}

function ModalShell({ children, onCancel }) {
  return (
    <div className="no-print" onClick={onCancel}
      style={{
        position: 'fixed', inset: 0, zIndex: 100,
        background: 'rgba(43,35,25,0.4)',
        backdropFilter: 'blur(4px)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 20,
      }}>
      <div onClick={e => e.stopPropagation()}
        style={{
          background: C.card, borderRadius: 14,
          boxShadow: '0 20px 60px rgba(43,35,25,0.25)',
          animation: 'fdsaFade .2s ease',
          position: 'relative',
        }}>
        <button onClick={onCancel} aria-label="Close"
          style={{
            position: 'absolute', top: 12, right: 12,
            width: 28, height: 28, borderRadius: 6,
            background: 'transparent', border: 'none', cursor: 'pointer',
            color: C.mute, display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
          <X size={16} />
        </button>
        {children}
      </div>
    </div>
  );
}

function Toast({ kind, msg }) {
  const isErr = kind === 'error';
  return (
    <div style={{
      position: 'fixed', bottom: 24, right: 24, zIndex: 200,
      padding: '14px 18px', maxWidth: 380,
      background: isErr ? C.danger : C.ink,
      color: isErr ? '#FAF8F3' : C.bg,
      borderRadius: 10,
      boxShadow: '0 8px 28px rgba(43,35,25,0.3)',
      display: 'flex', alignItems: 'center', gap: 10,
      fontSize: 13, fontWeight: 500, fontFamily: FONT_BODY,
      animation: 'fdsaFade .25s ease',
    }}>
      {isErr ? <AlertCircle size={16} /> : <Check size={16} style={{ color: C.gold }} />}
      {msg}
    </div>
  );
}

function ReceiptView({ day, onClose }) {
  const notes = day.notes || { payouts: [], clears: [], custom: [] };
  const sales = day.sales || [];
  const sortedSales = [...sales].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  
  const itemMap = {};
  ALL_ITEMS.forEach(it => { itemMap[it.id] = it; });

  let salesWithItems;
  if (sales.length > 0) {
    salesWithItems = sortedSales.map(s => ({
      ...s,
      item: itemMap[s.itemId],
    }));
  } else {
    salesWithItems = Object.entries(day.quantities || {})
      .filter(([_, qty]) => qty > 0)
      .map(([id, qty]) => ({
        itemId: id,
        delta: qty,
        timestamp: day.lastEdited || day.savedAt,
        item: itemMap[id],
      }));
  }

  const handlePrint = () => {
    const doc = new jsPDF();
    let y = 20;
    
    doc.setFontSize(22);
    doc.setFont('helvetica', 'bold');
    doc.text('Receipt', 105, y, { align: 'center' });
    y += 10;
    
    doc.setFontSize(12);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100);
    doc.text(day.date + (day.signature ? ' · Signed by ' + day.signature : ' · Unsigned'), 105, y, { align: 'center' });
    y += 15;
    
    doc.setDrawColor(200);
    doc.line(20, y, 190, y);
    y += 10;
    
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(100);
    doc.text('ITEM', 20, y);
    doc.text('TIME', 160, y);
    y += 5;
    doc.line(20, y, 190, y);
    y += 8;
    
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(0);
    salesWithItems.forEach(s => {
      const time = new Date(s.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      const qty = (s.delta > 0 ? '+' : '') + s.delta;
      const name = s.item?.label || s.itemId;
      
      doc.setTextColor(90, 180, 106);
      doc.text(qty, 22, y);
      doc.setTextColor(0);
      doc.text(name, 35, y);
      doc.setTextColor(100);
      doc.text(time, 165, y);
      y += 8;
    });
    
    y += 5;
    doc.setFillColor(250, 248, 243);
    doc.roundedRect(20, y, 170, 20, 3, 3, 'F');
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0);
    doc.text('Total', 30, y + 13);
    doc.text('$' + (day.total || 0).toFixed(2), 170, y + 13, { align: 'right' });
    y += 30;
    
    if (day.signature) {
      doc.setDrawColor(200);
      doc.setLineDashPattern([3, 3], 0);
      doc.line(60, y, 150, y);
      y += 8;
      doc.setFontSize(9);
      doc.setTextColor(100);
      doc.text('Signed by', 105, y, { align: 'center' });
      y += 6;
      doc.setFontSize(14);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0);
      doc.text(day.signature, 105, y, { align: 'center' });
    }

    if (notes.payouts.length > 0 || notes.clears.length > 0 || notes.custom.length > 0) {
      y += 15;
      doc.setDrawColor(200);
      doc.setLineDashPattern([2, 2], 0);
      doc.line(20, y, 190, y);
      y += 8;
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(100);
      doc.text('NOTES', 20, y);
      y += 8;
      doc.setFont('helvetica', 'normal');
      notes.payouts.forEach(n => {
        doc.setTextColor(200, 50, 50);
        doc.text('Payout', 20, y);
        doc.setTextColor(0);
        doc.text('$' + n.value.toFixed(2), 170, y, { align: 'right' });
        y += 7;
      });
      notes.clears.forEach(n => {
        doc.setTextColor(50, 160, 80);
        doc.text('Clear', 20, y);
        doc.setTextColor(0);
        doc.text('$' + n.value.toFixed(2), 170, y, { align: 'right' });
        y += 7;
      });
      notes.custom.forEach(n => {
        doc.setTextColor(100);
        doc.text('Note', 20, y);
        doc.setTextColor(0);
        doc.text(n.value, 50, y);
        y += 7;
      });
    }
    
    doc.save(`Receipt_${day.date}.pdf`);
  };

  return (
    <div className="receipt-modal receipt-print" style={{ padding: 24, maxWidth: 500, maxHeight: '80vh', display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20, gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontFamily: FONT_DISPLAY, fontSize: 22, fontWeight: 500, letterSpacing: '-0.015em' }}>
            Receipt
          </h2>
          <p style={{ margin: '4px 0 0', color: C.mute, fontSize: 13 }}>
            {formatDateLong(day.date)} · {day.signature || 'Unsigned'}
          </p>
        </div>
        <button onClick={handlePrint}
          style={{
            padding: '8px 14px', background: C.ink, color: C.bg,
            border: 'none', borderRadius: 6, cursor: 'pointer',
            fontSize: 12, fontWeight: 500, fontFamily: FONT_BODY,
            display: 'flex', alignItems: 'center', gap: 6,
          }}>
          <Printer size={14} /> Download PDF
        </button>
      </div>

      <div style={{ flex: 1, overflow: 'auto', border: `1px solid ${C.line}`, borderRadius: 10 }}>
        <div style={{ padding: '12px 16px', borderBottom: `1px solid ${C.line}`, background: C.softBg, position: 'sticky', top: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.2em', color: C.mute }}>
            <span>Item</span>
            <span>Time</span>
          </div>
        </div>
        
        {salesWithItems.length === 0 ? (
          <div style={{ padding: 32, textAlign: 'center', color: C.mute }}>
            No sales recorded for this day.
          </div>
        ) : (
          <div>
            {salesWithItems.map((s, idx) => (
              <div className="item-row" key={idx} style={{ display: 'flex', justifyContent: 'space-between', padding: '10px 16px', borderBottom: `1px solid ${C.faint}` }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span className="item-qty" style={{ color: s.delta > 0 ? C.cardinal : C.danger, fontSize: 11, fontWeight: 600, minWidth: 24 }}>
                    {s.delta > 0 ? '+' : ''}{s.delta}
                  </span>
                  <span style={{ color: C.ink, fontSize: 13 }}>{s.item?.label || s.itemId}</span>
                </div>
                <span style={{ color: C.mute, fontSize: 12, fontFamily: FONT_MONO }}>
                  {new Date(s.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ marginTop: 16, padding: '14px 16px', background: C.softBg, borderRadius: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span style={{ fontSize: 13, fontWeight: 500, color: C.mute }}>Total</span>
        <span style={{ fontFamily: FONT_MONO, fontSize: 20, fontWeight: 600, color: C.ink }}>{formatMoney(day.total)}</span>
      </div>
      
      <div style={{ marginTop: 24, paddingTop: 16, borderTop: `1px dashed ${C.lineStrong}`, textAlign: 'center' }}>
        <div style={{ fontSize: 11, color: C.mute, marginBottom: 4 }}>Signed by</div>
        <div style={{ fontSize: 16, fontWeight: 500, color: C.ink, fontFamily: FONT_DISPLAY }}>{day.signature || '—'}</div>
      </div>

      {(notes.payouts.length > 0 || notes.clears.length > 0 || notes.custom.length > 0) && (
        <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px dashed ${C.lineStrong}` }}>
          <div style={{ fontSize: 11, color: C.mute, marginBottom: 12, textTransform: 'uppercase', letterSpacing: '0.2em', fontWeight: 600 }}>Notes</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {notes.payouts.map(n => (
              <div key={n.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 12px', background: '#FFF5F5', borderRadius: 6, fontSize: 13 }}>
                <span style={{ color: C.danger, fontWeight: 500 }}>Payout</span>
                <span style={{ fontFamily: FONT_MONO }}>{formatMoney(n.value)}</span>
              </div>
            ))}
            {notes.clears.map(n => (
              <div key={n.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 12px', background: '#F0FFF4', borderRadius: 6, fontSize: 13 }}>
                <span style={{ color: C.success, fontWeight: 500 }}>Clear</span>
                <span style={{ fontFamily: FONT_MONO }}>{formatMoney(n.value)}</span>
              </div>
            ))}
            {notes.custom.map(n => (
              <div key={n.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 12px', background: C.softBg, borderRadius: 6, fontSize: 13 }}>
                <span style={{ color: C.ink, fontWeight: 500 }}>Note</span>
                <span style={{ color: C.mute }}>{n.value}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}