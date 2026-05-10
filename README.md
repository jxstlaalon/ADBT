# A & D's Bakery Tracker
### Daily Sales Tracking for Aalon & Deklon Bakery

A dedicated daily tally system designed for bakery front-desk operations. This tracker streamlines the tracking of bakery items, pastries, and refreshments with automated reporting, receipt generation, and secure cloud storage.

## 🚀 Features
- **Daily Sales Tally:** Real-time counter for Savoury, Specialty, Sweet, Cookie, and Drinks categories.
- **Dynamic Pricing:** Automatically calculates subtotals based on category-specific rates.
- **Sales Timestamps:** Each sale is recorded with timestamp for tracking when items were sold.
- **Receipt Generation:** Generate and download PDF receipts with full sales history.
- **Secure Architecture:** Integrated with Firebase Authentication and Firestore Security Rules.
- **Past Record Management:** Ability to view, edit, or delete historical sales data with edit-tracking.
- **Excel Exports:** Generates professional multi-sheet workbooks (.xlsx) including:
  - **Summary:** Daily totals and grand totals.
  - **Itemized:** Line-by-line breakdown of every transaction.

## 🛠 Tech Stack
- **Frontend:** React.js, Lucide-React (Icons)
- **Styling:** Custom CSS with warm ivory/bakery design tokens.
- **Backend:** Firebase (Firestore & Authentication)
- **Hosting:** Cloudflare Pages
- **Data Processing:** SheetJS (XLSX), jsPDF

## 🔐 Security & Privacy
- **Access Control:** Restricted to authorized personnel via Email/Password authentication.
- **Database Safety:** Firestore rules prevent unauthorized data modification or external API access.
- **Data Integrity:** Tracks edit counts and requires digital signatures for every saved day.

## 📜 License
**Proprietary - All Rights Reserved.**
Copyright (c) 2026 Aalon Peters.

This software is private. No permission is granted to any person to use, copy, modify, or distribute this software for any purpose. Authorized use is restricted to Aalon & Deklon Bakery operations.

---
*Created by Aalon Peters · 2026*