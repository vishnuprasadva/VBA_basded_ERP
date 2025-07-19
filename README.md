# Excel VBA ERP System

This is a lightweight, Excel VBA-based ERP solution designed for small to mid-scale manufacturing processes. It automates workflows like work order creation, BOM generation, raw material tracking, job card creation, and material issuance — all through interactive Excel UserForms.

## 📦 Features

- 📋 **Work Order Management**  
  Create structured work orders with auto-filled customer, tender, and scope details.

- 📄 **Quotation Generator**  
  Generate printable PDF quotations using selected customer and tender information.

- 🧱 **Bill of Materials (BOM)**  
  Generate BOMs linked to work orders with component-level data including drawing numbers, quantities, and materials.

- 🛠️ **Job Card System**  
  Auto-generate job cards with multiple operation entries, workstation selection, and drawing references.

- 📑 **Raw Material Master**  
  Record raw materials with fields like name, type, condition, size, quantity, and test certificate number. Data is organized by DC numbers.

- 📦 **Material Issue Tracker**  
  Issue materials linked to a work order. Auto-fills fields like RMC, RMCC, size, and quantity. Supports up to 3 materials per voucher.

## 🗂️ File Structure

- `customer_master.xlsx` — Master list of customers.  
- `tender_master.xlsx` — Master list of tenders.  
- `material_master.xlsm` — BOM and material linking with work orders.  
- `raw_material_master.xlsm` — Entry form for raw material records.  
- `raw_records.xlsx` — Stores submitted raw material data.  
- `JobCard_Template.xlsx` — Format used to generate printable job cards.

## 🛠️ Technologies Used

- Microsoft Excel (Office 365 or 2016+)  
- VBA (Visual Basic for Applications)  
- Excel UserForms & Macros  
- PDF generation and Excel file automation

## 💡 Getting Started

1. Open the `.xlsm` files in Excel.
2. Enable macros when prompted.
3. Use the Home form to navigate through available actions.
4. Ensure linked files like `customer_master.xlsx` and `tender_master.xlsx` are in the same directory.

## 📌 Notes

- This system is designed for offline use within a local folder structure.
- Data is stored within `.xlsx` or `.xlsm` files, no external database is required.
- Make sure to keep file names consistent and in the same folder to avoid broken links.

---


