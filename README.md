# Issue Tracking Analytics Dashboard - VBA (Apps.csv)

## 🎯 What This Does

Apps.csv ke data par ek **fully interactive dashboard** banata hai with:
- ✅ 6 KPI cards (auto-update with slicers)
- ✅ 5 charts (Donut, Bar, Column)
- ✅ 4 interactive slicers (Issue Type, Severity, Status, Department)
- ✅ All 24 columns from Apps.csv supported

---

## 📋 Required Data Sheet Columns (24 columns - exact order from Apps.csv)

| # | Column Name | Sample Values |
|---|------------|---------------|
| 1 | Reporter | Person name |
| 2 | Reporter Identifier | SPBLRHO126 |
| 3 | Reporter Designation | Manager, Engineer |
| 4 | **Report Department** ⚠️ | Operations, Store Operations |
| 5 | Reporter Division | Karnataka |
| 6 | Reporter Sub Division | Bangalore |
| 7 | Reporter Location | HQ, BLR-DN-Haralur |
| 8 | Reported At | Date/Time |
| 9 | Issue ID | 6707 |
| 10 | Issue Title | Brief title |
| 11 | Issue Type | IT, Marketing, Repair and Maintenance |
| 12 | Severity | Critical, High, Medium, Low |
| 13 | Current Status | open, Closed |
| 14 | Issue Location | Where the issue is |
| 15 | Issue Description | Details |
| 16 | Resolver | Person who resolved |
| 17 | Resolver Identifier | Employee ID |
| 18 | Resolver Designation | IT executive |
| 19 | Resolver Department | Operations |
| 20 | Resolver Division | Karnataka |
| 21 | Resolver Sub Division | Bangalore |
| 22 | Resolver Location | Corporate Office |
| 23 | Resolved At | Resolution date |
| 24 | Resolved Remarks | Resolution notes |

> ⚠️ **CRITICAL:** Column #4 is **"Report Department"** (NOT "Reporter Department")

---

## 🚀 Setup Steps

### Step 1: Excel Trust Settings (REQUIRED for KPI auto-update!)
1. **File → Options → Trust Center → Trust Center Settings**
2. **Macro Settings**: "Enable all macros"
3. ✅ **Check: "Trust access to the VBA project object model"**
4. Click OK, restart Excel

### Step 2: Import Data
1. Open new Excel workbook
2. Save as **.xlsm** (macro-enabled)
3. Create sheet named **"Data"**
4. Import Apps.csv data into it (with headers in row 1)

### Step 3: Import VBA Module
1. Press **Alt + F11** (VBA Editor)
2. **File → Import File**
3. Select **IssueDashboardModule.bas**

### Step 4: Run Dashboard
1. Press **Alt + F8**
2. Select **BuildIssueDashboard**
3. Click **Run**

---

## 📊 KPI Cards (Auto-update with Slicers!)

| KPI | Calculation | Color |
|-----|-------------|-------|
| **Total Issues** | All filtered issues | Blue |
| **Open Issues** | Status = "open" | Orange |
| **Closed Issues** | Status = "Closed" | Green |
| **Critical/High** | Severity = Critical or High | Red |
| **Closure %** | Closed / Total | Purple |
| **Departments** | Department count | Navy |

### How Auto-Update Works:
1. User clicks a slicer (e.g., "Critical")
2. All pivot tables filter automatically
3. `Worksheet_PivotTableUpdate` event fires
4. `RefreshDashboard` runs
5. KPI text values update with filtered data ✨

---

## 📈 Charts

1. **Severity Distribution** - Donut chart (Critical/High/Medium/Low)
2. **Current Status Breakdown** - Bar chart (open vs Closed)
3. **Issues by Department** - Column chart (Operations, Store Operations)
4. **Top 10 Resolvers** - Bar chart
5. **Issues by Location** - Column chart (31 locations)

---

## 🎯 Slicers (4 Interactive Filters)

1. **Issue Type** - IT, Marketing, Repair and Maintenance
2. **Severity** - Critical, High, Medium, Low
3. **Current Status** - open, Closed
4. **Report Department** - Operations, Store Operations

---

## 🔧 Available Macros

| Macro | Purpose |
|-------|---------|
| `BuildIssueDashboard` | Main: builds entire dashboard |
| `RefreshDashboard` | Refresh KPIs after data change |

---

## ❌ Common Issues & Fixes

### Issue 1: KPI shows 0 even when data is filtered
**Cause:** "Trust access to VBA project" not enabled  
**Fix:** Enable it (see Step 1 above)

### Issue 2: "Run-time error - PivotTable field not found"
**Cause:** Column names don't match exactly  
**Fix:** Check column #4 is **"Report Department"** (no 'er')

### Issue 3: KPI not updating on slicer change
**Cause:** Event injection failed  
**Fix:** 
1. Enable VBA Project access (Step 1)
2. Re-run `BuildIssueDashboard`
3. Or use `RefreshDashboard` macro manually

### Issue 4: "open" vs "Open" mismatch
**Cause:** Apps.csv uses lowercase "open"  
**Fix:** ✅ Already handled - formula checks both cases

---

## 💡 Tips

- **70% zoom** automatically applied for best view
- **Excel Table** auto-created from data
- **Pivot tables** stored in hidden "Pivot" sheet
- **All formulas use GETPIVOTDATA** = respects filters

---

## 📁 Files

| File | Description |
|------|-------------|
| `IssueDashboardModule.bas` | Main VBA code |
| `Apps.csv` | Your data (64 rows × 24 columns) |
| `SAMPLE_DATA.csv` | Old sample (reference only) |
| `README.md` | This file |

---

🔗 **Repository:** https://github.com/khalidengineer/tests

**Made for Apps.csv data analysis** ✨
