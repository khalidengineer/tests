# Issue Tracking Analytics Dashboard - VBA

## 🐛 Error Fix: `.DrawingObject.Formula` Issue

**Problem:** `.DrawingObject.Formula = kpiCells(k)` - Excel ke new versions me `AddTextbox` se banaye gaye shapes par ye property work nahi karti.

**Solution:** Hum hidden cells me formula likhte hain aur fir textbox ka `.TextFrame.Characters.Text` directly update karte hain via `UpdateKPIText` sub.

```vba
' OLD (Error aata tha):
valBox.DrawingObject.Formula = kpiCells(k)

' NEW (Working):
' Step 1: Formula hidden cell me likho
wsDash.Cells(KPI_ROW, k + 1).Formula = "=COUNTA(Data!A2:A10000)"

' Step 2: Textbox text directly set karo
wsDash.Shapes("KPI_Value_" & k).TextFrame.Characters.Text = displayText
```

---

## 📋 Required Data Sheet Columns (24 columns)

Aapka **Data** sheet mein ye columns honi chahiye (exact order zaruri nahi, but exact name match hona chahiye):

| # | Column Name | Description |
|---|------------|-------------|
| 1 | Issue Type | Hardware, Software, Network, Security |
| 2 | Severity | Critical, High, Medium, Low |
| 3 | Current Status | Open, In Progress, Resolved, Closed |
| 4 | Issue Location | Building/Floor info |
| 5 | Issue Description | Brief description |
| 6 | Resolver | Resolver's name |
| 7 | Resolver Identifier | Employee ID |
| 8 | Reporter | Reporter's name |
| 9 | Reporter Identifier | Employee ID |
| 10 | Reporter Designation | Job title |
| 11 | Reporter Department | Department name |
| 12 | Reporter Division | Division |
| 13 | Reporter Sub Division | Sub division |
| 14 | Reporter Location | City/Office |
| 15 | Resolver Designation | Job title |
| 16 | Resolver Department | Department |
| 17 | Resolver Division | Division |
| 18 | Resolver Sub Division | Sub division |
| 19 | Resolver Location | City/Office |
| 20 | Resolved At | Resolution timestamp |
| 21 | Resolved Remarks | Resolution notes |
| 22 | Reported At | Report timestamp |
| 23 | Issue ID | Unique ID |
| 24 | Issue Title | Brief title |

---

## 🚀 Installation Steps

### Step 1: Excel Settings
1. **File → Options → Trust Center → Trust Center Settings**
2. **Macro Settings:** "Enable all macros"
3. ✅ **"Trust access to the VBA project object model"** (REQUIRED)

### Step 2: Import Module
1. **Alt + F11** (VBA Editor open karo)
2. **File → Import File**
3. Select **IssueDashboardModule.bas**

### Step 3: Prepare Data
1. **"Data"** naam ki sheet banao
2. Sample data load karo (SAMPLE_DATA.csv use kar sakte ho)
3. Ensure all 24 columns are present with exact names

### Step 4: Run Dashboard
1. **Alt + F8**
2. Select **BuildIssueDashboard**
3. Click **Run**

---

## 📊 Dashboard Components

### 6 KPI Cards
| KPI | Description | Color |
|-----|-------------|-------|
| Total Issues | All issues count | Blue |
| Open Issues | Open + In Progress + Pending | Orange |
| Resolved | Resolved + Closed + Completed | Green |
| Critical/High | High priority issues | Red |
| Resolution % | Resolved / Total ratio | Purple |
| Total Reporters | Unique reporters count | Navy |

### 5 Charts
1. **Severity Distribution** - Donut chart
2. **Current Status Breakdown** - Bar chart
3. **Issues by Department** - Column chart
4. **Top 10 Resolvers** - Bar chart
5. **Issues by Location** - Column chart

### 4 Slicers
1. Issue Type
2. Severity
3. Current Status
4. Reporter Department

---

## 🔧 Available Macros

After running once, ye macros available honge:

| Macro | Purpose |
|-------|---------|
| `BuildIssueDashboard` | Main macro - dashboard rebuild karta hai |
| `RefreshDashboard` | KPI values refresh karta hai (data change ke baad) |

---

## ❌ Common Errors & Fixes

### Error 1: `.DrawingObject.Formula` Error
**Reason:** New Excel versions support nahi karte
**Fix:** ✅ Already fixed using `UpdateKPIText` method

### Error 2: "Unable to get the PivotTable property"
**Reason:** Pivot table create nahi hua
**Fix:** Column names exact match karo (case-sensitive)

### Error 3: "Type mismatch"
**Reason:** Date columns mein text values
**Fix:** Reported At aur Resolved At columns ko Date format mein convert karo

### Error 4: Slicer error
**Reason:** Pivot tables exist nahi karte
**Fix:** Pehle pivots create karo, fir slicers

---

## 💡 Pro Tips

1. **Data update karne ke baad:**
   ```
   Alt + F8 → RefreshDashboard → Run
   ```

2. **New columns add karne hain to:**
   - `IssueDashboardModule.bas` mein additional pivot subs add karo
   - Charts function mein chart add karo

3. **Custom colors change karne hain:**
   - `accentColors()` array edit karo
   - RGB values modify karo

4. **Dashboard slow hai to:**
   - Data ko Excel Table format mein rakho (auto-applied by macro)
   - Conditional formatting kam rakho

---

## 📁 Files in this Repo

| File | Description |
|------|-------------|
| `IssueDashboardModule.bas` | Main VBA module |
| `SAMPLE_DATA.csv` | 50 sample issues for testing |
| `README.md` | This documentation |

---

## ✨ Key Features

- ✅ Auto-creates Excel Table from data
- ✅ 7 different pivot tables
- ✅ 6 KPI cards with formulas
- ✅ 5 interactive charts
- ✅ 4 connected slicers
- ✅ Professional design
- ✅ Auto-refresh capability
- ✅ Error-free execution
- ✅ **Fixed `.DrawingObject.Formula` error**

---

**Made with ❤️ for Issue Analytics**
