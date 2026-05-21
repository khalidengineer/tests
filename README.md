# Ticket Analyst Dashboard - VBA Code

## 📋 Data Sheet Requirements

Your Excel workbook must have a sheet named **"Data"** with the following columns:

1. **Ticket ID** - Unique identifier for each ticket
2. **Date** - Ticket creation date
3. **Status** - Current status (e.g., Open, In Progress, Resolved, Closed)
4. **Priority** - Priority level (e.g., Critical, High, Medium, Low)
5. **Category** - Ticket category (e.g., Technical, Billing, Account, Support)
6. **Assigned Agent** - Name of the agent handling the ticket
7. **Resolution Time (Hours)** - Time taken to resolve in hours
8. **Source** - Channel through which ticket was created (e.g., Email, Phone, Chat, Portal)
9. **SLA Status** - Whether SLA was met (Met, Breached)
10. **Satisfaction Rating** - Customer satisfaction score (1-5)

---

## 🚀 Installation Steps

### Step 1: Enable Macro Settings
1. Open Excel
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Click **Macro Settings**
4. Select **"Enable all macros"** (temporarily for development)
5. Check **"Trust access to the VBA project object model"** ✅ (REQUIRED)
6. Click OK

### Step 2: Import the VBA Module
1. Press **Alt + F11** to open VBA Editor
2. Go to **File → Import File**
3. Select **TicketDashboardModule.bas**
4. The module will appear in your VBA Project

### Step 3: Prepare Your Data
1. Ensure you have a sheet named **"Data"** with all required columns
2. Make sure the data starts from Row 1 with headers

### Step 4: Run the Dashboard
1. Press **Alt + F8** to open Macro dialog
2. Select **BuildTicketDashboard**
3. Click **Run**
4. Wait for the process to complete (you'll see a success message)

---

## 📊 Dashboard Features

### KPI Cards (6 Cards)
1. **Total Tickets** - Count of all tickets
2. **Open Tickets** - Count of open tickets
3. **Resolved Tickets** - Count of resolved + closed tickets
4. **Avg Resolution (Hrs)** - Average resolution time
5. **Avg CSAT Score** - Average customer satisfaction
6. **SLA Compliance %** - Percentage of tickets meeting SLA

### Charts (5 Charts)
1. **Ticket Status Trend Over Time** - Stacked area chart showing status changes by month
2. **Priority Distribution** - Donut chart showing ticket breakdown by priority
3. **Tickets by Category** - Horizontal bar chart
4. **Agent Performance** - Column chart showing tickets per agent
5. **Customer Satisfaction Rating** - Column chart showing CSAT distribution

### Interactive Slicers (4 Slicers)
1. **Status** - Filter by ticket status
2. **Priority** - Filter by priority level
3. **Category** - Filter by category
4. **Assigned Agent** - Filter by agent name

---

## 🎨 Dashboard Layout

- **Header**: Professional banner with title "Ticket Support Analytics Dashboard"
- **KPI Row**: 6 colorful KPI cards with real-time formulas
- **Row 1**: Status Trend Chart + Priority Donut Chart
- **Row 2**: Category Chart + Agent Performance Chart + CSAT Chart
- **Right Panel**: 4 interactive slicers for filtering

---

## ⚙️ Customization

### Change Colors
Edit the RGB values in the `CreateKPICards` and `CreateCharts` subroutines:
```vba
RGB(37, 99, 235)  ' Blue
RGB(245, 158, 11) ' Orange
RGB(5, 150, 105)  ' Green
RGB(220, 38, 38)  ' Red
```

### Adjust Layout
Modify these constants in `CreateCharts`:
```vba
Dim cht1W As Single: cht1W = Int(contentW * 0.48)  ' Chart 1 width (48% of total)
Dim cht2W As Single: cht2W = contentW - cht1W - cGap  ' Chart 2 width (remaining)
```

### Change Chart Types
Edit the `.ChartType` property:
- `xlLine` - Line chart
- `xlColumnClustered` - Column chart
- `xlBarClustered` - Bar chart
- `xlAreaStacked` - Stacked area chart
- `xlDoughnut` - Donut chart
- `xlPie` - Pie chart

---

## 🔧 Troubleshooting

### Error: "Compile Error: User-defined type not defined"
- Enable **Microsoft Scripting Runtime** in VBA Editor (Tools → References)

### Error: "Run-time error '1004': PivotTable field name is not valid"
- Check that all column names in your Data sheet match exactly
- Ensure there are no typos or extra spaces

### Dashboard doesn't update automatically
- Make sure **"Trust access to the VBA project object model"** is enabled
- The `InjectDashboardEvent` sub requires this permission

### Charts look distorted
- Run the macro again
- The zoom sequence at the end fixes rendering issues

### Slicers not working
- Ensure all pivot tables are created successfully
- Check that field names match your data columns

---

## 📝 Notes

- The dashboard is set to **70% zoom** for optimal viewing
- All data is linked via formulas - updates automatically when data changes
- Pivot tables are recreated each time you run the macro
- The original Data sheet is never modified

---

## 🆘 Support

If you encounter any issues:
1. Check that your Data sheet has all required columns
2. Verify that "Trust access to VBA project" is enabled
3. Make sure your Excel version supports PivotTables and Slicers (Excel 2010+)

---

## ✨ One-Click Copy

Simply select all the code from **TicketDashboardModule.bas** and copy it into your VBA editor!

**Made with ❤️ for Ticket Analytics**
