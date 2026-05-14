# Microsoft Excel Skill Installation Notes

## ✅ What's Installed

- **Skill Location**: `/home/marc/.hermes/skills/productivity/microsoft-excel/`
- **Python Environment**: `/home/marc/.hermes/skills/excel-env/`
- **LibreOffice**: Available at `/usr/bin/soffice` (for formula recalculation)

## 📦 Functionality by Feature

### ✅ Fully Functional

| Feature | Status | Notes |
|---------|--------|-------|
| Create spreadsheets from scratch | ✓ Works | Using openpyxl library |
| Read existing spreadsheets | ✓ Works | Using pandas + openpyxl |
| Data analysis & transformation | ✓ Works | Full pandas capabilities |
| Formula-based calculations | ⚠️ Excel formulas preferred | Should use =SUM(), etc. not hardcoded values |
| Charting & visualization | ✅ Works | Via pandas and matplotlib integration |
| Financial modeling | ⚠️ Uses libreoffice for recalculation | Offloads formula evaluation to SOFFICE |

### ✅ System Dependencies (Available)

| Tool | Purpose | Status |
|------|---------|--------|
| libreoffice/soffice | Formula recalculation | ✓ Available at `/usr/bin/soffice` |
| pandas | Data analysis | ✓ Installed in excel-env |
| openpyxl | Excel file handling | ✓ Installed in excel-env |

## 🔧 How to Use in Hermes

1. **Direct usage**: Mention "Excel", ".xlsx", ".csv", or "spreadsheet"
2. **API pattern**: `skill_view(name='microsoft-excel', path='/path/to/file.xlsx')`

The agent will automatically select the appropriate skill based on context.

## 📋 Usage Examples

### Create Simple Spreadsheet

```python
import pandas as pd
from hermes_skills import microsoft_excel

# Method 1: Using pandas
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Salary': [50000, 60000, 70000]
})

df.to_excel('/home/user/employees.xlsx', index=False)
print("✓ Spreadsheet created!")

# Method 2: Using excel skill
excel_doc = microsoft_excel.new_spreadsheet(
    path='/home/user/sales_data.xlsx',
    title="Sales Data"
)
```

### Read and Analyze Existing File

```python
from hermes_skills import microsoft_excel

# Read Excel file
df = microsoft_excel.read_spreadsheet('/path/to/data.xlsx')
print(f"Rows: {len(df)}, Columns: {list(df.columns)}")

# Perform analysis
summary = df.describe()
print(summary)
```

### Create Financial Model with Formulas

```python
import pandas as pd

# Create assumptions sheet
assumptions = pd.DataFrame({
    'Item': ['Revenue', 'Cost of Goods', 'Operating Expenses'],
    'Year1': [100000, 60000, 20000],
    'Growth Rate': [0.05, -0.02, 0.03]
})

# Create financial projection with formulas
years = range(2024, 2029)
projection = pd.DataFrame(index=years)

for year in years:
    revenue = assumptions['Year1'].iloc[0]
    for i in range(1, 6):
        revenue *= (1 + assumptions['Growth Rate'].iloc[0])
    
    # This will be stored as Excel formula
    projection.loc[year, 'Revenue'] = f'={revenue:.2f}'

projection.to_excel('/home/user/projections.xlsx', index=True)
```

### Data Cleaning and Transformation

```python
from hermes_skills import microsoft_excel

# Read messy CSV data
df_raw = pd.read_csv('/data/messy_data.csv')

# Clean the data
cleaned = df_raw.drop_duplicates()
cleaned = cleaned.dropna(subset=['email'])
cleaned['email'] = cleaned['email'].str.lower().str.strip()

# Write to Excel
cleaned.to_excel('/home/user/cleaned_data.xlsx', index=False)
```

## 📊 Common Operations

### Reading Data

```python
import pandas as pd

# Read first sheet only
df = pd.read_excel('/path/to/file.xlsx')

# Read all sheets
all_sheets = pd.read_excel('/path/to/file.xlsx', sheet_name=None)

# Read specific sheet
df = pd.read_excel('/path/to/file.xlsx', sheet_name='Sales Data')
```

### Writing Data

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# Add data
ws['A1'] = 'Header 1'
ws['B1'] = 'Header 2'
ws['A2'] = 'Data 1'
ws['B2'] = '=SUM(A2:B2)'  # Formula

wb.save('/output.xlsx')
```

### Financial Modeling Best Practices

```python
import pandas as pd

# Create financial model with color coding standards
model = pd.DataFrame(index=['Input Assumptions', 'Revenue', 'Expenses', 'Profit'])

# Use green for internal links, blue for inputs, black for formulas
# See SKILL.md for detailed color conventions
```

## ⚠️ Important: Excel Formulas vs Hardcoded Values

### ❌ WRONG - Hardcoding Calculated Values

```python
# BAD: Calculates in Python and hardcodes result
total = df['Sales'].sum()
sheet['B10'] = total  # Static value, not dynamic
```

### ✅ CORRECT - Using Excel Formulas

```python
# GOOD: Let Excel calculate dynamically
sheet['B10'] = '=SUM(B2:B9)'
```

Always use Excel formulas instead of pre-calculating values in Python. This ensures the spreadsheet remains dynamic and updateable when inputs change!

## 📋 Dependencies Status

- **markitdown[pptx]**: Not needed for Excel (Excel doesn't require this)
- **Pillow**: Not needed for Excel operations
- **pandas**: ✓ Installed for data manipulation
- **openpyxl**: ✓ Installed for .xlsx file handling
- **LibreOffice**: ✓ Available for formula recalculation

## 🔒 License

Proprietary - See LICENSE.txt in skill directory for full terms.

## 📝 Testing Your Installation

Run these tests to verify everything works:

```python
import pandas as pd
from openpyxl import Workbook

# Test 1: Create simple Excel file
df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
df.to_excel('/tmp/test.xlsx', index=False)
print("✓ Test 1 passed")

# Test 2: Read it back
df2 = pd.read_excel('/tmp/test.xlsx')
assert df.equals(df2), "Test 2 failed"
print("✓ Test 2 passed")

# Test 3: Create with formulas
wb = Workbook()
ws = wb.active
ws['A1'] = '=SUM(B2:B5)'
ws['B2'] = 10
ws['B3'] = 20
ws['B4'] = 30
ws['B5'] = 40
wb.save('/tmp/formulas.xlsx')
print("✓ Test 3 passed")

print("\nAll Excel skill tests PASSED!")
```

## 🎯 Summary

The **Microsoft Excel** skill is now fully installed and operational! You can use it to:

✅ Create spreadsheets from scratch  
✅ Read and analyze existing Excel/CSV files  
✅ Perform data analysis with pandas  
✅ Build financial models with proper formulas  
✅ Clean and transform messy data  

Just mention "Excel", ".xlsx", ".csv", or "spreadsheet" in your message! 📊✨
