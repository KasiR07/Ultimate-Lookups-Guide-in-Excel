
# 📘 Advanced Excel Formulas – Theory, Syntax, and Limitations

---

## 🔎 VLOOKUP

### 📖 Theory
`VLOOKUP` searches for a value in the **first column** of a table and returns a value from a **specified column** in the same row.

### 🧾 Syntax
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

### ⚠️ Limitations
- Only searches **left to right**
- Lookup column must be **first**
- Slower on large datasets  
- Cannot return value to the **left** of the match

---

## 🔎 HLOOKUP

### 📖 Theory
`HLOOKUP` searches for a value in the **first row** and returns a value from a specified **row below it**.

### 🧾 Syntax
```excel
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

### ⚠️ Limitations
- Searches only **horizontally**
- Cannot look **above** the first row
- Table row must be sorted if using `TRUE`

---

## 🔎 XLOOKUP

### 📖 Theory
`XLOOKUP` is a modern lookup function that replaces `VLOOKUP` and `HLOOKUP`, supporting both vertical and horizontal search with built-in error handling.

### 🧾 Syntax
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

### ⚠️ Limitations
- Available only in **Excel 365 and 2019+**
- Slightly complex with nested or dynamic arrays

---

## ➕ SUM(FILTER(FILTER()))

### 📖 Theory
Nested `FILTER` functions can extract data meeting multiple criteria; wrapping them in `SUM()` totals those values.

### 🧾 Syntax
```excel
=SUM(FILTER(FILTER(data_range, condition1), condition2))
```

### ⚠️ Limitations
- Returns error if no match → Use `IFERROR`
- Complex nesting can affect readability

---

## 🔁 XLOOKUP(XLOOKUP())

### 📖 Theory
Nesting `XLOOKUP` functions allows multi-dimensional lookups (e.g., row + column matching).

### 🧾 Syntax
```excel
=XLOOKUP(row_input, row_labels, XLOOKUP(col_input, col_labels, data_matrix))
```

### ⚠️ Limitations
- Ranges must be properly aligned
- Complex to debug if not documented

---

## 🔢 Match Modes: 0, 1, -1, 2, 3

| Mode | Meaning                        | Used In      |
|------|--------------------------------|--------------|
| 0    | Exact match                    | MATCH, XLOOKUP |
| 1    | Exact or next smaller (sorted) | MATCH        |
| -1   | Exact or next larger (sorted)  | MATCH        |
| 2    | Wildcard match (`*`, `?`)      | XLOOKUP      |
| 3    | Reserved for future use        | N/A          |

---

## 📝 Summary

| Function            | Direction  | Error Handling | Approx Match | Exact Match | Notes                                 |
|---------------------|------------|----------------|--------------|-------------|----------------------------------------|
| VLOOKUP             | Vertical   | ❌ Use IFERROR | ✅            | ✅           | Cannot look left of lookup column      |
| HLOOKUP             | Horizontal | ❌ Use IFERROR | ✅            | ✅           | Only searches rows                     |
| XLOOKUP             | Both       | ✅ Built-in    | ✅            | ✅           | Most flexible, modern lookup           |
| SUM(FILTER(...))    | N/A        | ❌ Use IFERROR | ❌            | ✅           | Filters and aggregates dynamic data    |
| XLOOKUP(XLOOKUP())  | 2D Lookup  | ✅ Built-in    | ✅            | ✅           | Matrix-style lookup                    |

---

## ❗ Error Handling: IFERROR vs IF_NOT_FOUND

### ✅ `IFERROR()`

#### 📖 Theory
`IFERROR` handles common Excel errors (e.g., `#N/A`, `#DIV/0!`, `#VALUE!`) by returning a fallback value when an error is detected.

#### 🧾 Syntax
```excel
=IFERROR(expression, value_if_error)
```

#### ✅ Example
```excel
=IFERROR(A2/B2, "Invalid")
```
If B2 is 0 (division error), returns `"Invalid"` instead of `#DIV/0!`.

#### ⚠️ Limitations
- Wraps the entire expression
- Does not distinguish between types of errors

---

### ✅ `if_not_found` in XLOOKUP

#### 📖 Theory
`XLOOKUP` includes built-in error handling via its `if_not_found` parameter, allowing cleaner formulas without needing `IFERROR`.

#### 🧾 Syntax
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, if_not_found)
```

#### ✅ Example
```excel
=XLOOKUP("Sai", A2:A10, B2:B10, "Not Found")
```

If `"Sai"` is not in column A, returns `"Not Found"` instead of `#N/A`.

#### ⚠️ Limitations
- Only works within `XLOOKUP`
- Cannot handle other types of formula errors (e.g., divide by zero)

---

## 🔍 When to Use

| Function       | Purpose                                | Use When...                                |
|----------------|----------------------------------------|--------------------------------------------|
| `IFERROR()`    | Handles all types of formula errors     | You want broad error-catching              |
| `if_not_found` | Handles only missing lookup values      | You're using `XLOOKUP` and want cleaner code |



