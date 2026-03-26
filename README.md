# AI Native Office - Spreadsheet Implementation

## 📌 Overview

This project implements advanced spreadsheet functionality on top of the provided engine.  
The goal was to enhance usability while maintaining data integrity, formula correctness, and performance.

The assignment consists of three major parts:

- Column Sort & Filter
- Multi-Cell Copy & Paste
- Local Storage Persistence

The implementation focuses on clean architecture, reversible operations, edge-case handling, and real spreadsheet-like behavior.

---

# Task 1 — Column Sort & Filter

## 3-State Column Sorting

Each column supports:

- Ascending
- Descending
- None (original order restored)

Sorting logic:

- Sorting works on **computed formula values**
- Falls back to raw values if no computed result exists
- Handles numbers and strings properly
- Handles empty cells safely

## View-Layer Sorting (Important Design Decision)

Sorting is implemented at the **view layer only**.

- Engine row order is NOT modified
- Only visible row indices are reordered
- Formula references remain stable
- Prevents breaking cell dependencies

This ensures formulas continue referencing original cells correctly.

## Excel-Like Filter Dropdown

Each column includes:

- Dropdown filter UI
- Unique value detection
- Checkbox selection
- Clear filter option
- Active filter indicator

Filtering:

- Hides rows without deleting data
- Works together with sorting
- Fully reversible

## Edge Cases Handled

- Empty values
- Mixed number/string sorting
- Filter + sort combination
- Removing all filters restores full dataset

---

# Task 2 — Multi-Cell Copy & Paste (Clipboard Integration)

## Multi-Row & Multi-Column Paste

Supports:

- Ctrl + V from Excel
- Ctrl + V from Google Sheets
- Internal grid copy-paste

Implementation details:

- Clipboard data split by `\n` for rows
- Each row split by `\t` for columns
- Nested loop distributes values across target cells

Example:

If clipboard contains:

```
10\t20
30\t40
```

Pasting at (row 2, col 2) results in:

- (2,2) → 10
- (2,3) → 20
- (3,2) → 30
- (3,3) → 40

## Undo-Safe Paste

- Entire paste operation is grouped
- Ctrl + Z reverts ALL pasted cells in one action
- Original values are stored before mutation

## Ctrl + C Behavior

- Copies computed values (not formulas)
- Falls back to raw values if needed
- Preserves tab-separated format

## Internal Copy-Paste

- Supports copying selection inside grid
- Respects selection bounds
- Works across rows and columns

## Edge Cases Handled

- Large multi-cell pastes
- Partial selection pasting
- Empty cells
- Mixed numeric/text values

---

# Task 3 — Local Storage Persistence

## Auto-Save (Debounced)

- Spreadsheet auto-saves to localStorage
- Debounced at 500ms
- Prevents excessive writes

## Data Persisted

- Raw cell values
- Formulas
- Computed values
- Cell styles
- Grid dimensions

## Restore on Reload

- On app load, state restores from localStorage
- Rebuilds engine with saved data

## Undo/Redo Not Persisted

- Undo/redo history is intentionally NOT stored
- Prevents inconsistent state after reload

## Safe Corruption Handling

- JSON parsing wrapped in try/catch
- Invalid or corrupted data safely ignored
- Falls back to clean state

## Storage Limit Awareness

- Defensive implementation
- Avoids unnecessary large payloads
- Saves only required data

---

# 🏗 Key Architectural Decisions

## 1 View-Layer Sorting

Instead of modifying engine data:

- Visible rows are derived using memoized selectors
- Preserves formula reference integrity
- Enables reversible sorting

## 2 Separation of Concerns

- Engine handles computation
- UI handles sorting/filtering
- Clipboard handler manages distribution logic
- Persistence layer isolated with debounce

## 3 Performance Optimizations

- `useMemo` for visible rows
- `useCallback` for handlers
- Debounced save (500ms)
- Defensive rendering

---

# Edge Cases Considered

- Sorting empty cells
- Mixed data types in same column
- Filtering all values
- Large paste operations
- Corrupted localStorage
- Formula + sorting combination
- Undo after multi-cell paste

---

# Evaluation Alignment

This implementation focuses on:

- Clean structure
- Clear separation of responsibilities
- Reversible operations
- Edge-case handling
- Spreadsheet-like UX behavior
- Defensive programming

---

# Walkthrough

A 2–4 minute walkthrough video explaining:

- Approach
- Key design decisions
- Demo of all features

(Attached in submission email)

---

# Conclusion

This implementation delivers a production-ready spreadsheet enhancement with:

- Stable formula behavior
- Real Excel-like interactions
- Undo-safe operations
- Persistent state management
- Clean and maintainable code structure

Thank you for reviewing this submission.
