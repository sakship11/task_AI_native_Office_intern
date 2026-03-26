import { useState, useRef, useCallback, useMemo, useEffect } from "react";
import "./App.css";
import { createEngine } from "./engine/core.js";

const TOTAL_ROWS = 50;
const TOTAL_COLS = 50;

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const persistedState = useMemo(() => {
    try {
      const rawState = localStorage.getItem("spreadsheetState");
      if (!rawState) return null;
      const parsed = JSON.parse(rawState);
      if (
        !parsed ||
        typeof parsed !== "object" ||
        !Number.isInteger(parsed.rows) ||
        !Number.isInteger(parsed.cols)
      )
        return null;
      return parsed;
    } catch (err) {
      console.warn("Invalid persisted spreadsheet state:", err);
      return null;
    }
  }, []);

  const [engine] = useState(() => {
    const rows =
      persistedState &&
      Number.isInteger(persistedState.rows) &&
      persistedState.rows > 0
        ? persistedState.rows
        : TOTAL_ROWS;
    const cols =
      persistedState &&
      Number.isInteger(persistedState.cols) &&
      persistedState.cols > 0
        ? persistedState.cols
        : TOTAL_COLS;
    const e = createEngine(rows, cols);
    if (persistedState?.cells && typeof persistedState.cells === "object") {
      for (const key of Object.keys(persistedState.cells)) {
        const [rStr, cStr] = key.split(",");
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        const value = persistedState.cells[key];
        if (!Number.isNaN(r) && !Number.isNaN(c) && typeof value === "string") {
          e.setCell(r, c, value);
        }
      }
    }
    return e;
  });

  const [_version, setVersion] = useState(0);
  const [selectedCell, setSelectedCell] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");
  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState(
    persistedState?.cellStyles && typeof persistedState.cellStyles === "object"
      ? persistedState.cellStyles
      : {},
  );
  const cellInputRef = useRef(null);

  // Sorting / filtering
  const [columnSort, setColumnSort] = useState({}); // { colIndex: 'none'|'asc'|'desc' }
  const [columnFilters, setColumnFilters] = useState({}); // { colIndex: Set }
  const [openFilterDropdown, setOpenFilterDropdown] = useState(null);

  // Multi-cell selection + clipboard support
  const [selectionStart, setSelectionStart] = useState(null);
  const [selectionEnd, setSelectionEnd] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [_pasteUndoData, setPasteUndoData] = useState(null);

  const forceRerender = useCallback(() => setVersion((v) => v + 1), []);

  useEffect(() => {
    const saveState = () => {
      try {
        const cells = {};
        for (let r = 0; r < engine.rows; r++) {
          for (let c = 0; c < engine.cols; c++) {
            const cell = engine.getCell(r, c);
            if (cell.raw && cell.raw.trim() !== "") {
              cells[`${r},${c}`] = cell.raw;
            }
          }
        }
        const state = {
          rows: engine.rows,
          cols: engine.cols,
          cells,
          cellStyles,
        };
        localStorage.setItem("spreadsheetState", JSON.stringify(state));
      } catch (err) {
        console.warn("Could not save spreadsheet state to localStorage:", err);
      }
    };

    const timer = setTimeout(saveState, 500);
    return () => clearTimeout(timer);
  }, [engine, cellStyles, _version]);

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback(
    (row, col) => {
      const key = `${row},${col}`;
      return (
        cellStyles[key] || {
          bold: false,
          italic: false,
          underline: false,
          bg: "white",
          color: "#202124",
          align: "left",
          fontSize: 13,
        }
      );
    },
    [cellStyles],
  );

  const updateCellStyle = useCallback(
    (row, col, updates) => {
      const key = `${row},${col}`;
      setCellStyles((prev) => ({
        ...prev,
        [key]: { ...getCellStyle(row, col), ...updates },
      }));
    },
    [getCellStyle],
  );

  // ────── Selection helpers ──────

  const getSelectionBounds = useCallback(() => {
    if (!selectionStart || !selectionEnd) return null;
    return {
      minRow: Math.min(selectionStart.r, selectionEnd.r),
      maxRow: Math.max(selectionStart.r, selectionEnd.r),
      minCol: Math.min(selectionStart.c, selectionEnd.c),
      maxCol: Math.max(selectionStart.c, selectionEnd.c),
    };
  }, [selectionStart, selectionEnd]);

  const isCellInSelection = useCallback(
    (row, col) => {
      const bounds = getSelectionBounds();
      if (!bounds) return false;
      return (
        row >= bounds.minRow &&
        row <= bounds.maxRow &&
        col >= bounds.minCol &&
        col <= bounds.maxCol
      );
    },
    [getSelectionBounds],
  );

  const clearSelection = useCallback(() => {
    setSelectionStart(null);
    setSelectionEnd(null);
  }, []);

  // ────── Cell editing ──────

  const startEditing = useCallback(
    (row, col) => {
      setSelectedCell({ r: row, c: col });
      setEditingCell({ r: row, c: col });
      const cellData = engine.getCell(row, col);
      setEditValue(cellData.raw);
      setTimeout(() => cellInputRef.current?.focus(), 0);
    },
    [engine],
  );

  const commitEdit = useCallback(
    (row, col) => {
      // Only commit if the value actually changed to avoid unnecessary recalculations
      const currentCell = engine.getCell(row, col);
      if (currentCell.raw !== editValue) {
        engine.setCell(row, col, editValue);
        forceRerender();
      }
      setEditingCell(null);
    },
    [engine, editValue, forceRerender],
  );

  const handleCellClick = useCallback(
    (row, col) => {
      if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
        commitEdit(editingCell.r, editingCell.c);
      }
      if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
        startEditing(row, col);
      }
    },
    [editingCell, commitEdit, startEditing],
  );

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback(
    (event, row, col) => {
      if (event.key === "Enter") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.min(row + 1, engine.rows - 1), col);
      } else if (event.key === "Tab") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(row, Math.min(col + 1, engine.cols - 1));
      } else if (event.key === "Escape") {
        setEditValue(engine.getCell(row, col).raw);
        setEditingCell(null);
      } else if (event.key === "ArrowDown") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.min(row + 1, engine.rows - 1), col);
      } else if (event.key === "ArrowUp") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.max(row - 1, 0), col);
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        commitEdit(row, col);
        if (col > 0) {
          startEditing(row, col - 1);
        } else if (row > 0) {
          startEditing(row - 1, engine.cols - 1);
        }
      } else if (event.key === "ArrowRight") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(row, Math.min(col + 1, engine.cols - 1));
      }
    },
    [engine, commitEdit, startEditing],
  );

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback(
    (event) => {
      if (!editingCell) return;
      handleKeyDown(event, editingCell.r, editingCell.c);
    },
    [editingCell, handleKeyDown],
  );

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell);
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw);
    }
  }, [selectedCell, editingCell, engine]);

  const handleFormulaBarChange = useCallback(
    (value) => {
      if (!editingCell && selectedCell) setEditingCell(selectedCell);
      setEditValue(value);
    },
    [editingCell, selectedCell],
  );

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => {
    if (_pasteUndoData?.originalValues?.length > 0) {
      _pasteUndoData.originalValues.forEach(({ r, c, val }) => {
        engine.setCell(r, c, val);
      });
      setPasteUndoData(null);
      forceRerender();
      return;
    }
    if (engine.undo()) {
      forceRerender();
    }
  }, [engine, forceRerender, _pasteUndoData]);

  const handleRedo = useCallback(() => {
    if (engine.redo()) {
      forceRerender();
    }
  }, [engine, forceRerender]);

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, {
      underline: !style.underline,
    });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const changeFontSize = useCallback(
    (size) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size });
    },
    [selectedCell, updateCellStyle],
  );

  const changeAlignment = useCallback(
    (align) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { align });
    },
    [selectedCell, updateCellStyle],
  );

  const changeFontColor = useCallback(
    (color) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { color });
    },
    [selectedCell, updateCellStyle],
  );

  const changeBackgroundColor = useCallback(
    (color) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { bg: color });
    },
    [selectedCell, updateCellStyle],
  );

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return;
    engine.setCell(selectedCell.r, selectedCell.c, "");
    forceRerender();
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`;
    setCellStyles((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    setEditValue("");
  }, [selectedCell, engine, forceRerender]);

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, "");
      }
    }
    forceRerender();
    setCellStyles({});
    setSelectedCell(null);
    setEditingCell(null);
    setEditValue("");
  }, [engine, forceRerender]);

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return;
    engine.insertRow(selectedCell.r);
    forceRerender();
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c });
  }, [selectedCell, engine, forceRerender]);

  const deleteRow = useCallback(() => {
    if (!selectedCell) return;
    engine.deleteRow(selectedCell.r);
    forceRerender();
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c });
    }
  }, [selectedCell, engine, forceRerender]);

  const insertColumn = useCallback(() => {
    if (!selectedCell) return;
    engine.insertColumn(selectedCell.c);
    forceRerender();
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 });
  }, [selectedCell, engine, forceRerender]);

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return;
    engine.deleteColumn(selectedCell.c);
    forceRerender();
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 });
    }
  }, [selectedCell, engine, forceRerender]);

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null;
  }, [selectedCell, getCellStyle]);

  const getColumnLabel = useCallback((col) => {
    let label = "";
    let num = col + 1;
    while (num > 0) {
      num--;
      label = String.fromCharCode(65 + (num % 26)) + label;
      num = Math.floor(num / 26);
    }
    return label;
  }, []);

  const toggleColumnSort = useCallback((colIndex) => {
    setColumnSort((prev) => {
      const current = prev[colIndex] || "none";
      const next =
        current === "none" ? "asc" : current === "asc" ? "desc" : "none";
      return { ...prev, [colIndex]: next };
    });
  }, []);

  const setColumnFilter = useCallback((colIndex, allowedValues) => {
    setColumnFilters((prev) => ({ ...prev, [colIndex]: allowedValues }));
  }, []);

  const clearColumnFilter = useCallback((colIndex) => {
    setColumnFilters((prev) => {
      const next = { ...prev };
      delete next[colIndex];
      return next;
    });
  }, []);

  const getColumnUniqueValues = useCallback(
    (colIndex) => {
      const values = new Set();
      for (let r = 0; r < engine.rows; r++) {
        const cell = engine.getCell(r, colIndex);
        const value = cell.error
          ? cell.error
          : cell.computed !== null && cell.computed !== ""
            ? String(cell.computed)
            : cell.raw;
        values.add(value);
      }
      return Array.from(values).sort();
    },
    [engine],
  );

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : "No cell";

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : selectedCell
      ? engine.getCell(selectedCell.r, selectedCell.c).raw
      : "";

  const getVisibleRows = useMemo(() => {
    let rows = Array.from({ length: engine.rows }, (_, i) => i);

    // filtering
    Object.entries(columnFilters).forEach(([colStr, allowed]) => {
      const colIndex = parseInt(colStr, 10);
      if (allowed && allowed.size > 0) {
        rows = rows.filter((row) => {
          const cell = engine.getCell(row, colIndex);
          const value = cell.error
            ? cell.error
            : cell.computed !== null && cell.computed !== ""
              ? String(cell.computed)
              : cell.raw;
          return allowed.has(value);
        });
      }
    });

    // sorting (applied in order of key definition)
    Object.entries(columnSort).forEach(([colStr, order]) => {
      const colIndex = parseInt(colStr, 10);
      if (order !== "none") {
        rows.sort((a, b) => {
          const cellA = engine.getCell(a, colIndex);
          const cellB = engine.getCell(b, colIndex);
          const valA = cellA.error
            ? cellA.error
            : cellA.computed !== null && cellA.computed !== ""
              ? cellA.computed
              : cellA.raw;
          const valB = cellB.error
            ? cellB.error
            : cellB.computed !== null && cellB.computed !== ""
              ? cellB.computed
              : cellB.raw;
          if (
            (valA === "" || valA === null || valA === undefined) &&
            (valB === "" || valB === null || valB === undefined)
          )
            return 0;
          if (valA === "" || valA === null || valA === undefined) return 1;
          if (valB === "" || valB === null || valB === undefined) return -1;
          const nA = typeof valA === "number" ? valA : parseFloat(valA);
          const nB = typeof valB === "number" ? valB : parseFloat(valB);
          if (!isNaN(nA) && !isNaN(nB))
            return order === "asc" ? nA - nB : nB - nA;
          const sA = String(valA).toLowerCase();
          const sB = String(valB).toLowerCase();
          return order === "asc" ? sA.localeCompare(sB) : sB.localeCompare(sA);
        });
      }
    });

    return rows;
  }, [engine, columnSort, columnFilters]);

  const parseClipboardData = useCallback((text) => {
    if (typeof text !== "string") return [];

    // Excel/Google Sheets clipboard is usually tab-delimited rows, but support CSV fallback
    const rows = text.split(/\r\n|\n|\r/);

    const parsed = rows.map((row) => {
      if (row === "") return [""];

      // Prefer tab delimiting for spreadsheet style
      const tabParts = row.split("\t");
      if (tabParts.length > 1) {
        return tabParts;
      }

      // Fallback to comma-separated data
      const commaParts = row.split(",");
      return commaParts;
    });

    // Keep empty leading/trailing cells, but drop fully empty trailing row if it comes from final newline
    if (
      parsed.length > 1 &&
      parsed[parsed.length - 1].every((cell) => cell === "")
    ) {
      parsed.pop();
    }

    return parsed;
  }, []);

  const copySelection = useCallback(() => {
    const bounds = getSelectionBounds();
    if (bounds) {
      const data = [];
      for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
        const row = [];
        for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
          const cell = engine.getCell(r, c);
          const value = cell.error
            ? cell.error
            : cell.computed !== null && cell.computed !== ""
              ? String(cell.computed)
              : cell.raw;
          row.push(value);
        }
        data.push(row);
      }
      const tsv = data.map((row) => row.join("\t")).join("\n");
      navigator.clipboard.writeText(tsv).catch(() => {});
      return;
    }
    if (selectedCell) {
      const cell = engine.getCell(selectedCell.r, selectedCell.c);
      const value = cell.error
        ? cell.error
        : cell.computed !== null && cell.computed !== ""
          ? String(cell.computed)
          : cell.raw;
      navigator.clipboard.writeText(value).catch(() => {});
    }
  }, [engine, selectedCell, getSelectionBounds]);

  const pasteToSelection = useCallback(
    (data) => {
      if (!Array.isArray(data) || data.length === 0) return;

      // When a range is selected, paste starts from top-left of selection.
      const bounds = getSelectionBounds();
      const startRow = bounds ? bounds.minRow : selectedCell?.r;
      const startCol = bounds ? bounds.minCol : selectedCell?.c;
      if (
        startRow === undefined ||
        startCol === undefined ||
        startRow === null ||
        startCol === null
      )
        return;

      const originalValues = [];
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < data[r].length; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          if (rr < engine.rows && cc < engine.cols) {
            originalValues.push({
              r: rr,
              c: cc,
              val: engine.getCell(rr, cc).raw,
            });
            engine.setCell(rr, cc, data[r][c] ?? "");
          }
        }
      }
      setPasteUndoData({ originalValues });
      forceRerender();
    },
    [engine, selectedCell, getSelectionBounds, forceRerender],
  );

  const handlePrimitivePaste = useCallback(async () => {
    try {
      const text = await navigator.clipboard.readText();
      const data = parseClipboardData(text);
      pasteToSelection(data);
    } catch {
      /* ignore */
    }
  }, [parseClipboardData, pasteToSelection]);

  useEffect(() => {
    const onKeyDown = (event) => {
      const isCmd = event.ctrlKey || event.metaKey;
      if (isCmd) {
        switch (event.key.toLowerCase()) {
          case "c":
            event.preventDefault();
            copySelection();
            break;
          case "v":
            event.preventDefault();
            if (editingCell) {
              commitEdit(editingCell.r, editingCell.c);
            }
            handlePrimitivePaste();
            break;
          case "z":
            event.preventDefault();
            if (editingCell) {
              // when editing, keep default undo in input for simple text editing
              break;
            }
            handleUndo();
            break;
          case "y":
            event.preventDefault();
            if (editingCell) {
              break;
            }
            handleRedo();
            break;
          default:
            break;
        }
        return;
      }

      // standard navigation only when not editing
      if (editingCell) return;
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [
    editingCell,
    copySelection,
    handlePrimitivePaste,
    handleUndo,
    handleRedo,
  ]);

  useEffect(() => {
    const clickOutside = (event) => {
      if (!event.target.closest(".filter-container")) {
        setOpenFilterDropdown(null);
        clearSelection();
      }
    };
    const handleUp = () => setIsDragging(false);
    document.addEventListener("mousedown", clickOutside);
    document.addEventListener("mouseup", handleUp);
    return () => {
      document.removeEventListener("mousedown", clickOutside);
      document.removeEventListener("mouseup", handleUp);
    };
  }, [clearSelection]);

  // ────── Render ──────

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">
        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button
              className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? "active" : ""}`}
              onClick={toggleBold}
              title="Bold"
            >
              B
            </button>
            <button
              className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? "active" : ""}`}
              onClick={toggleItalic}
              title="Italic"
            >
              I
            </button>
            <button
              className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? "active" : ""}`}
              onClick={toggleUnderline}
              title="Underline"
            >
              U
            </button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select
              className="toolbar-select"
              value={selectedCellStyle?.fontSize || 13}
              onChange={(e) => changeFontSize(parseInt(e.target.value))}
            >
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button
              className={`align-btn ${selectedCellStyle?.align === "left" ? "active" : ""}`}
              onClick={() => changeAlignment("left")}
              title="Align Left"
            >
              ⬤←
            </button>
            <button
              className={`align-btn ${selectedCellStyle?.align === "center" ? "active" : ""}`}
              onClick={() => changeAlignment("center")}
              title="Align Center"
            >
              ⬤
            </button>
            <button
              className={`align-btn ${selectedCellStyle?.align === "right" ? "active" : ""}`}
              onClick={() => changeAlignment("right")}
              title="Align Right"
            >
              ⬤→
            </button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || "#000000"}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{
                width: "32px",
                height: "32px",
                border: "1px solid #dadce0",
                cursor: "pointer",
                borderRadius: "4px",
              }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select
              className="toolbar-select"
              value={selectedCellStyle?.bg || "white"}
              onChange={(e) => changeBackgroundColor(e.target.value)}
            >
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button
              className="toolbar-btn"
              onClick={handleUndo}
              disabled={!engine.canUndo()}
              title="Undo"
            >
              ↶ Undo
            </button>
            <button
              className="toolbar-btn"
              onClick={handleRedo}
              disabled={!engine.canRedo()}
              title="Redo"
            >
              ↷ Redo
            </button>
          </div>

          <div className="toolbar-group">
            <button
              className="toolbar-btn"
              onClick={insertRow}
              title="Insert Row"
            >
              + Row
            </button>
            <button
              className="toolbar-btn"
              onClick={deleteRow}
              title="Delete Row"
            >
              - Row
            </button>
            <button
              className="toolbar-btn"
              onClick={insertColumn}
              title="Insert Column"
            >
              + Col
            </button>
            <button
              className="toolbar-btn"
              onClick={deleteColumn}
              title="Delete Column"
            >
              - Col
            </button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>
              ✕ Cell
            </button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>
              ✕ All
            </button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const sortState = columnSort[colIndex] || "none";
                  const hasFilter =
                    columnFilters[colIndex] && columnFilters[colIndex].size > 0;
                  return (
                    <th key={colIndex} className="col-header">
                      <div className="col-header-content">
                        <span>{getColumnLabel(colIndex)}</span>
                        <div className="col-header-controls">
                          <button
                            className={`sort-btn ${sortState}`}
                            onClick={() => toggleColumnSort(colIndex)}
                          >
                            {sortState === "asc"
                              ? "↑"
                              : sortState === "desc"
                                ? "↓"
                                : "↕"}
                          </button>
                          <div className="filter-container">
                            <button
                              className={`filter-btn ${hasFilter ? "active" : ""}`}
                              onClick={(e) => {
                                e.stopPropagation();
                                setOpenFilterDropdown(
                                  openFilterDropdown === colIndex
                                    ? null
                                    : colIndex,
                                );
                              }}
                            >
                              ⚲
                            </button>
                            {openFilterDropdown === colIndex && (
                              <div className="filter-dropdown">
                                {getColumnUniqueValues(colIndex).map(
                                  (value) => {
                                    const isChecked =
                                      !columnFilters[colIndex] ||
                                      columnFilters[colIndex].has(value);
                                    return (
                                      <label key={value}>
                                        <input
                                          type="checkbox"
                                          checked={isChecked}
                                          onChange={(e) => {
                                            const existing = columnFilters[
                                              colIndex
                                            ]
                                              ? new Set(columnFilters[colIndex])
                                              : new Set(
                                                  getColumnUniqueValues(
                                                    colIndex,
                                                  ),
                                                );
                                            if (e.target.checked) {
                                              existing.add(value);
                                            } else {
                                              existing.delete(value);
                                            }
                                            if (
                                              existing.size ===
                                              getColumnUniqueValues(colIndex)
                                                .length
                                            ) {
                                              clearColumnFilter(colIndex);
                                            } else {
                                              setColumnFilter(
                                                colIndex,
                                                existing,
                                              );
                                            }
                                          }}
                                        />
                                        <span>{value || "(empty)"}</span>
                                      </label>
                                    );
                                  },
                                )}
                                <div>
                                  <button
                                    onClick={() => {
                                      clearColumnFilter(colIndex);
                                      setOpenFilterDropdown(null);
                                    }}
                                  >
                                    Clear
                                  </button>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {getVisibleRows.map((rowIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{rowIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected =
                      selectedCell?.r === rowIndex &&
                      selectedCell?.c === colIndex;
                    const isInSelection = isCellInSelection(rowIndex, colIndex);
                    const isEditing =
                      editingCell?.r === rowIndex &&
                      editingCell?.c === colIndex;
                    const cellData = engine.getCell(rowIndex, colIndex);
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {};
                    const displayValue = cellData.error
                      ? cellData.error
                      : cellData.computed !== null && cellData.computed !== ""
                        ? String(cellData.computed)
                        : cellData.raw;

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? "selected" : ""} ${isInSelection ? "in-selection" : ""}`}
                        style={{ background: style.bg || "white" }}
                        onMouseDown={(e) => {
                          e.preventDefault();
                          setIsDragging(true);
                          setSelectionStart({ r: rowIndex, c: colIndex });
                          setSelectionEnd({ r: rowIndex, c: colIndex });
                          handleCellClick(rowIndex, colIndex);
                        }}
                        onMouseEnter={() => {
                          if (isDragging)
                            setSelectionEnd({ r: rowIndex, c: colIndex });
                        }}
                        onMouseUp={() => setIsDragging(false)}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) =>
                              handleKeyDown(e, rowIndex, colIndex)
                            }
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? "bold" : "normal",
                              fontStyle: style.italic ? "italic" : "normal",
                              textDecoration: style.underline
                                ? "underline"
                                : "none",
                              color: style.color || "#202124",
                              fontSize: (style.fontSize || 13) + "px",
                              textAlign: style.align || "left",
                              background: style.bg || "white",
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || "left"} ${cellData.error ? "error" : ""}`}
                            style={{
                              fontWeight: style.bold ? "bold" : "normal",
                              fontStyle: style.italic ? "italic" : "normal",
                              textDecoration: style.underline
                                ? "underline"
                                : "none",
                              color: cellData.error
                                ? "#d93025"
                                : style.color || "#202124",
                              fontSize: (style.fontSize || 13) + "px",
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas:
          =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  );
}
