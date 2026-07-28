# Center Manager App - Default Coding & UI Rules

The following mandatory rules MUST be strictly followed across all features, tables, search indices, UI components, and git workflows in this repository:

---

## 🚀 1. Git Auto Quick-Commit, Version Bump & Push Standard
- **Automatic Sync to Remote Repository**:
  - Remote Origin URL: `https://github.com/CallMeMrPenguin/Center_manager.git` (Branch: `main`).
  - **Mandatory End-of-Task Action**: After making code updates or completing features in any coding session, the agent MUST automatically perform a quick-commit & push to GitHub:
    ```bash
    git add . ; git commit -m "feat/fix: [Brief summary of changes made]" ; git push origin main
    ```
- **Auto Push Update & Version Bump Standard**:
  - Whenever asked to release an update, bump version, or publish changes:
    1. Bump the version string in `VERSION` file (e.g., `1.0.0` → `1.0.1`).
    2. Rebuild the frontend (`npm run build`) and Windows installer (`python installer/build_installer.py`).
    3. Commit and push commits along with git tags:
       ```bash
       git add . ; git commit -m "release: v1.0.1" ; git tag v1.0.1 ; git push origin main --tags
       ```
    4. All installed user app instances will automatically detect the new tag via `updater.py` and update in-place without manual reinstall.

---

## 📋 2. Table Pagination & Sorting Standards
- **Strict 20 Rows / Page**: Every table in the application MUST stick to **20 rows per page** by default.
- **No Row-Count Dropdowns**: Selection dropdowns allowing row-count changing are strictly forbidden.
- **Pagination Controls**: Tables exceeding 20 rows must feature standard `Trước`, `Trang X / Y`, `Sau` pagination controls.

---

## 🔍 3. Table Column Heading Filter Pattern (Ngân Hàng Câu Hỏi Standard)
Every data table MUST feature a filter icon next to each column title in the table header (`<th>`), working identically to the **Ngân Hàng Câu Hỏi** column filter system:

### ⚙️ How Column Heading Filters Work:
1. **Filter Icon Trigger**: Next to each heading text, render a small `Filter` icon. If a column has an active filter or sort applied, the filter icon highlights in active blue (`text-blue-400 bg-blue-500/10`).
2. **Interactive Filter Popover**:
   - Clicking the filter icon opens a floating popover card (`filter-dropdown-menu z-50`).
   - **Title Bar**: Shows `Lọc: [Tên Cột]` with a close (`X`) button.
   - **Sort Buttons**: Includes `Sắp xếp Tăng dần (A-Z)` and `Sắp xếp Giảm dần (Z-A)` options.
   - **Search Input**: Includes a search input to quickly filter unique column values.
   - **Value List Checkboxes**: Renders a scrollable list of unique values extracted from table records, allowing multi-select checking.
   - **Clear Button**: Includes a `Xóa bộ lọc cột` button to reset that column's filter.
3. **Optimized Performance**: Unique column values MUST be cached with `useMemo` / `useCallback` to ensure instantaneous filtering with zero UI lag across thousands of records.

---

## 🔍 4. Search & Index Filtering
- **Case-Insensitive Search**: All search queries and index filters MUST be case-insensitive (e.g. using `.toLowerCase().trim()`). Capitalization must never prevent a search match.
- **Live Filtering**: Tables must provide responsive live-search filtering for student names, nicknames, class names, or grades.

---

## 📈 5. Graph & Chart Visual Standards
- **Liquid Circular Glow**: SVG data point glow filters MUST use expansive filter bounds (`x="-100%" y="-100%" width="300%" height="300%"`) or SVG radial gradients so that point glows are perfectly circular with ZERO clipped square edges.
- **SVG Circle Scale Center Origin**: Any SVG data point `<circle>` with hover scaling MUST specify `style={{ transformBox: 'fill-box', transformOrigin: 'center' }}` so hover scaling expands symmetrically from the dot's exact center rather than jumping to the top-left origin (0,0).
- **Data Point Hover Values**: Hovering over any data point on line/progress charts MUST render an interactive tooltip card displaying the exact numerical score values (`Check 1`, `Check 2`, `Homework`).
- **Generous Vertical Spacing**: Progress charts must maintain generous height (min 560px) to provide clear vertical separation between score grid lines.

---

## 📅 6. Date Picker & Calendar Styling
- **Dark Theme Calendar**: Any `<input type="date">` or calendar picker MUST match the application's dark theme (`color-scheme: dark !important`).
- **White Calendar Icon**: The calendar picker icon (`::-webkit-calendar-picker-indicator`) MUST match the text color (pure white `#ffffff` with `filter: invert(1) brightness(2) !important`).

---

## 🎨 7. Segmented Control UI & Sliding Indicator Standard
- **Unified Summary Card Aesthetics**: All KPI summary cards must use glowing glassmorphic backdrops with curated gradient glow borders (`.kpi-card-blue`, `.kpi-card-purple`, `.kpi-card-green`, `.kpi-card-amber`).
- **Sliding Pill Indicator Segmented Control**:
  All Segmented Control option groups (e.g. `Điểm Danh & Điểm / Sơ Đồ Lớp / Nhóm Bạn & Xung Đột`, `Lịch Tháng / Lịch Tuần / Danh Sách`, `1 Tháng / 2 Tháng / 3 Tháng / Tất Cả`) MUST implement an absolute sliding indicator pill backdrop (matching the `FlowTab.tsx` reference standard):
  1. **Outer Container**: `relative flex bg-[#0d1018] p-1 rounded-xl border border-white/10 text-xs shrink-0 font-bold select-none`.
  2. **Sliding Indicator Backdrop**:
     - Class: `absolute top-1 bottom-1 rounded-lg bg-[#5c36f5] shadow-[0_0_14px_rgba(92,54,245,0.5)] transition-all duration-300 ease-[cubic-bezier(0.4,0,0.2,1)] pointer-events-none`
     - Inline style for smooth physical sliding animation:
       - `left`: `calc( (100% / N) * index + 1px )` (or `4px` for index 0).
       - `width`: `calc( (100% / N) - 4px )`.
  3. **Option Buttons**: `flex-1 relative z-10 py-1 text-center transition-colors cursor-pointer`. Text transitions cleanly: `active ? 'text-white font-black' : 'text-slate-400 hover:text-white'`.

---

## 🚫 8. Data Integrity: No Grade 0 Defaults
- **No Grade 0 Entries**: Unentered or missing grades must NEVER be recorded as numeric `0` or `0.0`. Missing grades are recorded as `NULL` so they do not artificially drag down student averages.

---

## ⚡ 9. Zero-Blur GPU Performance Standard
- **No CSS Blur Effects (`backdrop-blur` / `blur-*`)**: CSS blur filters create severe GPU rendering and compositing bottlenecks on scrollable lists, review cards, and large data tables. All elements MUST use crisp, solid high-contrast dark theme surfaces (`bg-[#0c0f1d]`, `bg-[#121626]`, `bg-[#0d1018]`) without any `backdrop-blur` or `blur-*` classes.

---

## 🚀 10. High-Performance Table & List Rendering Rule
- **High-Performance Component Memoization**: All table rows (`QuestionTableRow`, `VocabTableRow`, `StudentTableRow`) and review card items MUST be extracted into standalone `React.memo()` components with pre-memoized text parsers (`useMemo`) to prevent unnecessary row re-renders on hover, checkbox selection, or filter updates.

---

## 🎨 11. Custom Dark Theme Dropdown Select Standard
- **Custom Dark Theme Select Component Required (`<CustomSelect />`)**: Native HTML `<select>` elements are strictly forbidden because OS-level popovers override CSS styles and display ugly light grey options. All dropdown selection MUST use the custom `<CustomSelect />` component matching the app's dark indigo space theme (`#0c0f1e` background, `#212c4b` border, `#5c36f5` selected highlight, white text, and chevron animation).

---

## ☀️ 12. White Icons & Date Picker Standard
- **Pure White Icons**: No icon or SVG may use black or near-black colors (`text-black`, `text-slate-900`). All date pickers (`<input type="date">`) MUST enforce `color-scheme: dark !important` and `::-webkit-calendar-picker-indicator { filter: invert(1) brightness(2) !important; }` so that calendar picker icons display in crisp pure white.

---

## ⚡ 13. Lightweight Modal Overlays (Zero Blur Lag)
- **Solid Dark Backdrop Overlays**: Small pop-up modal windows and context menus MUST NOT use `backdrop-blur-sm` or `backdrop-blur-xl`. All modal backdrops MUST use solid dark overlays (`bg-black/85`) to ensure 0-lag input typing and instant responsiveness.

---

## 📅 14. Custom Dark Theme DatePicker Component Standard
- **Custom Dark Theme DatePicker Required**: Native browser `<input type="date">` elements are strictly forbidden as OS-native widgets violate the dark UI theme with grey popovers. All date selection MUST use the custom `<CustomDatePicker />` component matching the app's dark indigo space theme (`#0c0f1e` background, `#212c4b` border, `#5c36f5` highlight, white text, and white calendar icon).

---

## 📐 15. Professional Data Separators (No Bullet Dots)
- **Vertical Pipe Separators (`|`)**: Bullet dot characters (`•`) MUST NOT be used to separate text items or metadata fields. Data items MUST be separated using clean vertical pipe lines (`|`) or structured flex badges.

---

## 🚫 16. Strictly No Emojis In UI & Parsed Content Standard
- **No Injected Emojis**: Emojis (e.g. `👉`, `✨`, `🎨`, `📌`, `↔️`, `↕️`, `👁️`, `🎯`, `📄`, `⚡`, `📤`, `☑️`, `📱`, `🚀`, `🚫`) MUST NOT be injected by serializers, text formatters, parsers, or UI components unless explicitly present in raw source content. All parsed text and UI code must remain strictly faithful to source text without inserting inferential symbols or emojis.

---

## 🔢 17. Single Decimal Truncation Standard (No Rounding Up/Down)
- **1-Decimal Figure Truncation (`format1Dec` / `trunc1Dec`)**: All calculated scores, session averages, academic predictions, quiz results, and Excel/PDF exports MUST be formatted/truncated to strictly **1 figure after the decimal point (`.`) without rounding up or down** (e.g. `6.21` -> `6.2`, `6.28` -> `6.2`, `8.86` -> `8.8`).

---

## ✏️ 16. Single Pen Action Button (Merged Delete inside Edit Modal)
- **Merged Single Pen Button**: Cards and tables MUST NOT display a separate Trash Bin icon (`<Trash2 />`) alongside an Edit Pen icon (`<Edit3 />`). Cards/rows MUST feature a single clean Pen action button (`<Edit3 />`). Deletion functionality MUST be provided as a red Delete button inside the Edit modal popup.

---

## 📊 17. TanStack Table Standard — MANDATORY FOR ALL TABLES

> **THIS IS A HARD RULE. ANY TABLE IN ANY PAGE OR COMPONENT MUST USE TANSTACK TABLE VIA THE `<DataTable />` WRAPPER. NO EXCEPTIONS.**

### 17.1 — Core Mandate

- **ALWAYS use `<DataTable />`** from `src/components/DataTable.tsx` for every data table. Never write a raw `<table>` with manual filtering, sorting, or pagination.
- The `DataTable` component is a **self-contained TanStack Table wrapper** with all 15 features built-in. Pages only define `data` and `columns`; they never manage table state.
- **Never add manual search state** (e.g. `const [search, setSearch] = useState('')`) in a page that uses `<DataTable />`. TanStack's `globalFilter` handles it internally.
- **Never add manual pagination state** (e.g. `currentPage`, `pageSize`) in a page that uses `<DataTable />`. TanStack's `getPaginationRowModel()` handles it internally.
- **Never add manual sort state** (e.g. `sortConfig`, `sortField`, `setSortDirection`) in a page that uses `<DataTable />`.
- **Never add manual column visibility state** (e.g. `visibleCols`, `showColPicker`) in a page. The built-in `ColumnVisibilityDropdown` handles it.

---

### 17.2 — Required Packages

These packages MUST be installed and used:

```bash
npm install @tanstack/react-table @tanstack/react-virtual @dnd-kit/core @dnd-kit/sortable @dnd-kit/utilities xlsx jspdf jspdf-autotable
```

| Package | Purpose |
|---|---|
| `@tanstack/react-table` | Core table engine |
| `@tanstack/react-virtual` | Virtual scrolling for 500+ rows |
| `@dnd-kit/core` + `@dnd-kit/sortable` + `@dnd-kit/utilities` | Drag-and-drop column reorder |
| `xlsx` | Excel (.xlsx) export |
| `jspdf` + `jspdf-autotable` | PDF export |

---

### 17.3 — `DataTableProps<TData>` — Full API Reference

```tsx
import { DataTable } from '@/components/DataTable';

<DataTable<MyType>
  // ── Core (required) ──────────────────────────────────────────
  data={rows}                         // TData[]
  columns={columns}                   // ColumnDef<TData, any>[]

  // ── Loading & Empty ──────────────────────────────────────────
  loading={false}                     // shows spinner
  loadingMessage="Đang tải..."
  emptyMessage="Không có dữ liệu"     // string | ReactNode

  // ── Pagination ───────────────────────────────────────────────
  pageSize={20}                       // default 20 rows/page
  showPagination={true}               // hides pagination bar if false

  // ── Feature Toggles (all default to sensible values) ─────────
  enableGlobalSearch={true}           // 🔍 search input in toolbar
  enableColumnVisibility={true}       // 👁 show/hide columns dropdown
  enableRowSelection={false}          // ☑️ checkbox column
  enableColumnResizing={true}         // ↔️ drag-to-resize column borders
  enableColumnReorder={true}          // ↕️ drag column headers to reorder
  enableGrouping={false}              // grouping rows by column value
  enableRowExpansion={false}          // expand rows with sub-content
  enableColumnPinning={false}         // pin columns left/right
  enableMultiSort={true}              // 🎯 click multiple headers to sort
  enableVirtualization={undefined}    // ⚡ auto-on when data.length > 500
  enableExport={true}                 // 📤 Excel + PDF export dropdown

  // ── Sticky Layout ─────────────────────────────────────────────
  stickyHeader={true}                 // 📌 header stays on scroll
  stickyFirstColumn={false}           // 📌 pin first column sticky

  // ── Initial State ─────────────────────────────────────────────
  initialSorting={[{ id: 'name', desc: false }]}
  initialColumnVisibility={{ phone: false }}
  initialColumnPinning={{ left: ['name'] }}

  // ── Callbacks ─────────────────────────────────────────────────
  onRowClick={(row) => handleSelect(row)}
  onSelectionChange={(rows) => setSelected(rows)}
  renderSubComponent={({ row }) => <Detail row={row} />}

  // ── Toolbar Slots (inject custom buttons beside search) ───────
  toolbarLeft={<MyCustomFilter />}
  toolbarRight={<RefreshButton />}
  searchPlaceholder="Tìm theo tên..."

  // ── Export ────────────────────────────────────────────────────
  exportFilename="students_export"
/>
```

---

### 17.4 — Column Definition Patterns (`ColumnDef<T>`)

**Always define columns with `useMemo` in the page component:**

```tsx
import { ColumnDef } from '@tanstack/react-table';

const columns = useMemo<ColumnDef<Student>[]>(() => [
  // Simple accessor column
  {
    accessorKey: 'full_name',
    header: 'Họ và Tên',
    cell: (info) => <span className="font-bold">{info.getValue<string>()}</span>,
  },

  // Computed accessor column (not a raw field)
  {
    id: 'attendance_pct',
    header: () => <div className="text-center w-full">Điểm Danh %</div>,
    accessorFn: (row) => row.present / row.total * 100,
    cell: ({ getValue }) => <div className="text-center">{getValue<number>().toFixed(0)}%</div>,
  },

  // Disable sorting/filtering on a column
  {
    id: 'actions',
    header: 'Thao Tác',
    enableSorting: false,
    enableGlobalFilter: false,
    cell: ({ row }) => <ActionButtons row={row.original} />,
  },
], [/* reactive deps like callbacks */]);
```

---

### 17.5 — TanStack APIs Used Inside `DataTable.tsx`

The `DataTable` component internally uses **all** of these TanStack Table APIs. Never re-implement them in pages:

**Row Models:**
- `getCoreRowModel()` — base model (required)
- `getSortedRowModel()` — column sort
- `getFilteredRowModel()` — global + column filter
- `getPaginationRowModel()` — page slicing
- `getGroupedRowModel()` — row grouping
- `getExpandedRowModel()` — expanded rows / sub-rows

**State managed internally:**
- `globalFilter` / `onGlobalFilterChange` — global text search
- `sorting` / `onSortingChange` — sort state (supports multi-column)
- `columnFilters` / `onColumnFiltersChange` — per-column filters
- `columnVisibility` / `onColumnVisibilityChange` — show/hide columns
- `rowSelection` / `onRowSelectionChange` — checkbox selection
- `grouping` / `onGroupingChange` — grouping state
- `expanded` / `onExpandedChange` — row expansion
- `columnPinning` / `onColumnPinningChange` — sticky columns
- `columnOrder` / `onColumnOrderChange` — drag reorder order

**Column resize:**
- `columnResizeMode: 'onChange'` — live resize as you drag
- `header.getResizeHandler()` — mouse + touch resize handler
- `header.column.getIsResizing()` — resize active indicator
- `table.getTotalSize()` — total table width for resize layout

**Render helpers:**
- `flexRender(header.column.columnDef.header, header.getContext())` — header render
- `flexRender(cell.column.columnDef.cell, cell.getContext())` — cell render
- `table.getHeaderGroups()` — header rows
- `table.getRowModel().rows` — current visible rows (after filter+sort+page)
- `table.getFilteredRowModel().rows` — all filtered rows (for export count)

**Sort helpers (per column):**
- `header.column.getCanSort()` — is sortable?
- `header.column.getToggleSortingHandler()` — click handler
- `header.column.getIsSorted()` — `'asc' | 'desc' | false`

**Visibility helpers (per column):**
- `col.getIsVisible()` — is column visible?
- `col.getToggleVisibilityHandler()` — toggle handler
- `table.toggleAllColumnsVisible(bool)` — show/hide all

**Pinning helpers:**
- `header.column.getIsPinned()` — `'left' | 'right' | false`
- `header.column.getStart('left')` — sticky left offset
- `cell.column.getAfter('right')` — sticky right offset

**Pagination helpers:**
- `table.getState().pagination.pageIndex` — current page (0-indexed)
- `table.getPageCount()` — total pages
- `table.getCanPreviousPage()` / `table.getCanNextPage()`
- `table.previousPage()` / `table.nextPage()`
- `table.setPageIndex(n)` / `table.setPageSize(n)`

**Selection helpers:**
- `table.getIsAllPageRowsSelected()` — all on page selected?
- `table.getIsSomePageRowsSelected()` — some selected?
- `table.getToggleAllPageRowsSelectedHandler()` — toggle all handler
- `row.getIsSelected()` / `row.getIsSomeSelected()`
- `row.getToggleSelectedHandler()`

**Virtual scroll (`@tanstack/react-virtual`):**
- `useVirtualizer({ count, getScrollElement, estimateSize, overscan })`
- `virtualizer.getVirtualItems()` — only render these indices
- `virtualizer.getTotalSize()` — total scrollable height
- Auto-activates when `data.length > 500`

**DnD column reorder (`@dnd-kit`):**
- `DndContext` + `SortableContext` wraps the table header
- `useSortable({ id: header.id })` on each `<DraggableHeader>`
- `onDragEnd` → `arrayMove(columnOrder, oldIndex, newIndex)` → `setColumnOrder`
- `GripVertical` icon on each header as the drag handle

---

### 17.6 — 15 Built-In Features Checklist

When any new table is created, ALL of the following are available by default via `<DataTable />`:

| # | Feature | How |
|---|---|---|
| 1 | ✨ Smooth row hover | `transition-colors duration-150` + alternating row bg |
| 2 | 🎨 Rounded header & cells | `border-separate` + `rounded-tl/tr/bl/br-xl` on corners |
| 3 | 🌙 Custom dark theme | `#0d1018` / `#111827` / `#090c16` palette |
| 4 | 📌 Sticky header | `sticky top-0 z-20` on `<thead>` |
| 5 | 📌 Sticky first column | `stickyFirstColumn` prop → auto column pin |
| 6 | ↔️ Drag-to-resize | Resize handle on `<th>` right edge |
| 7 | ↕️ Drag column reorder | `@dnd-kit` + `GripVertical` handle on headers |
| 8 | 👁️ Show/hide columns | `ColumnVisibilityDropdown` in toolbar |
| 9 | 🔍 Global search | Search input in toolbar, TanStack `globalFilter` |
| 10 | 🎯 Multi-column sort | `enableMultiSort: true`, numbered priority badges |
| 11 | 📄 Pagination | Page pills, prev/next/first/last, page-size selector |
| 12 | ⚡ Virtual scrolling | `@tanstack/react-virtual`, auto-on at 500+ rows |
| 13 | 📤 Excel/PDF export | `xlsx` + `jspdf-autotable`, exports filtered data |
| 14 | ☑️ Row selection | Indeterminate checkboxes, selection badge, clear |
| 15 | 📱 Responsive layout | Toolbar wraps, table scrolls, labels collapse on mobile |

---

### 17.7 — Strict Anti-Patterns (FORBIDDEN)

```tsx
// ❌ FORBIDDEN — raw HTML table without TanStack
<table>
  <thead><tr><th>Name</th></tr></thead>
  <tbody>{data.map(row => <tr>...</tr>)}</tbody>
</table>

// ❌ FORBIDDEN — manual search state in a page with DataTable
const [search, setSearch] = useState('');
const filtered = data.filter(d => d.name.includes(search));
<DataTable data={filtered} ... />

// ❌ FORBIDDEN — manual pagination in a page with DataTable
const [page, setPage] = useState(1);
const paged = data.slice((page-1)*20, page*20);
<DataTable data={paged} ... />

// ❌ FORBIDDEN — manual sort in a page with DataTable
const [sortKey, setSortKey] = useState('name');
const sorted = [...data].sort((a,b) => a[sortKey].localeCompare(b[sortKey]));
<DataTable data={sorted} ... />

// ❌ FORBIDDEN — custom column visibility in a page with DataTable
const [visibleCols, setVisibleCols] = useState({ name: true, phone: false });
// ... custom dropdown UI ...
<DataTable data={data} columnVisibility={visibleCols} ... />

// ❌ FORBIDDEN — useReactTable called directly in a page
const table = useReactTable({ data, columns, getCoreRowModel: getCoreRowModel() });
// ... manual table rendering ...

// ✅ CORRECT — page passes raw data, DataTable handles everything
<DataTable
  data={students}          // raw, unfiltered, unsorted, unpaged
  columns={columns}        // just ColumnDef definitions
  searchPlaceholder="Tìm học sinh..."
  pageSize={20}
  enableRowSelection={true}
  onSelectionChange={setSelected}
/>
```

---

### 17.8 — When Adding a New Table (Checklist)

When any agent session adds a new table to any page:

1. ✅ Import `DataTable` from `src/components/DataTable`
2. ✅ Import `ColumnDef` from `@tanstack/react-table`
3. ✅ Define `columns` with `useMemo<ColumnDef<T>[]>(() => [...], [deps])`
4. ✅ Load data once with no server-side search/sort/page parameters
5. ✅ Pass raw `data` array directly — never pre-filter/sort/page it
6. ✅ Use `toolbarLeft` / `toolbarRight` for any page-specific controls (e.g. role filter dropdown, refresh button)
7. ✅ Set `exportFilename` to a descriptive name
8. ✅ Run `npm run build` — 0 TypeScript errors required
9. ✅ Commit with `git add . && git commit -m "feat: ..." && git push origin main`

---

## 💬 18. Zero Code Dumps in Conversation — MANDATORY

- **NO CODE DUMPING IN CHAT**: Agents must NEVER output long code blocks, full source code files, or extensive multi-line code diffs inside conversation text responses.
- **WRITE CODE DIRECTLY TO FILES**: Use file editing tools (`replace_file_content`, `write_to_file`) to apply all code changes directly to project files.
- **CONCISE SUMMARY WITH FILE LINKS**: Chat responses must contain only a concise high-level summary of what was accomplished, highlighting key decisions and providing clickable markdown links to modified files (`[filename](file:///path/to/file)`).

---

## 🧠 19. Core Behavioral Guidelines (Merged from RULE.md)

### 19.1 — Think Before Coding
- State assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them — do not pick silently.
- If a simpler approach exists, propose it.
- If requirements are ambiguous, clarify before mutating source code.

### 19.2 — Simplicity First
- Minimum code that solves the problem. Nothing speculative.
- No unrequested abstractions or unused configurability.
- If 50 lines solves a problem, do not write 200.

### 19.3 — Surgical Changes
- Touch only what is necessary to fulfill the request.
- Match existing repository style and patterns.
- Remove orphaned imports, variables, or functions created by your changes.

### 19.4 — Goal-Driven Execution
- Define verifiable success criteria before starting.
- Verify changes with concrete runtime commands (`npm run build`, API tests, git checks).
- Persevere through errors until verified success is achieved.

---

## 📄 20. DOCX Generation — Use pywin32 Exclusively

Any feature that creates, edits, or exports `.docx` files **MUST** use **`pywin32` (`win32com.client`)** — no exceptions.

### ❌ FORBIDDEN
- `python-docx` / `pydocx` or any other third-party DOCX library.
- Generating `.docx` via raw XML manipulation.

### ✅ REQUIRED
- Use `win32com.client.Dispatch("Word.Application")` (COM automation via `pywin32`).
- Keep `Visible = False` and `DisplayAlerts = False` on the Word Application instance.
- Always call `doc.Close(False)` and `word.Quit()` in a `finally` block to prevent orphaned Word processes.

### Why pywin32?
`pywin32` drives the full Microsoft Word engine via COM, giving access to every Word feature (styles, tables, mail merge, headers/footers, tracked changes, etc.) that `python-docx` / `pydocx` cannot replicate. It is the only correct choice for production-quality Word document generation on Windows.

### Quick Pattern
```python
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    doc = word.Documents.Add()
    # ... build document via COM API ...
    doc.SaveAs(os.path.abspath("output.docx"), FileFormat=16)  # 16 = wdFormatDocx
    doc.Close(False)
finally:
    word.Quit()
```
