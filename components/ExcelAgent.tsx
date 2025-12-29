"use client";

import { useCallback, useMemo, useState } from "react";
import { create } from "zustand";
import clsx from "clsx";
import * as XLSX from "xlsx";

type RawSheetData = {
  name: string;
  headers: string[];
  rows: (string | number | boolean | null | undefined)[][];
};

type AgentState = {
  workbookName: string;
  sheets: RawSheetData[];
  activeSheetIndex: number;
  selectedColumns: Record<string, boolean>;
  columnAliases: Record<string, string>;
  searchTerm: string;
  setWorkbook: (payload: {
    workbookName: string;
    sheets: RawSheetData[];
  }) => void;
  setActiveSheetIndex: (index: number) => void;
  toggleColumn: (columnKey: string) => void;
  setColumnAlias: (columnKey: string, alias: string) => void;
  setSearchTerm: (term: string) => void;
  reset: () => void;
};

const useAgentStore = create<AgentState>((set, get) => ({
  workbookName: "",
  sheets: [],
  activeSheetIndex: 0,
  selectedColumns: {},
  columnAliases: {},
  searchTerm: "",
  setWorkbook: ({ workbookName, sheets }) => {
    const headers =
      sheets.at(0)?.headers ?? sheets.at(0)?.rows.at(0)?.map(String) ?? [];
    const initialSelectedColumns: Record<string, boolean> = {};
    headers.forEach((header, idx) => {
      const key = buildColumnKey(header, idx);
      initialSelectedColumns[key] = true;
    });
    set({
      workbookName,
      sheets,
      activeSheetIndex: 0,
      selectedColumns: initialSelectedColumns,
      columnAliases: {},
      searchTerm: ""
    });
  },
  setActiveSheetIndex: (index: number) =>
    set(() => {
      const sheet = get().sheets.at(index);
      if (!sheet) {
        return {};
      }
      const initialSelectedColumns: Record<string, boolean> = {};
      sheet.headers.forEach((header, idx) => {
        const key = buildColumnKey(header, idx);
        initialSelectedColumns[key] = true;
      });
      return {
        activeSheetIndex: index,
        selectedColumns: initialSelectedColumns,
        columnAliases: {},
        searchTerm: ""
      };
    }),
  toggleColumn: (columnKey: string) =>
    set((state) => ({
      selectedColumns: {
        ...state.selectedColumns,
        [columnKey]: !state.selectedColumns[columnKey]
      }
    })),
  setColumnAlias: (columnKey: string, alias: string) =>
    set((state) => ({
      columnAliases: {
        ...state.columnAliases,
        [columnKey]: alias
      }
    })),
  setSearchTerm: (term: string) => set({ searchTerm: term }),
  reset: () =>
    set({
      workbookName: "",
      sheets: [],
      activeSheetIndex: 0,
      selectedColumns: {},
      columnAliases: {},
      searchTerm: ""
    })
}));

const buildColumnKey = (header: string | undefined, index: number) =>
  header?.trim() ? `${header.trim()}::${index}` : `ستون-${index + 1}`;

const normalizeForSearch = (value: unknown) =>
  String(value ?? "")
    .toLocaleLowerCase("fa")
    .normalize("NFKD");

const useAgent = () => {
  const store = useAgentStore();
  const sheet = store.sheets.at(store.activeSheetIndex);
  const visibleRows = useMemo(() => {
    if (!sheet) {
      return [];
    }
    if (!store.searchTerm.trim()) {
      return sheet.rows;
    }
    const searchValue = normalizeForSearch(store.searchTerm);
    return sheet.rows.filter((row) =>
      row.some((cell) => normalizeForSearch(cell).includes(searchValue))
    );
  }, [sheet, store.searchTerm]);

  const headers = useMemo(
    () => sheet?.headers ?? [],
    [sheet]
  );
  const columnKeys = headers.map(buildColumnKey);

  const selectedColumnIndices = columnKeys
    .map((key, idx) => ({ key, idx }))
    .filter(({ key }) => store.selectedColumns[key])
    .map(({ idx }) => idx);

  const rowsForExport = useMemo(() => {
    if (!sheet) return [];
    if (!selectedColumnIndices.length) return [];
    const headerRow = selectedColumnIndices.map((idx) => {
      const key = columnKeys[idx];
      return store.columnAliases[key]?.trim() || headers[idx] || `ستون ${idx + 1}`;
    });
    const dataRows = visibleRows.map((row) =>
      selectedColumnIndices.map((idx) => row[idx] ?? "")
    );
    return [headerRow, ...dataRows];
  }, [
    sheet,
    selectedColumnIndices,
    visibleRows,
    store.columnAliases,
    headers,
    columnKeys
  ]);

  return {
    ...store,
    activeSheet: sheet,
    columnKeys,
    rowsForExport,
    hasWorkbook: Boolean(store.sheets.length),
    selectedColumnCount: selectedColumnIndices.length,
    visibleRowCount: visibleRows.length
  };
};

const readWorkbookFile = async (file: File): Promise<RawSheetData[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  return workbook.SheetNames.map((name) => {
    const worksheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json<(string | number | boolean)[]>(
      worksheet,
      { header: 1, defval: "" }
    );
    const headers = (rows.shift() ?? []).map((cell, idx) => {
      const value = String(cell ?? "").trim();
      return value || `ستون ${idx + 1}`;
    });
    return {
      name,
      headers,
      rows
    };
  });
};

const downloadWorkbook = (
  rows: (string | number | boolean | null | undefined)[][],
  filename: string
) => {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "نتیجه");
  XLSX.writeFile(workbook, filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`);
};

export default function ExcelAgent() {
  const {
    workbookName,
    hasWorkbook,
    sheets,
    activeSheet,
    activeSheetIndex,
    setActiveSheetIndex,
    setWorkbook,
    reset,
    columnKeys,
    selectedColumns,
    toggleColumn,
    columnAliases,
    setColumnAlias,
    searchTerm,
    setSearchTerm,
    rowsForExport,
    selectedColumnCount,
    visibleRowCount
  } = useAgent();

  const [exportFilename, setExportFilename] = useState("نتیجه");
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileChange = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;
      setError(null);
      try {
        const sheets = await readWorkbookFile(file);
        setWorkbook({
          workbookName: file.name,
          sheets
        });
      } catch (err) {
        console.error(err);
        setError("مشکلی در خواندن فایل رخ داد. لطفاً دوباره تلاش کنید.");
      }
    },
    [setWorkbook]
  );

  const handleExport = useCallback(() => {
    if (!rowsForExport.length) {
      setError(
        "ستونی برای خروجی انتخاب نشده است یا داده‌ای برای ذخیره وجود ندارد."
      );
      return;
    }
    setIsExporting(true);
    setError(null);
    try {
      downloadWorkbook(rowsForExport, exportFilename);
    } catch (err) {
      console.error(err);
      setError("در ذخیره فایل مشکلی پیش آمد. دوباره تلاش کنید.");
    } finally {
      setIsExporting(false);
    }
  }, [rowsForExport, exportFilename]);

  return (
    <section className="mx-auto flex w-full max-w-6xl flex-col gap-6 px-4 py-10">
      <header className="flex flex-col gap-2 text-right">
        <h1 className="text-3xl font-semibold">ایجنت مدیریت فایل اکسل</h1>
        <p className="text-sm text-slate-600">
          فایل اکسل خود را بارگذاری کنید، ستون‌های دلخواه را انتخاب یا نام‌گذاری
          مجدد کنید و خروجی را به سرعت ذخیره نمایید.
        </p>
      </header>

      <div className="flex flex-col gap-3 rounded-xl border border-slate-200 bg-white p-6 shadow-sm">
        <label
          htmlFor="excel-upload"
          className="flex cursor-pointer flex-col items-center justify-center gap-3 rounded-lg border border-dashed border-slate-300 bg-slate-50 px-6 py-8 transition hover:border-slate-400 hover:bg-slate-100"
        >
          <span className="text-lg font-medium">فایل اکسل من را انتخاب کن</span>
          <span className="text-xs text-slate-500">
            فرمت‌های پشتیبانی‌شده: .xlsx، .xls
          </span>
          <input
            id="excel-upload"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileChange}
            className="sr-only"
          />
        </label>

        {hasWorkbook && (
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg bg-slate-100 px-4 py-3 text-sm">
            <span>
              فایل بارگذاری‌شده:{" "}
              <strong className="font-semibold text-slate-800">
                {workbookName}
              </strong>
            </span>
            <button
              type="button"
              className="rounded-md border border-slate-300 px-3 py-1 text-xs text-slate-600 transition hover:border-rose-400 hover:text-rose-600"
              onClick={reset}
            >
              حذف و بارگذاری فایل جدید
            </button>
          </div>
        )}
      </div>

      {hasWorkbook && activeSheet && (
        <div className="grid gap-6 lg:grid-cols-[1.1fr_0.9fr]">
          <div className="flex flex-col gap-4">
            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <span className="block text-sm font-semibold text-slate-500">
                انتخاب شیت
              </span>
              <div className="mt-3 flex flex-wrap gap-2">
                {sheets.map((sheet, index) => (
                  <button
                    key={sheet.name}
                    type="button"
                    onClick={() => setActiveSheetIndex(index)}
                    className={clsx(
                      "rounded-lg border px-4 py-2 text-sm transition",
                      index === activeSheetIndex
                        ? "border-blue-500 bg-blue-50 text-blue-600 shadow"
                        : "border-slate-300 bg-slate-50 text-slate-700 hover:border-blue-300 hover:text-blue-500"
                    )}
                  >
                    {sheet.name}
                  </button>
                ))}
              </div>
            </div>

            <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="flex flex-col gap-3">
                <div className="flex flex-col gap-1">
                  <label htmlFor="search" className="text-sm font-semibold">
                    جستجو در داده‌ها
                  </label>
                  <input
                    id="search"
                    type="search"
                    value={searchTerm}
                    onChange={(event) => setSearchTerm(event.target.value)}
                    placeholder="کلمه یا عبارتی را وارد کنید..."
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                  />
                </div>
                <div className="flex items-center justify-between text-xs text-slate-500">
                  <span>تعداد ردیف‌های قابل رؤیت: {visibleRowCount}</span>
                  <span>ستون‌های انتخاب‌شده: {selectedColumnCount}</span>
                </div>
              </div>

              <div className="mt-4 overflow-x-auto">
                <table className="w-full min-w-[640px] table-fixed border-separate border-spacing-y-1 text-left text-xs">
                  <thead>
                    <tr>
                      {columnKeys.map((key, index) => (
                        <th
                          key={key}
                          className="whitespace-nowrap rounded-md bg-slate-100 px-3 py-2 font-semibold text-slate-600"
                        >
                          <label className="flex items-center gap-2">
                            <input
                              type="checkbox"
                              checked={selectedColumns[key] ?? false}
                              onChange={() => toggleColumn(key)}
                              className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                            />
                            <span>{activeSheet.headers[index]}</span>
                          </label>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {activeSheet.rows.slice(0, 8).map((row, rowIndex) => (
                      <tr key={rowIndex} className="bg-white shadow-sm">
                        {row.map((cell, colIndex) => (
                          <td key={colIndex} className="px-3 py-2">
                            {String(cell ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                    {!activeSheet.rows.length && (
                      <tr>
                        <td
                          colSpan={columnKeys.length || 1}
                          className="px-3 py-6 text-center text-slate-400"
                        >
                          داده‌ای برای نمایش وجود ندارد.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div className="flex flex-col gap-4">
            <div className="flex flex-col gap-3 rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
              <h2 className="text-base font-semibold text-slate-700">
                تنظیمات خروجی
              </h2>
              <div className="flex flex-col gap-2">
                <label
                  htmlFor="export-filename"
                  className="text-xs font-medium text-slate-500"
                >
                  نام فایل خروجی
                </label>
                <input
                  id="export-filename"
                  type="text"
                  value={exportFilename}
                  placeholder="مثلاً: گزارش-فروش"
                  onChange={(event) => setExportFilename(event.target.value)}
                  className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                />
              </div>
              <div className="flex flex-col gap-4">
                <span className="text-xs font-medium text-slate-500">
                  نام‌گذاری مجدد ستون‌ها
                </span>
                <div className="flex max-h-80 flex-col gap-3 overflow-y-auto rounded-lg border border-slate-200 bg-slate-50 p-3">
                  {columnKeys.map((key, index) => (
                    <div
                      key={key}
                      className={clsx(
                        "flex items-center gap-2 rounded-md border px-3 py-2 text-xs transition",
                        selectedColumns[key]
                          ? "border-blue-200 bg-white"
                          : "border-slate-200 bg-slate-100 text-slate-400"
                      )}
                    >
                      <span className="min-w-[120px] text-slate-600">
                        {activeSheet.headers[index]}
                      </span>
                      <input
                        type="text"
                        placeholder="نام جدید ستون"
                        value={columnAliases[key] ?? ""}
                        onChange={(event) =>
                          setColumnAlias(key, event.target.value)
                        }
                        disabled={!selectedColumns[key]}
                        className="flex-1 rounded-md border border-slate-200 px-2 py-1 focus:border-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:bg-slate-100"
                      />
                    </div>
                  ))}
                </div>
              </div>
              <button
                type="button"
                onClick={handleExport}
                disabled={isExporting}
                className="mt-2 inline-flex items-center justify-center rounded-lg bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-300 disabled:cursor-not-allowed disabled:bg-blue-300"
              >
                {isExporting ? "در حال ذخیره..." : "ساخت فایل اکسل جدید"}
              </button>
              {error && (
                <p className="text-xs text-rose-500">
                  {error}
                </p>
              )}
            </div>

            <div className="rounded-xl border border-slate-200 bg-white p-4 text-sm text-slate-600 shadow-sm">
              <h3 className="mb-2 text-base font-semibold text-slate-700">
                نکات کار با ایجنت
              </h3>
              <ul className="list-disc space-y-2 pr-4">
                <li>برای دقت بیشتر، ستون‌های غیرضروری را غیرفعال کنید.</li>
                <li>قادر هستید نام ستون‌ها را پیش از خروجی شخصی‌سازی کنید.</li>
                <li>
                  خروجی به صورت خودکار با فرمت اکسل ذخیره می‌شود؛ برای فرمت‌های
                  دیگر می‌توانید از ابزارهای تبدیل استفاده کنید.
                </li>
              </ul>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
