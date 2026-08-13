import AsyncStorage from "@react-native-async-storage/async-storage";
import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as DocumentPicker from "expo-document-picker";
import * as FileSystem from "expo-file-system/legacy";
import * as Sharing from "expo-sharing";
import { useEffect, useMemo, useRef, useState } from "react";
import { FlatList, KeyboardAvoidingView, PanResponder, Platform, Pressable, ScrollView, StyleSheet, Text, TextInput, View } from "react-native";
import Svg, { Line, Path, Rect, Text as SvgText } from "react-native-svg";

import { ScreenContainer } from "@/components/screen-container";
import { appendFormulaDraftToken, buildRelativeFormula, getRelativeFormulaReferences, type RelativeFormulaPreset } from "@/lib/formula-editor";
import { displayCellValue } from "@/lib/formula-engine";
import {
  createEmptySheet,
  createInitialWorkbook,
  createTemplateSheet,
  DEFAULT_COLUMN_WIDTH,
  formatCellDisplay,
  formatLabel,
  getColumnWidth,
  getMergedRangeForAddress,
  getRangeBounds,
  isAddressInRange,
  isMergedChild,
  normalizeWorkbook,
  rangeAddresses,
  rangeNumericSummary,
  TEMPLATE_DETAILS,
  type CellNumberFormat,
  type CellRange,
  type SpreadsheetSheet,
  type SpreadsheetWorkbook,
  type TemplateKind,
} from "@/lib/spreadsheet-model";
import { exportSheetToCsv, exportWorkbookToXlsx, importSpreadsheetFile } from "@/lib/workbook-file-utils";

const WORKBOOK_KEY = "jadwali.workbook.v1";
const FORMAT_OPTIONS: CellNumberFormat[] = ["general", "currency", "percent", "decimal"];
const COLOR_OPTIONS = ["#FFFFFF", "#FFF7D6", "#E8F7F0", "#E9EEFF", "#FDECEC"];
const TOOLBAR_SECTIONS = [
  { id: "file", label: "ملف", icon: "folder" },
  { id: "edit", label: "تحرير", icon: "edit" },
  { id: "formulas", label: "صيغ", icon: "functions" },
  { id: "table", label: "جدول", icon: "table-view" },
  { id: "analysis", label: "تحليل", icon: "insert-chart-outlined" },
] as const;
type ToolbarSection = (typeof TOOLBAR_SECTIONS)[number]["id"];

function safeFileName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "-").slice(0, 40) || "jadwali";
}

function arrayBufferToBase64(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return globalThis.btoa(binary);
}

function base64ToArrayBuffer(base64: string) {
  const binary = globalThis.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes.buffer;
}

function columnLabel(index: number) {
  let label = "";
  let current = index;
  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);
  return label;
}

function arabicToLatinDigits(value: string) {
  const arabic = "٠١٢٣٤٥٦٧٨٩";
  const persian = "۰۱۲۳۴۵۶۷۸۹";
  return value
    .replace(/[٠-٩]/g, (digit) => String(arabic.indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String(persian.indexOf(digit)));
}

function columnIndex(label: string) {
  return label.split("").reduce((sum, character) => sum * 26 + character.charCodeAt(0) - 64, 0) - 1;
}

function parseAddress(value: string) {
  const normalized = arabicToLatinDigits(value.trim()).toUpperCase().replace(/\s/g, "");
  const match = normalized.match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) return null;
  return { address: normalized, column: columnIndex(match[1]), row: Number(match[2]) - 1 };
}

export default function HomeScreen() {
  const [workbook, setWorkbook] = useState<SpreadsheetWorkbook>(createInitialWorkbook);
  const [selectedCell, setSelectedCell] = useState("A1");
  const [draft, setDraft] = useState("");
  const [savedMessage, setSavedMessage] = useState("اختر خلية لبدء الإدخال");
  const [hasLoaded, setHasLoaded] = useState(false);
  const [lastWorkbook, setLastWorkbook] = useState<SpreadsheetWorkbook | null>(null);
  const [isFileAction, setIsFileAction] = useState(false);
  const [showGuide, setShowGuide] = useState(false);
  const [showFormatting, setShowFormatting] = useState(false);
  const [showTemplates, setShowTemplates] = useState(false);
  const [showCharts, setShowCharts] = useState(false);
  const [showSearch, setShowSearch] = useState(false);
  const [activeToolbarSection, setActiveToolbarSection] = useState<ToolbarSection>("formulas");
  const [chartType, setChartType] = useState<"bar" | "line">("bar");
  const [isRangeSelecting, setIsRangeSelecting] = useState(false);
  const [selectionRange, setSelectionRange] = useState<CellRange | null>(null);
  const [goToValue, setGoToValue] = useState("");
  const [searchValue, setSearchValue] = useState("");
  const horizontalGridRef = useRef<ScrollView>(null);
  const verticalGridRef = useRef<ScrollView>(null);
  const pageScrollRef = useRef<ScrollView>(null);
  const rangeStartRef = useRef<string | null>(null);

  const activeSheet = useMemo(
    () => workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0],
    [workbook],
  );
  const columns = useMemo(() => Array.from({ length: activeSheet.columnCount }, (_, index) => columnLabel(index)), [activeSheet.columnCount]);
  const rows = useMemo(() => Array.from({ length: activeSheet.rowCount }, (_, index) => index + 1), [activeSheet.rowCount]);
  const selectedFormat = activeSheet.cellFormats[selectedCell] ?? {};
  const formulaReferences = useMemo(() => getRelativeFormulaReferences(selectedCell), [selectedCell]);
  const additionFormula = useMemo(() => buildRelativeFormula(selectedCell, "add"), [selectedCell]);
  const rangeSummary = useMemo(() => (selectionRange ? rangeNumericSummary(activeSheet, selectionRange) : null), [activeSheet, selectionRange]);
  const chartData = useMemo(() => {
    if (!selectionRange) return [];
    return rangeAddresses(selectionRange.start, selectionRange.end)
      .map((address) => ({ address, value: Number(arabicToLatinDigits(displayCellValue(address, activeSheet.cells)).replace(/[^0-9.-]/g, "")) }))
      .filter((point) => Number.isFinite(point.value))
      .slice(0, 12);
  }, [activeSheet.cells, selectionRange]);
  const chartMax = useMemo(() => Math.max(1, ...chartData.map((point) => Math.abs(point.value))), [chartData]);
  const selectionSummary = useMemo(() => {
    const numericValues = Object.keys(activeSheet.cells)
      .map((address) => Number(displayCellValue(address, activeSheet.cells).replace(/,/g, "")))
      .filter((value) => Number.isFinite(value));
    const value = Number(formatCellDisplay(selectedCell, activeSheet).replace(/[^0-9.-]/g, ""));
    return { cellValue: Number.isFinite(value) ? value : null, total: numericValues.reduce((sum, item) => sum + item, 0), count: numericValues.length };
  }, [activeSheet, selectedCell]);

  const rangeResponder = useMemo(
    () =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => isRangeSelecting,
        onMoveShouldSetPanResponder: () => isRangeSelecting,
        onPanResponderGrant: (event) => {
          const x = event.nativeEvent.locationX;
          const y = event.nativeEvent.locationY;
          let offset = 37;
          const column = columns.find((label) => {
            const width = getColumnWidth(activeSheet, label);
            const matches = x >= offset && x < offset + width;
            offset += width;
            return matches;
          });
          const row = Math.floor(y / 42) + 1;
          if (!column || row < 1 || row > activeSheet.rowCount) return;
          const address = `${column}${row}`;
          rangeStartRef.current = address;
          setSelectionRange({ start: address, end: address });
          setSelectedCell(address);
          setDraft(activeSheet.cells[address] ?? "");
        },
        onPanResponderMove: (event) => {
          if (!rangeStartRef.current) return;
          const x = event.nativeEvent.locationX;
          const y = event.nativeEvent.locationY;
          let offset = 37;
          const column = columns.find((label) => {
            const width = getColumnWidth(activeSheet, label);
            const matches = x >= offset && x < offset + width;
            offset += width;
            return matches;
          });
          const row = Math.max(1, Math.min(activeSheet.rowCount, Math.floor(y / 42) + 1));
          if (column) setSelectionRange({ start: rangeStartRef.current, end: `${column}${row}` });
        },
        onPanResponderRelease: () => {
          rangeStartRef.current = null;
          setIsRangeSelecting(false);
          setSavedMessage("تم تحديد النطاق؛ يظهر المجموع والمتوسط في شريط الحالة");
        },
        onPanResponderTerminate: () => {
          rangeStartRef.current = null;
          setIsRangeSelecting(false);
        },
      }),
    [activeSheet, columns, isRangeSelecting],
  );

  useEffect(() => {
    let isActive = true;
    AsyncStorage.getItem(WORKBOOK_KEY)
      .then((stored) => {
        if (!stored || !isActive) return;
        const restored = normalizeWorkbook(JSON.parse(stored));
        if (restored) {
          setWorkbook(restored);
          setSavedMessage("تمت استعادة المصنف المحلي");
        }
      })
      .catch(() => setSavedMessage("تعذر استعادة النسخة السابقة"))
      .finally(() => isActive && setHasLoaded(true));
    return () => {
      isActive = false;
    };
  }, []);

  useEffect(() => {
    if (!hasLoaded) return;
    AsyncStorage.setItem(WORKBOOK_KEY, JSON.stringify(workbook)).catch(() => setSavedMessage("تعذر حفظ التغييرات محلياً"));
  }, [hasLoaded, workbook]);

  function updateActiveSheet(updater: (sheet: SpreadsheetSheet) => SpreadsheetSheet) {
    setLastWorkbook(workbook);
    setWorkbook((current) => ({
      ...current,
      sheets: current.sheets.map((sheet) => (sheet.id === current.activeSheetId ? updater(sheet) : sheet)),
    }));
  }

  function selectToolbarSection(section: ToolbarSection) {
    setActiveToolbarSection(section);
    if (section !== "file") setShowTemplates(false);
    if (section !== "edit") setShowSearch(false);
    if (section !== "table") setShowFormatting(false);
    if (section !== "analysis") setShowCharts(false);
  }

  function revealCell(address: string) {
    const destination = parseAddress(address);
    if (!destination) return;
    requestAnimationFrame(() => {
      const offset = columns.slice(0, destination.column).reduce((sum, column) => sum + getColumnWidth(activeSheet, column), 0);
      horizontalGridRef.current?.scrollTo({ x: Math.max(0, offset - DEFAULT_COLUMN_WIDTH), animated: true });
      verticalGridRef.current?.scrollTo({ y: Math.max(0, destination.row * 42 - 84), animated: true });
    });
  }

  function selectCell(address: string) {
    if (!isRangeSelecting) setSelectionRange(null);
    setSelectedCell(address);
    setDraft(activeSheet.cells[address] ?? "");
    setSavedMessage(`الخلية ${address} جاهزة للتحرير`);
    revealCell(address);
  }

  function saveCell() {
    const value = draft.trim();
    updateActiveSheet((sheet) => {
      const cells = { ...sheet.cells };
      if (value) cells[selectedCell] = value;
      else delete cells[selectedCell];
      return { ...sheet, cells };
    });
    setSavedMessage(`تم حفظ ${selectedCell} محلياً`);
  }

  function clearSelectedCell() {
    setDraft("");
    updateActiveSheet((sheet) => {
      const cells = { ...sheet.cells };
      delete cells[selectedCell];
      return { ...sheet, cells };
    });
    setSavedMessage(`تم مسح محتوى ${selectedCell}`);
  }

  function changeTableSize(axis: "rowCount" | "columnCount") {
    updateActiveSheet((sheet) => ({ ...sheet, [axis]: sheet[axis] + 1 }));
    setSavedMessage(axis === "rowCount" ? "تمت إضافة صف جديد" : "تمت إضافة عمود جديد");
  }

  function undoLastChange() {
    if (!lastWorkbook) {
      setSavedMessage("لا يوجد تعديل سابق للتراجع عنه");
      return;
    }
    const restored = lastWorkbook.sheets.find((sheet) => sheet.id === lastWorkbook.activeSheetId) ?? lastWorkbook.sheets[0];
    setWorkbook(lastWorkbook);
    setLastWorkbook(null);
    setSelectedCell("A1");
    setDraft(restored.cells.A1 ?? "");
    setSavedMessage("تم التراجع عن آخر تعديل");
  }

  function addSheet() {
    const nextNumber = workbook.sheets.length + 1;
    const next = createEmptySheet(`sheet-${Date.now()}`, `ورقة ${nextNumber}`);
    setLastWorkbook(workbook);
    setWorkbook((current) => ({ ...current, sheets: [...current.sheets, next], activeSheetId: next.id }));
    setSelectedCell("A1");
    setDraft("");
    setSavedMessage(`تم إنشاء ${next.name}`);
  }

  function chooseSheet(sheet: SpreadsheetSheet) {
    setWorkbook((current) => ({ ...current, activeSheetId: sheet.id }));
    setSelectedCell("A1");
    setDraft(sheet.cells.A1 ?? "");
    setSavedMessage(`تم فتح ${sheet.name}`);
  }

  function applyRelativeFormula(preset: RelativeFormulaPreset, label: string) {
    const formula = buildRelativeFormula(selectedCell, preset);
    setDraft(formula);
    updateActiveSheet((sheet) => ({ ...sheet, cells: { ...sheet.cells, [selectedCell]: formula } }));
    setSavedMessage(`تم تطبيق ${label} للصف ${formulaReferences.row} في ${selectedCell}`);
  }

  function appendFormulaToken(token: string) {
    setDraft((current) => appendFormulaDraftToken(current, token));
    setSavedMessage("استخدم المراجع والرموز، ثم اضغط ✓ لحساب النتيجة");
  }

  function keepFormulaVisible() {
    requestAnimationFrame(() => pageScrollRef.current?.scrollTo({ y: 0, animated: true }));
  }

  function loadGradeExample() {
    updateActiveSheet((sheet) => ({
      ...sheet,
      cells: { ...sheet.cells, A2: "أحمد محمد", B2: "٥", C2: "٥", D2: "=B2+C2" },
    }));
    setSelectedCell("D2");
    setDraft("=B2+C2");
    setShowGuide(false);
    setSavedMessage("تمت إضافة مثال الدرجات؛ النتيجة في D2 هي 10");
  }

  function updateCellFormat(patch: Partial<typeof selectedFormat>) {
    updateActiveSheet((sheet) => ({
      ...sheet,
      cellFormats: { ...sheet.cellFormats, [selectedCell]: { ...sheet.cellFormats[selectedCell], ...patch } },
    }));
    setSavedMessage(`تم تطبيق تنسيق على ${selectedCell}`);
  }

  function clearCellFormat() {
    updateActiveSheet((sheet) => {
      const cellFormats = { ...sheet.cellFormats };
      delete cellFormats[selectedCell];
      return { ...sheet, cellFormats };
    });
    setSavedMessage(`تمت إزالة تنسيق ${selectedCell}`);
  }

  function startRangeSelection() {
    setSelectionRange(null);
    rangeStartRef.current = null;
    setIsRangeSelecting(true);
    setSavedMessage("اسحب فوق الخلايا لتحديد نطاق وتحليله");
  }

  function mergeSelection() {
    if (!selectionRange) {
      setSavedMessage("حدد نطاقاً أفقياً أولاً ثم اختر دمج");
      return;
    }
    const bounds = getRangeBounds(selectionRange);
    if (!bounds || bounds.startColumn === bounds.endColumn || bounds.startRow !== bounds.endRow) {
      setSavedMessage("الدمج متاح لخليتين أو أكثر ضمن الصف نفسه");
      return;
    }
    const overlaps = activeSheet.mergedCells.some((merge) => rangeAddresses(merge.start, merge.end).some((address) => isAddressInRange(address, selectionRange)));
    if (overlaps) {
      setSavedMessage("يتداخل النطاق مع خلايا مدمجة مسبقاً");
      return;
    }
    updateActiveSheet((sheet) => ({ ...sheet, mergedCells: [...sheet.mergedCells, selectionRange] }));
    setSelectedCell(selectionRange.start);
    setSavedMessage(`تم دمج ${selectionRange.start}:${selectionRange.end}`);
  }

  function unmergeSelection() {
    const merge = getMergedRangeForAddress(activeSheet, selectedCell);
    if (!merge) {
      setSavedMessage("اختر خلية ضمن نطاق مدمج لإلغاء الدمج");
      return;
    }
    updateActiveSheet((sheet) => ({ ...sheet, mergedCells: sheet.mergedCells.filter((item) => item !== merge) }));
    setSelectionRange(null);
    setSavedMessage(`تم إلغاء دمج ${merge.start}:${merge.end}`);
  }

  function adjustColumnWidth(delta: number) {
    const match = selectedCell.match(/^([A-Z]+)/);
    if (!match) return;
    const column = match[1];
    const current = getColumnWidth(activeSheet, column);
    updateActiveSheet((sheet) => ({ ...sheet, columnWidths: { ...sheet.columnWidths, [column]: current + delta } }));
    setSavedMessage(`تم تعديل عرض العمود ${column}`);
  }

  function goToCell() {
    const destination = parseAddress(goToValue);
    if (!destination) {
      setSavedMessage("اكتب عنواناً صحيحاً مثل D25");
      return;
    }
    if (destination.column > 60 || destination.row > 500) {
      setSavedMessage("عنوان الخلية خارج النطاق المدعوم");
      return;
    }
    if (destination.column >= activeSheet.columnCount || destination.row >= activeSheet.rowCount) {
      updateActiveSheet((sheet) => ({ ...sheet, columnCount: Math.max(sheet.columnCount, destination.column + 1), rowCount: Math.max(sheet.rowCount, destination.row + 1) }));
    }
    selectCell(destination.address);
    setGoToValue("");
  }

  function findCellByText() {
    const query = searchValue.trim().toLocaleLowerCase("ar");
    if (!query) {
      setSavedMessage("اكتب اسماً أو قيمة للبحث");
      return;
    }
    const found = Object.entries(activeSheet.cells).find(([, value]) => value.toLocaleLowerCase("ar").includes(query));
    if (!found) {
      setSavedMessage("لا توجد خلية تطابق البحث في الورقة الحالية");
      return;
    }
    selectCell(found[0]);
    setSavedMessage(`تم العثور على النص في ${found[0]}`);
  }

  function applyTemplate(kind: TemplateKind) {
    const sheet = createTemplateSheet(kind, `template-${kind}-${Date.now()}`);
    setLastWorkbook(workbook);
    setWorkbook((current) => ({ ...current, sheets: [...current.sheets, sheet], activeSheetId: sheet.id }));
    setSelectedCell("A1");
    setDraft(sheet.cells.A1 ?? "");
    setShowTemplates(false);
    setSavedMessage(`تم إنشاء قالب ${TEMPLATE_DETAILS[kind].title} في ورقة جديدة`);
  }

  async function importFile() {
    if (isFileAction) return;
    setIsFileAction(true);
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: ["text/csv", "application/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
        copyToCacheDirectory: true,
        multiple: false,
      });
      if (result.canceled) {
        setSavedMessage("تم إلغاء اختيار الملف");
        return;
      }
      const asset = result.assets[0];
      const name = asset.name.toLowerCase();
      if (!name.endsWith(".csv") && !name.endsWith(".xlsx")) {
        setSavedMessage("الملف غير مدعوم؛ اختر CSV أو XLSX");
        return;
      }
      const content = name.endsWith(".csv")
        ? await FileSystem.readAsStringAsync(asset.uri, { encoding: FileSystem.EncodingType.UTF8 })
        : base64ToArrayBuffer(await FileSystem.readAsStringAsync(asset.uri, { encoding: FileSystem.EncodingType.Base64 }));
      const imported = importSpreadsheetFile(asset.name, content);
      if (!imported.sheets.length) {
        setSavedMessage("لا يحتوي الملف على أوراق قابلة للاستيراد");
        return;
      }
      const nextWorkbook: SpreadsheetWorkbook = { sheets: imported.sheets, activeSheetId: imported.sheets[0].id };
      setLastWorkbook(workbook);
      setWorkbook(nextWorkbook);
      setSelectedCell("A1");
      setDraft(imported.sheets[0].cells.A1 ?? "");
      setSavedMessage(imported.truncated ? "تم الاستيراد مع تقليص البيانات الكبيرة" : `تم استيراد ${imported.sheets.length} ورقة بنجاح`);
    } catch {
      setSavedMessage("تعذر قراءة الملف؛ تأكد من سلامة صيغة CSV أو XLSX");
    } finally {
      setIsFileAction(false);
    }
  }

  async function exportFile(format: "csv" | "xlsx") {
    if (isFileAction) return;
    setIsFileAction(true);
    try {
      if (!FileSystem.cacheDirectory) throw new Error("لا تتوفر مساحة ملفات مؤقتة");
      const extension = format === "csv" ? "csv" : "xlsx";
      const fileName = `${safeFileName(format === "csv" ? activeSheet.name : "jadwali-workbook")}.${extension}`;
      const fileUri = `${FileSystem.cacheDirectory}${fileName}`;
      const content = format === "csv" ? exportSheetToCsv(activeSheet) : arrayBufferToBase64(exportWorkbookToXlsx(workbook.sheets));
      await FileSystem.writeAsStringAsync(fileUri, content, { encoding: format === "csv" ? FileSystem.EncodingType.UTF8 : FileSystem.EncodingType.Base64 });
      if (!(await Sharing.isAvailableAsync())) {
        setSavedMessage("تم تجهيز الملف، لكن المشاركة غير متاحة على هذا الجهاز");
        return;
      }
      await Sharing.shareAsync(fileUri, { dialogTitle: format === "csv" ? "تصدير ورقة CSV" : "تصدير مصنف Excel", mimeType: format === "csv" ? "text/csv" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      setSavedMessage(format === "csv" ? "تم تجهيز CSV للتحميل أو المشاركة" : "تم تجهيز ملف Excel للتحميل أو المشاركة");
    } catch {
      setSavedMessage("تعذر تصدير الملف؛ حاول مرة أخرى");
    } finally {
      setIsFileAction(false);
    }
  }

  return (
    <ScreenContainer className="flex-1" containerClassName="bg-background">
      <KeyboardAvoidingView style={styles.keyboardAvoider} behavior={Platform.OS === "ios" ? "padding" : "height"} keyboardVerticalOffset={0}>
        <ScrollView ref={pageScrollRef} keyboardShouldPersistTaps="handled" keyboardDismissMode="interactive" contentContainerStyle={styles.pageScrollContent}>
          <View style={styles.page}>
        <View style={styles.header}>
          <View style={styles.brandGroup}>
            <View style={styles.logoTile}><MaterialIcons name="grid-on" size={21} color="#FFFFFF" /></View>
            <View><Text style={styles.brand}>جدولي</Text><Text style={styles.documentName}>مصنف محلي</Text></View>
          </View>
          <View style={styles.headerBadge}><View style={styles.headerBadgeDot} /><Text style={styles.headerBadgeText}>{hasLoaded ? "محفوظ محلياً" : "جارٍ التحميل"}</Text></View>
        </View>

        <View style={styles.toolbarCard}>
          <View style={styles.toolbarIntro}><Text style={styles.toolbarTitle}>أدوات الجدول</Text><Text style={styles.toolbarHint}>اختر عنواناً لعرض أدواته</Text></View>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.primaryToolbarRow}>
            {TOOLBAR_SECTIONS.map((section) => <Pressable key={section.id} onPress={() => selectToolbarSection(section.id)} style={({ pressed }) => [styles.primaryToolbarButton, activeToolbarSection === section.id && styles.activePrimaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name={section.icon} size={17} color={activeToolbarSection === section.id ? "#FFFFFF" : "#2457E5"} /><Text style={[styles.primaryToolbarText, activeToolbarSection === section.id && styles.activePrimaryToolbarText]}>{section.label}</Text></Pressable>)}
          </ScrollView>
          <View style={styles.toolbarDivider} />
          <ScrollView horizontal showsHorizontalScrollIndicator={false} keyboardShouldPersistTaps="handled" contentContainerStyle={styles.secondaryToolbarRow}>
            {activeToolbarSection === "file" && <>
              <Pressable onPress={() => setShowTemplates((visible) => !visible)} style={({ pressed }) => [styles.secondaryToolbarButton, showTemplates && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="dashboard-customize" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>قوالب</Text></Pressable>
              <Pressable onPress={importFile} disabled={isFileAction} style={({ pressed }) => [styles.secondaryToolbarButton, (pressed || isFileAction) && styles.toolPressed]}><MaterialIcons name="file-upload" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>استيراد</Text></Pressable>
              <Pressable onPress={() => exportFile("xlsx")} disabled={isFileAction} style={({ pressed }) => [styles.secondaryToolbarButton, (pressed || isFileAction) && styles.toolPressed]}><MaterialIcons name="table-view" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>Excel</Text></Pressable>
              <Pressable onPress={() => exportFile("csv")} disabled={isFileAction} style={({ pressed }) => [styles.secondaryToolbarButton, (pressed || isFileAction) && styles.toolPressed]}><MaterialIcons name="download" size={17} color="#16865B" /><Text style={[styles.secondaryToolbarText, styles.csvActionText]}>CSV</Text></Pressable>
            </>}
            {activeToolbarSection === "edit" && <>
              <Pressable onPress={undoLastChange} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="undo" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>تراجع</Text></Pressable>
              <Pressable onPress={clearSelectedCell} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="backspace" size={17} color="#C24141" /><Text style={[styles.secondaryToolbarText, styles.clearText]}>مسح الخلية</Text></Pressable>
              <Pressable onPress={() => setShowSearch((visible) => !visible)} style={({ pressed }) => [styles.secondaryToolbarButton, showSearch && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="search" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>بحث وانتقال</Text></Pressable>
              <Pressable onPress={() => setShowGuide((visible) => !visible)} style={({ pressed }) => [styles.secondaryToolbarButton, showGuide && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="help-outline" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>مساعدة</Text></Pressable>
            </>}
            {activeToolbarSection === "formulas" && <>
              <Pressable onPress={() => appendFormulaToken(formulaReferences.firstReference)} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.referenceButtonText}>{formulaReferences.firstReference}</Text></Pressable>
              <Pressable onPress={() => appendFormulaToken("+")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>+</Text></Pressable>
              <Pressable onPress={() => appendFormulaToken(formulaReferences.secondReference)} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.referenceButtonText}>{formulaReferences.secondReference}</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("add", "جمع الدرجتين")} style={({ pressed }) => [styles.secondaryToolbarButton, styles.emphasizedFormulaButton, pressed && styles.toolPressed]}><MaterialIcons name="functions" size={17} color="#FFFFFF" /><Text style={styles.emphasizedFormulaText}>جمع</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("subtract", "الطرح")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>−</Text><Text style={styles.secondaryToolbarText}>طرح</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("multiply", "الضرب")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>×</Text><Text style={styles.secondaryToolbarText}>ضرب</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("divide", "القسمة")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>÷</Text><Text style={styles.secondaryToolbarText}>قسمة</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("sum", "SUM")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>Σ</Text><Text style={styles.secondaryToolbarText}>SUM</Text></Pressable>
              <Pressable onPress={() => applyRelativeFormula("average", "المتوسط")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><Text style={styles.operatorText}>Σ</Text><Text style={styles.secondaryToolbarText}>متوسط</Text></Pressable>
            </>}
            {activeToolbarSection === "table" && <>
              <Pressable onPress={startRangeSelection} style={({ pressed }) => [styles.secondaryToolbarButton, isRangeSelecting && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="select-all" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>نطاق</Text></Pressable>
              <Pressable onPress={() => { setShowFormatting((visible) => !visible); }} style={({ pressed }) => [styles.secondaryToolbarButton, showFormatting && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="format-paint" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>تنسيق</Text></Pressable>
              <Pressable onPress={() => changeTableSize("rowCount")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="view-agenda" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>صف</Text></Pressable>
              <Pressable onPress={() => changeTableSize("columnCount")} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="view-column" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>عمود</Text></Pressable>
              <Pressable onPress={mergeSelection} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="merge-type" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>دمج</Text></Pressable>
            </>}
            {activeToolbarSection === "analysis" && <>
              <Pressable onPress={() => setShowCharts((visible) => !visible)} style={({ pressed }) => [styles.secondaryToolbarButton, showCharts && styles.activeSecondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="insert-chart-outlined" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>مخطط</Text></Pressable>
              <Pressable onPress={unmergeSelection} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="vertical-split" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>إلغاء الدمج</Text></Pressable>
              <Pressable onPress={() => adjustColumnWidth(-16)} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="unfold-less" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>أضيق</Text></Pressable>
              <Pressable onPress={() => adjustColumnWidth(16)} style={({ pressed }) => [styles.secondaryToolbarButton, pressed && styles.toolPressed]}><MaterialIcons name="unfold-more" size={17} color="#2457E5" /><Text style={styles.secondaryToolbarText}>أوسع</Text></Pressable>
            </>}
          </ScrollView>
        </View>

        {showTemplates && (
          <View style={styles.templatesCard}>
            <Text style={styles.sectionTitle}>ابدأ بقالب جاهز</Text>
            <View style={styles.templateOptions}>
              {(Object.keys(TEMPLATE_DETAILS) as TemplateKind[]).map((kind) => {
                const template = TEMPLATE_DETAILS[kind];
                return <Pressable key={kind} onPress={() => applyTemplate(kind)} style={({ pressed }) => [styles.templateButton, pressed && styles.pressed]}><MaterialIcons name={template.icon} size={19} color="#2457E5" /><View style={styles.templateTextGroup}><Text style={styles.templateTitle}>{template.title}</Text><Text style={styles.templateDescription}>{template.description}</Text></View></Pressable>;
              })}
            </View>
          </View>
        )}

        <View style={styles.formulaCard}>
          <View style={styles.formulaMeta}><View style={styles.addressBadge}><Text style={styles.addressText}>{selectedCell}</Text></View><Text style={styles.formulaLabel}>تحرير الخلية النشطة</Text></View>
          <View style={styles.editorRow}>
            <TextInput accessibilityLabel="محرر الخلية" value={draft} onChangeText={setDraft} onFocus={keepFormulaVisible} onSubmitEditing={saveCell} placeholder="اكتب =B2+C2 أو قيمة" placeholderTextColor="#8B99AE" style={styles.formulaInput} textAlign="right" returnKeyType="done" autoCapitalize="characters" />
            <Pressable accessibilityLabel="تأكيد تعديل الخلية" onPress={saveCell} style={({ pressed }) => [styles.saveButton, pressed && styles.pressed]}><MaterialIcons name="check" size={20} color="#FFFFFF" /></Pressable>
          </View>
          <View style={styles.formulaHintRow}>
            <Text style={styles.formulaHintText}>صف {formulaReferences.row} · جمع هذا الصف: <Text style={styles.formulaExample}>{additionFormula}</Text></Text>
            <Pressable onPress={() => setShowGuide((visible) => !visible)} style={({ pressed }) => [styles.formulaHelpButton, pressed && styles.toolPressed]}><MaterialIcons name="help-outline" size={15} color="#2457E5" /><Text style={styles.formulaHelpText}>شرح</Text></Pressable>
          </View>
          <Text style={styles.formulaWorkflow}>من شريط «صيغ»: اضغط الدالة لتطبيقها فوراً على {selectedCell}، أو استخدم المراجع لبناء صيغة يدوية ثم احفظها.</Text>
        </View>

        <View style={styles.guideCard}>
          <Pressable accessibilityLabel="فتح أو إغلاق شرح استخدام الجدول" onPress={() => setShowGuide((visible) => !visible)} style={({ pressed }) => [styles.guideHeader, pressed && styles.toolPressed]}>
            <View style={styles.guideHeading}><View style={styles.guideIcon}><MaterialIcons name="school" size={18} color="#2457E5" /></View><View><Text style={styles.guideTitle}>كيف أحسب مجموع هذا الصف؟</Text><Text style={styles.guideSubtitle}>المراجع: {formulaReferences.firstReference} و{formulaReferences.secondReference}</Text></View></View>
            <MaterialIcons name={showGuide ? "expand-less" : "expand-more"} size={22} color="#2457E5" />
          </Pressable>
          {showGuide && <View style={styles.guideBody}><Text style={styles.guideStep}>١. في الصف {formulaReferences.row} اكتب القيم في {formulaReferences.firstReference} و{formulaReferences.secondReference}.</Text><Text style={styles.guideStep}>٢. حدّد خلية النتيجة {selectedCell}، ثم اختر «صيغ» واضغط «جمع» لتطبيق <Text style={styles.formulaExample}>{additionFormula}</Text>.</Text><View style={styles.guideResultRow}><Text style={styles.guideResultText}>لن تُستخدم مراجع أي صف آخر.</Text><Pressable onPress={loadGradeExample} style={({ pressed }) => [styles.loadExampleButton, pressed && styles.pressed]}><MaterialIcons name="play-circle-outline" size={17} color="#FFFFFF" /><Text style={styles.loadExampleText}>جرّب المثال</Text></Pressable></View></View>}
        </View>

        {showSearch && <View style={styles.searchCard}>
          <View style={styles.searchSection}><MaterialIcons name="my-location" size={17} color="#2457E5" /><TextInput value={goToValue} onChangeText={setGoToValue} onSubmitEditing={goToCell} placeholder="انتقل إلى D25" placeholderTextColor="#8B99AE" style={styles.searchInput} textAlign="right" autoCapitalize="characters" returnKeyType="go" /><Pressable onPress={goToCell} style={({ pressed }) => [styles.searchButton, pressed && styles.pressed]}><Text style={styles.searchButtonText}>انتقال</Text></Pressable></View>
          <View style={styles.searchDivider} />
          <View style={styles.searchSection}><MaterialIcons name="search" size={17} color="#2457E5" /><TextInput value={searchValue} onChangeText={setSearchValue} onSubmitEditing={findCellByText} placeholder="ابحث عن اسم" placeholderTextColor="#8B99AE" style={styles.searchInput} textAlign="right" returnKeyType="search" /><Pressable onPress={findCellByText} style={({ pressed }) => [styles.searchButton, pressed && styles.pressed]}><Text style={styles.searchButtonText}>بحث</Text></Pressable></View>
        </View>}

        {showFormatting && <View style={styles.formatCard}><View style={styles.formatHeader}><Text style={styles.formatTitle}>تنسيق {selectedCell}</Text><Pressable onPress={clearCellFormat} style={({ pressed }) => [styles.clearFormatButton, pressed && styles.toolPressed]}><Text style={styles.clearFormatText}>إزالة التنسيق</Text></Pressable></View><Text style={styles.formatLabel}>عرض الرقم</Text><View style={styles.formatOptions}>{FORMAT_OPTIONS.map((format) => <Pressable key={format} onPress={() => updateCellFormat({ numberFormat: format })} style={({ pressed }) => [styles.formatChip, selectedFormat.numberFormat === format && styles.activeFormatChip, pressed && styles.toolPressed]}><Text style={[styles.formatChipText, selectedFormat.numberFormat === format && styles.activeFormatChipText]}>{formatLabel(format)}</Text></Pressable>)}</View><Text style={styles.formatLabel}>لون الخلفية</Text><View style={styles.colorOptions}>{COLOR_OPTIONS.map((color) => <Pressable key={color} accessibilityLabel={`تطبيق اللون ${color}`} onPress={() => updateCellFormat({ backgroundColor: color })} style={({ pressed }) => [styles.colorSwatch, { backgroundColor: color }, selectedFormat.backgroundColor === color && styles.activeColorSwatch, pressed && styles.toolPressed]}>{selectedFormat.backgroundColor === color && <MaterialIcons name="check" size={15} color="#2457E5" />}</Pressable>)}</View></View>}

        {showCharts && <View style={styles.chartCard}><View style={styles.chartHeader}><View><Text style={styles.chartTitle}>مخطط بيانات النطاق</Text><Text style={styles.chartSubtitle}>{selectionRange ? `${selectionRange.start}:${selectionRange.end}` : "حدد نطاقاً أولاً"}</Text></View><View style={styles.chartTypeSwitch}><Pressable onPress={() => setChartType("bar")} style={({ pressed }) => [styles.chartTypeButton, chartType === "bar" && styles.activeChartTypeButton, pressed && styles.toolPressed]}><MaterialIcons name="bar-chart" size={16} color="#2457E5" /></Pressable><Pressable onPress={() => setChartType("line")} style={({ pressed }) => [styles.chartTypeButton, chartType === "line" && styles.activeChartTypeButton, pressed && styles.toolPressed]}><MaterialIcons name="show-chart" size={16} color="#2457E5" /></Pressable></View></View>{chartData.length ? <Svg width={320} height={178} viewBox="0 0 320 178"><Line x1="26" y1="144" x2="305" y2="144" stroke="#CBD5E1" strokeWidth="1" />{chartType === "bar" ? chartData.map((point, index) => { const step = 232 / chartData.length; const width = Math.max(13, step - 6); const height = Math.max(3, Math.abs(point.value) / chartMax * 105); return <Rect key={point.address} x={34 + index * step} y={144 - height} width={width} height={height} rx="4" fill="#2457E5" />; }) : <Path d={`M ${chartData.map((point, index) => `${38 + index * (232 / Math.max(chartData.length - 1, 1))} ${144 - Math.abs(point.value) / chartMax * 105}`).join(" L ")}`} fill="none" stroke="#16865B" strokeWidth="3" />}{chartData.map((point, index) => <SvgText key={`${point.address}-label`} x={42 + index * (232 / Math.max(chartData.length, 1))} y="164" fill="#64748B" fontSize="9">{point.address}</SvgText>)}</Svg> : <Text style={styles.emptyChartText}>حدد نطاقاً يحتوي أرقاماً، ثم افتح المخطط.</Text>}</View>}

        <View style={styles.sheetCard}>
          <View style={styles.sheetTopline}><View style={styles.sheetTitleGroup}><View style={styles.activeSheetDot} /><Text style={styles.sheetTitle}>{activeSheet.name}</Text><View style={styles.frozenPill}><MaterialIcons name="push-pin" size={12} color="#64748B" /><Text style={styles.frozenPillText}>العناوين ثابتة</Text></View></View><Text style={styles.sheetHint}>{savedMessage}</Text></View>
          <ScrollView ref={horizontalGridRef} horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.gridScrollContent}>
            <View>
              <View style={styles.headerRow}><View style={styles.cornerCell} />{columns.map((column) => <View key={column} style={[styles.columnHeader, { width: getColumnWidth(activeSheet, column) }]}><Text style={styles.columnHeaderText}>{column}</Text></View>)}</View>
              <ScrollView ref={verticalGridRef} showsVerticalScrollIndicator={false} nestedScrollEnabled>
                <View {...rangeResponder.panHandlers}>{rows.map((row) => <View key={row} style={styles.gridRow}><View style={styles.rowHeader}><Text style={styles.rowHeaderText}>{row}</Text></View>{columns.map((column) => { const address = `${column}${row}`; const selected = address === selectedCell; const inRange = Boolean(selectionRange && isAddressInRange(address, selectionRange)); const merged = getMergedRangeForAddress(activeSheet, address); if (isMergedChild(activeSheet, address)) return null; const bounds = merged ? getRangeBounds(merged) : null; const mergedWidth = bounds ? columns.slice(bounds.startColumn, bounds.endColumn + 1).reduce((sum, item) => sum + getColumnWidth(activeSheet, item), 0) : getColumnWidth(activeSheet, column); const cellStyle = activeSheet.cellFormats[address]; return <Pressable key={address} accessibilityLabel={`الخلية ${address}`} onLongPress={startRangeSelection} onPress={() => selectCell(address)} style={({ pressed }) => [styles.cell, { width: mergedWidth }, cellStyle?.backgroundColor ? { backgroundColor: cellStyle.backgroundColor } : null, inRange && styles.rangeCell, selected && styles.selectedCell, pressed && styles.cellPressed]}><Text numberOfLines={1} style={[styles.cellText, cellStyle?.textColor ? { color: cellStyle.textColor } : null, selected && styles.selectedCellText]}>{formatCellDisplay(address, activeSheet)}</Text></Pressable>; })}</View>)}</View>
              </ScrollView>
            </View>
          </ScrollView>
          <View style={styles.sheetTabsBar}><FlatList horizontal data={workbook.sheets} keyExtractor={(sheet) => sheet.id} showsHorizontalScrollIndicator={false} contentContainerStyle={styles.tabsList} renderItem={({ item }) => <Pressable onPress={() => chooseSheet(item)} style={({ pressed }) => [styles.sheetTab, item.id === activeSheet.id && styles.activeSheetTab, pressed && styles.toolPressed]}><Text style={[styles.sheetTabText, item.id === activeSheet.id && styles.activeSheetTabText]}>{item.name}</Text></Pressable>} /><Pressable accessibilityLabel="إضافة ورقة عمل" onPress={addSheet} style={({ pressed }) => [styles.addSheetButton, pressed && styles.pressed]}><MaterialIcons name="add" size={20} color="#2457E5" /></Pressable></View>
        </View>

        <View style={styles.statusBar}>{rangeSummary ? <><View style={styles.statusItem}><MaterialIcons name="select-all" size={16} color="#2457E5" /><Text style={styles.statusText}>{selectionRange?.start}:{selectionRange?.end}</Text></View><View style={styles.statusDivider} /><View style={styles.statusItem}><MaterialIcons name="functions" size={16} color="#16865B" /><Text style={styles.statusText}>المجموع</Text><Text style={styles.statusValue}>{rangeSummary.total}</Text></View><View style={styles.statusDivider} /><View style={styles.statusItem}><Text style={styles.statusText}>المتوسط</Text><Text style={styles.statusValue}>{rangeSummary.average.toFixed(2)}</Text></View></> : <><View style={styles.statusItem}><MaterialIcons name="pin-drop" size={16} color="#2457E5" /><Text style={styles.statusText}>{selectedCell}</Text><Text style={styles.statusValue}>{selectionSummary.cellValue === null ? "نص أو فارغ" : selectionSummary.cellValue}</Text></View><View style={styles.statusDivider} /><View style={styles.statusItem}><MaterialIcons name="functions" size={16} color="#16865B" /><Text style={styles.statusText}>مجموع القيم</Text><Text style={styles.statusValue}>{selectionSummary.total}</Text></View><View style={styles.statusDivider} /><View style={styles.statusItem}><Text style={styles.statusText}>{selectionSummary.count} قيمة رقمية</Text></View></>}</View>
          </View>
        </ScrollView>
      </KeyboardAvoidingView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  keyboardAvoider: { flex: 1 },
  pageScrollContent: { flexGrow: 1 },
  page: { flexGrow: 1, backgroundColor: "#F7F9FE", paddingHorizontal: 16, paddingTop: 12, paddingBottom: 24 },
  header: { flexDirection: "row-reverse", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  brandGroup: { flexDirection: "row-reverse", alignItems: "center", gap: 10 },
  logoTile: { width: 39, height: 39, borderRadius: 13, alignItems: "center", justifyContent: "center", backgroundColor: "#2457E5", elevation: 3 },
  brand: { color: "#13213A", fontSize: 21, fontWeight: "800", lineHeight: 27, textAlign: "right" },
  documentName: { color: "#64748B", fontSize: 12, fontWeight: "500", lineHeight: 17, textAlign: "right" },
  headerBadge: { flexDirection: "row-reverse", alignItems: "center", gap: 5, paddingHorizontal: 9, paddingVertical: 6, borderRadius: 99, backgroundColor: "#E8F7F0" },
  headerBadgeDot: { width: 6, height: 6, borderRadius: 3, backgroundColor: "#16865B" },
  headerBadgeText: { color: "#16865B", fontSize: 11, fontWeight: "700" },
  toolbarCard: { marginBottom: 10, padding: 10, borderRadius: 16, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5", elevation: 1 },
  toolbarIntro: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  toolbarTitle: { color: "#13213A", fontSize: 13, fontWeight: "800" },
  toolbarHint: { color: "#64748B", fontSize: 10, fontWeight: "600" },
  primaryToolbarRow: { flexDirection: "row-reverse", gap: 6, paddingHorizontal: 1 },
  primaryToolbarButton: { flexDirection: "row-reverse", alignItems: "center", gap: 5, minHeight: 35, paddingHorizontal: 10, borderRadius: 10, backgroundColor: "#F7F9FE", borderWidth: 1, borderColor: "#E3EAF5" },
  activePrimaryToolbarButton: { backgroundColor: "#2457E5", borderColor: "#2457E5" },
  primaryToolbarText: { color: "#2457E5", fontSize: 11, fontWeight: "800" },
  activePrimaryToolbarText: { color: "#FFFFFF" },
  toolbarDivider: { height: 1, marginVertical: 9, backgroundColor: "#E8EDF6" },
  secondaryToolbarRow: { flexDirection: "row-reverse", gap: 6, paddingHorizontal: 1 },
  secondaryToolbarButton: { flexDirection: "row-reverse", alignItems: "center", gap: 4, minHeight: 33, paddingHorizontal: 9, borderRadius: 9, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  activeSecondaryToolbarButton: { backgroundColor: "#E9EEFF", borderColor: "#9AB4FF" },
  secondaryToolbarText: { color: "#2457E5", fontSize: 10, fontWeight: "800" },
  referenceButtonText: { color: "#2457E5", fontSize: 12, fontWeight: "900" },
  emphasizedFormulaButton: { backgroundColor: "#2457E5", borderColor: "#2457E5" },
  emphasizedFormulaText: { color: "#FFFFFF", fontSize: 10, fontWeight: "900" },
  csvActionText: { color: "#16865B" },
  templatesCard: { marginBottom: 10, padding: 12, borderRadius: 16, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  sectionTitle: { color: "#13213A", fontSize: 13, fontWeight: "800", textAlign: "right", marginBottom: 8 },
  templateOptions: { gap: 7 },
  templateButton: { flexDirection: "row-reverse", alignItems: "center", gap: 9, padding: 9, borderRadius: 11, backgroundColor: "#F7F9FE" },
  templateTextGroup: { flex: 1 },
  templateTitle: { color: "#2457E5", fontSize: 12, fontWeight: "800", textAlign: "right" },
  templateDescription: { color: "#64748B", fontSize: 10, fontWeight: "600", marginTop: 1, textAlign: "right" },
  guideCard: { marginBottom: 10, padding: 11, borderRadius: 16, backgroundColor: "#F0F4FF", borderWidth: 1, borderColor: "#D8E2FF" },
  guideHeader: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between" },
  guideHeading: { flexDirection: "row-reverse", alignItems: "center", gap: 9, flex: 1 },
  guideIcon: { width: 34, height: 34, alignItems: "center", justifyContent: "center", borderRadius: 11, backgroundColor: "#FFFFFF" },
  guideTitle: { color: "#13213A", fontSize: 14, fontWeight: "800", lineHeight: 20, textAlign: "right" },
  guideSubtitle: { color: "#64748B", fontSize: 11, fontWeight: "600", lineHeight: 17, textAlign: "right" },
  guideBody: { marginTop: 8, gap: 3 },
  guideStep: { color: "#40516D", fontSize: 11, fontWeight: "600", lineHeight: 18, textAlign: "right" },
  formulaExample: { color: "#2457E5", fontWeight: "900" },
  guideResultRow: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginTop: 3 },
  guideResultText: { color: "#40516D", fontSize: 12, fontWeight: "700" },
  guideResultValue: { color: "#16865B", fontSize: 15, fontWeight: "900" },
  loadExampleButton: { flexDirection: "row-reverse", alignItems: "center", gap: 5, paddingHorizontal: 10, paddingVertical: 7, borderRadius: 9, backgroundColor: "#2457E5" },
  loadExampleText: { color: "#FFFFFF", fontSize: 11, fontWeight: "800" },
  searchCard: { flexDirection: "row-reverse", alignItems: "center", marginBottom: 10, paddingHorizontal: 9, paddingVertical: 7, borderRadius: 13, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#E3EAF5" },
  searchSection: { flexDirection: "row-reverse", alignItems: "center", gap: 5, flex: 1 },
  searchInput: { minWidth: 0, flex: 1, paddingVertical: 4, color: "#13213A", fontSize: 11, fontWeight: "600" },
  searchButton: { paddingHorizontal: 7, paddingVertical: 5, borderRadius: 7, backgroundColor: "#E9EEFF" },
  searchButtonText: { color: "#2457E5", fontSize: 10, fontWeight: "800" },
  searchDivider: { width: 1, height: 24, marginHorizontal: 6, backgroundColor: "#E3EAF5" },
  formulaCard: { marginBottom: 10, padding: 11, borderRadius: 17, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#E3EAF5", elevation: 1 },
  formulaMeta: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  addressBadge: { minWidth: 47, paddingHorizontal: 9, paddingVertical: 4, alignItems: "center", borderRadius: 7, backgroundColor: "#E9EEFF" },
  addressText: { color: "#2457E5", fontSize: 12, fontWeight: "800" },
  formulaLabel: { color: "#64748B", fontSize: 12, fontWeight: "600" },
  editorRow: { flexDirection: "row-reverse", alignItems: "center", gap: 9 },
  formulaInput: { flex: 1, minHeight: 42, paddingHorizontal: 11, color: "#13213A", fontSize: 14, fontWeight: "600", backgroundColor: "#F8FAFE", borderRadius: 11, borderWidth: 1, borderColor: "#E3EAF5" },
  saveButton: { width: 42, height: 42, alignItems: "center", justifyContent: "center", borderRadius: 12, backgroundColor: "#2457E5" },
  formulaHintRow: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", gap: 8, marginTop: 8 },
  formulaHintText: { flex: 1, color: "#64748B", fontSize: 11, fontWeight: "600", textAlign: "right" },
  formulaHelpButton: { flexDirection: "row-reverse", alignItems: "center", gap: 3, paddingHorizontal: 7, paddingVertical: 5, borderRadius: 8, backgroundColor: "#EAF0FF" },
  formulaHelpText: { color: "#2457E5", fontSize: 10, fontWeight: "800" },
  formulaWorkflow: { marginTop: 7, color: "#64748B", fontSize: 10, fontWeight: "600", lineHeight: 16, textAlign: "right" },
  formulaToolsScroller: { marginHorizontal: -2 },
  tokenToolsRow: { flexDirection: "row-reverse", gap: 6, paddingHorizontal: 2, paddingTop: 8, paddingBottom: 2 },
  tokenTool: { minWidth: 39, alignItems: "center", justifyContent: "center", paddingHorizontal: 9, paddingVertical: 7, borderRadius: 9, backgroundColor: "#F7F9FE", borderWidth: 1, borderColor: "#DDE6F5" },
  tokenText: { color: "#2457E5", fontSize: 13, fontWeight: "900" },
  symbolToken: { color: "#2457E5", fontSize: 18, fontWeight: "900", lineHeight: 18 },
  exampleTokenTool: { flexDirection: "row-reverse", alignItems: "center", gap: 4, paddingHorizontal: 10, paddingVertical: 7, borderRadius: 9, backgroundColor: "#2457E5" },
  exampleTokenText: { color: "#FFFFFF", fontSize: 12, fontWeight: "900" },
  suggestionsHeader: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginTop: 9 },
  suggestionsTitle: { flexDirection: "row-reverse", alignItems: "center", gap: 4 },
  suggestionsLabel: { color: "#2457E5", fontSize: 10, fontWeight: "800" },
  swipeHint: { flexDirection: "row-reverse", alignItems: "center", gap: 3, paddingHorizontal: 5, paddingVertical: 2, borderRadius: 6, backgroundColor: "#F5F7FC" },
  swipeHintText: { color: "#64748B", fontSize: 9, fontWeight: "700" },
  formulaToolsRow: { flexDirection: "row-reverse", gap: 7, paddingHorizontal: 2, paddingTop: 6, paddingBottom: 2 },
  formulaTool: { flexDirection: "row-reverse", alignItems: "center", gap: 4, paddingHorizontal: 9, paddingVertical: 7, borderRadius: 9, backgroundColor: "#EAF0FF", borderWidth: 1, borderColor: "#D8E2FF" },
  formulaToolText: { color: "#2457E5", fontSize: 11, fontWeight: "800" },
  sigmaText: { color: "#2457E5", fontSize: 16, fontWeight: "900", lineHeight: 17 },
  operatorText: { color: "#2457E5", fontSize: 16, fontWeight: "900", lineHeight: 17 },
  toolsRow: { flexDirection: "row-reverse", alignItems: "center", gap: 6, marginBottom: 10 },
  analysisToolsRow: { flexDirection: "row-reverse", gap: 6, paddingBottom: 10 },
  analysisTool: { flexDirection: "row-reverse", alignItems: "center", gap: 4, paddingHorizontal: 9, paddingVertical: 7, borderRadius: 9, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  analysisToolText: { color: "#2457E5", fontSize: 10, fontWeight: "800" },
  toolButton: { flexDirection: "row-reverse", alignItems: "center", gap: 4, paddingHorizontal: 8, paddingVertical: 8, borderRadius: 10, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#E3EAF5" },
  activeToolButton: { backgroundColor: "#E9EEFF", borderColor: "#BFD0FF" },
  toolText: { color: "#2457E5", fontSize: 11, fontWeight: "700" },
  clearText: { color: "#C24141" },
  formatCard: { marginBottom: 10, padding: 12, borderRadius: 15, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  chartCard: { marginBottom: 10, padding: 12, borderRadius: 15, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  chartHeader: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  chartTitle: { color: "#13213A", fontSize: 13, fontWeight: "800", textAlign: "right" },
  chartSubtitle: { color: "#64748B", fontSize: 10, fontWeight: "600", textAlign: "right", marginTop: 2 },
  chartTypeSwitch: { flexDirection: "row-reverse", gap: 4 },
  chartTypeButton: { width: 32, height: 29, alignItems: "center", justifyContent: "center", borderRadius: 8, backgroundColor: "#F3F6FB" },
  activeChartTypeButton: { backgroundColor: "#E9EEFF", borderWidth: 1, borderColor: "#BFD0FF" },
  emptyChartText: { paddingVertical: 18, color: "#64748B", fontSize: 11, fontWeight: "600", textAlign: "center" },
  formatHeader: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  formatTitle: { color: "#13213A", fontSize: 13, fontWeight: "800" },
  clearFormatButton: { paddingHorizontal: 8, paddingVertical: 5, borderRadius: 7, backgroundColor: "#FFF3F3" },
  clearFormatText: { color: "#C24141", fontSize: 10, fontWeight: "800" },
  formatLabel: { marginTop: 5, marginBottom: 6, color: "#64748B", fontSize: 10, fontWeight: "800", textAlign: "right" },
  formatOptions: { flexDirection: "row-reverse", flexWrap: "wrap", gap: 6 },
  formatChip: { paddingHorizontal: 9, paddingVertical: 6, borderRadius: 8, backgroundColor: "#F7F9FE", borderWidth: 1, borderColor: "#E3EAF5" },
  activeFormatChip: { backgroundColor: "#E9EEFF", borderColor: "#2457E5" },
  formatChipText: { color: "#64748B", fontSize: 10, fontWeight: "700" },
  activeFormatChipText: { color: "#2457E5", fontWeight: "900" },
  colorOptions: { flexDirection: "row-reverse", gap: 9 },
  colorSwatch: { width: 28, height: 28, alignItems: "center", justifyContent: "center", borderRadius: 9, borderWidth: 1, borderColor: "#C8D4E5" },
  activeColorSwatch: { borderWidth: 2, borderColor: "#2457E5" },
  sheetCard: { flex: 1, minHeight: 200, overflow: "hidden", backgroundColor: "#FFFFFF", borderTopLeftRadius: 18, borderTopRightRadius: 18, borderWidth: 1, borderColor: "#DDE6F5" },
  sheetTopline: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", paddingHorizontal: 12, paddingVertical: 9, backgroundColor: "#FAFBFF", borderBottomWidth: 1, borderColor: "#E8EDF6" },
  sheetTitleGroup: { flexDirection: "row-reverse", alignItems: "center", gap: 6, flex: 1 },
  activeSheetDot: { width: 7, height: 7, borderRadius: 4, backgroundColor: "#2457E5" },
  sheetTitle: { color: "#13213A", fontSize: 12, fontWeight: "800" },
  frozenPill: { flexDirection: "row-reverse", alignItems: "center", gap: 3, paddingHorizontal: 6, paddingVertical: 3, borderRadius: 6, backgroundColor: "#F0F4FF" },
  frozenPillText: { color: "#64748B", fontSize: 9, fontWeight: "700" },
  sheetHint: { maxWidth: 150, color: "#7A879B", fontSize: 9, fontWeight: "500", textAlign: "left" },
  gridScrollContent: { paddingBottom: 2 },
  headerRow: { flexDirection: "row" },
  cornerCell: { width: 37, height: 33, backgroundColor: "#F1F5FB", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#DDE6F5" },
  columnHeader: { width: 78, height: 33, alignItems: "center", justifyContent: "center", backgroundColor: "#F1F5FB", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#DDE6F5" },
  columnHeaderText: { color: "#53627A", fontSize: 12, fontWeight: "800" },
  gridRow: { flexDirection: "row" },
  rowHeader: { width: 37, height: 42, alignItems: "center", justifyContent: "center", backgroundColor: "#F8FAFD", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#E2E8F2" },
  rowHeaderText: { color: "#708098", fontSize: 11, fontWeight: "700" },
  cell: { width: DEFAULT_COLUMN_WIDTH, height: 42, paddingHorizontal: 7, alignItems: "flex-end", justifyContent: "center", backgroundColor: "#FFFFFF", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#E2E8F2" },
  rangeCell: { backgroundColor: "#EAF0FF", borderColor: "#8EACFF" },
  selectedCell: { backgroundColor: "#EAF0FF", borderWidth: 2, borderColor: "#2457E5", marginLeft: -1, marginTop: -1, elevation: 2 },
  cellPressed: { opacity: 0.72 },
  cellText: { width: "100%", color: "#24344F", fontSize: 12, fontWeight: "600", textAlign: "right" },
  selectedCellText: { color: "#1D4ED8" },
  sheetTabsBar: { minHeight: 46, flexDirection: "row-reverse", alignItems: "center", gap: 7, paddingHorizontal: 8, borderTopWidth: 1, borderColor: "#E8EDF6", backgroundColor: "#FAFBFF" },
  tabsList: { flexDirection: "row-reverse", alignItems: "center", gap: 6, paddingVertical: 6 },
  sheetTab: { paddingHorizontal: 11, paddingVertical: 7, borderRadius: 8 },
  activeSheetTab: { backgroundColor: "#E9EEFF" },
  sheetTabText: { color: "#71809A", fontSize: 11, fontWeight: "700" },
  activeSheetTabText: { color: "#2457E5" },
  addSheetButton: { width: 31, height: 31, alignItems: "center", justifyContent: "center", borderRadius: 9, backgroundColor: "#E9EEFF" },
  statusBar: { flexDirection: "row-reverse", alignItems: "center", minHeight: 45, paddingHorizontal: 4 },
  statusItem: { flexDirection: "row-reverse", alignItems: "center", gap: 4, flex: 1 },
  statusText: { color: "#64748B", fontSize: 10, fontWeight: "700" },
  statusValue: { color: "#13213A", fontSize: 11, fontWeight: "900" },
  statusDivider: { width: 1, height: 20, marginHorizontal: 7, backgroundColor: "#DDE6F5" },
  pressed: { transform: [{ scale: 0.97 }], opacity: 0.9 },
  toolPressed: { opacity: 0.68 },
});
