import AsyncStorage from "@react-native-async-storage/async-storage";
import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as DocumentPicker from "expo-document-picker";
import * as FileSystem from "expo-file-system/legacy";
import * as Sharing from "expo-sharing";
import { useEffect, useMemo, useState } from "react";
import {
  FlatList,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  View,
} from "react-native";

import { ScreenContainer } from "@/components/screen-container";
import { displayCellValue, type CellValues } from "@/lib/formula-engine";
import { exportSheetToCsv, exportWorkbookToXlsx, importSpreadsheetFile } from "@/lib/workbook-file-utils";

type Sheet = {
  id: string;
  name: string;
  cells: CellValues;
  rowCount: number;
  columnCount: number;
};

type Workbook = { sheets: Sheet[]; activeSheetId: string };

const WORKBOOK_KEY = "jadwali.workbook.v1";
const INITIAL_ROWS = 14;
const INITIAL_COLUMNS = 6;
const FIRST_SHEET: Sheet = { id: "sheet-1", name: "ورقة 1", cells: {}, rowCount: INITIAL_ROWS, columnCount: INITIAL_COLUMNS };
const INITIAL_WORKBOOK: Workbook = { sheets: [FIRST_SHEET], activeSheetId: FIRST_SHEET.id };

function safeFileName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "-").slice(0, 40) || "jadwali";
}

function arrayBufferToBase64(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
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

export default function HomeScreen() {
  const [workbook, setWorkbook] = useState<Workbook>(INITIAL_WORKBOOK);
  const [selectedCell, setSelectedCell] = useState("A1");
  const [draft, setDraft] = useState("");
  const [savedMessage, setSavedMessage] = useState("اختر خلية لبدء الإدخال");
  const [hasLoaded, setHasLoaded] = useState(false);
  const [lastWorkbook, setLastWorkbook] = useState<Workbook | null>(null);
  const [isFileAction, setIsFileAction] = useState(false);

  const activeSheet = useMemo(
    () => workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0],
    [workbook],
  );
  const columns = useMemo(
    () => Array.from({ length: activeSheet.columnCount }, (_, index) => columnLabel(index)),
    [activeSheet.columnCount],
  );
  const rows = useMemo(
    () => Array.from({ length: activeSheet.rowCount }, (_, index) => index + 1),
    [activeSheet.rowCount],
  );

  useEffect(() => {
    let isActive = true;
    AsyncStorage.getItem(WORKBOOK_KEY)
      .then((stored) => {
        if (!stored || !isActive) return;
        const parsed = JSON.parse(stored) as Workbook;
        if (parsed.sheets?.length && parsed.activeSheetId) {
          setWorkbook(parsed);
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
    AsyncStorage.setItem(WORKBOOK_KEY, JSON.stringify(workbook)).catch(() => {
      setSavedMessage("تعذر حفظ التغييرات محلياً");
    });
  }, [hasLoaded, workbook]);

  function updateActiveSheet(updater: (sheet: Sheet) => Sheet) {
    setWorkbook((current) => ({
      ...current,
      sheets: current.sheets.map((sheet) => (sheet.id === current.activeSheetId ? updater(sheet) : sheet)),
    }));
    setLastWorkbook(workbook);
  }

  function selectCell(address: string) {
    setSelectedCell(address);
    setDraft(activeSheet.cells[address] ?? "");
    setSavedMessage(`الخلية ${address} جاهزة للتحرير`);
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

  function addSheet() {
    const nextNumber = workbook.sheets.length + 1;
    const id = `sheet-${nextNumber}`;
    const nextSheet: Sheet = { id, name: `ورقة ${nextNumber}`, cells: {}, rowCount: INITIAL_ROWS, columnCount: INITIAL_COLUMNS };
    setLastWorkbook(workbook);
    setWorkbook((current) => ({ ...current, sheets: [...current.sheets, nextSheet], activeSheetId: id }));
    setSelectedCell("A1");
    setDraft("");
    setSavedMessage(`تم إنشاء ${nextSheet.name}`);
  }

  function chooseSheet(sheet: Sheet) {
    setWorkbook((current) => ({ ...current, activeSheetId: sheet.id }));
    setSelectedCell("A1");
    setDraft(sheet.cells.A1 ?? "");
    setSavedMessage(`تم فتح ${sheet.name}`);
  }

  function insertFormula(template: string, label: string) {
    setDraft(template);
    setSavedMessage(`أُدرجت صيغة ${label}؛ عدّل مراجع الخلايا ثم اضغط التأكيد`);
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
    const restoredSheet = lastWorkbook.sheets.find((sheet) => sheet.id === lastWorkbook.activeSheetId) ?? lastWorkbook.sheets[0];
    setWorkbook(lastWorkbook);
    setLastWorkbook(null);
    setSelectedCell("A1");
    setDraft(restoredSheet.cells.A1 ?? "");
    setSavedMessage("تم التراجع عن آخر تعديل");
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

      const nextWorkbook: Workbook = { sheets: imported.sheets, activeSheetId: imported.sheets[0].id };
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
      await Sharing.shareAsync(fileUri, {
        dialogTitle: format === "csv" ? "تصدير ورقة CSV" : "تصدير مصنف Excel",
        mimeType: format === "csv" ? "text/csv" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      setSavedMessage(format === "csv" ? "تم تجهيز CSV للتحميل أو المشاركة" : "تم تجهيز ملف Excel للتحميل أو المشاركة");
    } catch {
      setSavedMessage("تعذر تصدير الملف؛ حاول مرة أخرى");
    } finally {
      setIsFileAction(false);
    }
  }

  return (
    <ScreenContainer className="flex-1" containerClassName="bg-background">
      <View style={styles.page}>
        <View style={styles.header}>
          <View style={styles.brandGroup}>
            <View style={styles.logoTile}>
              <MaterialIcons name="grid-on" size={21} color="#FFFFFF" />
            </View>
            <View>
              <Text style={styles.brand}>جدولي</Text>
              <Text style={styles.documentName}>مصنف جديد</Text>
            </View>
          </View>
          <View style={styles.headerBadge}>
            <View style={styles.headerBadgeDot} />
            <Text style={styles.headerBadgeText}>{hasLoaded ? "محفوظ محلياً" : "جارٍ التحميل"}</Text>
          </View>
        </View>

        <View style={styles.fileActionsRow}>
          <Pressable accessibilityLabel="استيراد ملف CSV أو Excel" disabled={isFileAction} onPress={importFile} style={({ pressed }) => [styles.fileActionButton, (pressed || isFileAction) && styles.fileActionPressed]}>
            <MaterialIcons name="file-upload" size={18} color="#2457E5" /><Text style={styles.fileActionText}>استيراد</Text>
          </Pressable>
          <Pressable accessibilityLabel="تصدير الورقة الحالية إلى CSV" disabled={isFileAction} onPress={() => exportFile("csv")} style={({ pressed }) => [styles.fileActionButton, (pressed || isFileAction) && styles.fileActionPressed]}>
            <MaterialIcons name="download" size={18} color="#16865B" /><Text style={[styles.fileActionText, styles.csvActionText]}>CSV</Text>
          </Pressable>
          <Pressable accessibilityLabel="تصدير المصنف إلى Excel" disabled={isFileAction} onPress={() => exportFile("xlsx")} style={({ pressed }) => [styles.fileActionButton, (pressed || isFileAction) && styles.fileActionPressed]}>
            <MaterialIcons name="table-view" size={18} color="#2457E5" /><Text style={styles.fileActionText}>Excel</Text>
          </Pressable>
          <Text style={styles.fileHint}>{isFileAction ? "جارٍ تجهيز الملف…" : "CSV أو XLSX"}</Text>
        </View>

        <View style={styles.formulaCard}>
          <View style={styles.formulaMeta}>
            <View style={styles.addressBadge}>
              <Text style={styles.addressText}>{selectedCell}</Text>
            </View>
            <Text style={styles.formulaLabel}>شريط الصيغ</Text>
          </View>
          <View style={styles.editorRow}>
            <TextInput
              accessibilityLabel="محرر الخلية"
              value={draft}
              onChangeText={setDraft}
              onSubmitEditing={saveCell}
              placeholder="قيمة أو صيغة تبدأ بـ ="
              placeholderTextColor="#8B99AE"
              style={styles.formulaInput}
              textAlign="right"
              returnKeyType="done"
              autoCapitalize="characters"
            />
            <Pressable accessibilityLabel="تأكيد تعديل الخلية" onPress={saveCell} style={({ pressed }) => [styles.saveButton, pressed && styles.pressed]}>
              <MaterialIcons name="check" size={20} color="#FFFFFF" />
            </Pressable>
          </View>
        </View>

        <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.formulaToolsRow}>
          <Pressable onPress={() => insertFormula("=SUM(A1:A3)", "SUM")} style={({ pressed }) => [styles.formulaTool, pressed && styles.toolPressed]}>
            <Text style={styles.sigmaText}>Σ</Text><Text style={styles.formulaToolText}>جمع</Text>
          </Pressable>
          <Pressable onPress={() => insertFormula("=A1-B1", "الطرح")} style={({ pressed }) => [styles.formulaTool, pressed && styles.toolPressed]}>
            <Text style={styles.operatorText}>−</Text><Text style={styles.formulaToolText}>طرح</Text>
          </Pressable>
          <Pressable onPress={() => insertFormula("=A1*B1", "الضرب")} style={({ pressed }) => [styles.formulaTool, pressed && styles.toolPressed]}>
            <Text style={styles.operatorText}>×</Text><Text style={styles.formulaToolText}>ضرب</Text>
          </Pressable>
          <Pressable onPress={() => insertFormula("=A1/B1", "القسمة")} style={({ pressed }) => [styles.formulaTool, pressed && styles.toolPressed]}>
            <Text style={styles.operatorText}>÷</Text><Text style={styles.formulaToolText}>قسمة</Text>
          </Pressable>
          <Pressable onPress={() => insertFormula("=AVERAGE(A1:A3)", "AVERAGE")} style={({ pressed }) => [styles.formulaTool, pressed && styles.toolPressed]}>
            <MaterialIcons name="functions" size={17} color="#2457E5" /><Text style={styles.formulaToolText}>متوسط</Text>
          </Pressable>
        </ScrollView>

        <View style={styles.toolsRow}>
          <Pressable onPress={() => changeTableSize("columnCount")} style={({ pressed }) => [styles.toolButton, pressed && styles.toolPressed]}>
            <MaterialIcons name="view-column" size={18} color="#2457E5" /><Text style={styles.toolText}>عمود</Text>
          </Pressable>
          <Pressable onPress={() => changeTableSize("rowCount")} style={({ pressed }) => [styles.toolButton, pressed && styles.toolPressed]}>
            <MaterialIcons name="view-agenda" size={18} color="#2457E5" /><Text style={styles.toolText}>صف</Text>
          </Pressable>
          <Pressable onPress={clearSelectedCell} style={({ pressed }) => [styles.toolButton, pressed && styles.toolPressed]}>
            <MaterialIcons name="backspace" size={18} color="#C24141" /><Text style={[styles.toolText, styles.clearText]}>مسح</Text>
          </Pressable>
          <Pressable onPress={undoLastChange} style={({ pressed }) => [styles.toolButton, pressed && styles.toolPressed]}>
            <MaterialIcons name="undo" size={18} color="#2457E5" /><Text style={styles.toolText}>تراجع</Text>
          </Pressable>
          <View style={styles.tipPill}>
            <MaterialIcons name="touch-app" size={16} color="#64748B" /><Text style={styles.tipText}>المس أي خلية</Text>
          </View>
        </View>

        <View style={styles.sheetCard}>
          <View style={styles.sheetTopline}>
            <View style={styles.sheetTitleGroup}>
              <View style={styles.activeSheetDot} /><Text style={styles.sheetTitle}>{activeSheet.name}</Text>
            </View>
            <Text style={styles.sheetHint}>{savedMessage}</Text>
          </View>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={styles.gridScrollContent}>
            <View>
              <View style={styles.headerRow}>
                <View style={styles.cornerCell} />
                {columns.map((column) => <View key={column} style={styles.columnHeader}><Text style={styles.columnHeaderText}>{column}</Text></View>)}
              </View>
              <ScrollView showsVerticalScrollIndicator={false} nestedScrollEnabled>
                {rows.map((row) => (
                  <View key={row} style={styles.gridRow}>
                    <View style={styles.rowHeader}><Text style={styles.rowHeaderText}>{row}</Text></View>
                    {columns.map((column) => {
                      const address = `${column}${row}`;
                      const isSelected = address === selectedCell;
                      return (
                        <Pressable key={address} accessibilityLabel={`الخلية ${address}`} onPress={() => selectCell(address)} style={({ pressed }) => [styles.cell, isSelected && styles.selectedCell, pressed && styles.cellPressed]}>
                          <Text numberOfLines={1} style={[styles.cellText, isSelected && styles.selectedCellText]}>{displayCellValue(address, activeSheet.cells)}</Text>
                        </Pressable>
                      );
                    })}
                  </View>
                ))}
              </ScrollView>
            </View>
          </ScrollView>
          <View style={styles.sheetTabsBar}>
            <FlatList
              horizontal
              data={workbook.sheets}
              keyExtractor={(sheet) => sheet.id}
              showsHorizontalScrollIndicator={false}
              contentContainerStyle={styles.tabsList}
              renderItem={({ item }) => (
                <Pressable onPress={() => chooseSheet(item)} style={({ pressed }) => [styles.sheetTab, item.id === activeSheet.id && styles.activeSheetTab, pressed && styles.toolPressed]}>
                  <Text style={[styles.sheetTabText, item.id === activeSheet.id && styles.activeSheetTabText]}>{item.name}</Text>
                </Pressable>
              )}
            />
            <Pressable accessibilityLabel="إضافة ورقة عمل" onPress={addSheet} style={({ pressed }) => [styles.addSheetButton, pressed && styles.pressed]}>
              <MaterialIcons name="add" size={20} color="#2457E5" />
            </Pressable>
          </View>
        </View>

        <View style={styles.bottomBar}>
          <View style={styles.bottomInfo}>
            <MaterialIcons name="info-outline" size={17} color="#64748B" />
            <Text style={styles.bottomInfoText}>النتائج تُحدّث تلقائياً عند تعديل المراجع</Text>
          </View>
          <View style={styles.cellCounter}><Text style={styles.cellCounterText}>{Object.keys(activeSheet.cells).length} خلية</Text></View>
        </View>
      </View>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  page: { flex: 1, backgroundColor: "#F7F9FE", paddingHorizontal: 16, paddingTop: 12 },
  header: { flexDirection: "row-reverse", justifyContent: "space-between", alignItems: "center", marginBottom: 14 },
  brandGroup: { flexDirection: "row-reverse", alignItems: "center", gap: 10 },
  logoTile: { width: 39, height: 39, borderRadius: 13, alignItems: "center", justifyContent: "center", backgroundColor: "#2457E5", shadowColor: "#2457E5", shadowOpacity: 0.18, shadowRadius: 10, shadowOffset: { width: 0, height: 5 }, elevation: 3 },
  brand: { color: "#13213A", fontSize: 21, fontWeight: "800", lineHeight: 27, textAlign: "right" },
  documentName: { color: "#64748B", fontSize: 12, fontWeight: "500", lineHeight: 17, textAlign: "right" },
  headerBadge: { flexDirection: "row-reverse", alignItems: "center", gap: 5, paddingHorizontal: 9, paddingVertical: 6, borderRadius: 99, backgroundColor: "#E8F7F0" },
  headerBadgeDot: { width: 6, height: 6, borderRadius: 3, backgroundColor: "#16865B" },
  headerBadgeText: { color: "#16865B", fontSize: 11, fontWeight: "700" },
  fileActionsRow: { flexDirection: "row-reverse", alignItems: "center", gap: 7, marginBottom: 12 },
  fileActionButton: { flexDirection: "row-reverse", alignItems: "center", gap: 5, minHeight: 36, paddingHorizontal: 10, borderRadius: 10, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#DDE6F5" },
  fileActionPressed: { opacity: 0.55, transform: [{ scale: 0.98 }] },
  fileActionText: { color: "#2457E5", fontSize: 12, fontWeight: "800" },
  csvActionText: { color: "#16865B" },
  fileHint: { marginRight: "auto", color: "#7A879B", fontSize: 10, fontWeight: "600" },
  formulaCard: { backgroundColor: "#FFFFFF", borderRadius: 18, padding: 12, borderWidth: 1, borderColor: "#E3EAF5", shadowColor: "#183B74", shadowOpacity: 0.06, shadowRadius: 14, shadowOffset: { width: 0, height: 5 }, elevation: 2 },
  formulaMeta: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", marginBottom: 9 },
  addressBadge: { minWidth: 47, paddingHorizontal: 9, paddingVertical: 4, alignItems: "center", borderRadius: 7, backgroundColor: "#E9EEFF" },
  addressText: { color: "#2457E5", fontSize: 12, fontWeight: "800" },
  formulaLabel: { color: "#64748B", fontSize: 12, fontWeight: "600" },
  editorRow: { flexDirection: "row-reverse", alignItems: "center", gap: 9 },
  formulaInput: { flex: 1, minHeight: 43, paddingHorizontal: 12, color: "#13213A", fontSize: 14, fontWeight: "600", backgroundColor: "#F8FAFE", borderRadius: 11, borderWidth: 1, borderColor: "#E3EAF5" },
  saveButton: { width: 43, height: 43, alignItems: "center", justifyContent: "center", borderRadius: 12, backgroundColor: "#2457E5" },
  pressed: { transform: [{ scale: 0.97 }], opacity: 0.92 },
  formulaToolsRow: { flexDirection: "row-reverse", gap: 7, paddingVertical: 11 },
  formulaTool: { flexDirection: "row-reverse", alignItems: "center", gap: 4, paddingHorizontal: 10, paddingVertical: 8, borderRadius: 10, backgroundColor: "#EAF0FF", borderWidth: 1, borderColor: "#D8E2FF" },
  formulaToolText: { color: "#2457E5", fontSize: 12, fontWeight: "800" },
  sigmaText: { color: "#2457E5", fontSize: 18, fontWeight: "900", lineHeight: 18 },
  operatorText: { color: "#2457E5", fontSize: 17, fontWeight: "900", lineHeight: 18 },
  toolsRow: { flexDirection: "row-reverse", alignItems: "center", gap: 7, paddingBottom: 12 },
  toolButton: { flexDirection: "row-reverse", alignItems: "center", gap: 5, paddingHorizontal: 10, paddingVertical: 8, borderRadius: 10, backgroundColor: "#FFFFFF", borderWidth: 1, borderColor: "#E3EAF5" },
  toolPressed: { opacity: 0.72 },
  toolText: { color: "#2457E5", fontSize: 12, fontWeight: "700" },
  clearText: { color: "#C24141" },
  tipPill: { flexDirection: "row-reverse", alignItems: "center", gap: 4, marginRight: "auto" },
  tipText: { color: "#64748B", fontSize: 11, fontWeight: "600" },
  sheetCard: { flex: 1, minHeight: 310, overflow: "hidden", backgroundColor: "#FFFFFF", borderTopLeftRadius: 18, borderTopRightRadius: 18, borderWidth: 1, borderColor: "#DDE6F5" },
  sheetTopline: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", paddingHorizontal: 13, paddingVertical: 10, backgroundColor: "#FAFBFF", borderBottomWidth: 1, borderColor: "#E8EDF6" },
  sheetTitleGroup: { flexDirection: "row-reverse", alignItems: "center", gap: 7 },
  activeSheetDot: { width: 7, height: 7, borderRadius: 4, backgroundColor: "#2457E5" },
  sheetTitle: { color: "#13213A", fontSize: 13, fontWeight: "800" },
  sheetHint: { maxWidth: 180, color: "#7A879B", fontSize: 10, fontWeight: "500", textAlign: "left" },
  gridScrollContent: { paddingBottom: 2 },
  headerRow: { flexDirection: "row" },
  cornerCell: { width: 37, height: 34, backgroundColor: "#F1F5FB", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#DDE6F5" },
  columnHeader: { width: 78, height: 34, alignItems: "center", justifyContent: "center", backgroundColor: "#F1F5FB", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#DDE6F5" },
  columnHeaderText: { color: "#53627A", fontSize: 12, fontWeight: "800" },
  gridRow: { flexDirection: "row" },
  rowHeader: { width: 37, height: 43, alignItems: "center", justifyContent: "center", backgroundColor: "#F8FAFD", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#E2E8F2" },
  rowHeaderText: { color: "#708098", fontSize: 11, fontWeight: "700" },
  cell: { width: 78, height: 43, paddingHorizontal: 7, alignItems: "flex-end", justifyContent: "center", backgroundColor: "#FFFFFF", borderRightWidth: 1, borderBottomWidth: 1, borderColor: "#E2E8F2" },
  selectedCell: { backgroundColor: "#EEF2FF", borderWidth: 2, borderColor: "#2457E5", marginLeft: -1, marginTop: -1 },
  cellPressed: { backgroundColor: "#F4F7FF" },
  cellText: { width: "100%", color: "#24344F", fontSize: 12, fontWeight: "600", textAlign: "right" },
  selectedCellText: { color: "#1D4ED8" },
  sheetTabsBar: { minHeight: 47, flexDirection: "row-reverse", alignItems: "center", gap: 7, paddingHorizontal: 8, borderTopWidth: 1, borderColor: "#E8EDF6", backgroundColor: "#FAFBFF" },
  tabsList: { flexDirection: "row-reverse", alignItems: "center", gap: 6, paddingVertical: 6 },
  sheetTab: { paddingHorizontal: 11, paddingVertical: 7, borderRadius: 8 },
  activeSheetTab: { backgroundColor: "#E9EEFF" },
  sheetTabText: { color: "#71809A", fontSize: 11, fontWeight: "700" },
  activeSheetTabText: { color: "#2457E5" },
  addSheetButton: { width: 31, height: 31, alignItems: "center", justifyContent: "center", borderRadius: 9, backgroundColor: "#E9EEFF" },
  bottomBar: { flexDirection: "row-reverse", alignItems: "center", justifyContent: "space-between", minHeight: 52, paddingHorizontal: 3 },
  bottomInfo: { flexDirection: "row-reverse", alignItems: "center", gap: 5, flex: 1 },
  bottomInfoText: { flexShrink: 1, color: "#64748B", fontSize: 11, fontWeight: "500", textAlign: "right" },
  cellCounter: { paddingHorizontal: 9, paddingVertical: 5, backgroundColor: "#E9EEFF", borderRadius: 8 },
  cellCounterText: { color: "#2457E5", fontSize: 11, fontWeight: "800" },
});
