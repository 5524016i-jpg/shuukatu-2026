import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class ShiftScheduler {

    static class Staff {
        String name;
        int startRow;
        Map<Integer, String> requests = new HashMap<>(); // 日付(1~16) -> 希望シフト
        Map<Integer, String> assigned = new HashMap<>(); // 日付(1~16) -> 確定シフト

        Staff(String name, int startRow) {
            this.name = name;
            this.startRow = startRow;
        }

        int getRequestCount() {
            return requests.size();
        }

        int getAssignedCount() {
            return assigned.size();
        }

        double getAssignmentRatio() {
            if (getRequestCount() == 0) return 1.0;
            return (double) getAssignedCount() / getRequestCount();
        }
    }

    public static void main(String[] args) {
        String inputFilePath = "シフト表_千葉_fixed.xlsx";
        String outputFilePath = "完成シフト表_千葉_v2.xlsx";

        // 目標人数の設定 (ピーク日: 6名 / 通常日: 4名)
        int PEAK_TARGET_CAPACITY = 6;
        int REGULAR_TARGET_CAPACITY = 4;

        // 祝前日(11/2)および週末(金・土)
        Set<Integer> eveOfHolidays = new HashSet<>(Arrays.asList(2));

        // 除外する文字（出勤希望ではない記号）
        Set<String> offKeywords = new HashSet<>(Arrays.asList(
            "", "-", "－", "ー", "×", "公休", "希望", "休", "有給", "欠"
        ));

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("千葉");
            if (sheet == null) {
                System.out.println("エラー: シート『千葉』が見つかりません。");
                return;
            }

            List<Staff> staffList = new ArrayList<>();
            int[] targetRows = {47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 61, 62, 63, 64, 65};

            // 1. 読み込み
            for (int r : targetRows) {
                Row row = sheet.getRow(r - 1);
                if (row == null) continue;

                Cell nameCell = row.getCell(1); // B列（インデックス1）
                String staffName = getCellValueAsString(nameCell).trim();
                if (staffName.isEmpty()) continue;

                Staff staff = new Staff(staffName, r - 1);

                for (int day = 1; day <= 16; day++) {
                    Cell cell = row.getCell(day + 1); // C列（インデックス2）が1日
                    String val = getCellValueAsString(cell).trim();
                    if (!offKeywords.contains(val)) {
                        staff.requests.put(day, val);
                    }
                }
                staffList.add(staff);
            }

            // 2. シフトカット・割当処理
            for (int day = 1; day <= 16; day++) {
                int dayOfWeek = day % 7;
                boolean isPeakDay = (dayOfWeek == 6) || (dayOfWeek == 0) || eveOfHolidays.contains(day);
                int targetCapacity = isPeakDay ? PEAK_TARGET_CAPACITY : REGULAR_TARGET_CAPACITY;

                List<Staff> candidates = new ArrayList<>();
                for (Staff staff : staffList) {
                    if (staff.requests.containsKey(day)) {
                        candidates.add(staff);
                    }
                }

                // 確定率が低いスタッフを優先採用（同率の場合は元の並び順を維持）
                candidates.sort(Comparator.comparingDouble(Staff::getAssignmentRatio));

                int countToAssign = Math.min(candidates.size(), targetCapacity);
                for (int i = 0; i < countToAssign; i++) {
                    Staff selected = candidates.get(i);
                    selected.assigned.put(day, selected.requests.get(day));
                }
            }

            // カットセル用のスタイル（薄いグレー背景・グレー太字・中央揃え）
            CellStyle cutStyle = workbook.createCellStyle();
            cutStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cutStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cutStyle.setAlignment(HorizontalAlignment.CENTER);
            cutStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cutStyle.setBorderTop(BorderStyle.THIN);
            cutStyle.setBorderBottom(BorderStyle.THIN);
            cutStyle.setBorderLeft(BorderStyle.THIN);
            cutStyle.setBorderRight(BorderStyle.THIN);

            Font cutFont = workbook.createFont();
            cutFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
            cutFont.setBold(true);
            cutStyle.setFont(cutFont);

            // 3. 「千葉」シートへ結果の書き戻し
            for (Staff staff : staffList) {
                Row row = sheet.getRow(staff.startRow);
                for (int day = 1; day <= 16; day++) {
                    Cell cell = row.getCell(day + 1);
                    if (cell == null) cell = row.createCell(day + 1);

                    if (staff.assigned.containsKey(day)) {
                        cell.setCellValue(staff.assigned.get(day));
                    } else if (staff.requests.containsKey(day)) {
                        cell.setCellValue("-");
                        cell.setCellStyle(cutStyle);
                    }
                }
            }

            // 4. 「千葉_集計サマリー」シートの作成
            createSummarySheet(workbook, staffList);

            // 5. 保存
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            System.out.println("出力完了: " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 集計サマリーシートを作成するメソッド
    private static void createSummarySheet(Workbook workbook, List<Staff> staffList) {
        String sheetName = "千葉_集計サマリー";
        int existingIdx = workbook.getSheetIndex(sheetName);
        if (existingIdx != -1) {
            workbook.removeSheetAt(existingIdx);
        }

        Sheet summarySheet = workbook.createSheet(sheetName);

        // スタイル定義
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 13);
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFont(headerFont);
        setThinBorders(headerStyle);

        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setAlignment(HorizontalAlignment.LEFT);
        textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setThinBorders(textStyle);

        CellStyle numStyle = workbook.createCellStyle();
        numStyle.setAlignment(HorizontalAlignment.RIGHT);
        numStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setThinBorders(numStyle);

        CellStyle pctStyle = workbook.createCellStyle();
        pctStyle.setAlignment(HorizontalAlignment.RIGHT);
        pctStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        pctStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        setThinBorders(pctStyle);

        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        CellStyle totalStyle = workbook.createCellStyle();
        totalStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        totalStyle.setAlignment(HorizontalAlignment.RIGHT);
        totalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        totalStyle.setFont(boldFont);
        setThinBorders(totalStyle);

        CellStyle totalPctStyle = workbook.createCellStyle();
        totalPctStyle.cloneStyleFrom(totalStyle);
        totalPctStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));

        // タイトル
        Row titleRow = summarySheet.createRow(1);
        Cell titleCell = titleRow.createCell(1);
        titleCell.setCellValue("千葉店 シフト確定・カット集計（11月1日〜16日）");
        titleCell.setCellStyle(titleStyle);

        // ヘッダー
        String[] headers = {"スタッフ名", "希望日数", "確定日数", "カット日数", "採用率", "カットされた日"};
        Row headerRow = summarySheet.createRow(3);
        headerRow.setHeightInPoints(24);
        for (int i = 0; i < headers.length; i++) {
            Cell c = headerRow.createCell(i + 1);
            c.setCellValue(headers[i]);
            c.setCellStyle(headerStyle);
        }

        // データ行
        for (int i = 0; i < staffList.size(); i++) {
            Staff s = staffList.get(i);
            int rowIdx = 4 + i;
            int excelRowNum = rowIdx + 1;
            Row r = summarySheet.createRow(rowIdx);
            r.setHeightInPoints(20);

            List<String> cutDays = new ArrayList<>();
            for (int day = 1; day <= 16; day++) {
                if (s.requests.containsKey(day) && !s.assigned.containsKey(day)) {
                    cutDays.add(day + "日");
                }
            }
            String cutStr = cutDays.isEmpty() ? "なし" : String.join(", ", cutDays);

            Cell cName = r.createCell(1);
            cName.setCellValue(s.name);
            cName.setCellStyle(textStyle);

            Cell cReq = r.createCell(2);
            cReq.setCellValue(s.getRequestCount());
            cReq.setCellStyle(numStyle);

            Cell cAss = r.createCell(3);
            cAss.setCellValue(s.getAssignedCount());
            cAss.setCellStyle(numStyle);

            Cell cCut = r.createCell(4);
            cCut.setCellFormula("C" + excelRowNum + "-D" + excelRowNum);
            cCut.setCellStyle(numStyle);

            Cell cRat = r.createCell(5);
            cRat.setCellFormula("IF(C" + excelRowNum + "=0, 0, D" + excelRowNum + "/C" + excelRowNum + ")");
            cRat.setCellStyle(pctStyle);

            Cell cDays = r.createCell(6);
            cDays.setCellValue(cutStr);
            cDays.setCellStyle(textStyle);
        }

        // 合計行
        int totalRowIdx = 4 + staffList.size();
        int lastDataRowNum = totalRowIdx;
        Row totalRow = summarySheet.createRow(totalRowIdx);
        totalRow.setHeightInPoints(22);

        Cell tLabel = totalRow.createCell(1);
        tLabel.setCellValue("合計 / 平均");
        tLabel.setCellStyle(totalStyle);

        Cell tReq = totalRow.createCell(2);
        tReq.setCellFormula("SUM(C5:C" + lastDataRowNum + ")");
        tReq.setCellStyle(totalStyle);

        Cell tAss = totalRow.createCell(3);
        tAss.setCellFormula("SUM(D5:D" + lastDataRowNum + ")");
        tAss.setCellStyle(totalStyle);

        Cell tCut = totalRow.createCell(4);
        tCut.setCellFormula("SUM(E5:E" + lastDataRowNum + ")");
        tCut.setCellStyle(totalStyle);

        Cell tRat = totalRow.createCell(5);
        tRat.setCellFormula("AVERAGEIF(C5:C" + lastDataRowNum + ", \">0\", F5:F" + lastDataRowNum + ")");
        tRat.setCellStyle(totalPctStyle);

        Cell tEmpty = totalRow.createCell(6);
        tEmpty.setCellValue("");
        tEmpty.setCellStyle(totalStyle);

        // 列幅の調整
        summarySheet.setColumnWidth(1, 16 * 256);
        summarySheet.setColumnWidth(2, 12 * 256);
        summarySheet.setColumnWidth(3, 12 * 256);
        summarySheet.setColumnWidth(4, 12 * 256);
        summarySheet.setColumnWidth(5, 12 * 256);
        summarySheet.setColumnWidth(6, 28 * 256);
    }

    private static void setThinBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultType();
        }
        switch (type) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double num = cell.getNumericCellValue();
                if (num == (long) num) {
                    return String.valueOf((long) num);
                } else {
                    return String.valueOf(num);
                }
            default:
                return "";
        }
    }
}