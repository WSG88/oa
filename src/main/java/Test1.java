import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class Test1 {


    public static void main(String[] args) throws Exception {

        Utils.YEAR_MONTH = "202008";
        Utils.FILE_NAME = "1.8.xls";
        Utils.ROOM = 1;
        Utils.arrayNamesList.clear();

        InputStream in = new FileInputStream(Utils.FILE_PATH + Utils.FILE_NAME);
        Workbook wbs = WorkbookFactory.create(in);
        for (int i = 0; i < wbs.getNumberOfSheets(); i++) {
            Sheet childSheet = getSheet(wbs, null, i);
            setData(childSheet);
        }
        Test2.getData(Utils.arrayNamesList);
    }

    public static void setData(Sheet childSheet) throws Exception {

        String name1 = ExcelUtil.readExcelData(childSheet, 3, 9);
        String name2 = ExcelUtil.readExcelData(childSheet, 3, 24);
        String name3 = ExcelUtil.readExcelData(childSheet, 3, 39);
        Utils.arrayNamesList.add(name1);
        Utils.arrayNamesList.add(name2);
        Utils.arrayNamesList.add(name3);

        List<String> list11 = new ArrayList<>();
        List<String> list22 = new ArrayList<>();
        List<String> list33 = new ArrayList<>();

        for (int row = 12; row < 43; row++) {
            String day = ExcelUtil.readExcelData(childSheet, row, 0);
            if (!"".equals(day)) {
                day = day.substring(0, 2);
            }
            for (int cell = 0; cell < 44; cell++) {
                if (cell == 2 - 1
                        || cell == 4 - 1
                        || cell == 7 - 1
                        || cell == 9 - 1
                        || cell == 11 - 1
                        || cell == 13 - 1
                        ) {
                    list11.add(readExcelDataGetNumericCellValue(childSheet, row, cell));
                    if (list11.size() == 6) {
                        insertData(list11, name1, day);
                        list11.clear();
                    }
                }
                if (cell == 17 - 1
                        || cell == 19 - 1
                        || cell == 22 - 1
                        || cell == 24 - 1
                        || cell == 26 - 1
                        || cell == 28 - 1
                        ) {
                    list22.add(readExcelDataGetNumericCellValue(childSheet, row, cell));

                    if (list22.size() == 6) {
                        insertData(list22, name2, day);
                        list22.clear();
                    }

                }
                if (cell == 32 - 1
                        || cell == 34 - 1
                        || cell == 37 - 1
                        || cell == 39 - 1
                        || cell == 41 - 1
                        || cell == 43 - 1
                        ) {
                    list33.add(readExcelDataGetNumericCellValue(childSheet, row, cell));

                    if (list33.size() == 6) {
                        insertData(list33, name3, day);
                        list33.clear();
                    }
                }
            }
        }

    }

    public static void insertData(List<String> l, String name, String day) {
        if (name == null
                || "".equals(name)
                || "正常".equals(name)
                || name.length() < 2) {
            return;
        }

        String date = Utils.YEAR_MONTH + day;
        List<String> list = new ArrayList<>();
        List<String> list1 = new ArrayList<>();
        for (String s : l) {
            list.add(calculateTime(s));
            if ("缺勤".equals(s)) {
                list1.add("缺勤");
            }
        }
        if (list1.size() == 6) {
            return;
        }
        Data data = new Data(name, date, list);
        Test2.saveToDatabase(data, Utils.ROOM);
    }

    public static String calculateTime(String string) {
        if (!"缺勤".equals(string)) {
            double d0 = Double.parseDouble(string);
            double dd = d0 * 24;
            int hour = (int) dd;
            double dou = (dd - hour) * 60;
            int minuter = (int) dou;
            if (dou > minuter) {
                minuter = minuter + 1;
            }
            StringBuilder stringBuilder = new StringBuilder();
            if (hour < 10) {
                stringBuilder.append("0");
            }
            stringBuilder.append(hour);
            stringBuilder.append(":");
            if (minuter < 10) {
                stringBuilder.append("0");
            }
            stringBuilder.append(minuter);
            return stringBuilder.toString();
        } else {
            return string;
        }
    }

    public static String readExcelDataGetNumericCellValue(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
//                    System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + String.valueOf(cell.getNumericCellValue()));
                    return String.valueOf(cell.getNumericCellValue());
                }
            }
        }
        return "缺勤";
    }

    public static Sheet getSheet(Workbook wbs, String sheetName, int sheetIndex) {
        Sheet childSheet;
        if (sheetName == null) {
            childSheet = wbs.getSheetAt(sheetIndex);
        } else {
            childSheet = wbs.getSheet(sheetName);
        }
        return childSheet;
    }
}
