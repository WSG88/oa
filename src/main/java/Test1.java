import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class Test1 {

    public static String YEAR = "202008";

    public static void main(String[] args) throws Exception {
        InputStream in = new FileInputStream("C:/Work/oa/file/1.7.xls");
        Workbook wbs = WorkbookFactory.create(in);
        for (int i = 0; i < wbs.getNumberOfSheets(); i++) {

            Sheet childSheet = getSheet(wbs, null, i);

            main1(childSheet);
        }
    }

    public static void main1(Sheet childSheet) throws Exception {

        String name1 = readExcelData(childSheet, 3, 9);
        String name2 = readExcelData(childSheet, 3, 24);
        String name3 = readExcelData(childSheet, 3, 39);

        List<String> list11 = new ArrayList<>();
        List<String> list22 = new ArrayList<>();
        List<String> list33 = new ArrayList<>();

        for (int row = 12; row < 43; row++) {
            String day = readExcelData(childSheet, row, 0);
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

    public static void insertData(List<String> list, String name, String day) {
        try {
            if (name == null
                    || "".equals(name) || "正常".equals(name) || name.length() == 1) {
                return;
            }

            Db.use().insert(
                    Entity.create("oatime")
                            .set("id", SecureUtil.md5(name + YEAR + day))
                            .set("name", name)
                            .set("day", YEAR + day)
                            .set("d1", ppp(list.get(0)))
                            .set("d2", ppp(list.get(1)))
                            .set("d3", ppp(list.get(2)))
                            .set("d4", ppp(list.get(3)))
                            .set("d5", ppp(list.get(4)))
                            .set("d6", ppp(list.get(5)))
                            .set("room", 1)
            );
        } catch (Exception e) {
        }
    }

    public static String ppp(String string) {
        if (!"缺勤".equals(string)) {
            double d0 = Double.parseDouble(string);
            double dd = d0 * 24;
            int hour = (int) dd;
            int minuter = (int) ((dd - (int) dd) * 60) + 1;
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


    public static String readExcelData(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    return "";
                } else {
//                    System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + cell.getStringCellValue());
                    return cell.getStringCellValue();
                }
            }
        }
        return "";
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
