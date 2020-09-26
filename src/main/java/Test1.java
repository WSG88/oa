import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class Test1 {

    public static void main(String[] args) throws Exception {

        Utils.YEAR_MONTH = "202008";
        Utils.FILE_NAME = "1.8.xls";
        Utils.ROOM = 1;
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        for (int i = 0; i < wbs.getNumberOfSheets(); i++) {
            Sheet childSheet = Utils.getSheet(wbs, null, i);
            setData(childSheet);
        }
        Utils.getData(Utils.arrayNamesList);
    }


    public static void setData(Sheet childSheet) throws Exception {

        String name1 = Utils.readExcelData(childSheet, 3, 9);
        String name2 = Utils.readExcelData(childSheet, 3, 24);
        String name3 = Utils.readExcelData(childSheet, 3, 39);
        Utils.add(name1);
        Utils.add(name2);
        Utils.add(name3);

        List<String> list11 = new ArrayList<>();
        List<String> list22 = new ArrayList<>();
        List<String> list33 = new ArrayList<>();

        for (int row = 12; row < 43; row++) {
            String day = Utils.readExcelData(childSheet, row, 0);
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
                    list11.add(Utils.readExcelDataGetNumericCellValue(childSheet, row, cell));
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
                    list22.add(Utils.readExcelDataGetNumericCellValue(childSheet, row, cell));

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
                    list33.add(Utils.readExcelDataGetNumericCellValue(childSheet, row, cell));

                    if (list33.size() == 6) {
                        insertData(list33, name3, day);
                        list33.clear();
                    }
                }
            }
        }

    }

    public static void insertData(List<String> l, String name, String day) {
        if (!Utils.isAdd(name)) {
            return;
        }
        String date = Utils.YEAR_MONTH + day;
        List<String> list = new ArrayList<>();
        for (String s : l) {
            list.add(Utils.calculateTime(s));
        }
        Utils.saveToDatabase(new Data(name, date, list), Utils.ROOM);
    }

}
