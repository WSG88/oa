import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class AttendanceCalculation {

    public static void main(String[] args) throws Exception {
        Utils.clear();
        Utils.clearList();
        Utils.YEAR_MONTH = "202010";
        Utils.ROOM = 1;
        Utils.FILE_NAME = Utils.ROOM + ".10.xls";

        List<Data> dataArrayList = new ArrayList<>();
        Workbook wbs = Utils.getWorkbook();
        if (wbs != null) {
            for (int i = 0; i < wbs.getNumberOfSheets(); i++) {
                Sheet childSheet = Utils.getSheet(wbs, null, i);
                dataArrayList.addAll(setDataOne(childSheet));
            }
            Utils.getData(Utils.arrayNamesList, dataArrayList);
        }
    }

    private static List<Data> setDataOne(Sheet childSheet) throws Exception {
        List<Data> dataArrayList = new ArrayList<>();

        //姓名
        String name1 = Utils.readExcelData(childSheet, 3, 9);
        String name2 = Utils.readExcelData(childSheet, 3, 24);
        String name3 = Utils.readExcelData(childSheet, 3, 39);
        Utils.add(name1);
        Utils.add(name2);
        Utils.add(name3);

        //考勤记录
        List<String> list11 = new ArrayList<>();
        List<String> list22 = new ArrayList<>();
        List<String> list33 = new ArrayList<>();

        for (int row = 12; row < 43; row++) {
            //日期
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
                    getData(dataArrayList, name1, day, list11);
                }
                if (cell == 17 - 1
                        || cell == 19 - 1
                        || cell == 22 - 1
                        || cell == 24 - 1
                        || cell == 26 - 1
                        || cell == 28 - 1
                        ) {
                    list22.add(Utils.readExcelDataGetNumericCellValue(childSheet, row, cell));
                    getData(dataArrayList, name2, day, list22);

                }
                if (cell == 32 - 1
                        || cell == 34 - 1
                        || cell == 37 - 1
                        || cell == 39 - 1
                        || cell == 41 - 1
                        || cell == 43 - 1
                        ) {
                    list33.add(Utils.readExcelDataGetNumericCellValue(childSheet, row, cell));
                    getData(dataArrayList, name3, day, list33);
                }
            }
        }
        return dataArrayList;
    }

    private static void getData(List<Data> dataArrayList, String name, String day, List<String> list) {
        if (!Utils.isAdd(name)) {
            return;
        }
        if (list.size() == 6) {
            String date = Utils.YEAR_MONTH + day;
            List<String> arrayList = new ArrayList<>();
            for (String string : list) {
                arrayList.add(Utils.calculateTime(string));
            }
            Data data = new Data(name, date, arrayList);
            dataArrayList.add(data);
            list.clear();
        }
    }

}
