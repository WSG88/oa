import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class AttendanceCalculation {


    public static void main(String[] args) throws Exception {
        Utils.clearList();
        System.out.println("----------------11111111111111111111111111111\n\n");
        testOne();
        Thread.sleep(5000);
        System.out.println("----------------22222222222222222222222222222\n\n");
        testTwo();
        Utils.printList();
    }

    public static void testOne() throws Exception {

        Utils.YEAR_MONTH = "202009";
        Utils.FILE_NAME = "1.9.xls";
        Utils.ROOM = 1;
        Utils.clear();
        List<Data> dataArrayList11 = new ArrayList<>();

        Workbook wbs = Utils.getWorkbook();
        for (int i = 0; i < wbs.getNumberOfSheets(); i++) {
            Sheet childSheet = Utils.getSheet(wbs, null, i);
            List<Data> dataArrayList = setDataOne(childSheet);
            for (Data data : dataArrayList) {
                Utils.saveToDatabase(data, Utils.ROOM);
            }
            dataArrayList11.addAll(dataArrayList);
        }

        //元数据保存到EXCEL
        Utils.copyData(Utils.arrayNamesList, dataArrayList11);

        //计算并保存
        Utils.getData(Utils.arrayNamesList, Utils.YEAR_MONTH);
    }

    public static List<Data> setDataOne(Sheet childSheet) throws Exception {
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
        if (list.size() == 6) {
            Data data = insertData(list, name, day);
            if (data != null) {
                dataArrayList.add(data);
            }
            list.clear();
        }
    }

    public static Data insertData(List<String> list, String name, String day) {
        if (!Utils.isAdd(name)) {
            return null;
        }
        String date = Utils.YEAR_MONTH + day;
        List<String> arrayList = new ArrayList<>();
        for (String string : list) {
            arrayList.add(Utils.calculateTime(string));
        }
        return new Data(name, date, arrayList);
    }

    public static void testTwo() throws Exception {
        Utils.YEAR_MONTH = "202009";
        Utils.FILE_NAME = "2.9.xls";
        Utils.ROOM = 2;
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);

        List<Data> dataArrayList = setDataTwo(childSheet);
        for (Data data : dataArrayList) {
            Utils.saveToDatabase(data, Utils.ROOM);
        }

        Utils.getData(Utils.arrayNamesList, Utils.YEAR_MONTH);
    }

    public static List<Data> setDataTwo(Sheet childSheet) throws Exception {
        List<Data> dataArrayList = new ArrayList<>();
        for (int index = 5; index < childSheet.getLastRowNum() + 1; index = index + 2) {
            //姓名
            String name = Utils.readExcelData(childSheet, index - 1, 10);
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);
            Row row = childSheet.getRow(index);
            if (row != null) {
                int kk = row.getLastCellNum();
                for (int i = 0; i < kk; i++) {
                    Cell cell = row.getCell(i);
                    //日期
                    String date = Utils.YEAR_MONTH + String.format("%02d", i + 1);
                    //考勤记录
                    List<String> list = new ArrayList<>();
                    if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
                        String string = cell.getStringCellValue();
                        int len = string.length();
                        int ll = len / 5;
                        for (int ii = 0; ii < ll; ii++) {
                            String sss = string.substring(ii * 5, ii * 5 + 5);
                            list.add(sss);
                        }
                    }
                    if (!list.isEmpty()) {
                        //补全考勤数据
                        Utils.completeQueQing(list);
                        dataArrayList.add(new Data(name, date, list));
                    }
                }
            }
        }
        return dataArrayList;
    }
}
