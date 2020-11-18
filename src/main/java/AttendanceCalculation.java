import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class AttendanceCalculation {
    public static String YEAR = "2020";
    public static String MONTH = "10";

    public static void main(String[] args) throws Exception {
        Utils.YEAR_MONTH = YEAR + MONTH;

        Utils.clearList();
        System.out.println("----------------11111111111111111111111111111\n\n");
        testOne();
        Thread.sleep(5000);
        System.out.println("----------------22222222222222222222222222222\n\n");
        testTwo();
        Utils.printList();
    }

    public static void testOne() throws Exception {
        Utils.ROOM = 1;
        Utils.FILE_NAME = Utils.ROOM + "." + MONTH + ".xls";
        Utils.clear();
        List<Data> dataArrayList11 = new ArrayList<>();

        Workbook wbs = Utils.getWorkbook();
        for (int i = 0; i < wbs.getNumberOfSheets(); i++) {
            Sheet childSheet = Utils.getSheet(wbs, null, i);
            dataArrayList11.addAll(setDataOne(childSheet));
        }
        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList11, dataArrayList11);
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
        Utils.ROOM = 2;
        Utils.FILE_NAME = Utils.ROOM + "." + MONTH + ".xls";
        Utils.clear();
        List<Data> dataArrayList11 = new ArrayList<>();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);
        dataArrayList11.addAll(setDataTwo(childSheet));
        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList11, dataArrayList11);
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

    public static void again() throws Exception {
        Utils.YEAR_MONTH = "202009";
        Utils.FILE_NAME = "202009_1车间补.xlsx";
        Utils.ROOM = 1;
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);
        List<Data> dataArrayList = new ArrayList<>();
        List<Data> dataArrayList11 = new ArrayList<>();
        for (int rowNumber = 0; rowNumber < childSheet.getLastRowNum() + 1; rowNumber = rowNumber + 13) {
            String name = Utils.readExcel(childSheet, rowNumber, 0);
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);
            Row row = childSheet.getRow(rowNumber + 1);
            int totalCellNumber = row.getLastCellNum();
            for (int cellNumber = 0; cellNumber < totalCellNumber; cellNumber++) {
                //日期
                String date = Utils.YEAR_MONTH + String.format("%02d", cellNumber + 1);
                //考勤记录
                List<String> list = new ArrayList<>();
                for (int j = rowNumber + 2; j < rowNumber + 8; j++) {
                    String str = Utils.readExcel(childSheet, j, cellNumber);
                    list.add(str);
                }
                dataArrayList11.add(new Data(name, date, list));

                List<String> listNew = new ArrayList<>();
                for (String s : list) {
                    if (StrUtil.isEmpty(s) || Utils.QUE_QING.equals(s)) {

                    } else {
                        listNew.add(s);
                    }
                }
                //处理夜班数据
                long l2 = cn.hutool.core.date.DateUtil.parse(date + "08:06", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l3 = cn.hutool.core.date.DateUtil.parse(date + "19:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l4 = cn.hutool.core.date.DateUtil.parse(date + "20:06", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l5 = cn.hutool.core.date.DateUtil.parse(date + "24:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                if (listNew.size() == 1 || listNew.size() == 2) {
                    if (listNew.size() == 1) {
                        String str0 = listNew.get(0);
                        long l0 = cn.hutool.core.date.DateUtil.parse(date + str0, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                        if (l0 < l2) {
                            listNew.clear();
                            listNew.add("00:00");
                            listNew.add(str0);
                        } else if (l0 < l4 && l0 > l3) {
                            listNew.clear();
                            listNew.add(str0);
                            listNew.add("24:00");
                        } else {
                            System.out.println(name + date + listNew);
                        }
                    }
                    if (listNew.size() == 2) {
                        String str0 = listNew.get(0);
                        String str1 = listNew.get(1);
                        long l0 = cn.hutool.core.date.DateUtil.parse(date + str0, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                        long l1 = DateUtil.parse(date + str1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                        if (l0 < l2 && ((l1 < l4 && l1 > l3) || (l1 <= l5 && l1 > l3))) {
                            listNew.clear();
                            listNew.add("00:00");
                            listNew.add(str0);
                            listNew.add(str1);
                            listNew.add("24:00");
                        } else if (l0 > l3 && l1 <= l5 && l1 > l3) {

                        } else if (l1 < l2) {

                        } else {
                            System.out.println(name + date + listNew);
                        }
                    }
                }

                if (!listNew.isEmpty()) {
                    //补全考勤数据
                    Utils.completeQueQing(listNew);
                    dataArrayList.add(new Data(name, date, listNew));
                }
            }
        }
        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList, dataArrayList11);
    }
}
