package data;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class TTT {
    public static void main(String[] args) throws Exception {
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
                long l2 = DateUtil.parse(date + "08:06", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l3 = DateUtil.parse(date + "19:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l4 = DateUtil.parse(date + "20:06", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                long l5 = DateUtil.parse(date + "24:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                if (listNew.size() == 1 || listNew.size() == 2) {
                    if (listNew.size() == 1) {
                        String str0 = listNew.get(0);
                        long l0 = DateUtil.parse(date + str0, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
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
                        long l0 = DateUtil.parse(date + str0, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
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