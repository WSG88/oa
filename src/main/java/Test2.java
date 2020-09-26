import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class Test2 {

    public static void main(String[] args) throws Exception {
        Utils.YEAR_MONTH = "202008";
        Utils.FILE_NAME = "2.81.xls";
        Utils.ROOM = 2;
        Utils.clear();

        setData();
        Utils.getData(Utils.arrayNamesList);
    }

    public static void setData() throws Exception {
        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);
        List<String> list;//考勤记录
        for (int index = 7; index < childSheet.getLastRowNum() + 1; index = index + 2) {
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
                    list = new ArrayList<>();
                    if (cell != null) {
                        if (cell.getCellTypeEnum() == CellType.STRING) {
                            String string = cell.getStringCellValue();
                            int len = string.length();
                            int ll = len / 5;
                            for (int ii = 0; ii < ll; ii++) {
                                String sss = string.substring(ii * 5, ii * 5 + 5);
                                list.add(sss);
                            }
                            //检查考勤数据是否完整

//                            //早上多打卡
//                            if (list.size() > 1) {
//                                int l = 0;
//                                for (String s : list) {
//                                    if (s.startsWith("07")) {
//                                        l++;
//                                    }
//                                }
//                                if (l > 1) {
//                                    System.out.println(name + "   " + date + " " + list);
//                            return;
//                                }
//                            }
//                            //打卡次数缺失
//                            if (list.size() == 1 || list.size() == 3 || list.size() == 5) {
//                                System.out.println(name + "   " + date + " " + list);
//                            return;
//                            }

//                            //是否请假
//                            if (list.size() == 2) {
//                                System.out.println(name + "   " + date + " " + list);
//                            } else if (list.size() == 4) {
//                                if (!(list.get(0).startsWith("07") || list.get(0).startsWith("08"))) {
//                                    System.out.println(name + "   " + date + " " + list);
//                                }
//                            }

                            //清空一次打卡
                            if (list.size() == 1) {
                                list.clear();
                                continue;
                            }
                            //补全考勤数据
                            Utils.completeQueQing(list);
                            //检查各区间值是否正常

                            Utils.saveToDatabase(new Data(name, date, list), Utils.ROOM);

                        }

                    }

                }
            }

        }


    }


}
