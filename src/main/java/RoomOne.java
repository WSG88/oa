import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class RoomOne {
    static int DAY = 32;

    public static void main(String[] args) throws Exception {
        Utils.clearList();
        Utils.ROOM = 1;
        Utils.YEAR_MONTH = "202010";
        Utils.FILE_PATH = "d:\\Work\\oa\\file\\";
        Utils.FILE_NAME = "员工刷卡记录表.xls";
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);

        List<Data> dataArrayList = new ArrayList<>();

        List<Integer> lineNumberList = new ArrayList<>();
        //姓名所在的行号集合
        for (int line = 0; line < childSheet.getLastRowNum() + 1; line++) {
            String value = Utils.readExcelData(childSheet, line - 1, 11);
            if (value.length() > 0 && !value.contains(":")) {
                lineNumberList.add(line);
            }
        }

        for (int i = 0; i < lineNumberList.size(); i++) {
            int lineStart = lineNumberList.get(i) - 1;
            int lineEnd;
            if (i == lineNumberList.size() - 1) {
                lineEnd = 4;
            } else {
                lineEnd = lineNumberList.get(i + 1) - 1;
            }
            String name = Utils.readExcelData(childSheet, lineStart, 11);
//            System.out.println(lineStart + "  " + name);
            //
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);

            Map<String, ArrayList<String>> map = new HashMap<>();
            for (int i1 = 0; i1 < DAY; i1++) {
                String key = name + "," + i1;
                ArrayList<String> arrayList = map.get(key);
                if (arrayList == null) {
                    arrayList = new ArrayList<>();
                    map.put(key, arrayList);
                }
                for (int i2 = 1; i2 < lineEnd - lineStart; i2++) {
                    String value = Utils.readExcelData(childSheet, lineStart + i2, i1);
                    if (value.length() > 0) {
                        String[] strings = value.split("\n");
                        for (String s : strings) {
                            if (s != null && s.trim().length() > 0) {
                                arrayList.add(s.trim());
                            }
                        }
                        map.put(key, arrayList);
                    }
                }
            }
//            System.out.println(map);
            //
            for (int i1 = 0; i1 < DAY; i1++) {
                String key = name + "," + i1;
                ArrayList<String> arrayList = map.get(key);
                if (!arrayList.isEmpty()) {
                    if (arrayList.size() == 7) {

                        if (arrayList.get(0).startsWith("07") && arrayList.get(1).startsWith("08")) {
                            arrayList.remove(0);
                        }

                        if (arrayList.get(3).startsWith("16:36") && arrayList.get(4).startsWith("17:31")) {
                            arrayList.remove("16:36");
                        }
                    }
                    if (arrayList.size() == 7) {
                        ArrayList<String> arrayList7 = new ArrayList<>();
                        ArrayList<String> arrayList12 = new ArrayList<>();
                        ArrayList<String> arrayList17 = new ArrayList<>();
                        ArrayList<String> arrayList21 = new ArrayList<>();
                        for (String s : arrayList) {
                            if (s.startsWith("07")) {
                                arrayList7.add(s);
                            }
                            if (s.startsWith("12")) {
                                arrayList12.add(s);
                            }
                            if (s.startsWith("17")) {
                                arrayList17.add(s);
                            }
                            if (s.startsWith("21")) {
                                arrayList21.add(s);
                            }
                        }
                        if (arrayList7.size() > 1) {
                            arrayList.removeAll(arrayList7);
                            arrayList.add(0, arrayList7.get(0));
                        }
                        if (arrayList12.size() > 2) {
                            arrayList.remove(arrayList12.get(arrayList12.size() - 1));
                        }
                        if (arrayList17.size() > 2) {
                            arrayList.remove(arrayList17.get(arrayList17.size() - 1));
                        }
                        if (arrayList21.size() > 1) {
                            arrayList.remove(arrayList21.get(0));
                        }
                    }

//                    if (arrayList.size() == 5) {
//                        System.out.println(key + "," + arrayList);
//                    }

                    Utils.completeQueQing(arrayList);
                    String date = Utils.YEAR_MONTH + String.format("%02d", i1);
                    dataArrayList.add(new Data(name, date, arrayList));
                }
            }
        }
//        System.out.println(dataArrayList);

        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList);

        Utils.printList();
    }
}
