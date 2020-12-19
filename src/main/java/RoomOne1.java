import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class RoomOne1 {
    static int DAY = 32;

    public static void main(String[] args) throws Exception {
        Utils.clearList();
        Utils.ROOM = 1;
        Utils.YEAR_MONTH = "202011";
        Utils.FILE_PATH = "d:\\Work\\oa\\file\\";
        Utils.FILE_NAME = "1111.xlsx";
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);
        Utils.clear();

        List<Data> dataArrayList = new ArrayList<>();
        Map<String, ArrayList<String>> map = new HashMap<>();
        for (int line = 0; line < childSheet.getLastRowNum() + 1; line++) {
            String name = Utils.readExcelData(childSheet, line - 1, 0);
            String date = Utils.readExcelData(childSheet, line - 1, 1);
            String s = Utils.readExcelData(childSheet, line - 1, 2);
            //
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);

            String key = name + "," + date;
            if (key.contains("2020-12-") || key.contains("2020-10-")) {
                continue;
            }
            ArrayList<String> arrayList = map.get(key);
            if (arrayList == null) {
                arrayList = new ArrayList<>();
                map.put(key, arrayList);
            }

            if (s != null && s.trim().length() > 0) {
                arrayList.add(s.trim());
            }
            map.put(key, arrayList);
        }
        System.out.println(map);

        List<String> newList = Utils.arrayNamesList.stream().distinct().collect(Collectors.toList());
        Utils.arrayNamesList=newList;

        System.out.println(Utils.arrayNamesList);

        for (String name : Utils.arrayNamesList) {

            for (int i1 = 0; i1 < DAY; i1++) {
                String key = name + ",2020-11-" + i1;
                if (i1 < 10) {
                    key = name + ",2020-11-0" + i1;
                }
                ArrayList<String> arrayList = map.get(key);
                if (arrayList != null && !arrayList.isEmpty()) {
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
        System.out.println(dataArrayList);

        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList);

        Utils.printList();
    }
}
