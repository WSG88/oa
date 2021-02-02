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
        Utils.YEAR_MONTH = "202101";
        Utils.FILE_PATH = "F:\\WORK\\oa\\file\\";
        Utils.FILE_NAME = "1.202101.xlsx";
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);
        Utils.clear();

        List<Data> dataArrayList = new ArrayList<>();
        Map<String, ArrayList<String>> map = new HashMap<>();
        for (int line = 1; line < childSheet.getLastRowNum() + 1; line++) {
            String name = Utils.readExcelData(childSheet, line , 2);
            String date = Utils.readExcelData(childSheet, line , 4);
            String s = Utils.readExcelData(childSheet, line , 5);
            //
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);

            String key = name + "," + date.replace("-","");
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
//        System.out.println(map);

        List<String> newList = Utils.arrayNamesList.stream().distinct().collect(Collectors.toList());
        Utils.arrayNamesList=newList;

//        System.out.println(Utils.arrayNamesList);

        for (String name : Utils.arrayNamesList) {

            for (int i1 = 1; i1 < DAY; i1++) {
                String date = Utils.YEAR_MONTH + String.format("%02d", i1);
                String key = name + "," + date;
//                System.out.println(key);
                ArrayList<String> arrayList = map.get(key);
                //培训
                if ("20201226".equals(date)) {
                    if (
                            "	蓝天航	".trim().equals(name) ||
                                    "	肖立军	".trim().equals(name) ||
                                    "	程宇文	".trim().equals(name) ||
                                    "	危发强	".trim().equals(name) ||
                                    "	陈震宇	".trim().equals(name) ||
                                    "	刘小龙	".trim().equals(name) ||
                                    "	万立妹	".trim().equals(name) ||
                                    "	万运来	".trim().equals(name) ||
                                    "	曹明光	".trim().equals(name) ||
                                    "	张康文	".trim().equals(name) ||
                                    "	葛银保	".trim().equals(name) ||
                                    "	张新丽	".trim().equals(name) ||
                                    "	王思刚	".trim().equals(name) ||
                                    "	苏金丽	".trim().equals(name) ||
                                    "	程厚阳	".trim().equals(name) ||
                                    "	严命坤	".trim().equals(name) ||
                                    "	黄亦龙	".trim().equals(name) ||
                                    "	王章美	".trim().equals(name) ||
                                    "	江鹏	".trim().equals(name) ||
                                    "	虞涛	".trim().equals(name) ||
                                    "	陈亚军	".trim().equals(name) ||
                                    "	张志旗	".trim().equals(name) ||
                                    "	陈妍	".trim().equals(name) ||
                                    "	徐靖	".trim().equals(name) ||
                                    "	苏芳华	".trim().equals(name) ||
                                    "	刘镇	".trim().equals(name) ||
                                    "	张学成	".trim().equals(name) ||
                                    "	方兴兴	".trim().equals(name) ||
                                    "	沈长征	".trim().equals(name) ||
                                    "	侯木财	".trim().equals(name) ||
                                    "	王金宝	".trim().equals(name) ||
                                    "	蒋婵	".trim().equals(name) ||
                                    "	罗飞	".trim().equals(name) ||
                                    "	邵泽球	".trim().equals(name) ||
                                    "	张涛	".trim().equals(name) ||
                                    "	张祖胜	".trim().equals(name) ||
                                    "	张欢	".trim().equals(name) ||
                                    "	李开立	".trim().equals(name) ||
                                    "	周谟林	".trim().equals(name) ||
                                    "	杜传国	".trim().equals(name) ||
                                    "	刘太华	".trim().equals(name) ||
                                    "	许凌玉	".trim().equals(name) ||
                                    "	徐慧林	".trim().equals(name) ||
                                    "	苏俊潇	".trim().equals(name) ||
                                    "	章水根	".trim().equals(name) ||
                                    "	方智鑫	".trim().equals(name) ||
                                    "	郑彦俊	".trim().equals(name) ||
                                    "	张志成	".trim().equals(name)
                    ) {
                        arrayList=new ArrayList<>();
                        arrayList.add("07:59");
                        arrayList.add("11:31");
                        arrayList.add("11:59");
                        arrayList.add("16:31");
                    }
                }

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
