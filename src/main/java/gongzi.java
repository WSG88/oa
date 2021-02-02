import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import data.Gz;
import data.Utils;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.text.DecimalFormat;
import java.util.*;

public class gongzi {
    public static List<File> PATH_LIST = new ArrayList<>();
    public static List<List<String>> SAVE_LIST = new ArrayList<>();
    public static String PATH_DIR = "C:\\Users\\Administrator\\Desktop\\";
    public static String PATH_NAME = "";

    public static void main(String[] args) throws Exception {
        PATH_NAME = "aaa\\";

        File[] files = FileUtil.ls(PATH_DIR + PATH_NAME);
        for (int i = 0; i < files.length; i++) {
            File file = files[i];
            if (file.isDirectory()) {
                File[] filess = FileUtil.ls(file.getAbsolutePath());
                for (int j = 0; j < filess.length; j++) {
                    File fileFile = filess[j];
                    Utils.FILE_PATH = fileFile.getParent();
                    Utils.FILE_NAME = fileFile.getName();
                    PATH_LIST.add(fileFile);
                }
            } else {
                Utils.FILE_PATH = file.getParent();
                Utils.FILE_NAME = file.getName();
                PATH_LIST.add(file);
            }
        }

        List<List<File>> subs = ListUtils.partition(PATH_LIST, 100);
        for (List<File> sub : subs) {
            for (int i = 0; i < sub.size(); i++) {
                File file = sub.get(i);
                Utils.FILE_PATH = file.getParent() + "\\";
                Utils.FILE_NAME = file.getName();
                AAA();
            }
        }
//        toSave();
    }


    private static void toSave() {
        List<String> stringArrayList = new ArrayList<>();
        stringArrayList.add("姓名");
        stringArrayList.add("实付工资");
        SAVE_LIST.add(0, stringArrayList);
        ExcelWriter writer = ExcelUtil.getWriter(PATH_DIR + System.currentTimeMillis() + ".xls");
        writer.write(SAVE_LIST, true);
        writer.close();
    }


    private static void AAA() throws Exception {
        Map<String, ArrayList<Gz>> map = new HashMap<>();
        Workbook wbs = Utils.getWorkbook();
        if (wbs == null) {
            return;
        }
        int cnt = wbs.getNumberOfSheets();
        for (int jj = 0; jj < cnt; jj++) {
            Sheet childSheet = wbs.getSheetAt(jj);
            String sheetName = childSheet.getSheetName();
            String month = sheetName.replace("月", "");
            int rowNum = childSheet.getPhysicalNumberOfRows();
            int columnNum = childSheet.getRow(0).getPhysicalNumberOfCells();
            for (int j = 0; j < rowNum; j++) {
                int i1 = 0, i2 = 0;
                for (int i = 0; i < columnNum; i++) {
                    String s = Utils.readExcel(childSheet, j, i);
                    if ("姓名".equals(s)||"姓  名".equals(s)) {
                        i1 = i;
                    }
                    if ("实付工资".equals(s)) {
                        i2 = i;
                    }
                    if (i1 > 0 && i2 > 0) {
                        for (int k = j; k < 200; k++) {
                            String name = Utils.readExcel(childSheet, k, i1);
                            String sale = Utils.readExcel(childSheet, k, i2);
                            String key = name;
                            if (name.length() > 0 && sale.length() > 0 &&!"姓名".equals(name) && !"姓  名".equals(name) && !"实付工资".equals(sale)) {
                                ArrayList<Gz> arrayList = map.computeIfAbsent(key, k1 -> new ArrayList<>());
                                Gz gz = new Gz();
                                gz.setName(name);
                                gz.setSale(sale);
                                gz.setMonth(month);
                                if (!arrayList.contains(gz)) {
                                    arrayList.add(gz);
                                }
                                map.put(key, arrayList);
                            }
                        }
                    }
                }

            }
        }
        wbs.close();
//        System.out.println(map);

        List<Gz> listList = new ArrayList<>();

        Object key[] = map.keySet().toArray();
        for (int i = 0; i < key.length; i++) {
            List<Gz> list = map.get(key[i]);
            String name = (String) key[i];
            double saleD = 0.0;
            for (int j = 0; j < list.size(); j++) {
                Gz gz = list.get(j);
                System.out.println(gz);
                String sale = gz.getSale();
                if (sale == null || sale.length() == 0) {
                    sale = "0.0";
                }
                saleD += Double.parseDouble(sale);
            }

            Gz gg = new Gz();
            gg.setName(name);
            DecimalFormat df = new DecimalFormat("0.00");
            gg.setSale(df.format(saleD));
            listList.add(gg);
        }

        listList.sort(new Comparator<Gz>() {
            @Override
            public int compare(Gz o1, Gz o2) {
                double d1 = Double.parseDouble(o1.getSale());
                double d2 = Double.parseDouble(o2.getSale());
                if (d1 == d2) {
                    return 0;
                }
                if (d1 > d2) {
                    return -1;
                }
                return 1;
            }
        });
        for (int i = 0; i < listList.size(); i++) {
            System.out.println(listList.get(i).getString());

        }
    }

    //字母转数字  A-Z ：1-26
    public static int letterToNumber(String letter) {
        int length = letter.length();
        int num = 0;
        int number = 0;
        for (int i = 0; i < length; i++) {
            char ch = letter.charAt(length - i - 1);
            num = (int) (ch - 'A' + 1);
            num *= Math.pow(26, i);
            number += num;
        }
        return number;
    }

    //数字转字母 1-26 ： A-Z
    public static String numberToLetter(int num) {
        if (num <= 0) {
            return null;
        }
        String letter = "";
        num--;
        do {
            if (letter.length() > 0) {
                num--;
            }
            letter = ((char) (num % 26 + (int) 'A')) + letter;
            num = (int) ((num - num % 26) / 26);
        } while (num > 0);

        return letter;
    }
}
