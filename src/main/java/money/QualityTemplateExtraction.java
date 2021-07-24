package money;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import data.Utils;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class QualityTemplateExtraction {
    public static List<File> PATH_LIST = new ArrayList<>();
    public static List<List<String>> SAVE_LIST = new ArrayList<>();
    public static String PATH_NAME = "";

    public static void main(String[] args) throws Exception {
        PATH_NAME = "239指令\\";
        PATH_NAME = "550厂工艺指令\\";
        PATH_NAME = "602指令\\";
        PATH_NAME = "安达维尔\\";
        PATH_NAME = "昌飞指令\\";
        PATH_NAME = "错误\\";
        PATH_NAME = "";

//        File[] files = FileUtil.ls("D:\\ERP&MES\\天一指令\\" + PATH_NAME);
        File[] files = FileUtil.ls("E:\\信息办\\123123\\" + PATH_NAME);
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
                System.out.println(file.getAbsolutePath());
                AAA();
            }
        }
        toSave();
    }


    private static void toSave() {
        List<String> stringArrayList = new ArrayList<>();
        stringArrayList.add("零件图号");
        stringArrayList.add("工序号");
        stringArrayList.add("工序名称");
        stringArrayList.add("特征描述");
        stringArrayList.add("特征分类");
        stringArrayList.add("特征等级");
        stringArrayList.add("标准值");
        stringArrayList.add("最大值");
        stringArrayList.add("最小值");
        stringArrayList.add("检测方法");
        stringArrayList.add("检验数量");
        stringArrayList.add("检验频次");
        stringArrayList.add("检验单位");
        SAVE_LIST.add(0, stringArrayList);
        ExcelWriter writer = ExcelUtil.getWriter("C:\\Work\\oa\\file\\_品保_" + System.currentTimeMillis() + ".xls");
        writer.write(SAVE_LIST, true);
        writer.close();

        System.out.println(hashSet);
        System.out.println(hashSet1);
        System.out.println(hashSet2);
        List list = Arrays.asList(hashSet2.toArray());
        ExcelWriter writer1 = ExcelUtil.getWriter("C:\\Work\\oa\\file\\品保数据解释.xls");
        writer1.write(list, true);
        writer1.close();
    }

    static String 零件图号 = "";
    static String 工序号 = "";
    static String 工序名称 = "";
    static String 标准值 = "";
    static String 最大值 = "";
    static String 最小值 = "";
    static String 检测方法 = "";
    static String 检验数量 = "";
    static String 检验单位 = "";
    static HashSet hashSet = new HashSet();
    static HashSet hashSet1 = new HashSet();
    static HashSet hashSet2 = new HashSet();

    private static void AAA() throws Exception {
        Workbook wbs = Utils.getWorkbook();
        if (wbs == null) {
            return;
        }
        int cnt = wbs.getNumberOfSheets();
        for (int jj = 0; jj < cnt; jj++) {
            Sheet childSheet = wbs.getSheetAt(jj);
            String sheetName = childSheet.getSheetName();
            int rowNum = childSheet.getPhysicalNumberOfRows();
            int columnNum = childSheet.getRow(0).getPhysicalNumberOfCells();

            if ("流程表".equals(sheetName)) {
                for (int j = 0; j < rowNum; j++) {
                    for (int i = 0; i < columnNum; i++) {
                        String s = Utils.readExcel(childSheet, j, i);
                        if ("数 量".equals(s)) {
                            String s1 = Utils.readExcel(childSheet, j, i + 1);
                            检验数量 = getNumber(s1);
                            检验单位 = getUnit(s1);
                        }
                        if ("零 件 图 号".equals(s)) {
                            零件图号 = Utils.readExcel(childSheet, j, i + 2);
                        }
                    }
                }
            }
            if ("工序表".equals(sheetName)) {
                for (int j = 0; j < rowNum; j++) {
                    for (int i = 0; i < columnNum; i++) {
                        String s = Utils.readExcel(childSheet, j, i);
                        if ("工序号".equals(s)) {
                            工序号 = Utils.readExcel(childSheet, j, i + 1);
                        }
                        if ("工序名称".equals(s)) {
                            工序名称 = Utils.readExcel(childSheet, j, i + 1);
                        }
                        if ("序号".equals(s)) {
                            List<List<String>> list = getDataList(childSheet, rowNum, i, j, s);
                            for (int i1 = 0; i1 < list.size(); i1++) {
                                标准值 = list.get(i1).get(1);

                                标准值 = 标准值.replace("±±", "±");
                                标准值 = 标准值.replace("0.0.12", "0.12");
                                标准值 = 标准值.replace("0.0.43", "0.43");
                                标准值 = 标准值.replace("88.5.5", "88.5");
                                标准值 = 标准值.replace("\"", "″");
                                标准值 = 标准值.replace("〞", "″");
                                标准值 = 标准值.replace("'", "′");
                                标准值 = 标准值.replace("(", "（");
                                标准值 = 标准值.replace(")", "）");
                                标准值 = 标准值.replace(":", "：");
                                标准值 = 标准值.replace("..", ".");
                                标准值 = 标准值.replace("Φ", "∅");
                                标准值 = 标准值.replace("φ", "∅");
                                标准值 = 标准值.replace("<", "＜");
                                标准值 = 标准值.replace(">", "＞");

//                                char[] ch1 = 标准值.toCharArray();
//                                for (char c1 : ch1) {
//                                    hashSet1.add(c1);
//                                }

                                if (标准值.contains("±")
                                        && (标准值.contains("°") || 标准值.contains("′") || 标准值.contains("″"))
                                        && !标准值.contains("×")
                                        && !标准值.contains("～")
                                        && !标准值.contains("锥")
                                        && !标准值.contains("度")
                                        && !标准值.contains("*")
                                        && !标准值.contains("-")
                                        && !标准值.contains("≮")
                                        && !标准值.contains("/")
                                        && !标准值.contains("◁")
                                        && !标准值.contains(":")
                                        ) {

//                                    char[] ch = 标准值.toCharArray();
//                                    for (char c : ch) {
//                                        hashSet.add(c);
//                                    }

                                    String[] sss = 标准值.split("±");
                                    if (sss.length == 2) {
                                        String s0 = sss[0];
                                        String s1 = sss[1];
                                        if ((s0.contains("°") || s0.contains("′") || s0.contains("″"))
                                                && (s1.contains("°") || s1.contains("′") || s1.contains("″"))) {
                                            double d0 = Double.parseDouble(dms2D(sss[0], 标准值));
                                            double d1 = Double.parseDouble(dms2D(sss[1], 标准值));
                                            最大值 = getDoubleString(d0 + d1);
                                            最小值 = getDoubleString(d0 - d1);
                                        } else {
                                            hashSet2.add(标准值);
                                        }
                                    } else {
                                        hashSet2.add(标准值);
                                    }
                                } else if (标准值.contains("±")
                                        && !标准值.contains("°")
                                        && !标准值.contains("′")
                                        && !标准值.contains("R")
                                        && !标准值.contains("×")
                                        && !标准值.contains("C")
                                        && !标准值.contains("∅")
                                        && !标准值.contains("∅")
                                        && !标准值.contains("φ")
                                        && !标准值.contains("G")
                                        && !标准值.contains("J")
                                        && !标准值.contains("K")
                                        && !标准值.contains("S")
                                        && !标准值.contains("Φ")
                                        && !标准值.contains("g")
                                        && !标准值.contains("k")
                                        && !标准值.contains("≯")
                                        && !标准值.contains("-")
                                        ) {
                                    String[] sss = 标准值.split("±");
                                    if (sss.length == 2) {
                                        if (sss[0].length() == 0) {
                                            sss[0] = "0";
                                        }
                                        double d0 = Double.parseDouble(sss[0]);
                                        double d1 = Double.parseDouble(sss[1]);
                                        最大值 = String.format("%.2f", d0 + d1);
                                        最小值 = String.format("%.2f", d0 - d1);
                                    } else {
                                        hashSet2.add(标准值);
                                    }
                                } else if (标准值.contains("±")
                                        && (标准值.contains("R") || 标准值.contains("SR"))
                                        && !标准值.contains("-")
                                        && !标准值.contains("/")
                                        ) {
                                    String[] sss = 标准值.split("±");
                                    if (sss.length == 2) {
                                        double d0 = Double.parseDouble(sss[0].replace("SR", "").replace("R", ""));
                                        double d1 = Double.parseDouble(sss[1].replace("（典型）",""));
                                        最大值 = String.format("%.2f", d0 + d1);
                                        最小值 = String.format("%.2f", d0 - d1);
                                    } else {
                                        hashSet2.add(标准值);
                                    }
                                } else {
                                    hashSet2.add(标准值);
                                }

                                检测方法 = list.get(i1).get(2);
                                List<String> stringArrayList = new ArrayList<>();
                                stringArrayList.add(零件图号);
                                stringArrayList.add(工序号);
                                stringArrayList.add(工序名称);
                                stringArrayList.add("");
                                stringArrayList.add("");
                                stringArrayList.add("");
                                stringArrayList.add(标准值);
                                stringArrayList.add(最大值);
                                stringArrayList.add(最小值);
                                stringArrayList.add(检测方法);
                                stringArrayList.add(检验数量);
                                stringArrayList.add("");
                                stringArrayList.add(检验单位);
                                SAVE_LIST.add(stringArrayList);

                                标准值 = "";
                                最大值 = "";
                                最小值 = "";
                                检测方法 = "";
                            }
                        }
                    }
                }
                工序号 = "";
                工序名称 = "";
            }
        }
        wbs.close();
    }

    private static String getDoubleString(double dd) {
        int d = (int) dd;
        int m = (int) ((dd - d) * 60);
        int s = (int) (((dd - d) * 60 - m) * 60);
        if (s == 59) {
            m = m + 1;
            s = 0;
        }
        String s1 = d > 0 ? d + "°" : "";
        String s2 = m > 0 ? m + "′" : "";
        String s3 = s > 0 ? s + "″" : "";
        return s1 + s2 + s3;
    }

    private static List<List<String>> getDataList(Sheet childSheet, int rowNum, int i, int j, String s) {
        List<List<String>> listList = new ArrayList<>();
        for (int k = j + 1; k < rowNum; k++) {
            try {
                String ss = Utils.readExcel(childSheet, k, i);
                float in = Float.parseFloat(ss);
                String ss1 = Utils.readExcel(childSheet, k, i + 2);
                String ss3 = Utils.readExcel(childSheet, k, i + 3);
                List<String> list = new ArrayList<>();
                list.add(ss);
                list.add(ss1);
                list.add(ss3);
                listList.add(list);
            } catch (Exception e) {
                return listList;
            }
        }
        return listList;
    }

    public static String getNumber(String s) {
        String regEx = "[^0-9]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(s);
        return m.replaceAll("").trim();
    }

    public static String getUnit(String s) {
        String number = getNumber(s);
        return s.replace(number, "");
    }

    public static String dms2D(String string, String org) {
        double d = 0, m = 0, s = 0;
        if (string.contains("°") && string.contains("′") && string.contains("″")) {
            String[] dString = string.split("°");
            d = Double.parseDouble(dString[0]);
            String[] mString = dString[1].split("′");
            m = Double.parseDouble(mString[0]);
            s = Double.parseDouble(mString[1].replace("″", ""));
        } else if (string.contains("°") && string.contains("′")) {
            String[] dString = string.split("°");
            d = Double.parseDouble(dString[0]);
            m = Double.parseDouble(dString[1].replace("′", ""));
        } else if (string.contains("°") && string.contains("″")) {
            String[] dString = string.split("°");
            d = Double.parseDouble(dString[0]);
            s = Double.parseDouble(dString[1].replace("″", ""));
        } else if (string.contains("′") && string.contains("″")) {
            String[] dString = string.split("′");
            m = Double.parseDouble(dString[0]);
            s = Double.parseDouble(dString[1].replace("″", ""));
        } else if (string.contains("°")) {
            String[] dString = string.split("°");
            d = Double.parseDouble(dString[0]);
        } else if (string.contains("′")) {
            String[] dString = string.split("′");
            m = Double.parseDouble(dString[0]);
        } else if (string.contains("″")) {
            String[] dString = string.split("″");
            s = Double.parseDouble(dString[0]);
        } else {
            System.out.println(string + " " + org);
            return string;
        }
        return String.valueOf(d + m / 60 + s / 60 / 60);
    }
}
