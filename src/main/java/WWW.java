import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import data.Utils;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WWW {
    public static List<File> PATH_LIST = new ArrayList<>();
    public static List<List<String>> GX_LIST = new ArrayList<>();
    public static List<List<String>> SB_LIST = new ArrayList<>();
    public static List<List<String>> SAVE_LIST = new ArrayList<>();
    public static String NUMBER = "";

    public static void main(String[] args) throws Exception {
        String path = "D:\\ERP&MES\\天一指令\\昌飞指令\\";

        File[] files = FileUtil.ls(path);
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
                BBB();
                toAdd();
                System.out.println(file.getAbsolutePath());
            }
        }
        toSave();
    }


    private static void toSave() {
        List<String> stringArrayList = new ArrayList<>();
        stringArrayList.add("零件图号");
        stringArrayList.add("工序号");
        stringArrayList.add("工序名称");
        stringArrayList.add("工序内容");
        stringArrayList.add("加工设备");
        stringArrayList.add("准备时间");
        stringArrayList.add("加工时间");
        stringArrayList.add("是否MES管控");
        SAVE_LIST.add(0, stringArrayList);
        ExcelWriter writer = ExcelUtil.getWriter("C:\\Work\\oa\\file\\工艺批量导入数据" + System.currentTimeMillis() +
                ".xls");
        writer.write(SAVE_LIST, true);
        writer.close();
        System.out.println(SAVE_LIST);
    }

    private static void toAdd() {
        for (List<String> list : GX_LIST) {
            String a = list.get(0);
            String a1 = String.valueOf(Double.parseDouble(a));
            String b = list.get(1);
            String c = list.get(2);
            String d = "";
            String e = "60";
            String f = "60";
            String g = "是";
            for (List<String> strings : SB_LIST) {
                if (a.equals(strings.get(0))) {
                    d = strings.get(1);
                }
            }
            List<String> stringArrayList = new ArrayList<>();
            stringArrayList.add(NUMBER);
            stringArrayList.add(a1);
            stringArrayList.add(b);
            stringArrayList.add(c);
            stringArrayList.add(d);
            stringArrayList.add(e);
            stringArrayList.add(f);
            stringArrayList.add(g);
            SAVE_LIST.add(stringArrayList);
        }
        GX_LIST.clear();
        SB_LIST.clear();
        NUMBER = "";
    }

    private static void BBB() throws IOException, InvalidFormatException {
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
            for (int i = 0; i < columnNum; i++) {
                for (int j = 0; j < rowNum; j++) {
                    try {
                        if ("机床设备".equals(Utils.readExcelData(childSheet, j, i))) {
                            for (int k = j + 1; k < rowNum; k++) {
                                String ss = Utils.readExcelData(childSheet, k, i);
                                if (!Utils.QUE_QING.equals(ss) && ss.length() > 0) {
                                    String ss1 = Utils.readExcelDataGetNumericCellValue(childSheet, k, 1);
//                                    String ss2 = Utils.readExcelData(childSheet, k, 2);
//                                    int ii = (int) Double.parseDouble(ss1);
                                    List<String> list = new ArrayList<>();
                                    list.add(ss1);
                                    list.add(ss);
                                    SB_LIST.add(list);
                                }
                            }
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    try {
                        if ("零 件 图 号".equals(Utils.readExcelData(childSheet, j, i))) {
                            NUMBER = Utils.readExcelData(childSheet, j, i + 2);
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        wbs.close();
//        System.out.println(GX_LIST);
//        System.out.println(SB_LIST);
//        System.out.println(NUMBER);
    }

    private static void AAA() throws IOException, InvalidFormatException {
        List<List<String>> aaaArrayList = new ArrayList<>();
        Workbook wbs = Utils.getWorkbook();
        if (wbs == null) {
            return;
        }
        int cnt = wbs.getNumberOfSheets();
        for (int jj = 0; jj < cnt; jj++) {
            Sheet childSheet = wbs.getSheetAt(jj);
            String sheetName = childSheet.getSheetName();
            if ("目录".equals(sheetName)) {
                int rowNum = childSheet.getPhysicalNumberOfRows();
                int columnNum = childSheet.getRow(0).getPhysicalNumberOfCells();
                int startRom = 0;
                for (int i = 0; i < columnNum; i++) {
                    for (int j = 0; j < rowNum; j++) {
                        try {
                            if ("工序号".equals(Utils.readExcelData(childSheet, j, i))) {
                                startRom = j + 1;
                            }
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
                if (startRom > 0) {
                    for (int j = startRom; j < rowNum; j++) {
                        List<String> list = new ArrayList<>();
                        for (int i = 0; i < columnNum; i++) {
                            String s = null;
                            try {
                                s = Utils.readExcelData(childSheet, j, i);
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                            if (StrUtil.isEmpty(s)) {
                                try {
                                    s = Utils.readExcelDataGetNumericCellValue(childSheet, j, i);
                                } catch (Exception e) {
                                    e.printStackTrace();
                                }
                            }
                            if (!Utils.QUE_QING.equals(s)) {
                                list.add(s);
                            }
                        }
                        if (list.size() > 0) {
                            aaaArrayList.add(list);
                        }
                    }
                }
            }
        }

        int len = aaaArrayList.size();
        for (int i = 0; i < len; i++) {
            List<String> list1 = aaaArrayList.get(i);
            if (list1.size() == 3) {
                if ((i + 1) < len && aaaArrayList.get(i + 1).size() == 1) {
                    String s = list1.get(2);
                    for (int j = i + 1; j < len; j++) {
                        if (aaaArrayList.get(j).size() == 3) {
                            break;
                        }
                        if (aaaArrayList.get(j).size() == 1) {
                            s += aaaArrayList.get(j).get(0);
                        }
                    }
                    List<String> list11 = new ArrayList<>();
                    list11.add(list1.get(0));
                    list11.add(list1.get(1));
                    list11.add(s);
                    GX_LIST.add(list11);
                } else {
                    GX_LIST.add(list1);
                }
            }
        }
        wbs.close();
    }
}
