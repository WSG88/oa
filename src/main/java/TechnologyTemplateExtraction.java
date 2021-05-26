import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import data.Utils;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class TechnologyTemplateExtraction {
    public static List<File> PATH_LIST = new ArrayList<>();
    public static List<List<String>> SAVE_LIST = new ArrayList<>();
    public static List<String> NUMBER_LIST = new ArrayList<>();

    public static String PATH_NAME = "";

    public static void main(String[] args) throws Exception {
        PATH_NAME = "239指令\\";
        PATH_NAME = "550厂工艺指令\\";
        PATH_NAME = "602指令\\";
        PATH_NAME = "安达维尔\\";
        PATH_NAME = "昌飞指令\\";
        PATH_NAME = "常发指令\\";
        PATH_NAME = "九江船舶\\";
        PATH_NAME = "郑飞指令\\";
        PATH_NAME = "错误\\";
        PATH_NAME = "";

//        File[] files = FileUtil.ls("D:\\ERP&MES\\天一指令\\" + PATH_NAME);
        File[] files = FileUtil.ls("E:\\人力资源部\\00信息办\\天一指令\\" + PATH_NAME);
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
                try {
                    AAA();
                } catch (Exception e) {
                    e.printStackTrace();
                }
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
        ExcelWriter writer = ExcelUtil.getWriter("E:\\人力资源部\\00信息办\\天一指令\\" + "_工艺_" + System.currentTimeMillis() + ".xls");
        writer.write(SAVE_LIST, true);
        writer.close();
        System.out.println(NUMBER_LIST.size() + "_" + SAVE_LIST.size());
    }

    private static void AAA() throws Exception {
        Workbook wbs = Utils.getWorkbook();
        if (wbs == null) {
            return;
        }

        String NUMBER = "";
        List<List<String>> GX_LIST = new ArrayList<>();
        List<List<String>> SB_LIST = new ArrayList<>();
        List<List<String>> aaaArrayList = new ArrayList<>();

        int cnt = wbs.getNumberOfSheets();
        for (int jj = 0; jj < cnt; jj++) {
            Sheet childSheet = wbs.getSheetAt(jj);
            String sheetName = childSheet.getSheetName();
            int rowNum = childSheet.getPhysicalNumberOfRows();
            int columnNum = childSheet.getRow(0).getPhysicalNumberOfCells();

            for (int i = 0; i < columnNum; i++) {
                for (int j = 0; j < rowNum; j++) {
                    if ("机床设备".equals(Utils.readExcel(childSheet, j, i))) {
                        for (int k = j + 1; k < rowNum; k++) {
                            String ss = Utils.readExcel(childSheet, k, i);
                            if (!Utils.QUE_QING.equals(ss) && ss.length() > 0) {
                                String ss1 = Utils.readExcel(childSheet, k, 1);
                                List<String> list = new ArrayList<>();
                                list.add(ss1);
                                list.add(ss);
                                SB_LIST.add(list);
                            }
                        }
                    }
                    if ("零 件 图 号".equals(Utils.readExcel(childSheet, j, i))) {
                        NUMBER = Utils.readExcel(childSheet, j, i + 2);
                    }
                }
            }

            if ("目录".equals(sheetName)) {
                int startRom = 0;
                for (int i = 0; i < columnNum; i++) {
                    for (int j = 0; j < rowNum; j++) {
                        if ("工序号".equals(Utils.readExcel(childSheet, j, i))) {
                            startRom = j + 1;
                        }
                    }
                }
                if (startRom > 0) {
                    for (int j = startRom; j < rowNum; j++) {
                        List<String> list = new ArrayList<>();
                        for (int i = 0; i < columnNum; i++) {
                            String s = Utils.readExcel(childSheet, j, i);
                            if (!Utils.QUE_QING.equals(s) && s.length() > 0) {
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

//        System.out.println(NUMBER);
//        System.out.println(GX_LIST);
//        System.out.println(SB_LIST);

        for (List<String> list : GX_LIST) {
            //工序号
            String a = list.get(0);
            //
            String a1 = null;
            try {
                a1 = String.valueOf(Double.parseDouble(a));
            } catch (NumberFormatException e) {
                a1=a;
            }
            //工序名称
            String b = list.get(1);
            //工序内容
            String c = list.get(2);
            //加工设备
            String d = "";
            //准备时间
            String e = "60";
            //加工时间
            String f = "60";
            //是否MES管控
            String g = "是";
            if (b.contains("热处理") || b.contains("表面处理")) {
                g = "否";
            }
            for (List<String> strings : SB_LIST) {
                if (a.equals(strings.get(0))) {
                    d = strings.get(1);
                }
                if (b.equals(strings.get(0))) {
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
            NUMBER_LIST.add(NUMBER);
        }

    }
}
