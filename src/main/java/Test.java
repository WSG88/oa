import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import data.ExcelUtils;

import java.util.*;


public class Test {
    public static void main(String[] args) throws Exception {
        Map<String, Object> row1 = new LinkedHashMap<>();
        row1.put("姓名", "张三");
        row1.put("年龄", 23);
        row1.put("成绩", 88.32);
        row1.put("是否合格", true);
        row1.put("考试日期", DateUtil.date());

        Map<String, Object> row2 = new LinkedHashMap<>();
        row2.put("姓名", "李四");
        row2.put("年龄", 33);
        row2.put("成绩", 59.50);
        row2.put("是否合格", false);
        row2.put("考试日期", DateUtil.date());

        ArrayList<Map<String, Object>> rows = CollUtil.newArrayList(row1, row2);

        ExcelUtils.writeExcel("d:/writeMapTest.xlsx", "dd", new ArrayList<String>(), rows);


    }


    public static void main1(String[] args) {
//        double time = 24 * 60 * 60;
//        System.out.println((8 * 60 * 60 + 5 * 60) / time);
//        System.out.println((8 * 60 * 60 + 30 * 60) / time);
//        System.out.println((11 * 60 * 60 + 25 * 60) / time);
//        System.out.println((11 * 60 * 60 + 55 * 60) / time);
//        System.out.println((13 * 60 * 60 + 35 * 60) / time);
//        System.out.println((16 * 60 * 60 + 25 * 60) / time);
//        System.out.println((17 * 60 * 60) / time);
//        System.out.println((17 * 60 * 60 + 25 * 60) / time);
//        System.out.println((18 * 60 * 60 + 30 * 60) / time);
        String filePath = "C:/Work/oa/file/1.7.xls";
        List<String> names = new ArrayList<>();
        List<List<String>> times = new ArrayList<>();
        List<String> timess = new ArrayList<>();
        boolean isName = false;
        boolean isTime = false;
//        ExcelReader reader = ExcelUtil.getReader(filePath);
//通过sheet编号获取
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(filePath), 3);
//通过sheet名获取
//        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(filePath), "sheet1");
        int line = 0;
        List<List<Object>> readAll = reader.read();
        for (int i = 0; i < readAll.size(); i++) {
            List<Object> list = readAll.get(i);
            for (int j = 0; j < list.size(); j++) {
                Object object = list.get(j);
                if (isName) {
                    names.add(object.toString());
                }
                if ("姓名".equals(object)) {
                    isName = true;
                } else {
                    isName = false;
                }

                if (object == null || "".equals(object)) {
                } else {
                    String text = object.toString();
                    if (text.startsWith("01 三")) {
                        isTime = true;
                    }
                }
                if (isTime) {
                    if (object == null || "".equals(object)) {
                        timess.add(object.toString());
                    } else {
                        String text = object.toString();
                        if (text.length() > 15) {
                            timess.add(text.substring((text.length() - " 1899-12-31 ".length() + 4), text.length()));
                        } else {
                            timess.add(text);
                        }
                    }
                }
            }
        }

        List<List<String>> result = averageAssign(timess, 15);
        for (int i = 0; i < result.size(); i++) {
            List<String> l = result.get(i);
            for (int i1 = 0; i1 < l.size(); i1++) {
                System.out.printf("%s%s", l.get(i1), "    ");
            }
            System.out.println();
        }


    }

    /**
     * 将一个List均分成n个list,主要通过偏移量来实现的
     *
     * @param source 源集合
     * @param limit  最大值
     * @return
     */
    public static List<List<String>> averageAssign(List<String> source, int limit) {
        if (null == source || source.isEmpty()) {
            return Collections.emptyList();
        }
        List<List<java.lang.String>> result = new ArrayList<>();
        int listCount = (source.size() - 1) / limit + 1;
        int remaider = source.size() % listCount; // (先计算出余数)
        int number = source.size() / listCount; // 然后是商
        int offset = 0;// 偏移量
        for (int i = 0; i < listCount; i++) {
            List<java.lang.String> value;
            if (remaider > 0) {
                value = source.subList(i * number + offset, (i + 1) * number + offset + 1);
                remaider--;
                offset++;
            } else {
                value = source.subList(i * number + offset, (i + 1) * number + offset);
            }
            result.add(value);
        }
        return result;
    }

}
