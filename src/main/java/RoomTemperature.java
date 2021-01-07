import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

//温度流水记录
public class RoomTemperature {

    public static void main(String[] args) throws Exception {

        String fileName = "20210107";
        String fileEx = ".xlsx";
        try {
            InputStream in = new FileInputStream(Utils.FILE_PATH + fileName + fileEx);
            Workbook wbs = WorkbookFactory.create(in);
            Sheet childSheet = wbs.getSheetAt(0);
            Map<String, ArrayList<Tem>> map = new HashMap<>();
            HashSet<String> hashSet = new HashSet<>();
            for (int line = 1; line < childSheet.getLastRowNum() + 1; line++) {
                String name = Utils.readExcelData(childSheet, line, 5);
                String event = Utils.readExcelData(childSheet, line, 4);
//                if ("人脸认证失败".equals(event)) {
//                    System.out.println(Utils.readExcelData(childSheet, line, 0));
//                }
                if (name != null && name.length() > 0) {
                    Tem tem = new Tem();
                    tem.抓拍图片路径 = Utils.readExcelData(childSheet, line, 0);
                    tem.热图保存路径 = Utils.readExcelData(childSheet, line, 1);
                    tem.可见光图片保存路径 = Utils.readExcelData(childSheet, line, 2);
                    tem.序号 = Utils.readExcelData(childSheet, line, 3);
                    tem.事件类型 = Utils.readExcelData(childSheet, line, 4);
                    tem.持卡人员 = Utils.readExcelData(childSheet, line, 5);
                    tem.卡号 = Utils.readExcelData(childSheet, line, 6);
                    tem.温度 = Utils.readExcelData(childSheet, line, 7);
                    tem.温度异常 = Utils.readExcelData(childSheet, line, 8);
                    tem.事件时间 = Utils.readExcelData(childSheet, line, 9);
                    tem.佩戴口罩 = Utils.readExcelData(childSheet, line, 10);
                    tem.佩戴安全帽 = Utils.readExcelData(childSheet, line, 11);
                    tem.设备名称 = Utils.readExcelData(childSheet, line, 12);
                    tem.事件源 = Utils.readExcelData(childSheet, line, 13);
                    tem.方向 = Utils.readExcelData(childSheet, line, 14);
                    tem.物理地址 = Utils.readExcelData(childSheet, line, 15);
                    tem.认证方式 = Utils.readExcelData(childSheet, line, 16);
                    tem.卡类型 = Utils.readExcelData(childSheet, line, 17);
                    tem.读卡器型号 = Utils.readExcelData(childSheet, line, 18);
                    tem.事件级别 = Utils.readExcelData(childSheet, line, 19);
                    tem.状态 = Utils.readExcelData(childSheet, line, 20);
                    if ("人脸认证通过".equals(event)) {
                        String key = name + "," + tem.事件时间.substring(0, 10);
                        hashSet.add(key);
                        ArrayList<Tem> arrayList = map.computeIfAbsent(key, k -> new ArrayList<>());
                        arrayList.add(tem);
                        map.put(key, arrayList);
                    }
                }
            }
//            System.out.println(hashSet);
//            System.out.println(map);
//            System.out.println(map.entrySet().size());
            List<String> list0 = new ArrayList<>();
            list0.add("持卡人员");
            list0.add("事件时间");
            list0.add("温度");
            list0.add("温度异常");
            List<List<String>> rowsList1 = new ArrayList<>();
            rowsList1.add(list0);
            List<List<String>> rowsList2 = new ArrayList<>();
            rowsList2.add(list0);
            for (String s : hashSet) {
                ArrayList<Tem> list = map.computeIfAbsent(s, k -> new ArrayList<>());
                for (Tem tem1 : list) {
                    List<String> list1 = new ArrayList<>();
                    list1.add(tem1.持卡人员);
                    list1.add(tem1.事件时间);
                    list1.add(tem1.温度);
                    list1.add(tem1.温度异常);
                    rowsList1.add(list1);
                }
                Collections.sort(list, new SortByTem());
                Tem tem2 = list.get(0);
                List<String> list2 = new ArrayList<>();
                list2.add(tem2.持卡人员);
                list2.add(tem2.事件时间);
                list2.add(tem2.温度);
                list2.add(tem2.温度异常);
                rowsList2.add(list2);
            }
            Utils.toExcel(rowsList1, Utils.FILE_PATH + fileName + "_体温记录流水_" + System.currentTimeMillis() + fileEx);
            Utils.toExcel(rowsList2, Utils.FILE_PATH + fileName + "_体温记录汇总_" + System.currentTimeMillis() + fileEx);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static class SortByTem implements Comparator {
        public int compare(Object o1, Object o2) {
            Tem s1 = (Tem) o1;
            Tem s2 = (Tem) o2;
            if (Double.parseDouble(s1.温度.replace("℃", "")) > Double.parseDouble(s2.温度.replace("℃", ""))) {
                return -1;
            }
            return 1;
        }
    }

    static class Tem {

        public Tem() {
        }

        public Tem(String 抓拍图片路径, String 热图保存路径, String 可见光图片保存路径, String 序号, String 事件类型, String 持卡人员, String 卡号, String 温度, String 温度异常, String 事件时间, String 佩戴口罩, String 佩戴安全帽, String 设备名称, String 事件源, String 方向, String 物理地址, String 认证方式, String 卡类型, String 读卡器型号, String 事件级别, String 状态) {
            this.抓拍图片路径 = 抓拍图片路径;
            this.热图保存路径 = 热图保存路径;
            this.可见光图片保存路径 = 可见光图片保存路径;
            this.序号 = 序号;
            this.事件类型 = 事件类型;
            this.持卡人员 = 持卡人员;
            this.卡号 = 卡号;
            this.温度 = 温度;
            this.温度异常 = 温度异常;
            this.事件时间 = 事件时间;
            this.佩戴口罩 = 佩戴口罩;
            this.佩戴安全帽 = 佩戴安全帽;
            this.设备名称 = 设备名称;
            this.事件源 = 事件源;
            this.方向 = 方向;
            this.物理地址 = 物理地址;
            this.认证方式 = 认证方式;
            this.卡类型 = 卡类型;
            this.读卡器型号 = 读卡器型号;
            this.事件级别 = 事件级别;
            this.状态 = 状态;
        }

        /**
         * 抓拍图片路径 : C:/Users/Administrator/Desktop/20210107-2/1_20210107090008_1.jpg
         * 热图保存路径 : C:/Users/Administrator/Desktop/20210107-2/EF41ABA9C69C45E680969AE29D4BAFD6_HeatImage.png
         * 可见光图片保存路径 :
         * 序号 : 1
         * 事件类型 : 人脸认证失败
         * 持卡人员 :
         * 卡号 :
         * 温度 : 36.3℃
         * 温度异常 : 否
         * 事件时间 : 2021-01-07 09:00:08
         * 佩戴口罩 : 否
         * 佩戴安全帽 : 未知
         * 设备名称 : 体温
         * 事件源 : 进门读卡器1
         * 方向 : 进入
         * 物理地址 :
         * 认证方式 : 刷卡或人脸
         * 卡类型 : 普通卡
         * 读卡器型号 : 无效
         * 事件级别 : 未归类
         * 状态 : 未处理
         */

        public String 抓拍图片路径;
        public String 热图保存路径;
        public String 可见光图片保存路径;
        public String 序号;
        public String 事件类型;
        public String 持卡人员;
        public String 卡号;
        public String 温度;
        public String 温度异常;
        public String 事件时间;
        public String 佩戴口罩;
        public String 佩戴安全帽;
        public String 设备名称;
        public String 事件源;
        public String 方向;
        public String 物理地址;
        public String 认证方式;
        public String 卡类型;
        public String 读卡器型号;
        public String 事件级别;
        public String 状态;

        @Override
        public String toString() {
            return "\n{" +
                    " 持卡人员='" + 持卡人员 + '\'' +
//                ",抓拍图片路径='" + 抓拍图片路径 + '\'' +
//                ", 热图保存路径='" + 热图保存路径 + '\'' +
//                ", 可见光图片保存路径='" + 可见光图片保存路径 + '\'' +
//                ", 序号='" + 序号 + '\'' +
//                ", 卡号='" + 卡号 + '\'' +
                    ", 事件时间='" + 事件时间 + '\'' +
                    ", 温度='" + 温度 + '\'' +
                    ", 温度异常='" + 温度异常 + '\'' +
                    ", 事件类型='" + 事件类型 + '\'' +
//                ", 佩戴口罩='" + 佩戴口罩 + '\'' +
//                ", 佩戴安全帽='" + 佩戴安全帽 + '\'' +
//                ", 设备名称='" + 设备名称 + '\'' +
//                ", 事件源='" + 事件源 + '\'' +
//                ", 方向='" + 方向 + '\'' +
//                ", 物理地址='" + 物理地址 + '\'' +
//                ", 认证方式='" + 认证方式 + '\'' +
//                ", 卡类型='" + 卡类型 + '\'' +
//                ", 读卡器型号='" + 读卡器型号 + '\'' +
//                ", 事件级别='" + 事件级别 + '\'' +
//                ", 状态='" + 状态 + '\'' +
                    '}';
        }
    }
}
