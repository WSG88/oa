package money;

import data.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Formatter;
import java.util.List;

public class Money {
    static Formatter formatter = new Formatter(System.out);

    public static void main(String[] args) throws Exception {
        Utils.clear();
        Utils.clearList();

//        Utils.FILE_NAME = "2020年工资表.xlsx";
//        Workbook wbs1 = Utils.getWorkbook();
//        Sheet childSheet1 = wbs1.getSheetAt(11);
//
//        Utils.FILE_NAME = "（江西）2021年工资表(1)(1).xlsx";
//        Workbook wbs2 = Utils.getWorkbook();
//        Sheet childSheet2 = wbs2.getSheetAt(0);
//
//        List<MoneyData> list1 = extracted(childSheet1);
//        List<MoneyData> list2 = extracted(childSheet2);
//
//        for (MoneyData moneyData1 : list1) {
//            for (MoneyData moneyData2 : list2) {
//                if (moneyData1.姓名.equals(moneyData2.姓名) && (moneyData2.应付工资小计 > moneyData1.应付工资小计)&&moneyData2.应付工资小计<20000) {
////                    formatter.format("%-6s %-10s %-10s %-10s\n", moneyData1.姓名 ,moneyData1.应付工资小计 , moneyData2.应付工资小计 , (moneyData2.应付工资小计 - moneyData1.应付工资小计));
////                    formatter.format("%-6s\n", moneyData1.姓名 , (moneyData2.应付工资小计 - moneyData1.应付工资小计));
////                    formatter.format("%-10s\n",  (moneyData2.应付工资小计-moneyData2.工龄工资 - (moneyData1.应付工资小计-moneyData1.工龄工资)));
//                    formatter.format("%-10s\n",  moneyData2.工龄工资 - moneyData1.工龄工资);
//                }
//            }
//        }


        Utils.FILE_NAME = "111111111111111.xlsx";
        Workbook wbs3 = Utils.getWorkbook();
        Sheet childSheet3 = wbs3.getSheetAt(0);
        List<MoneyData> list = new ArrayList<>();
        List<MoneyData> list1 = new ArrayList<>();
        List<MoneyData> list2 = new ArrayList<>();

        for (int i = 0; i < childSheet3.getLastRowNum() + 1; i++) {
            for (int j = 0; j < 10; j++) {
                String s0 = Utils.readExcel(childSheet3, i, 0);
                String s1 = Utils.readExcel(childSheet3, i, 1);
                String s2 = Utils.readExcel(childSheet3, i, 3);
                String s3 = Utils.readExcel(childSheet3, i, 5);
                if(s0.length()>0){
                MoneyData moneyData1 = new MoneyData();
                moneyData1.姓名=s0;
                moneyData1.应付工资小计=MoneyData.parseDouble(s1);
                list.add(moneyData1);

                MoneyData moneyData2= new MoneyData();
                moneyData2.姓名=s2;
                list1.add(moneyData2);

                MoneyData moneyData3 = new MoneyData();
                moneyData3.姓名=s3;
                list2.add(moneyData3);}
            }
        }
//        System.out.println(list);
//        System.out.println(list2);
        for (int i = 0; i < 10; i++) {
            MoneyData moneyData1 =list.get(0);
            System.out.println(moneyData1.姓名+moneyData1.应付工资小计);
            for (int j = 0; j <list1.size() ; j++) {
                MoneyData moneyData2 =list1.get(0);
                    System.out.println(moneyData2.姓名+moneyData2.应付工资小计);
                if(moneyData1.姓名.equals(moneyData2.姓名)){
                    moneyData2.应付工资小计= moneyData1.应付工资小计;
                }
            }
        }

    }

    private static List<MoneyData> extracted(Sheet childSheet) throws Exception {
        List<MoneyData> list = new ArrayList<>();
        MoneyData moneyData = new MoneyData();
//        for (int j = 0; j < 100; j++) {
//            String name = Utils.readExcel(childSheet, 2, j);
//        }
        for (int i = 3; i < childSheet.getLastRowNum() + 1; i++) {
            for (int j = 0; j < 1; j++) {
                moneyData = new MoneyData();
                moneyData.序号 = Utils.readExcel(childSheet, i, 0);
                moneyData.姓名 = Utils.readExcel(childSheet, i, 1);
                moneyData.应出勤天数 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 2));
                moneyData.实出勤天数 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 3));
                moneyData.基本工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 4));
                moneyData.技术工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 5));
                moneyData.满勤工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 6));
                moneyData.岗位津贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 7));
                moneyData.工龄工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 8));
                moneyData.应付工资小计 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 9));
                //补贴款项;
                moneyData.加班补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 10));


                moneyData.夜班补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 11));
                moneyData.操机补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 12));
                moneyData.绩效奖励 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 13));
                moneyData.住宿补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 14));
                moneyData.差旅补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 15));
                moneyData.高温补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 16));
                moneyData.劳保补贴 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 17));
                moneyData.补贴小计 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 18));
                //代扣款项;
                moneyData.缺勤工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 19));
                moneyData.养老险8 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 20));
                moneyData.医疗险2 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 21));
                moneyData.失业险 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 22));
                moneyData.个人所得税 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 23));
                moneyData.质量扣款 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 24));
                moneyData.绩效处罚 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 25));
                moneyData.电费 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 26));
                moneyData.代扣款小计 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 27));

                moneyData.实付工资 = MoneyData.parseDouble(Utils.readExcel(childSheet, i, 28));
                moneyData.备注 = Utils.readExcel(childSheet, i, 29);

//                System.out.println(moneyData);
//                System.out.println(moneyData.姓名 + " " + moneyData.应付工资小计);

//                //=J29/C29*(C29-D29)
//                if (moneyData.应出勤天数 > moneyData.实出勤天数) {
//                    double d = moneyData.应付工资小计 / moneyData.应出勤天数 * (moneyData.应出勤天数 - moneyData.实出勤天数);
//                    System.out.println("缺勤工资");
//                    System.out.println(moneyData.缺勤工资);
//                    System.out.println(d);
//                    System.out.println(moneyData);
//                }
//                //=(E30+F30+G30+H30)/C30*(D30-C30)*1.2
//                if (moneyData.应出勤天数 < moneyData.实出勤天数) {
//                    double d = (moneyData.基本工资 + moneyData.技术工资 + moneyData.满勤工资 + moneyData.岗位津贴) / moneyData.应出勤天数 * (moneyData.实出勤天数 - moneyData.应出勤天数) * 1.2;
//                    DecimalFormat df = new DecimalFormat("#.00");
//                    String s = df.format(d);
//                    d = Double.parseDouble(s);
//                    if (moneyData.加班补贴 - d != 0) {
//                        System.out.println("加班补贴");
//                        System.out.println(moneyData.加班补贴);
//                        System.out.println(d);
//                        System.out.println(moneyData);
//                    }
//                }
                list.add(moneyData);
            }
        }
        return list;
    }


}
