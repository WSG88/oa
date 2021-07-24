package money;

import cn.hutool.core.util.NumberUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static java.util.stream.Collectors.toList;

public class Money {

    public static void main(String[] args) throws Exception {
        String path1 = "E:\\人力资源部\\资料\\工资\\2021工资\\202105\\202105工资表(1).xlsx";
        List<MoneyData> dataArrayList1 = cal(path1);
        String path2 = "E:\\人力资源部\\资料\\工资\\2021工资\\202106\\202106工资表计算.xlsx";
        List<MoneyData> dataArrayList2 = cal(path2);
//        for (MoneyData moneyData : dataArrayList2) {
//            System.out.println(moneyData.姓名+"----------------------------------------------------------------------");
//        }

        for (int i = 0; i < dataArrayList1.size(); i++) {
            MoneyData moneyData1 = dataArrayList1.get(i);
            String name1 = moneyData1.姓名;
//            System.out.println(name1+"----------------------------------------------------------------------");
            String dept1 = moneyData1.部门;
            for (MoneyData moneyData2 : dataArrayList2) {
                if (moneyData2.姓名.equals(name1) && moneyData2.部门.equals(dept1)) {
                    double dd, dd1, dd2, dd3, dd4, dd5;

                    if ("技术部".equals(dept1)) {
                        dd5 = moneyData2.应付工资小计 + moneyData2.加班补贴 - moneyData1.应付工资小计 - moneyData1.加班补贴;
                    } else {
                        dd5 = moneyData2.应付工资小计 - moneyData1.应付工资小计;
                    }

                    double dddd = add(add(add(add(moneyData1.养老险8, moneyData1.医疗险2), moneyData1.大病医疗6), moneyData1.失业险), moneyData1.个人所得税);

                    dd = sub(moneyData2.加班补贴, moneyData1.加班补贴);
                    dd3 = sub(moneyData2.操机补贴, moneyData1.操机补贴);
                    dd4 = sub(moneyData2.夜班补贴, moneyData1.夜班补贴);
                    dd1 = sub(moneyData2.实出勤天数, moneyData1.实出勤天数);
                    dd2 = sub(moneyData2.实付工资, add(moneyData1.实付工资, dddd));
                    if (dd2 != 0) {
                        System.out.println(
                                getStringLen(name1) +
                                        getStringLen(" 工资差 " + dd5) +
                                        getStringLen(" 实付差 " + dd2) +
                                        getStringLen(" 操机差 " + dd3) +
                                        getStringLen(" 夜班差 " + dd4) +
                                        getStringLen(" 上班费差 " + dd) +
                                        getStringLen(" 上班时差 " + dd1)
                        );
                        System.out.println("------------------");
                    }

                }
            }
        }

        List<String> list1 = getList(dataArrayList1);
        List<String> list2 = getList(dataArrayList2);
        // 差集 (list1 - list2)
        List<String> reduce1 = list1.stream().filter(item -> !list2.contains(item)).collect(toList());
        System.out.println("---离职人员---");
        reduce1.parallelStream().forEach(System.out::println);

        // 差集 (list2 - list1)
        List<String> reduce2 = list2.stream().filter(item -> !list1.contains(item)).collect(toList());
        System.out.println("---入职人员---");
        reduce2.parallelStream().forEach(System.out::println);
    }

    private static String getStringLen(Object str) {
        StringBuilder ss = new StringBuilder(str.toString());
        if (ss.length() < 20) {
            for (int i = 0; i < 20 - ss.length(); i++) {
                ss.append(" ");
            }
            return ss.toString();
        } else {
            return ss.toString();
        }
    }

    private static List<String> getList(List<MoneyData> dataArrayList1) {
        Map<String, List<MoneyData>> map = dataArrayList1.stream().collect(Collectors.groupingBy(MoneyData::get姓名));
        Set<String> keySet = map.keySet();
        List<String> list = new ArrayList<>(keySet);
        return list;
    }

    public static List<MoneyData> cal(String path) throws Exception {
        Sheet childSheet = WorkbookFactory.create(new FileInputStream(path)).getSheetAt(0);
        double allDouble = 0;//总金额
        double dddd = 0;//代缴金额
        boolean jump = false;//结束标志
        List<MoneyData> dataArrayList = new ArrayList<>();
        MoneyData moneyData;
        for (int i = 3; i < childSheet.getLastRowNum() + 1; i++) {
            if (jump) {
                break;
            }
            for (int j = 0; j < 1; j++) {
                moneyData = new MoneyData();
                moneyData.序号 = readExcel(childSheet, i, 0);
                moneyData.部门 = readExcel(childSheet, i, 1);
                moneyData.姓名 = readExcel(childSheet, i, 2);
                if (moneyData.姓名 == null || moneyData.姓名.length() == 0 || moneyData.姓名.length() > 5 || moneyData.姓名.startsWith("姓")) {
                    break;
                }
                moneyData.应出勤天数 = readExcelDouble(childSheet, i, 3);
                moneyData.实出勤天数 = readExcelDouble(childSheet, i, 4);
                moneyData.基本工资 = readExcelDouble(childSheet, i, 5);
                moneyData.技术工资 = readExcelDouble(childSheet, i, 6);
                moneyData.满勤工资 = readExcelDouble(childSheet, i, 7);
                moneyData.岗位津贴 = readExcelDouble(childSheet, i, 8);
                moneyData.工龄工资 = readExcelDouble(childSheet, i, 9);
                moneyData.应付工资小计 = readExcelDouble(childSheet, i, 10);
                //补贴款项;
                moneyData.加班补贴 = readExcelDouble(childSheet, i, 11);
                moneyData.操机补贴 = readExcelDouble(childSheet, i, 12);
                moneyData.劳保补贴 = readExcelDouble(childSheet, i, 13);
                moneyData.住宿补贴 = readExcelDouble(childSheet, i, 14);
                moneyData.夜班补贴 = readExcelDouble(childSheet, i, 15);
                moneyData.高温补贴 = readExcelDouble(childSheet, i, 16);
                moneyData.差旅补贴 = readExcelDouble(childSheet, i, 17);
                moneyData.餐补 = readExcelDouble(childSheet, i, 18);
                moneyData.假补 = readExcelDouble(childSheet, i, 19);
                moneyData.绩效奖励 = readExcelDouble(childSheet, i, 20);
                moneyData.补贴小计 = readExcelDouble(childSheet, i, 21);


                //代扣款项;
                moneyData.缺勤工资 = readExcelDouble(childSheet, i, 22);
                moneyData.养老险8 = readExcelDouble(childSheet, i, 23);
                moneyData.医疗险2 = readExcelDouble(childSheet, i, 24);
                moneyData.失业险 = readExcelDouble(childSheet, i, 25);
                moneyData.大病医疗6 = readExcelDouble(childSheet, i, 26);
                moneyData.个人所得税 = readExcelDouble(childSheet, i, 27);
                moneyData.电费 = readExcelDouble(childSheet, i, 28);
                moneyData.质量扣款 = readExcelDouble(childSheet, i, 29);
                moneyData.餐费 = readExcelDouble(childSheet, i, 30);
                moneyData.绩效处罚 = readExcelDouble(childSheet, i, 31);
                moneyData.代扣款小计 = readExcelDouble(childSheet, i, 32);

                moneyData.实付工资 = readExcelDouble(childSheet, i, 33);
                moneyData.备注 = readExcel(childSheet, i, 34);


                allDouble = add(moneyData.实付工资, allDouble);
                dddd = add(add(add(add(add(moneyData.养老险8, moneyData.医疗险2), moneyData.大病医疗6), moneyData.失业险), moneyData.个人所得税), dddd);

//                System.out.println(moneyData);

//                System.out.println(moneyData.姓名 + " " + moneyData.应付工资小计 + " " + moneyData.补贴小计 + " " + moneyData.代扣款小计 + " " + moneyData.实付工资 + " " +
//                        sub(add(moneyData.应付工资小计, moneyData.补贴小计), add(moneyData.实付工资, moneyData.代扣款小计)));

//                //=J29/C29*(C29-D29)
                if (moneyData.应出勤天数 > moneyData.实出勤天数) {
                    double d;
                    if ("技术部".equals(moneyData.部门)) {
                        d = (moneyData.应付工资小计 + moneyData.加班补贴) / moneyData.应出勤天数 * (moneyData.应出勤天数 - moneyData.实出勤天数);
                    } else {
                        d = moneyData.应付工资小计 / moneyData.应出勤天数 * (moneyData.应出勤天数 - moneyData.实出勤天数);
                    }
                    double dd = sub(moneyData.缺勤工资, d);
                    if (dd != 0) {
                        System.out.println(moneyData.姓名 + moneyData.缺勤工资);
                        System.out.println("计算差 " + dd);
                        System.out.println("----------------");
                    }
                }
//                //=(E30+F30+G30+H30)/C30*(D30-C30)*1.2
                if (moneyData.应出勤天数 < moneyData.实出勤天数) {
                    double total = moneyData.基本工资 + moneyData.技术工资 + moneyData.满勤工资 + moneyData.岗位津贴;
                    total = moneyData.基本工资 + moneyData.技术工资;
                    double d = mul(mul(divide(total, moneyData.应出勤天数), sub(moneyData.实出勤天数, moneyData.应出勤天数)), 1.2);
                    double dd = sub(moneyData.加班补贴, d);
                    if (!"技术部".equals(moneyData.部门) && dd != 0) {
                        System.out.println(moneyData.姓名 + "  " + moneyData.加班补贴 + "  " + d);
                        System.out.println("计算差 " + dd);
                        System.out.println("----------------");
                    }
                }

                dataArrayList.add(moneyData);

                if ("★标黄处为社保全扣人员".equals(moneyData.姓名)) {
                    jump = true;
                }
            }
        }
        System.out.println(allDouble);
        System.out.println(dddd);
        return dataArrayList;
    }


    public static double readExcelDouble(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        double dou;
        BigDecimal bigDecimal;
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case NUMERIC:
                    case FORMULA:
                        dou = cell.getNumericCellValue();
                        bigDecimal = new BigDecimal(dou).setScale(4, RoundingMode.HALF_UP);
                        return bigDecimal.doubleValue();
                    case BLANK:
                        return 0.0;
                    case STRING:
                        String str = cell.getStringCellValue();
                        if (NumberUtil.isNumber(str)) {
                            dou = Double.parseDouble(str);
                            bigDecimal = new BigDecimal(dou).setScale(4, RoundingMode.HALF_UP);
                            return bigDecimal.doubleValue();
                        }
                    default:
                        throw new Exception();

                }
            }
        }
        return 0.0;
    }

    public static String readExcel(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue());
                    case FORMULA:
                        DecimalFormat df = new DecimalFormat("0.00");
                        return df.format(cell.getNumericCellValue());
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                    case STRING:
                        return cell.getStringCellValue();
                    default:
                        return "";
                }
                //System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + s);
            }
        }
        return "";
    }


    /**
     * 提供精确的加法运算。
     *
     * @param value1 被加数
     * @param value2 加数
     * @return 两个参数的和
     */
    public static Double add(Double value1, Double value2) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1));
        BigDecimal b2 = new BigDecimal(Double.toString(value2));
        return b1.add(b2).doubleValue();
    }

    /**
     * 提供精确的减法运算。
     *
     * @param value1 被减数
     * @param value2 减数
     * @return 两个参数的差
     */
    public static double sub(Double value1, Double value2) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1));
        BigDecimal b2 = new BigDecimal(Double.toString(value2));
        return b1.subtract(b2).doubleValue();
    }

    /**
     * 提供精确的乘法运算。
     *
     * @param value1 被乘数
     * @param value2 乘数
     * @return 两个参数的积
     */
    public static Double mul(Double value1, Double value2) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1));
        BigDecimal b2 = new BigDecimal(Double.toString(value2));
        return b1.multiply(b2).doubleValue();
    }

    // 默认除法运算精度
    private static final Integer DEF_DIV_SCALE = 4;

    /**
     * 57      * 提供（相对）精确的除法运算，当发生除不尽的情况时， 精确到小数点以后10位，以后的数字四舍五入。
     * 58      *
     * 59      * @param dividend 被除数
     * 60      * @param divisor  除数
     * 61      * @return 两个参数的商
     * 62
     */
    public static Double divide(Double dividend, Double divisor) {
        return divide(dividend, divisor, DEF_DIV_SCALE);
    }

    /**
     * 68      * 提供（相对）精确的除法运算。 当发生除不尽的情况时，由scale参数指定精度，以后的数字四舍五入。
     * 69      *
     * 70      * @param dividend 被除数
     * 71      * @param divisor  除数
     * 72      * @param scale    表示表示需要精确到小数点以后几位。
     * 73      * @return 两个参数的商
     * 74
     */
    public static Double divide(Double dividend, Double divisor, Integer scale) {
        if (scale < 0) {
            throw new IllegalArgumentException("The scale must be a positive integer or zero");
        }
        BigDecimal b1 = new BigDecimal(Double.toString(dividend));
        BigDecimal b2 = new BigDecimal(Double.toString(divisor));
        return b1.divide(b2, scale, RoundingMode.HALF_UP).doubleValue();
    }

    /**
     * 提供指定数值的（精确）小数位四舍五入处理。
     * 86      *
     * 87      * @param value 需要四舍五入的数字
     * 88      * @param scale 小数点后保留几位
     * 89      * @return 四舍五入后的结果
     * 90
     */
    public static double round(double value, int scale) {
        if (scale < 0) {
            throw new IllegalArgumentException("The scale must be a positive integer or zero");
        }
        BigDecimal b = new BigDecimal(Double.toString(value));
        BigDecimal one = new BigDecimal("1");
        return b.divide(one, scale, RoundingMode.HALF_UP).doubleValue();
    }

    static class MoneyData {

        public String 序号;
        public String 部门;
        public String 姓名;
        public double 应出勤天数;
        public double 实出勤天数;
        public double 基本工资;
        public double 技术工资;
        public double 满勤工资;
        public double 岗位津贴;
        public double 工龄工资;
        public double 应付工资小计;
        //补贴款项;
        public double 加班补贴;
        public double 操机补贴;
        public double 劳保补贴;
        public double 住宿补贴;
        public double 夜班补贴;
        public double 高温补贴;
        public double 差旅补贴;
        public double 餐补;
        public double 假补;
        public double 绩效奖励;
        public double 补贴小计;
        //代扣款项;
        public double 缺勤工资;
        public double 养老险8;
        public double 医疗险2;
        public double 失业险;
        public double 大病医疗6;
        public double 个人所得税;
        public double 电费;
        public double 质量扣款;
        public double 餐费;
        public double 绩效处罚;
        public double 代扣款小计;

        public double 实付工资;
        public String 备注;

        public String get序号() {
            return 序号;
        }

        public void set序号(String 序号) {
            this.序号 = 序号;
        }

        public String get部门() {
            return 部门;
        }

        public void set部门(String 部门) {
            this.部门 = 部门;
        }

        public String get姓名() {
            return 姓名;
        }

        public void set姓名(String 姓名) {
            this.姓名 = 姓名;
        }

        public double get应出勤天数() {
            return 应出勤天数;
        }

        public void set应出勤天数(double 应出勤天数) {
            this.应出勤天数 = 应出勤天数;
        }

        public double get实出勤天数() {
            return 实出勤天数;
        }

        public void set实出勤天数(double 实出勤天数) {
            this.实出勤天数 = 实出勤天数;
        }

        public double get基本工资() {
            return 基本工资;
        }

        public void set基本工资(double 基本工资) {
            this.基本工资 = 基本工资;
        }

        public double get技术工资() {
            return 技术工资;
        }

        public void set技术工资(double 技术工资) {
            this.技术工资 = 技术工资;
        }

        public double get满勤工资() {
            return 满勤工资;
        }

        public void set满勤工资(double 满勤工资) {
            this.满勤工资 = 满勤工资;
        }

        public double get岗位津贴() {
            return 岗位津贴;
        }

        public void set岗位津贴(double 岗位津贴) {
            this.岗位津贴 = 岗位津贴;
        }

        public double get工龄工资() {
            return 工龄工资;
        }

        public void set工龄工资(double 工龄工资) {
            this.工龄工资 = 工龄工资;
        }

        public double get应付工资小计() {
            return 应付工资小计;
        }

        public void set应付工资小计(double 应付工资小计) {
            this.应付工资小计 = 应付工资小计;
        }

        public double get加班补贴() {
            return 加班补贴;
        }

        public void set加班补贴(double 加班补贴) {
            this.加班补贴 = 加班补贴;
        }

        public double get操机补贴() {
            return 操机补贴;
        }

        public void set操机补贴(double 操机补贴) {
            this.操机补贴 = 操机补贴;
        }

        public double get劳保补贴() {
            return 劳保补贴;
        }

        public void set劳保补贴(double 劳保补贴) {
            this.劳保补贴 = 劳保补贴;
        }

        public double get住宿补贴() {
            return 住宿补贴;
        }

        public void set住宿补贴(double 住宿补贴) {
            this.住宿补贴 = 住宿补贴;
        }

        public double get夜班补贴() {
            return 夜班补贴;
        }

        public void set夜班补贴(double 夜班补贴) {
            this.夜班补贴 = 夜班补贴;
        }

        public double get高温补贴() {
            return 高温补贴;
        }

        public void set高温补贴(double 高温补贴) {
            this.高温补贴 = 高温补贴;
        }

        public double get差旅补贴() {
            return 差旅补贴;
        }

        public void set差旅补贴(double 差旅补贴) {
            this.差旅补贴 = 差旅补贴;
        }

        public double get餐补() {
            return 餐补;
        }

        public void set餐补(double 餐补) {
            this.餐补 = 餐补;
        }

        public double get假补() {
            return 假补;
        }

        public void set假补(double 假补) {
            this.假补 = 假补;
        }

        public double get绩效奖励() {
            return 绩效奖励;
        }

        public void set绩效奖励(double 绩效奖励) {
            this.绩效奖励 = 绩效奖励;
        }

        public double get补贴小计() {
            return 补贴小计;
        }

        public void set补贴小计(double 补贴小计) {
            this.补贴小计 = 补贴小计;
        }

        public double get缺勤工资() {
            return 缺勤工资;
        }

        public void set缺勤工资(double 缺勤工资) {
            this.缺勤工资 = 缺勤工资;
        }

        public double get养老险8() {
            return 养老险8;
        }

        public void set养老险8(double 养老险8) {
            this.养老险8 = 养老险8;
        }

        public double get医疗险2() {
            return 医疗险2;
        }

        public void set医疗险2(double 医疗险2) {
            this.医疗险2 = 医疗险2;
        }

        public double get失业险() {
            return 失业险;
        }

        public void set失业险(double 失业险) {
            this.失业险 = 失业险;
        }

        public double get大病医疗6() {
            return 大病医疗6;
        }

        public void set大病医疗6(double 大病医疗6) {
            this.大病医疗6 = 大病医疗6;
        }

        public double get个人所得税() {
            return 个人所得税;
        }

        public void set个人所得税(double 个人所得税) {
            this.个人所得税 = 个人所得税;
        }

        public double get电费() {
            return 电费;
        }

        public void set电费(double 电费) {
            this.电费 = 电费;
        }

        public double get质量扣款() {
            return 质量扣款;
        }

        public void set质量扣款(double 质量扣款) {
            this.质量扣款 = 质量扣款;
        }

        public double get餐费() {
            return 餐费;
        }

        public void set餐费(double 餐费) {
            this.餐费 = 餐费;
        }

        public double get绩效处罚() {
            return 绩效处罚;
        }

        public void set绩效处罚(double 绩效处罚) {
            this.绩效处罚 = 绩效处罚;
        }

        public double get代扣款小计() {
            return 代扣款小计;
        }

        public void set代扣款小计(double 代扣款小计) {
            this.代扣款小计 = 代扣款小计;
        }

        public double get实付工资() {
            return 实付工资;
        }

        public void set实付工资(double 实付工资) {
            this.实付工资 = 实付工资;
        }

        public String get备注() {
            return 备注;
        }

        public void set备注(String 备注) {
            this.备注 = 备注;
        }

        @Override
        public String toString() {
            return "MoneyData{" +
                    "序号='" + 序号 + '\'' +
                    ", 部门='" + 部门 + '\'' +
                    ", 姓名='" + 姓名 + '\'' +
                    ", 应出勤天数=" + 应出勤天数 +
                    ", 实出勤天数=" + 实出勤天数 +
                    ", 基本工资=" + 基本工资 +
                    ", 技术工资=" + 技术工资 +
                    ", 满勤工资=" + 满勤工资 +
                    ", 岗位津贴=" + 岗位津贴 +
                    ", 工龄工资=" + 工龄工资 +
                    ", 应付工资小计=" + 应付工资小计 +
                    ", 加班补贴=" + 加班补贴 +
                    ", 操机补贴=" + 操机补贴 +
                    ", 劳保补贴=" + 劳保补贴 +
                    ", 住宿补贴=" + 住宿补贴 +
                    ", 夜班补贴=" + 夜班补贴 +
                    ", 高温补贴=" + 高温补贴 +
                    ", 差旅补贴=" + 差旅补贴 +
                    ", 餐补=" + 餐补 +
                    ", 假补=" + 假补 +
                    ", 绩效奖励=" + 绩效奖励 +
                    ", 补贴小计=" + 补贴小计 +
                    ", 缺勤工资=" + 缺勤工资 +
                    ", 养老险8=" + 养老险8 +
                    ", 医疗险2=" + 医疗险2 +
                    ", 失业险=" + 失业险 +
                    ", 大病医疗6=" + 大病医疗6 +
                    ", 个人所得税=" + 个人所得税 +
                    ", 电费=" + 电费 +
                    ", 质量扣款=" + 质量扣款 +
                    ", 餐费=" + 餐费 +
                    ", 绩效处罚=" + 绩效处罚 +
                    ", 代扣款小计=" + 代扣款小计 +
                    ", 实付工资=" + 实付工资 +
                    ", 备注='" + 备注 + '\'' +
                    '}';
        }

    }
}
