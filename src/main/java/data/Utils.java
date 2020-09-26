package data;

import cn.hutool.core.date.DateUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Utils {

    public static String DATABASE_NAME_2 = "oatime";
    public static String FILE_PATH = "C:\\Work\\oa\\file\\";
    public static String YEAR_MONTH = "202008";
    public static String FILE_NAME = "2.81.xls";
    public static int ROOM = 2;
    public static List<String> arrayNamesList = new ArrayList<>();

    /*名字汇总*/
    public static void clear() {
        arrayNamesList.clear();
    }

    public static void add(String name) {
        if (isAdd(name)) {
            arrayNamesList.add(name);
        }
    }

    public static boolean isAdd(String name) {
        if (name == null
                || "".equals(name)
                || "正常".equals(name)
                || "陈卫平".equals(name)
                || "何仁易".equals(name)
                || "王扬威".equals(name)
                || "陈鹏".equals(name)
                || "张冬生".equals(name)
                || "糜火峰".equals(name)
                || "周华栋".equals(name)
                || "汪有国".equals(name)
                || "洪志超".equals(name)
                || "程泉华".equals(name)
                || "聂玉光".equals(name)
                || "陈秋水".equals(name)
                || "陈卫峰".equals(name)
                || name.length() < 2) {
            return false;
        }
        return true;
    }

    public static void main(String[] args) throws Exception {

//        for (int j = 0; j < 23; j++) {
//            String s = j + ":";
//            if (j < 10) {
//                s = "0" + j + ":";
//            }
//            for (int i = 0; i < 61; i++) {
//                String ss = "";
//                if (i < 10) {
//                    ss = s + "0" + i;
//                } else {
//                    ss = s + i;
//                }
////                System.out.println(ss + "   " + getFirstTime("20200801", ss));
//            }
//        }
//        System.out.println(Math.floor(5.5));
//        System.out.println(m(-0.4f));

        String year = "202006";
        System.out.println(getDaysOfMonth(year));

    }

    //每月天数
    public static List<String> getDaysOfMonth(String year) {
        List<String> list = new ArrayList<>();
        try {
            Date date = new SimpleDateFormat("yyyyMM").parse(year);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);
            int days = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
            for (int i = 1; i < days + 1; i++) {
                String newString = String.format("%02d", i);
                list.add(year + newString);
            }
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return list;
    }


    public static void toExcel(List<List<String>> rows, String title, String path) {
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(path);
        // 合并单元格后的标题行，使用默认标题样式
//        writer.merge(rows.size() - 1, title);
        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    public static void toExcel(ArrayList<Map<String, Object>> rows, String title, String path) {
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(path);
        // 合并单元格后的标题行，使用默认标题样式
//        writer.merge(rows.size() - 1, title);
        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    public static float getDecimals(float f) {
        return Float.parseFloat(new DecimalFormat(".0").format(f));
    }

    /*时间间隔*/
    public static float timeDifference(String d1, String d2) {
        if (d1.contains(QUE_QING) || d2.contains(QUE_QING)) {
            return 0F;
        }
        long l1 = DateUtil.parse(d1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis() / 1000;
        long l2 = DateUtil.parse(d2, "yyyyMMddHH:mm").toCalendar().getTimeInMillis() / 1000;
        float f = (l2 - l1) / 3600F;
        return Float.parseFloat(new DecimalFormat(".00").format(f));
    }

    /*时间取值,允许6分钟，其他为半小时向上取整*/
    public static String getCompleteTime(String time) {
        String outTime = "00:00";
        StringTokenizer st = new StringTokenizer(time, ":");
        List<String> inTime = new ArrayList<String>();
        while (st.hasMoreElements()) {
            inTime.add(st.nextToken());
        }
        String hour = inTime.get(0).toString();
        String minutes = inTime.get(1).toString();
        if (Integer.parseInt(minutes) > SIX_T) {
            hour = (Integer.parseInt(hour) + 1) + "";
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) < SIX + 1) {
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) != 0) {
            outTime = hour + ":30";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");

            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return outTime;
    }

    /*时间取值,半小时向下取整*/
    public static String getLastCompleteTime(String time) {
        if (QUE_QING.equals(time)) {
            return time;
        }
        String outTime = "00:00";
        StringTokenizer st = new StringTokenizer(time, ":");
        List<String> inTime = new ArrayList<String>();
        while (st.hasMoreElements()) {
            inTime.add(st.nextToken());
        }
        String hour = inTime.get(0).toString();
        String minutes = inTime.get(1).toString();
        if (Integer.parseInt(minutes) >= 55) {
            hour = (Integer.parseInt(hour) + 1) + "";
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) < 25) {
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) < 55 && Integer.parseInt(minutes) >= 25) {
            outTime = hour + ":30";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");

            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return outTime;
    }

    public static int SIX = 6;
    public static int SIX_T = 36;
    public static String FIRST_TIME = "08:00";
    public static String FIRST_TIME_PRE = "08:0" + SIX;
    public static String QUE_QING = "缺勤";


    public static String calculateTime(String string) {
        try {
            double d0 = Double.parseDouble(string);
            double dd = d0 * 24;
            int hour = (int) dd;
            double dou = (dd - hour) * 60;
            int minuter = (int) dou;
            if (dou > minuter) {
                minuter = minuter + 1;
            }
            StringBuilder stringBuilder = new StringBuilder();
            if (hour < 10) {
                stringBuilder.append("0");
            }
            stringBuilder.append(hour);
            stringBuilder.append(":");
            if (minuter < 10) {
                stringBuilder.append("0");
            }
            stringBuilder.append(minuter);
            return stringBuilder.toString();
        } catch (Exception e) {
            return string;
        }
    }

    public static Workbook getWorkbook() throws IOException, InvalidFormatException {
        InputStream in = new FileInputStream(Utils.FILE_PATH + Utils.FILE_NAME);
        return WorkbookFactory.create(in);
    }

    public static Sheet getSheet(Workbook wbs, String sheetName, int sheetIndex) {
        Sheet childSheet;
        if (sheetName == null) {
            childSheet = wbs.getSheetAt(sheetIndex);
        } else {
            childSheet = wbs.getSheet(sheetName);
        }
        return childSheet;
    }

    public static String readExcelDataGetNumericCellValue(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
//                    System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + String.valueOf(cell.getNumericCellValue()));
                    return String.valueOf(cell.getNumericCellValue());
                }
            }
        }
        return QUE_QING;
    }

    public static String readExcelData(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    return "";
                } else {
//                    System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + cell.getStringCellValue());
                    return cell.getStringCellValue();
                }
            }
        }
        return "";
    }

    public static int getQueQing(List<String> list) {
        if (list == null) {
            return 0;
        }
        for (String string : list) {
            if (QUE_QING.equals(string)) {
                list.add(QUE_QING);
            }
        }
        return list.size();
    }

    public static String getFirstTime(String day, String time) {
        if (QUE_QING.equals(time)) {
            return time;
        }
        String dd = day + time;
        long l1 = DateUtil.parse(dd, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(day + FIRST_TIME_PRE, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();

        if (l1 > l2) {
            return getCompleteTime(time);
        } else {
            return FIRST_TIME;
        }
    }

    public static void completeQueQing(List<String> list) {
        int length = list.size();
        switch (length) {
            case 0:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 1:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 2:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 3:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 4:
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 5:
                list.add(QUE_QING);
                break;
        }
    }

    //检查考勤数据是否完整
    public static void checkData(Data data, int room) {
        String name = data.name;
        String date = data.date;
        List<String> list = data.list;

        //早上多打卡
        if (list.size() > 1) {
            int l = 0;
            for (String s : list) {
                if (s.startsWith("07")) {
                    l++;
                }
            }
            if (l > 1) {
                System.out.println(name + "   " + date + " " + list);
                return;
            }
        }
        //打卡次数缺失
        if (list.size() == 1 || list.size() == 3 || list.size() == 5) {
            System.out.println(name + "   " + date + " " + list);
            return;
        }

        //是否请假
        if (list.size() == 2) {
            System.out.println(name + "   " + date + " " + list);
        } else if (list.size() == 4) {
            if (!(list.get(0).startsWith("07") || list.get(0).startsWith("08"))) {
                System.out.println(name + "   " + date + " " + list);
            }
        }
    }

    public static void saveToDatabase(Data data, int room) {
        String name = data.name;
        String date = data.date;
        List<String> list = data.list;

        String d1 = list.get(0);
        String d2 = list.get(1);
        String d3 = list.get(2);
        String d4 = list.get(3);
        String d5 = list.get(4);
        String d6 = list.get(5);

        d1 = Utils.getFirstTime(date, d1);
        d3 = Utils.getFirstTime(date, d3);
        d5 = Utils.getFirstTime(date, d5);
        d6 = Utils.getLastCompleteTime(d6);

        String id = SecureUtil.md5(name + date);

        float f1 = Utils.timeDifference(date + d1, date + d2);
        float f2 = Utils.timeDifference(date + d3, date + d4);
        float f3 = Utils.timeDifference(date + d5, date + d6);

        try {
            //插入数据
            Db.use().insert(
                    Entity.create(Utils.DATABASE_NAME_2)
                            .set("id", id)
                            .set("name", name)
                            .set("day", date)
                            .set("d1", d1)
                            .set("d2", d2)
                            .set("d3", d3)
                            .set("d4", d4)
                            .set("d5", d5)
                            .set("d6", d6)
                            .set("m", f1)
                            .set("a", f2)
                            .set("n", f3)
                            .set("room", room));
        } catch (SQLException e) {
            try {
                //修改的数据
                Db.use().update(
                        Entity.create().set("d1", d1)
                                .set("name", name)
                                .set("day", date)
                                .set("d2", d2)
                                .set("d3", d3)
                                .set("d4", d4)
                                .set("d5", d5)
                                .set("d6", d6)
                                .set("m", f1)
                                .set("a", f2)
                                .set("n", f3),
                        Entity.create(Utils.DATABASE_NAME_2).set("id", id)
                );
            } catch (SQLException e1) {
            }
        }
    }


    public static void getData(List<String> arrayNamesList) throws Exception {
        LinkedHashMap<String, List<DataBean>> mapListDataBean = new LinkedHashMap<>();
        for (String name : arrayNamesList) {
            List<DataBean> arrayDataBean = new ArrayList<>();
            List<Entity> listEntity = Db.use().findAll(Entity.create(Utils.DATABASE_NAME_2).set("name", name));
            for (Entity e : listEntity) {
                String day = e.getStr("day");
                String d1 = e.getStr("d1");
                String d2 = e.getStr("d2");
                String d3 = e.getStr("d3");
                String d4 = e.getStr("d4");
                String d5 = e.getStr("d5");
                String d6 = e.getStr("d6");
                float f1 = e.getFloat("m");
                float f2 = e.getFloat("a");
                float f3 = e.getFloat("n");
                DataBean dataBean = new DataBean(name, day, d1, d2, d3, d4, d5, d6, f1, f2, f3);
                arrayDataBean.add(dataBean);
            }
            if (arrayDataBean.isEmpty()) {
                continue;
            }
            mapListDataBean.put(name, arrayDataBean);
        }

        saveToExcel(mapListDataBean);
    }

    private static void saveToExcel(Map<String, List<DataBean>> mapListDataBean) {
        List<List<String>> rowsList = new ArrayList<>();

        List<String> dayList = Utils.getDaysOfMonth(Utils.YEAR_MONTH);
        rowsList.add(dayList);

        Iterator iterator = mapListDataBean.keySet().iterator();
        while (iterator.hasNext()) {
            List<String> nameList = new ArrayList<>();
            String name = (String) iterator.next();
            List<DataBean> arrayDataBean1 = mapListDataBean.get(name);
            float n = 0;
            float c = 0;
            for (DataBean dataBean : arrayDataBean1) {
                c += dataBean.getDay();
                n += dataBean.getTimes();
            }
            nameList.add(name + "  ");
            nameList.add(c + "天 ");
            nameList.add(Utils.getDecimals(n) + "时  ");
            System.out.println(nameList);

            rowsList.add(nameList);
            rowsList.add(dayList);

            List<String> timeList1 = new ArrayList<>();
            List<String> timeList2 = new ArrayList<>();
            List<String> timeList3 = new ArrayList<>();
            List<String> timeList4 = new ArrayList<>();
            List<String> timeList5 = new ArrayList<>();
            List<String> timeList6 = new ArrayList<>();
            List<String> timeListAm = new ArrayList<>();
            List<String> timeListPm = new ArrayList<>();
            List<String> timeListNm = new ArrayList<>();
            List<String> timeListC = new ArrayList<>();
            List<String> timeListN = new ArrayList<>();

            for (int i = 0; i < dayList.size(); i++) {
                String tt = dayList.get(i);
                List<String> nameDay = new ArrayList<>();
                List<DataBean> arrayDataBean = mapListDataBean.get(name);
                for (DataBean dataBean : arrayDataBean) {
                    nameDay.add(dataBean.day);
                }
                if (nameDay.contains(tt)) {
                    for (DataBean dataBean : arrayDataBean) {
                        if (tt.equals(dataBean.day)) {
                            timeList1.add(dataBean.d1);
                            timeList2.add(dataBean.d2);
                            timeList3.add(dataBean.d3);
                            timeList4.add(dataBean.d4);
                            timeList5.add(dataBean.d5);
                            timeList6.add(dataBean.d6);
                            timeListAm.add(dataBean.am + "");
                            timeListPm.add(dataBean.pm + "");
                            timeListNm.add(dataBean.nm + "");
                            timeListC.add(dataBean.getDay() == 0 ? " " : dataBean.getDay() + "");
                            timeListN.add(dataBean.getTimes() == 0 ? " " : dataBean.getTimes() + "");
                        }
                    }
                } else {
                    timeList1.add(" ");
                    timeList2.add(" ");
                    timeList3.add(" ");
                    timeList4.add(" ");
                    timeList5.add(" ");
                    timeList6.add(" ");
                    timeListAm.add(" ");
                    timeListPm.add(" ");
                    timeListNm.add(" ");
                    timeListC.add("");
                    timeListN.add("");
                }
            }
            rowsList.add(timeList1);
            rowsList.add(timeList2);
            rowsList.add(timeList3);
            rowsList.add(timeList4);
            rowsList.add(timeList5);
            rowsList.add(timeList6);
            rowsList.add(timeListAm);
            rowsList.add(timeListPm);
            rowsList.add(timeListNm);
            rowsList.add(timeListC);
            rowsList.add(timeListN);
        }

        Utils.toExcel(rowsList, Utils.ROOM + "车间", Utils.FILE_PATH +
                Utils.ROOM + "车间" + "_" + Utils.YEAR_MONTH + "_" + new Date().getTime() + ".xlsx");
    }


}
