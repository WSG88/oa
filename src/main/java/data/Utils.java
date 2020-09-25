package data;

import cn.hutool.core.date.DateUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Utils {
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

    public static void test(Date date) {
        System.out.println(DateUtil.beginOfMonth(date));
        System.out.println(DateUtil.beginOfMonth(date).toJdkDate());
        System.out.println(DateUtil.beginOfMonth(date).toSqlDate());
        String format = DateUtil.format(date, "yyyyMMdd");

    }

    public static float m(float nm) {
        if (nm % 1 >= 0.5F) {
            return (float) Math.floor(nm) + 0.5F;
        } else {
            return (float) Math.floor(nm);
        }
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

    public static String getFirstTime(String day, String time) {
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
}
