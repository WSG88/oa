package data;

import cn.hutool.core.date.DateUtil;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.StringTokenizer;

public class Utils {
    public static void main(String[] args) {

        float fff = (float) Math.floor(3.4F + 4.5F - 8F);
        System.out.println(fff);
        System.out.print("小数部分是：" + 0.9f % 1);
        if ((0f % 1) >= 0.85f) {
            System.out.println("sssss");
        }
//        for (int i = 0; i < 61; i++) {
//            String s = "07:";
//            String ss = "";
//            if (i < 10) {
//                ss = s + "0" + i;
//            } else {
//                ss = s + i;
//            }
//            System.out.println(ss + "   " + getCompleteTime(ss));
//        }
//        System.out.println(Math.floor(5.5));
//
//        Map<String, Object> row1 = new LinkedHashMap<>();
//        row1.put("姓名", "张三");
//        row1.put("年龄", 23);
//        row1.put("成绩", 88.32);
//        row1.put("是否合格", true);
//        row1.put("考试日期", DateUtil.date());
//
//        Map<String, Object> row2 = new LinkedHashMap<>();
//        row2.put("姓名", "李四");
//        row2.put("年龄", 33);
//        row2.put("成绩", 59.50);
//        row2.put("是否合格", false);
//        row2.put("考试日期", DateUtil.date());
//
//        ArrayList<Map<String, Object>> rows = CollUtil.newArrayList(row1, row2);
    }

    public static float getDecimals(float f) {
        return Float.parseFloat(new DecimalFormat(".0").format(f));
    }

    /*时间间隔*/
    public static float timeDifference(String d1, String d2) {
        long l1 = DateUtil.parse(d1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(d2, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        float f = (0.0F + l2 - l1) / 1000F / 60F / 60F;
        System.out.println("timeDifference d1 = " + d1 + " d2 = " + d2 + " " + f);
        return f;
    }

    /*时间取值*/
    public static String getCompleteTime(String time) {
        String outTime = "00:00";
        StringTokenizer st = new StringTokenizer(time, ":");
        List<String> inTime = new ArrayList<String>();
        while (st.hasMoreElements()) {
            inTime.add(st.nextToken());
        }
        String hour = inTime.get(0).toString();
        String minutes = inTime.get(1).toString();
        if (Integer.parseInt(minutes) > 35) {
            hour = (Integer.parseInt(hour) + 1) + "";
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) < 6) {
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) <= 35 && Integer.parseInt(minutes) != 0) {
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
}
