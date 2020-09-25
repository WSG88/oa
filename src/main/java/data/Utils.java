package data;

import cn.hutool.core.date.DateUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

public class Utils {
    public static void main(String[] args) throws Exception {

        for (int j = 0; j < 23; j++) {
            String s = j + ":";
            if (j < 10) {
                s = "0" + j + ":";
            }
            for (int i = 0; i < 61; i++) {
                String ss = "";
                if (i < 10) {
                    ss = s + "0" + i;
                } else {
                    ss = s + i;
                }
//            System.out.println(ss + "   " + getCompleteTime(ss));
                System.out.println(ss + "   " + getFirstTime("20200801", ss));
            }
        }
//        System.out.println(Math.floor(5.5));

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
        long l1 = DateUtil.parse(d1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(d2, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        float f = (0.0F + l2 - l1) / 1000F / 60F / 60F;
//        System.out.println("timeDifference d1 = " + d1 + " d2 = " + d2 + " " + f);
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

    public static int SIX = 6;
    public static int SIX_T = 36;
    public static String FIRST_TIME = "08:00";
    public static String FIRST_TIME_PRE = "08:0" + SIX;

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
}
