package data;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.StringTokenizer;

public class TTT {
    public static void main(String[] args) {
        for (int i = 0; i < 61; i++) {
            String s = "07:";
            String ss = "";
            if (i < 10) {
                ss = s + "0" + i;
            } else {
                ss = s + i;
            }
            System.out.println(ss + "   " + getCompleteTime(ss));
        }
        System.out.println(Math.floor(5.5));
    }


    public static String getCompleteTime(String time) {
        String hour = "00";//小时
        String minutes = "00";//分钟
        String outTime = "00:00";
        StringTokenizer st = new StringTokenizer(time, ":");
        List<String> inTime = new ArrayList<String>();
        while (st.hasMoreElements()) {
            inTime.add(st.nextToken());
        }
        hour = inTime.get(0).toString();
        minutes = inTime.get(1).toString();
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
