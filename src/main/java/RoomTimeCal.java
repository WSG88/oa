import cn.hutool.core.date.DateUnit;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

import static java.util.Comparator.comparing;
import static java.util.stream.Collectors.collectingAndThen;
import static java.util.stream.Collectors.toCollection;

public class RoomTimeCal {

    public static List<DateTimeBean> getEmployeeListByDate1(String date) {
        List<DateTimeBean> employees = new ArrayList<>();
        try {
//            String sql = "select * from attendance1 where date like '2021-05%' order by date_time desc;";
            String sql = "select *\n" +
                    "from attendance1\n" +
                    "where date like '" + date +
                    "%'\n" +
                    "  and device_name like '车间%考勤'\n" +
                    "order by date_time desc;";
            List<Entity> list = Db.use().query(sql);
            for (Entity entity : list) {
                String employeeName = entity.getStr("employ_name");
                String employeeNo = entity.getStr("id");
                String time = entity.getStr("date") + " " + entity.getStr("time");
                if (StrUtil.isNotEmpty(employeeName) && StrUtil.isNotEmpty(employeeNo)) {
                    DateTimeBean employee = new DateTimeBean();
                    employee.setName(employeeName);
                    employee.setDatetime(time);
                    employee.setDate(time.substring(0, 10));
                    employee.setTime(time.substring(11, 19));
                    employees.add(employee);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return employees;
    }

    public static void main(String[] args) {
        //mysql
//        List<DateTimeBean> listDateTimeBean = employeeDao.getEmployeeListByDate("202105");

        //sqlserver
        RoomTimeCal.cal(getEmployeeListByDate1("2021-05"));

//        RoomOne4.calAgain(da,"F:\\WORK\\springboot-04-system\\target\\test-classes\\2021_06_17_17_21_42_2021_05_车间.xlsx");
    }

    static String QUE_QING = "--";

    public static void cal(List<DateTimeBean> listDateTimeBean) {
        Map<String, List<DateTimeBean>> mm = listDateTimeBean.stream().collect(Collectors.groupingBy(RoomTimeCal.DateTimeBean::getName));
        List<Data> dataArrayList = new ArrayList<>();
        List<String> arrayNamesList = getNameList2();
        for (String name : arrayNamesList) {
            List<DateTimeBean> lists = mm.get(name);
            if (lists == null) {
                continue;
            }
            Map<String, List<DateTimeBean>> map = lists.stream().collect(Collectors.groupingBy(DateTimeBean::getDate));
            Iterator<String> iterator = map.keySet().iterator();
            while (iterator.hasNext()) {
                String date = iterator.next();
                List<DateTimeBean> list = map.get(date);
                list.sort((o1, o2) -> o1.getTime().compareTo(o2.getTime()));
                ArrayList<String> arrayList = new ArrayList<>();
                for (DateTimeBean bean : list) {
                    arrayList.add(bean.getTime());
                }
                if (!arrayList.isEmpty()) {
                    arrayList.sort((o1, o2) -> o1.compareTo(o2));
                    HashSet<String> hashSet = new HashSet<>();
                    String tmp = null;
                    for (String s : arrayList) {
                        String string = date + " " + s;
                        if (tmp == null) {//第一个直接添加
                            tmp = string;
                            hashSet.add(s);
                        } else {//开始比较
                            if (DateUtil.between(DateUtil.parse(tmp, "yyyy-MM-dd HH:mm"),
                                    DateUtil.parse(string, "yyyy-MM-dd HH:mm"), DateUnit.MINUTE) > 6) {
                                tmp = string;
                                hashSet.add(s);
                            }
                        }
                    }
                    List<String> result = new ArrayList<>(hashSet);
                    result.sort((o1, o2) -> o1.compareTo(o2));
                    arrayList.clear();
                    arrayList.addAll(result);

                    //处理早上多打卡问题
                    if (arrayList.size() > 3) {
                        if (arrayList.get(0).startsWith("07:") && arrayList.get(1).startsWith("07:")) {
                            arrayList.remove(1);
                        }
                    }
                    //处理夏季中午多打卡问题
                    if (arrayList.size() > 4) {
                        if (arrayList.get(1).startsWith("12:")
                                && arrayList.get(2).startsWith("12:")
                                && arrayList.get(3).startsWith("13:")
                                && arrayList.get(4).startsWith("17:")) {
                            arrayList.remove(2);
                        }
                    }
                    //--------------------------------------------------

                    //处理晚上八点九点数据
                    if (arrayList.size() > 6) {
                        String s1 = arrayList.get(arrayList.size() - 1);
                        String s2 = arrayList.get(arrayList.size() - 2);
                        if (s1.startsWith("2") && s2.startsWith("2")) {
                            if (getLastTime(s1).equals(getLastTime(s2))) {
                                arrayList.remove(arrayList.size() - 2);
//                                arrayList.remove(arrayList.size() - 1);
//                                arrayList.add(getLastTime(s1));
                            }
                        }
                    }
                    //处理四点半下班数据
                    if (arrayList.size() > 2) {
                        ArrayList<String> arr1 = new ArrayList<>();
                        for (String s : arrayList) {
                            if (s.startsWith("16:")) {
                                arr1.add(s);
                            }
                        }
                        if (arr1.size() > 1) {
                            arrayList.remove(arr1.get(0));
                        }
                    }
                    if (arrayList.size() == 5) {
                        System.out.println(name + " " + date + " " + arrayList);
                    }
                    completeQueQing(arrayList);
                    dataArrayList.add(new Data(name, date, arrayList));
                }
            }

        }

        //计算并保存
        getData(arrayNamesList, dataArrayList);
    }

    //每月天数
    static List<String> getDaysOfMonth(String year) {
        List<String> list = new ArrayList<>();
        try {
            Date date = new SimpleDateFormat("yyyy-MM").parse(year);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);
            int days = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
            for (int i = 1; i < days + 1; i++) {
                String newString = String.format("%02d", i);
                list.add(year + "-" + newString);
            }
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return list;
    }

    static void toExcel(List<List<String>> rows, String path) {
        ExcelWriter writer = ExcelUtil.getWriter(path);
        writer.write(rows, true);
        writer.close();
    }

    static float getDecimals(float f) {
        return Float.parseFloat(new DecimalFormat(".0").format(f));
    }

    /*时间间隔*/
    static float timeDifference(String d1, String d2) {
        if (d1.contains(QUE_QING) || d2.contains(QUE_QING)) {
            return 0F;
        }
        long l1 = DateUtil.parse(d1, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis() / 1000;
        long l2 = DateUtil.parse(d2, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis() / 1000;
        float f = (l2 - l1) / 3600F;
        return Float.parseFloat(new DecimalFormat(".00").format(f));
    }

    //下班时间取三分钟内补齐整数
    static String getCompleteTime1(String time) {
        String outTime = time;
        if (time != null && time.contains(":") && time.split(":").length == 2) {
            String[] ss = time.split(":");
            int hour = Integer.parseInt(ss[0]);
            int min = Integer.parseInt(ss[1]);
            if (min > 26 && min < 30) {
                outTime = String.format("%02d", hour) + ":30";
//                System.out.println("time = " + time + " outTime = " + outTime);
            }
            if (min > 56 && min < 61) {
                outTime = String.format("%02d", hour + 1) + ":00";
//                System.out.println("time = " + time + " outTime = " + outTime);
            }
        }
        return outTime;
    }

    static String getLastTime(String time) {
        String outTime = time;
        if (time != null && time.contains(":") && time.split(":").length == 2) {
            String[] ss = time.split(":");
            int hour = Integer.parseInt(ss[0]);
            int min = Integer.parseInt(ss[1]);
            if (min > 0 && min < 27) {
                outTime = String.format("%02d", hour) + ":00";
            } else if (min >= 27 && min < 57) {
                outTime = String.format("%02d", hour) + ":30";
            } else if (min >= 57 && min < 61) {
                outTime = String.format("%02d", hour + 1) + ":00";
            }
        }
        return outTime;
    }

    /*时间取值,允许6分钟，其他为半小时向上取整*/
    static String getCompleteTime(String time) {
        String outTime = "00:00";
        StringTokenizer st = new StringTokenizer(time, ":");
        List<String> inTime = new ArrayList<String>();
        while (st.hasMoreElements()) {
            inTime.add(st.nextToken());
        }
        String hour = inTime.get(0).toString();
        String minutes = inTime.get(1).toString();
        if (Integer.parseInt(minutes) > 36) {
            hour = (Integer.parseInt(hour) + 1) + "";
            outTime = hour + ":00";
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
            try {
                outTime = sdf.format(sdf.parse(outTime));
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (Integer.parseInt(minutes) < 7) {
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
//        System.out.println("time = " + time + " outTime = " + outTime);
        return outTime;
    }

    /*时间取值,半小时向下取整*/
    static String getLastCompleteTime(String time) {
        if (QUE_QING.equals(time) || "24:00".equals(time)) {
            return time;
        }
        if ("00:00".equals(time)) {
            return "24:00";
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
//        System.out.println("time = " + time + " outTime = " + outTime);
        return outTime;
    }

    static String getFirstTime(String day, String time1, String time2, String time3, String time4, Data data) {
        if (QUE_QING.equals(time1)) {
            return time1;
        }
        String dd = day + time1;
        long l1 = DateUtil.parse(dd, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(day + "08:06", "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();

        if (l1 > l2) {
            return getCompleteTime(time1);
        } else {
            if (QUE_QING.equals(time2) && QUE_QING.equals(time3) && QUE_QING.equals(time4)) {
//                System.out.println(data);//夜班数据
                return time1;
            }
            return "08:00";
        }
    }

    static void completeQueQing(List<String> list) {
        int length = list.size();
        switch (length) {
            case 0:
                list.add(QUE_QING);
                list.add(QUE_QING);
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
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 2:
                list.add(QUE_QING);
                list.add(QUE_QING);
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
                list.add(QUE_QING);
                break;
            case 4:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 5:
                list.add(QUE_QING);
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 6:
                list.add(QUE_QING);
                list.add(QUE_QING);
                break;
            case 7:
                list.add(QUE_QING);
                break;
        }
    }

    static List<List<List<String>>> listList0 = new ArrayList<>();
    static List<List<List<String>>> listList1 = new ArrayList<>();
    static List<List<List<String>>> listList2 = new ArrayList<>();

    //type 0异常1迟到2早退3缺勤
    static void setListData(Data data, int room, int type) {
        List<String> list = new ArrayList<>();
        list.add(data.name);
        list.add(data.date);
        list.addAll(data.list);
        List<List<String>> rowsList = new ArrayList<>();
        rowsList.add(list);
        if (type == 0) {
            listList0.add(rowsList);
        } else if (type == 1) {
            listList1.add(rowsList);
        } else if (type == 2) {
            listList2.add(rowsList);
        }
    }

    //检查考勤数据是否完整
    static void checkData(Data data) {
        int room = 0;
        if (data == null) return;
        List<String> list = new ArrayList<>();
        for (int i = 0; i < data.list.size(); i++) {
            if (!"".equals(data.list.get(i)) && !QUE_QING.equals(data.list.get(i))) {
                list.add(data.list.get(i));
            }
        }
        //早上多打卡
        if (list.size() > 1) {
            int l = 0;
            for (String s : list) {
                if (s.startsWith("07")) {
                    l++;
                }
            }
            if (l > 1) {
                setListData(data, room, 0);
                data.error = 1;
                return;
            }
        }
        //打卡次数缺失
        if (list.size() == 1 || list.size() == 3 || list.size() == 5 || list.size() > 6) {
            setListData(data, room, 0);
            data.error = 1;
            return;
        }

        //是否请假
        if (list.size() == 2) {
            setListData(data, room, 2);
        } else if (list.size() == 4) {
            if (!(list.get(0).startsWith("07") || list.get(0).startsWith("08"))) {
                setListData(data, room, 2);
            }
        }
        //上午迟到
        if (list.size() > 1) {
            String dd = data.date + list.get(0);
            long l1 = DateUtil.parse(dd, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();
            long l2 = DateUtil.parse(data.date + "08:06", "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();
            if (l1 > l2) {
                setListData(data, room, 1);
            }
        }
    }

    static void getData(List<String> arrayNamesList, List<Data> dataArrayList) {
        LinkedHashMap<String, List<DataBean>> mapListDataBean = new LinkedHashMap<>();
        for (String name : arrayNamesList) {
            List<DataBean> arrayDataBean = new ArrayList<>();
            for (Data data : dataArrayList) {
                if (name.equals(data.name)) {
                    //检查数据
                    checkData(data);

                    //补全数据
                    String date = data.date;
                    List<String> list = data.list;

                    String d1 = list.get(0);
                    String d2 = list.get(1);
                    String d3 = list.get(2);
                    String d4 = list.get(3);
                    String d5 = list.get(4);
                    String d6 = list.get(5);
                    String d7 = list.get(6);
                    String d8 = list.get(7);

                    String dd1 = d1;
                    String dd2 = d2;
                    String dd3 = d3;
                    String dd4 = d4;
                    String dd5 = d5;
                    String dd6 = d6;
                    String dd7 = d7;
                    String dd8 = d8;

                    d1 = getFirstTime(date, d1, d2, d3, d4, data);
                    if (dd1.startsWith("00")) {
                        d1 = dd1;
                    }
                    d3 = getFirstTime(date, d3, d2, d3, d4, data);
                    d5 = getFirstTime(date, d5, d2, d3, d4, data);
                    d6 = getLastCompleteTime(d6);

                    d4 = getCompleteTime1(d4);

                    float f1 = timeDifference(date + d1, date + d2);
                    float f2 = timeDifference(date + d3, date + d4);
                    float f3 = timeDifference(date + d5, date + d6);

                    if (f1 > 6 && f2 == 0) {//夜班数据
//                        System.out.println("夜班数据 "+data+"/ "+f1+"/ "+f2+"/ "+f3);
                        f1 = 0;
                    }
                    if (f1 > 0 && f2 > 4 && d3.startsWith("12") && d4.startsWith("17")) {
                        f2 = 4;
                    }
                    if (f1 > 0 && f2 > 5.5 && (Double.parseDouble(d4.substring(0, 2)) > 17)) {//晚上加班不打卡扣0.5
                        f2 = f2 - 0.5f;
                    }
//                    f1 = 0;
//                    f2 = 0;
//                    f3 = 0;

                    DataBean dataBean = new DataBean(name, date, dd1, dd2, dd3, dd4, dd5, dd6, dd7, dd8, f1, f2, f3, 0);
                    dataBean.error = data.error;
                    arrayDataBean.add(dataBean);
                }
            }
            if (arrayDataBean.isEmpty()) {
                continue;
            }
            mapListDataBean.put(name, arrayDataBean);
        }

        saveToExcel(mapListDataBean);

    }

    static void saveToExcel(Map<String, List<DataBean>> mapListDataBean) {
        List<List<String>> rowsList = new ArrayList<>();
        List<List<String>> rowsList1 = new ArrayList<>();
        List<String> lll0 = new ArrayList<>();
        lll0.add("姓名");
        lll0.add("天数");
        lll0.add("加班时长");
        lll0.add("加班天数");
        lll0.add("总计");
        rowsList1.add(lll0);

        String YEAR_MONTH = "";
        Iterator<String> iterator = mapListDataBean.keySet().iterator();
        while (iterator.hasNext()) {
            List<String> nameList = new ArrayList<>();
            String name = iterator.next();
            List<DataBean> arrayDataBean1 = mapListDataBean.get(name);
            float n = 0;
            float c = 0;
            for (DataBean dataBean : arrayDataBean1) {
                YEAR_MONTH = dataBean.day.substring(0, 7);
                c += dataBean.getDay();
                n += dataBean.getTimes();
            }
            nameList.add("" + name);
            nameList.add("" + c + "d");
            nameList.add("" + getDecimals(n) + "h");
            nameList.add("" + (c + getDecimals(n) / 8));
            rowsList.add(nameList);

            //汇总数据
            List<String> lll = new ArrayList<>();
            lll.add(name);
            lll.add("" + c);
            lll.add("" + getDecimals(n));
            lll.add("" + getDecimals(n) / 8);
            lll.add("" + (c + getDecimals(n) / 8));
            rowsList1.add(lll);

            //日期数据
            List<String> dayList = getDaysOfMonth(YEAR_MONTH);
            List<String> dayList1 = new ArrayList<>();
            for (String s : dayList) {
                dayList1.add(s.substring(8, 10));
            }
            rowsList.add(dayList1);

            List<String> timeList1 = new ArrayList<>();
            List<String> timeList2 = new ArrayList<>();
            List<String> timeList3 = new ArrayList<>();
            List<String> timeList4 = new ArrayList<>();
            List<String> timeList5 = new ArrayList<>();
            List<String> timeList6 = new ArrayList<>();
            List<String> timeList7 = new ArrayList<>();
            List<String> timeList8 = new ArrayList<>();
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
                            timeList1.add(dataBean.d1.length() == 8 ? dataBean.d1.substring(0, 5) : dataBean.d1);
                            timeList2.add(dataBean.d2.length() == 8 ? dataBean.d2.substring(0, 5) : dataBean.d2);
                            timeList3.add(dataBean.d3.length() == 8 ? dataBean.d3.substring(0, 5) : dataBean.d3);
                            timeList4.add(dataBean.d4.length() == 8 ? dataBean.d4.substring(0, 5) : dataBean.d4);
                            timeList5.add(dataBean.d5.length() == 8 ? dataBean.d5.substring(0, 5) : dataBean.d5);
                            timeList6.add(dataBean.d6.length() == 8 ? dataBean.d6.substring(0, 5) : dataBean.d6);
                            timeList7.add(dataBean.d7.length() == 8 ? dataBean.d7.substring(0, 5) : dataBean.d7);
                            timeList8.add(dataBean.d8.length() == 8 ? dataBean.d8.substring(0, 5) : dataBean.d8);
                            timeListAm.add(dataBean.am > 0 ? dataBean.am + "" : "");
                            timeListPm.add(dataBean.pm > 0 ? dataBean.pm + "" : "");
                            timeListNm.add(dataBean.nm > 0 ? dataBean.nm + "" : "");
                            timeListC.add(dataBean.getDay() == 0 ? " " : dataBean.getDay() + "");
                            String X = dataBean.error == 1 ? "X" : "";
                            timeListN.add(dataBean.getTimes() == 0 ? " " + X : dataBean.getTimes() + "" + X);
//                            try {
//                                Db.use().insert(
//                                        Entity.create("employee_time_cal")
//                                                .set("employee", "")
//                                                .set("name", name)
//                                                .set("day_txt", dataBean.day)
//                                                .set("time_txt", dataBean.d1 + "," + dataBean.d2)
//                                                .set("am", dataBean.am)
//                                                .set("pm", dataBean.pm)
//                                                .set("nm", dataBean.nm)
//                                                .set("ds", dataBean.getDay())
//                                                .set("er", dataBean.error)
//                                                .set("md_id", SecureUtil.md5(name + dataBean.day))
//                                );
//                            } catch (SQLException e) {
//                                //e.printStackTrace();
//                            }
                        }
                    }
                } else {
                    timeList1.add("");
                    timeList2.add(" ");
                    timeList3.add(" ");
                    timeList4.add(" ");
                    timeList5.add(" ");
                    timeList6.add(" ");
                    timeList7.add(" ");
                    timeList8.add(" ");
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
            rowsList.add(timeList7);
            rowsList.add(timeList8);
            rowsList.add(timeListAm);
            rowsList.add(timeListPm);
            rowsList.add(timeListNm);
            rowsList.add(timeListC);
            rowsList.add(timeListN);
        }

        String name = (DateUtil.now() + "_" + YEAR_MONTH + "_" + "车间.xlsx").replace(" ", "_").replace("-", "_").replace(":", "_");
        toExcel(rowsList, name);
        toExcel(rowsList1, "汇总" + name);

    }

    public static List<String> getNameList2() {
        List<String> list = new ArrayList<>();
        list.addAll(Arrays.asList((

                "王振宇\n" +
                        "谭强\n" +
                        "徐柳根\n" +
                        "许凌玉\n" +
                        "蒋婵\n" +
                        "刘翔\n" +
                        "郑彦俊\n" +
                        "杨洁\n" +
                        "何晓晴\n" +

                        "吴兰\n" +
                        "徐育林\n" +
                        "王章美\n" +
                        "张新丽\n" +
                        "李安琦\n" +
                        "陈妍\n" +
                        "危欢\n" +
                        "张集成\n" +
                        "刘良红\n" +

                        "刘三波\n" +
                        "冯金平\n" +
                        "张世玉\n" +
                        "臧世荣\n" +
                        "刘太华\n" +
                        "余东湖\n" +
                        "宋圣林\n" +
                        "夏宇恒\n" +
                        "王之检\n" +
                        "周谟林\n" +
                        "李星星\n" +
                        "周左文\n" +
                        "杜传国\n" +
                        "汪胜利\n" +
                        "苏俊潇\n" +
                        "程后平\n" +
                        "危发强\n" +
                        "张康文\n" +
                        "程厚阳\n" +
                        "邓敏\n" +
                        "刘光辉\n" +
                        "虞涛\n" +
                        "张志成\n" +
                        "张志旗\n" +
                        "徐靖\n" +
                        "刘镇\n" +
                        "张学成\n" +
                        "方兴兴\n" +
                        "魏广益\n" +
                        "张玉军\n" +
                        "周国桃\n" +
                        "朱相雨\n" +
                        "朱礼涛\n" +
                        "汪彬\n" +
                        "方敏\n" +
                        "程建东\n" +
                        "郑凯琦\n" +
                        "周柯详\n" +
                        "万文才\n" +
                        "史龙飞\n" +
                        "朱正传\n" +

                        "冯明海\n" +
                        "徐锦军\n" +
                        "饶国辉\n" +
                        "沈长征\n" +

                        "程怡敏\n" +
                        "胡永星\n" +
                        "葛银保\n" +
                        "苏珍生\n" +
                        "吴想凤\n" +
                        "张祖胜\n" +
                        "张光华\n" +
                        "汪小英\n" +
                        "黄加粮\n" +
                        "陈秋生\n" +

                        "詹冬养\n" +
                        "胡镇红\n" +
                        "石治福\n" +
                        "夏晓龙\n" +
                        "苏芳华\n" +
                        "侯木财\n" +
                        "万立妹\n" +
                        "朱想梅\n" +
                        "汪月霞\n" +
                        "彭小英\n" +

                        "詹家景\n" +
                        "程征青\n" +
                        "黄颖超\n" +
                        "朱正飞\n" +
                        "李大江\n" +
                        "章水根\n" +
                        "田维仁\n" +
                        "方智鑫\n" +
                        "程宇文\n" +
                        "邹宇驰\n" +
                        "刘小龙\n" +
                        "何海贵\n" +
                        "万运来\n" +
                        "江伟\n" +
                        "李红海\n" +
                        "黄亦龙\n" +
                        "江鹏\n" +
                        "陈亚军\n" +
                        "王金宝\n" +
                        "张欢\n" +
                        "罗飞\n" +
                        "苏利民\n" +
                        "王国龙\n" +
                        "周铭\n" +
                        "刘敏\n" +
                        "李柳柳\n" +
                        "龚嘉诚\n" +

                        "何巧珍\n" +
                        "徐慧林\n" +
                        "周春\n"

//                        "徐琳\n" +
//                        "刘明巧\n" +
//                        "陈恋\n" +
//                        "樊玉明\n" +
//
//                        "糜火锋\n" +
//                        "周华栋\n" +
//                        "张冬生\n" +
//                        "王扬威\n" +
//                        "洪志超\n" +
//                        "汪有国\n" +
//                        "程泉华\n" +
//                        "方勇\n" +
//
//                        "张光宗\n" +
//                        "张仁爱\n" +
//                        "张金娥\n" +
//
//                        "陈鹏\n" +
//                        "陈彬\n"+
//                        "方卫华\n" +
//                        "王思刚\n" +
//                        "苏金丽\n" +
//                        "陈卫平\n" +
//
//                        "陈伟\n" +
//                        "何炳辉\n" +
//                        "苏阳阳\n" +
//                        "汪俊锋\n"

        ).split("\n")));
        return list;
    }

    public static String readExcel(Sheet childSheet, int rowNum, int cellNum) {
        String s = "";
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                switch (cell.getCellTypeEnum()) {
                    case _NONE:
                        break;
                    case NUMERIC:
                        s = String.valueOf(cell.getNumericCellValue());
                        break;
                    case STRING:
                        s = cell.getStringCellValue();
                        break;
                    case FORMULA:
                        DecimalFormat df = new DecimalFormat("0.00");
                        s = df.format(cell.getNumericCellValue());
                        break;
                    case BLANK:
                        break;
                    case BOOLEAN:
                        s = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case ERROR:
                        break;

                }
                //System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + s);
            }
        }
        return s;
    }

    public static void calAgain(String datetime, String filepath) {
        try {
            List<Data> dataArrayList = new ArrayList<>();
            InputStream in = new FileInputStream(filepath);
            Workbook wbs = WorkbookFactory.create(in);
            Sheet childSheet = wbs.getSheetAt(0);
            for (int rowNumber = 0; rowNumber < childSheet.getLastRowNum() + 1; rowNumber = rowNumber + 14) {
                String name = readExcel(childSheet, rowNumber, 0);
                Row row = childSheet.getRow(rowNumber + 1);
                int totalCellNumber = row.getLastCellNum();
                for (int cellNumber = 0; cellNumber < totalCellNumber; cellNumber++) {
                    //日期
                    String date = datetime + "-" + String.format("%02d", cellNumber + 1);
                    //考勤记录
                    List<String> list = new ArrayList<>();
                    for (int j = rowNumber + 2; j < rowNumber + 8; j++) {
                        String str = readExcel(childSheet, j, cellNumber);
                        list.add(str);
                    }

                    List<String> listNew = new ArrayList<>();
                    for (String s : list) {
                        s = s.trim();
                        if (StrUtil.isEmpty(s) || QUE_QING.equals(s)) {

                        } else {
                            listNew.add(s);
                        }
                    }

                    if (!listNew.isEmpty()) {
                        //补全考勤数据
                        completeQueQing(listNew);
                        dataArrayList.add(new Data(name, date, listNew));
                    }
                }
            }
            //计算并保存
            getData(getNameList2(), dataArrayList);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static class Data {
        public String name;
        public String date;
        public List<String> list;
        public int error;

        public Data(String name, String date, List<String> list) {
            this.name = name;
            this.date = date;
            this.list = list;
        }


        @Override
        public String toString() {
            return "Data{" +
                    "name='" + name + '\'' +
                    ", date='" + date + '\'' +
                    ", list=" + list +
                    '}';
        }

    }

    static class DataBean {

        public String name;
        public String day;
        public String d1;
        public String d2;
        public String d3;
        public String d4;
        public String d5;
        public String d6;
        public String d7;
        public String d8;

        public float am;
        public float pm;
        public float nm;
        public int room = 2;
        public int error;

        public DataBean(String name, String day, String d1, String d2, String d3, String d4, String d5, String d6, String d7, String d8,
                        float am, float pm, float nm, int room) {
            this.name = name;
            this.day = day;
            this.d1 = d1;
            this.d2 = d2;
            this.d3 = d3;
            this.d4 = d4;
            this.d5 = d5;
            this.d6 = d6;
            this.d7 = d7;
            this.d8 = d8;
            this.am = am;
            this.pm = pm;
            this.nm = nm;
            this.room = room;
        }

        @Override
        public String toString() {
            return "DataBean{" +
                    "name='" + name + '\'' +
                    ", day='" + day + '\'' +
                    ", d1='" + d1 + '\'' +
                    ", d2='" + d2 + '\'' +
                    ", d3='" + d3 + '\'' +
                    ", d4='" + d4 + '\'' +
                    ", d5='" + d5 + '\'' +
                    ", d6='" + d6 + '\'' +
                    ", d7='" + d7 + '\'' +
                    ", d8='" + d8 + '\'' +
                    ", am=" + am +
                    ", pm=" + pm +
                    ", nm=" + nm +
                    ", room=" + room +
                    '}';
        }

        public float m(float nm) {
            if (nm % 1 >= 0.5F) {
                return (float) Math.floor(nm) + 0.5F;
            } else {
                return (float) Math.floor(nm);
            }
        }


        public float NIGHT = 7.95F;

        public float getTimes() {
            float f1 = 4.5F;
            float f2 = 3.5F;
            if (room == 1) {
                f1 = 4.0F;
                f2 = 4.0F;
            } else if (room == 2) {
                f1 = 4.5F;
                f2 = 3.5F;
            }
            if (am + pm > NIGHT && pm > f1) {
                if (am < f2) {
                    return m(am + pm - NIGHT) + m(nm);
                }
                return m(pm - f1) + m(nm);
            } else {
                if (nm + am + pm - NIGHT > 0) {
                    return m(nm + am + pm - NIGHT);
                } else {
                    return m(nm + am + pm);
                }
            }
        }

        public int getDay() {
            if (am + pm > NIGHT) {
                return 1;
            }
            if (nm + am + pm - NIGHT > 0) {
                return 1;
            }
            return 0;
        }
    }

    public static class DateTimeBean {
        private String name;
        private String datetime;
        private String date;
        private String time;

        public DateTimeBean() {
        }

        public DateTimeBean(String name, String date, String time) {
            this.name = name;
            this.date = date;
            this.time = time;
        }

        public String getGroupBy() {
            return name + date;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getDatetime() {
            return datetime;
        }

        public void setDatetime(String datetime) {
            this.datetime = datetime;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        public String getTime() {
            return time;
        }

        public void setTime(String time) {
            this.time = time;
        }

        @Override
        public String toString() {
            return "DateTimeBean{" +
                    "name='" + name + '\'' +
                    ", datetime='" + datetime + '\'' +
                    ", date='" + date + '\'' +
                    ", time='" + time + '\'' +
                    '}';
        }
    }

    static class NameData {

        public static String NAME = "[\n" +
                "{\"employeeNo\":\"TY00001\",\"employeeName\":\"陈波\"},\n" +
                "{\"employeeNo\":\"TY00002\",\"employeeName\":\"金美兰\"},\n" +
                "{\"employeeNo\":\"TY00003\",\"employeeName\":\"陈伟\"},\n" +
                "{\"employeeNo\":\"TY00004\",\"employeeName\":\"陈秋生\"},\n" +
                "{\"employeeNo\":\"TY00005\",\"employeeName\":\"何仁易\"},\n" +
                "{\"employeeNo\":\"TY00006\",\"employeeName\":\"方卫华\"},\n" +
                "{\"employeeNo\":\"TY00007\",\"employeeName\":\"陈秋水\"},\n" +
                "{\"employeeNo\":\"TY00008\",\"employeeName\":\"陈卫平\"},\n" +
                "{\"employeeNo\":\"TY00009\",\"employeeName\":\"陈鹏\"},\n" +
                "{\"employeeNo\":\"TY00010\",\"employeeName\":\"聂玉光\"},\n" +
                "{\"employeeNo\":\"TY00011\",\"employeeName\":\"糜火锋\"},\n" +
                "{\"employeeNo\":\"TY00012\",\"employeeName\":\"陈卫锋\"},\n" +
                "{\"employeeNo\":\"TY00013\",\"employeeName\":\"陈金凤\"},\n" +
                "{\"employeeNo\":\"TY00014\",\"employeeName\":\"王扬威\"},\n" +
                "{\"employeeNo\":\"TY00015\",\"employeeName\":\"程泉华\"},\n" +
                "{\"employeeNo\":\"TY00016\",\"employeeName\":\"冯明海\"},\n" +
                "{\"employeeNo\":\"TY00017\",\"employeeName\":\"汪有国\"},\n" +
                "{\"employeeNo\":\"TY00018\",\"employeeName\":\"何炳辉\"},\n" +
                "{\"employeeNo\":\"TY00019\",\"employeeName\":\"何巧珍\"},\n" +
                "{\"employeeNo\":\"TY00020\",\"employeeName\":\"洪志超\"},\n" +
                "{\"employeeNo\":\"TY00021\",\"employeeName\":\"方卫兵\"},\n" +
                "{\"employeeNo\":\"TY00022\",\"employeeName\":\"苏阳阳\"},\n" +
                "{\"employeeNo\":\"TY00023\",\"employeeName\":\"胡淑平\"},\n" +
                "{\"employeeNo\":\"TY00024\",\"employeeName\":\"徐慧林\"},\n" +
                "{\"employeeNo\":\"TY00025\",\"employeeName\":\"张世玉\"},\n" +
                "{\"employeeNo\":\"TY00026\",\"employeeName\":\"刘三波\"},\n" +
                "{\"employeeNo\":\"TY00027\",\"employeeName\":\"程怡敏\"},\n" +
                "{\"employeeNo\":\"TY00028\",\"employeeName\":\"程征青\"},\n" +
                "{\"employeeNo\":\"TY00029\",\"employeeName\":\"臧世荣\"},\n" +
                "{\"employeeNo\":\"TY00030\",\"employeeName\":\"詹冬养\"},\n" +
                "{\"employeeNo\":\"TY00031\",\"employeeName\":\"程后盛\"},\n" +
                "{\"employeeNo\":\"TY00032\",\"employeeName\":\"刘太华\"},\n" +
                "{\"employeeNo\":\"TY00033\",\"employeeName\":\"冯金平\"},\n" +
                "{\"employeeNo\":\"TY00034\",\"employeeName\":\"王振宇\"},\n" +
                "{\"employeeNo\":\"TY00035\",\"employeeName\":\"胡林港\"},\n" +
                "{\"employeeNo\":\"TY00036\",\"employeeName\":\"张冬生\"},\n" +
                "{\"employeeNo\":\"TY00037\",\"employeeName\":\"胡永星\"},\n" +
                "{\"employeeNo\":\"TY00038\",\"employeeName\":\"余东湖\"},\n" +
                "{\"employeeNo\":\"TY00039\",\"employeeName\":\"吴兰\"},\n" +
                "{\"employeeNo\":\"TY00040\",\"employeeName\":\"胡镇红\"},\n" +
                "{\"employeeNo\":\"TY00041\",\"employeeName\":\"宋圣林\"},\n" +
                "{\"employeeNo\":\"TY00042\",\"employeeName\":\"詹家景\"},\n" +
                "{\"employeeNo\":\"TY00043\",\"employeeName\":\"许凌强\"},\n" +
                "{\"employeeNo\":\"TY00044\",\"employeeName\":\"夏宇恒\"},\n" +
                "{\"employeeNo\":\"TY00045\",\"employeeName\":\"许凌玉\"},\n" +
                "{\"employeeNo\":\"TY00046\",\"employeeName\":\"汪月霞\"},\n" +
                "{\"employeeNo\":\"TY00047\",\"employeeName\":\"徐锦军\"},\n" +
                "{\"employeeNo\":\"TY00048\",\"employeeName\":\"张仁爱\"},\n" +
                "{\"employeeNo\":\"TY00049\",\"employeeName\":\"王之检\"},\n" +
                "{\"employeeNo\":\"TY00050\",\"employeeName\":\"张光宗\"},\n" +
                "{\"employeeNo\":\"TY00051\",\"employeeName\":\"黄颖超\"},\n" +
                "{\"employeeNo\":\"TY00052\",\"employeeName\":\"谭强\"},\n" +
                "{\"employeeNo\":\"TY00053\",\"employeeName\":\"周谟林\"},\n" +
                "{\"employeeNo\":\"TY00054\",\"employeeName\":\"徐琳\"},\n" +
                "{\"employeeNo\":\"TY00055\",\"employeeName\":\"周华栋\"},\n" +
                "{\"employeeNo\":\"TY00056\",\"employeeName\":\"张光华\"},\n" +
                "{\"employeeNo\":\"TY00057\",\"employeeName\":\"汪小英\"},\n" +
                "{\"employeeNo\":\"TY00058\",\"employeeName\":\"徐育林\"},\n" +
                "{\"employeeNo\":\"TY00059\",\"employeeName\":\"夏晓龙\"},\n" +
                "{\"employeeNo\":\"TY00060\",\"employeeName\":\"李星星\"},\n" +
                "{\"employeeNo\":\"TY00061\",\"employeeName\":\"朱正飞\"},\n" +
                "{\"employeeNo\":\"TY00062\",\"employeeName\":\"洪建军\"},\n" +
                "{\"employeeNo\":\"TY00063\",\"employeeName\":\"周左文\"},\n" +
                "{\"employeeNo\":\"TY00064\",\"employeeName\":\"汪俊锋\"},\n" +
                "{\"employeeNo\":\"TY00065\",\"employeeName\":\"李大江\"},\n" +
                "{\"employeeNo\":\"TY00066\",\"employeeName\":\"汪林祥\"},\n" +
                "{\"employeeNo\":\"TY00067\",\"employeeName\":\"杜传国\"},\n" +
                "{\"employeeNo\":\"TY00068\",\"employeeName\":\"张军\"},\n" +
                "{\"employeeNo\":\"TY00069\",\"employeeName\":\"章水根\"},\n" +
                "{\"employeeNo\":\"TY00070\",\"employeeName\":\"沈永平\"},\n" +
                "{\"employeeNo\":\"TY00071\",\"employeeName\":\"汪胜利\"},\n" +
                "{\"employeeNo\":\"TY00072\",\"employeeName\":\"苏俊潇\"},\n" +
                "{\"employeeNo\":\"TY00073\",\"employeeName\":\"徐柳根\"},\n" +
                "{\"employeeNo\":\"TY00074\",\"employeeName\":\"田维仁\"},\n" +
                "{\"employeeNo\":\"TY00075\",\"employeeName\":\"程后平\"},\n" +
                "{\"employeeNo\":\"TY00076\",\"employeeName\":\"刘翔\"},\n" +
                "{\"employeeNo\":\"TY00077\",\"employeeName\":\"方智鑫\"},\n" +
                "{\"employeeNo\":\"TY00078\",\"employeeName\":\"严长凯\"},\n" +
                "{\"employeeNo\":\"TY00079\",\"employeeName\":\"熊雷\"},\n" +
                "{\"employeeNo\":\"TY00080\",\"employeeName\":\"郑彦俊\"},\n" +
                "{\"employeeNo\":\"TY00081\",\"employeeName\":\"危欢\"},\n" +
                "{\"employeeNo\":\"TY00082\",\"employeeName\":\"程晓栋\"},\n" +
                "{\"employeeNo\":\"TY00083\",\"employeeName\":\"钟镕骏\"},\n" +
                "{\"employeeNo\":\"TY00084\",\"employeeName\":\"刘玮\"},\n" +
                "{\"employeeNo\":\"TY00085\",\"employeeName\":\"蓝天航\"},\n" +
                "{\"employeeNo\":\"TY00086\",\"employeeName\":\"古佑强\"},\n" +
                "{\"employeeNo\":\"TY00087\",\"employeeName\":\"肖立军\"},\n" +
                "{\"employeeNo\":\"TY00088\",\"employeeName\":\"童广辉\"},\n" +
                "{\"employeeNo\":\"TY00089\",\"employeeName\":\"曾凡样\"},\n" +
                "{\"employeeNo\":\"TY00090\",\"employeeName\":\"李林康\"},\n" +
                "{\"employeeNo\":\"TY00091\",\"employeeName\":\"程裕良\"},\n" +
                "{\"employeeNo\":\"TY00092\",\"employeeName\":\"程宇文\"},\n" +
                "{\"employeeNo\":\"TY00093\",\"employeeName\":\"丁光华\"},\n" +
                "{\"employeeNo\":\"TY00094\",\"employeeName\":\"樊岍威\"},\n" +
                "{\"employeeNo\":\"TY00095\",\"employeeName\":\"危发强\"},\n" +
                "{\"employeeNo\":\"TY00096\",\"employeeName\":\"陈震宇\"},\n" +
                "{\"employeeNo\":\"TY00097\",\"employeeName\":\"石治福\"},\n" +
                "{\"employeeNo\":\"TY00098\",\"employeeName\":\"黄亦亮\"},\n" +
                "{\"employeeNo\":\"TY00099\",\"employeeName\":\"张年老\"},\n" +
                "{\"employeeNo\":\"TY00100\",\"employeeName\":\"朱佳星\"},\n" +
                "{\"employeeNo\":\"TY00101\",\"employeeName\":\"李石俊\"},\n" +
                "{\"employeeNo\":\"TY00102\",\"employeeName\":\"邹宇驰\"},\n" +
                "{\"employeeNo\":\"TY00103\",\"employeeName\":\"刘小龙\"},\n" +
                "{\"employeeNo\":\"TY00104\",\"employeeName\":\"周谟滔\"},\n" +
                "{\"employeeNo\":\"TY00105\",\"employeeName\":\"傅赞明\"},\n" +
                "{\"employeeNo\":\"TY00106\",\"employeeName\":\"朱合旺\"},\n" +
                "{\"employeeNo\":\"TY00107\",\"employeeName\":\"何海贵\"},\n" +
                "{\"employeeNo\":\"TY00108\",\"employeeName\":\"项国华\"},\n" +
                "{\"employeeNo\":\"TY00109\",\"employeeName\":\"陈鹏\"},\n" +
                "{\"employeeNo\":\"TY00110\",\"employeeName\":\"万立妹\"},\n" +
                "{\"employeeNo\":\"TY00111\",\"employeeName\":\"万运来\"},\n" +
                "{\"employeeNo\":\"TY00112\",\"employeeName\":\"朱想梅\"},\n" +
                "{\"employeeNo\":\"TY00113\",\"employeeName\":\"江伟\"},\n" +
                "{\"employeeNo\":\"TY00114\",\"employeeName\":\"曹明光\"},\n" +
                "{\"employeeNo\":\"TY00115\",\"employeeName\":\"何猷笑\"},\n" +
                "{\"employeeNo\":\"TY00116\",\"employeeName\":\"江涛\"},\n" +
                "{\"employeeNo\":\"TY00117\",\"employeeName\":\"王涛\"},\n" +
                "{\"employeeNo\":\"TY00118\",\"employeeName\":\"吴德富\"},\n" +
                "{\"employeeNo\":\"TY00119\",\"employeeName\":\"张康文\"},\n" +
                "{\"employeeNo\":\"TY00120\",\"employeeName\":\"程登宇\"},\n" +
                "{\"employeeNo\":\"TY00121\",\"employeeName\":\"徐剑海\"},\n" +
                "{\"employeeNo\":\"TY00122\",\"employeeName\":\"王典碧\"},\n" +
                "{\"employeeNo\":\"TY00123\",\"employeeName\":\"刘先慧\"},\n" +
                "{\"employeeNo\":\"TY00124\",\"employeeName\":\"程俊斌\"},\n" +
                "{\"employeeNo\":\"TY00125\",\"employeeName\":\"丁胜华\"},\n" +
                "{\"employeeNo\":\"TY00126\",\"employeeName\":\"李红海\"},\n" +
                "{\"employeeNo\":\"TY00127\",\"employeeName\":\"葛银保\"},\n" +
                "{\"employeeNo\":\"TY00128\",\"employeeName\":\"余俊\"},\n" +
                "{\"employeeNo\":\"TY00129\",\"employeeName\":\"张黔梅\"},\n" +
                "{\"employeeNo\":\"TY00130\",\"employeeName\":\"陈拥军\"},\n" +
                "{\"employeeNo\":\"TY00131\",\"employeeName\":\"李咸斌\"},\n" +
                "{\"employeeNo\":\"TY00132\",\"employeeName\":\"张新丽\"},\n" +
                "{\"employeeNo\":\"TY00133\",\"employeeName\":\"王思刚\"},\n" +
                "{\"employeeNo\":\"TY00134\",\"employeeName\":\"舒秦龙\"},\n" +
                "{\"employeeNo\":\"TY00135\",\"employeeName\":\"赵如意\"},\n" +
                "{\"employeeNo\":\"TY00136\",\"employeeName\":\"汪有强\"},\n" +
                "{\"employeeNo\":\"TY00137\",\"employeeName\":\"韩忠新\"},\n" +
                "{\"employeeNo\":\"TY00138\",\"employeeName\":\"韩淦\"},\n" +
                "{\"employeeNo\":\"TY00139\",\"employeeName\":\"徐建波\"},\n" +
                "{\"employeeNo\":\"TY00140\",\"employeeName\":\"丁国栋\"},\n" +
                "{\"employeeNo\":\"TY00141\",\"employeeName\":\"李英杰\"},\n" +
                "{\"employeeNo\":\"TY00142\",\"employeeName\":\"葛懿\"},\n" +
                "{\"employeeNo\":\"TY00143\",\"employeeName\":\"石志鹏\"},\n" +
                "{\"employeeNo\":\"TY00144\",\"employeeName\":\"张金娥\"},\n" +
                "{\"employeeNo\":\"TY00145\",\"employeeName\":\"虞童\"},\n" +
                "{\"employeeNo\":\"TY00146\",\"employeeName\":\"邹鹏\"},\n" +
                "{\"employeeNo\":\"TY00147\",\"employeeName\":\"苏金丽\"},\n" +
                "{\"employeeNo\":\"TY00148\",\"employeeName\":\"郑志鹏\"},\n" +
                "{\"employeeNo\":\"TY00149\",\"employeeName\":\"彭建\"},\n" +
                "{\"employeeNo\":\"TY00150\",\"employeeName\":\"余雪云\"},\n" +
                "{\"employeeNo\":\"TY00151\",\"employeeName\":\"刘波\"},\n" +
                "{\"employeeNo\":\"TY00152\",\"employeeName\":\"刘洋\"},\n" +
                "{\"employeeNo\":\"TY00153\",\"employeeName\":\"付文涛\"},\n" +
                "{\"employeeNo\":\"TY00154\",\"employeeName\":\"李华昭\"},\n" +
                "{\"employeeNo\":\"TY00155\",\"employeeName\":\"程厚阳\"},\n" +
                "{\"employeeNo\":\"TY00156\",\"employeeName\":\"胡亮\"},\n" +
                "{\"employeeNo\":\"TY00157\",\"employeeName\":\"邓敏\"},\n" +
                "{\"employeeNo\":\"TY00158\",\"employeeName\":\"严命坤\"},\n" +
                "{\"employeeNo\":\"TY00159\",\"employeeName\":\"黄亦龙\"},\n" +
                "{\"employeeNo\":\"TY00160\",\"employeeName\":\"王章美\"},\n" +
                "{\"employeeNo\":\"TY00161\",\"employeeName\":\"江鹏\"},\n" +
                "{\"employeeNo\":\"TY00162\",\"employeeName\":\"刘光辉\"},\n" +
                "{\"employeeNo\":\"TY00163\",\"employeeName\":\"虞涛\"},\n" +
                "{\"employeeNo\":\"TY00164\",\"employeeName\":\"朱栋龙\"},\n" +
                "{\"employeeNo\":\"TY00165\",\"employeeName\":\"黄父爱\"},\n" +
                "{\"employeeNo\":\"TY00166\",\"employeeName\":\"陈亚军\"},\n" +
                "{\"employeeNo\":\"TY00167\",\"employeeName\":\"王仁焱\"},\n" +
                "{\"employeeNo\":\"TY00168\",\"employeeName\":\"曹煜\"},\n" +
                "{\"employeeNo\":\"TY00169\",\"employeeName\":\"张志成\"},\n" +
                "{\"employeeNo\":\"TY00170\",\"employeeName\":\"张志旗\"},\n" +
                "{\"employeeNo\":\"TY00171\",\"employeeName\":\"金绍雷\"},\n" +
                "{\"employeeNo\":\"TY00172\",\"employeeName\":\"李文洁\"},\n" +
                "{\"employeeNo\":\"TY00173\",\"employeeName\":\"刘安春\"},\n" +
                "{\"employeeNo\":\"TY00174\",\"employeeName\":\"陈妍\"},\n" +
                "{\"employeeNo\":\"TY00175\",\"employeeName\":\"徐靖\"},\n" +
                "{\"employeeNo\":\"TY00176\",\"employeeName\":\"江鸿洋\"},\n" +
                "{\"employeeNo\":\"TY00177\",\"employeeName\":\"苏芳华\"},\n" +
                "{\"employeeNo\":\"TY00178\",\"employeeName\":\"刘镇\"},\n" +
                "{\"employeeNo\":\"TY00179\",\"employeeName\":\"张学成\"},\n" +
                "{\"employeeNo\":\"TY00180\",\"employeeName\":\"方兴兴\"},\n" +
                "{\"employeeNo\":\"TY00181\",\"employeeName\":\"侯美佳\"},\n" +
                "{\"employeeNo\":\"TY00182\",\"employeeName\":\"陈典辉\"},\n" +
                "{\"employeeNo\":\"TY00183\",\"employeeName\":\"詹家杰\"},\n" +
                "{\"employeeNo\":\"TY00184\",\"employeeName\":\"沈长征\"},\n" +
                "{\"employeeNo\":\"TY00185\",\"employeeName\":\"彭春辉\"},\n" +
                "{\"employeeNo\":\"TY00186\",\"employeeName\":\"王添喜\"},\n" +
                "{\"employeeNo\":\"TY00187\",\"employeeName\":\"侯木财\"},\n" +
                "{\"employeeNo\":\"TY00188\",\"employeeName\":\"王金宝\"},\n" +
                "{\"employeeNo\":\"TY00189\",\"employeeName\":\"苏珍生\"},\n" +
                "{\"employeeNo\":\"TY00190\",\"employeeName\":\"臧世凯\"},\n" +
                "{\"employeeNo\":\"TY00191\",\"employeeName\":\"蒋婵\"},\n" +
                "{\"employeeNo\":\"TY00192\",\"employeeName\":\"李安琦\"},\n" +
                "{\"employeeNo\":\"TY00193\",\"employeeName\":\"吴想凤\"},\n" +
                "{\"employeeNo\":\"TY00194\",\"employeeName\":\"罗飞\"},\n" +
                "{\"employeeNo\":\"TY00195\",\"employeeName\":\"邵泽球\"},\n" +
                "{\"employeeNo\":\"TY00196\",\"employeeName\":\"张涛\"},\n" +
                "{\"employeeNo\":\"TY00197\",\"employeeName\":\"吴立平\"},\n" +
                "{\"employeeNo\":\"TY00198\",\"employeeName\":\"张祖胜\"},\n" +
                "{\"employeeNo\":\"TY00199\",\"employeeName\":\"李开立\"},\n" +
                "{\"employeeNo\":\"TY00200\",\"employeeName\":\"张欢\"},\n" +
                "{\"employeeNo\":\"TY00201\",\"employeeName\":\"金云\"},\n" +
                "{\"employeeNo\":\"TY00300\",\"employeeName\":\"江爱平\"},\n" +
                "{\"employeeNo\":1,\"employeeName\":\"陈卫槿\"},\n" +
                "{\"employeeNo\":2,\"employeeName\":\"陈卫楠\"},\n" +
                "{\"employeeNo\":\"TY00001\",\"employeeName\":\"陈波\"},\n" +
                "{\"employeeNo\":\"TY00002\",\"employeeName\":\"金美兰\"},\n" +
                "{\"employeeNo\":\"TY00003\",\"employeeName\":\"陈伟\"},\n" +
                "{\"employeeNo\":\"TY00004\",\"employeeName\":\"陈秋生\"},\n" +
                "{\"employeeNo\":\"TY00005\",\"employeeName\":\"何仁易\"},\n" +
                "{\"employeeNo\":\"TY00006\",\"employeeName\":\"方卫华\"},\n" +
                "{\"employeeNo\":\"TY00007\",\"employeeName\":\"陈秋水\"},\n" +
                "{\"employeeNo\":\"TY00008\",\"employeeName\":\"陈卫平\"},\n" +
                "{\"employeeNo\":\"TY00009\",\"employeeName\":\"陈鹏\"},\n" +
                "{\"employeeNo\":\"TY00011\",\"employeeName\":\"糜火锋\"},\n" +
                "{\"employeeNo\":\"TY00012\",\"employeeName\":\"陈卫锋\"},\n" +
                "{\"employeeNo\":\"TY00013\",\"employeeName\":\"陈金凤\"},\n" +
                "{\"employeeNo\":\"TY00014\",\"employeeName\":\"王扬威\"},\n" +
                "{\"employeeNo\":\"TY00015\",\"employeeName\":\"程泉华\"},\n" +
                "{\"employeeNo\":\"TY00016\",\"employeeName\":\"冯明海\"},\n" +
                "{\"employeeNo\":\"TY00017\",\"employeeName\":\"汪有国\"},\n" +
                "{\"employeeNo\":\"TY00018\",\"employeeName\":\"何炳辉\"},\n" +
                "{\"employeeNo\":\"TY00019\",\"employeeName\":\"何巧珍\"},\n" +
                "{\"employeeNo\":\"TY00020\",\"employeeName\":\"洪志超\"},\n" +
                "{\"employeeNo\":\"TY00021\",\"employeeName\":\"方卫兵\"},\n" +
                "{\"employeeNo\":\"TY00022\",\"employeeName\":\"苏阳阳\"},\n" +
                "{\"employeeNo\":\"TY00023\",\"employeeName\":\"胡淑平\"},\n" +
                "{\"employeeNo\":\"TY00024\",\"employeeName\":\"徐慧林\"},\n" +
                "{\"employeeNo\":\"TY00025\",\"employeeName\":\"张世玉\"},\n" +
                "{\"employeeNo\":\"TY00026\",\"employeeName\":\"刘三波\"},\n" +
                "{\"employeeNo\":\"TY00027\",\"employeeName\":\"程怡敏\"},\n" +
                "{\"employeeNo\":\"TY00028\",\"employeeName\":\"程征青\"},\n" +
                "{\"employeeNo\":\"TY00029\",\"employeeName\":\"臧世荣\"},\n" +
                "{\"employeeNo\":\"TY00030\",\"employeeName\":\"詹冬养\"},\n" +
                "{\"employeeNo\":\"TY00032\",\"employeeName\":\"刘太华\"},\n" +
                "{\"employeeNo\":\"TY00033\",\"employeeName\":\"冯金平\"},\n" +
                "{\"employeeNo\":\"TY00034\",\"employeeName\":\"王振宇\"},\n" +
                "{\"employeeNo\":\"TY00036\",\"employeeName\":\"张冬生\"},\n" +
                "{\"employeeNo\":\"TY00037\",\"employeeName\":\"胡永星\"},\n" +
                "{\"employeeNo\":\"TY00038\",\"employeeName\":\"余东湖\"},\n" +
                "{\"employeeNo\":\"TY00039\",\"employeeName\":\"吴兰\"},\n" +
                "{\"employeeNo\":\"TY00040\",\"employeeName\":\"胡镇红\"},\n" +
                "{\"employeeNo\":\"TY00041\",\"employeeName\":\"宋圣林\"},\n" +
                "{\"employeeNo\":\"TY00042\",\"employeeName\":\"詹家景\"},\n" +
                "{\"employeeNo\":\"TY00044\",\"employeeName\":\"夏宇恒\"},\n" +
                "{\"employeeNo\":\"TY00046\",\"employeeName\":\"汪月霞\"},\n" +
                "{\"employeeNo\":\"TY00047\",\"employeeName\":\"徐锦军\"},\n" +
                "{\"employeeNo\":\"TY00048\",\"employeeName\":\"张仁爱\"},\n" +
                "{\"employeeNo\":\"TY00049\",\"employeeName\":\"王之检\"},\n" +
                "{\"employeeNo\":\"TY00050\",\"employeeName\":\"张光宗\"},\n" +
                "{\"employeeNo\":\"TY00051\",\"employeeName\":\"黄颖超\"},\n" +
                "{\"employeeNo\":\"TY00052\",\"employeeName\":\"谭强\"},\n" +
                "{\"employeeNo\":\"TY00053\",\"employeeName\":\"周谟林\"},\n" +
                "{\"employeeNo\":\"TY00054\",\"employeeName\":\"徐琳\"},\n" +
                "{\"employeeNo\":\"TY00055\",\"employeeName\":\"周华栋\"},\n" +
                "{\"employeeNo\":\"TY00056\",\"employeeName\":\"张光华\"},\n" +
                "{\"employeeNo\":\"TY00057\",\"employeeName\":\"汪小英\"},\n" +
                "{\"employeeNo\":\"TY00058\",\"employeeName\":\"徐育林\"},\n" +
                "{\"employeeNo\":\"TY00059\",\"employeeName\":\"夏晓龙\"},\n" +
                "{\"employeeNo\":\"TY00060\",\"employeeName\":\"李星星\"},\n" +
                "{\"employeeNo\":\"TY00061\",\"employeeName\":\"朱正飞\"},\n" +
                "{\"employeeNo\":\"TY00063\",\"employeeName\":\"周左文\"},\n" +
                "{\"employeeNo\":\"TY00064\",\"employeeName\":\"汪俊锋\"},\n" +
                "{\"employeeNo\":\"TY00065\",\"employeeName\":\"李大江\"},\n" +
                "{\"employeeNo\":\"TY00067\",\"employeeName\":\"杜传国\"},\n" +
                "{\"employeeNo\":\"TY00069\",\"employeeName\":\"章水根\"},\n" +
                "{\"employeeNo\":\"TY00071\",\"employeeName\":\"汪胜利\"},\n" +
                "{\"employeeNo\":\"TY00072\",\"employeeName\":\"苏俊潇\"},\n" +
                "{\"employeeNo\":\"TY00073\",\"employeeName\":\"徐柳根\"},\n" +
                "{\"employeeNo\":\"TY00074\",\"employeeName\":\"田维仁\"},\n" +
                "{\"employeeNo\":\"TY00075\",\"employeeName\":\"程后平\"},\n" +
                "{\"employeeNo\":\"TY00076\",\"employeeName\":\"刘翔\"},\n" +
                "{\"employeeNo\":\"TY00077\",\"employeeName\":\"方智鑫\"},\n" +
                "{\"employeeNo\":\"TY00080\",\"employeeName\":\"郑彦俊\"},\n" +
                "{\"employeeNo\":\"TY00081\",\"employeeName\":\"危欢\"},\n" +
                "{\"employeeNo\":\"TY00092\",\"employeeName\":\"程宇文\"},\n" +
                "{\"employeeNo\":\"TY00095\",\"employeeName\":\"危发强\"},\n" +
                "{\"employeeNo\":\"TY00097\",\"employeeName\":\"石治福\"},\n" +
                "{\"employeeNo\":\"TY00102\",\"employeeName\":\"邹宇驰\"},\n" +
                "{\"employeeNo\":\"TY00103\",\"employeeName\":\"刘小龙\"},\n" +
                "{\"employeeNo\":\"TY00107\",\"employeeName\":\"何海贵\"},\n" +
                "{\"employeeNo\":\"TY00109\",\"employeeName\":\"陈鹏\"},\n" +
                "{\"employeeNo\":\"TY00110\",\"employeeName\":\"万立妹\"},\n" +
                "{\"employeeNo\":\"TY00111\",\"employeeName\":\"万运来\"},\n" +
                "{\"employeeNo\":\"TY00112\",\"employeeName\":\"朱想梅\"},\n" +
                "{\"employeeNo\":\"TY00113\",\"employeeName\":\"江伟\"},\n" +
                "{\"employeeNo\":\"TY00119\",\"employeeName\":\"张康文\"},\n" +
                "{\"employeeNo\":\"TY00126\",\"employeeName\":\"李红海\"},\n" +
                "{\"employeeNo\":\"TY00127\",\"employeeName\":\"葛银保\"},\n" +
                "{\"employeeNo\":\"TY00132\",\"employeeName\":\"张新丽\"},\n" +
                "{\"employeeNo\":\"TY00133\",\"employeeName\":\"王思刚\"},\n" +
                "{\"employeeNo\":\"TY00144\",\"employeeName\":\"张金娥\"},\n" +
                "{\"employeeNo\":\"TY00147\",\"employeeName\":\"苏金丽\"},\n" +
                "{\"employeeNo\":\"TY00155\",\"employeeName\":\"程厚阳\"},\n" +
                "{\"employeeNo\":\"TY00157\",\"employeeName\":\"邓敏\"},\n" +
                "{\"employeeNo\":\"TY00159\",\"employeeName\":\"黄亦龙\"},\n" +
                "{\"employeeNo\":\"TY00160\",\"employeeName\":\"王章美\"},\n" +
                "{\"employeeNo\":\"TY00161\",\"employeeName\":\"江鹏\"},\n" +
                "{\"employeeNo\":\"TY00162\",\"employeeName\":\"刘光辉\"},\n" +
                "{\"employeeNo\":\"TY00163\",\"employeeName\":\"虞涛\"},\n" +
                "{\"employeeNo\":\"TY00166\",\"employeeName\":\"陈亚军\"},\n" +
                "{\"employeeNo\":\"TY00169\",\"employeeName\":\"张志成\"},\n" +
                "{\"employeeNo\":\"TY00170\",\"employeeName\":\"张志旗\"},\n" +
                "{\"employeeNo\":\"TY00174\",\"employeeName\":\"陈妍\"},\n" +
                "{\"employeeNo\":\"TY00175\",\"employeeName\":\"徐靖\"},\n" +
                "{\"employeeNo\":\"TY00177\",\"employeeName\":\"苏芳华\"},\n" +
                "{\"employeeNo\":\"TY00178\",\"employeeName\":\"刘镇\"},\n" +
                "{\"employeeNo\":\"TY00179\",\"employeeName\":\"张学成\"},\n" +
                "{\"employeeNo\":\"TY00180\",\"employeeName\":\"方兴兴\"},\n" +
                "{\"employeeNo\":\"TY00184\",\"employeeName\":\"沈长征\"},\n" +
                "{\"employeeNo\":\"TY00187\",\"employeeName\":\"侯木财\"},\n" +
                "{\"employeeNo\":\"TY00188\",\"employeeName\":\"王金宝\"},\n" +
                "{\"employeeNo\":\"TY00189\",\"employeeName\":\"苏珍生\"},\n" +
                "{\"employeeNo\":\"TY00191\",\"employeeName\":\"蒋婵\"},\n" +
                "{\"employeeNo\":\"TY00192\",\"employeeName\":\"李安琦\"},\n" +
                "{\"employeeNo\":\"TY00193\",\"employeeName\":\"吴想凤\"},\n" +
                "{\"employeeNo\":\"TY00194\",\"employeeName\":\"罗飞\"},\n" +
                "{\"employeeNo\":\"TY00197\",\"employeeName\":\"吴立平\"},\n" +
                "{\"employeeNo\":\"TY00198\",\"employeeName\":\"张祖胜\"},\n" +
                "{\"employeeNo\":\"TY00200\",\"employeeName\":\"张欢\"},\n" +
                "{\"employeeNo\":\"TY00204\",\"employeeName\":\"刘明巧\"},\n" +
                "{\"employeeNo\":\"TY00207\",\"employeeName\":\"张玉军\"},\n" +
                "{\"employeeNo\":\"TY00208\",\"employeeName\":\"方勇\"},\n" +
                "{\"employeeNo\":\"TY00210\",\"employeeName\":\"饶国辉\"},\n" +
                "{\"employeeNo\":\"TY00211\",\"employeeName\":\"陈彬\"},\n" +
                "{\"employeeNo\":\"TY00212\",\"employeeName\":\"苏利民\"},\n" +
                "{\"employeeNo\":\"TY00213\",\"employeeName\":\"樊玉明\"},\n" +
                "{\"employeeNo\":\"TY00214\",\"employeeName\":\"王国龙\"},\n" +
                "{\"employeeNo\":\"TY00215\",\"employeeName\":\"周国桃\"},\n" +
                "{\"employeeNo\":\"TY00217\",\"employeeName\":\"朱礼涛\"},\n" +
                "{\"employeeNo\":\"TY00218\",\"employeeName\":\"汪彬\"},\n" +
                "{\"employeeNo\":\"TY00219\",\"employeeName\":\"陈恋\"},\n" +
                "{\"employeeNo\":\"TY00220\",\"employeeName\":\"黄加粮\"},\n" +
                "{\"employeeNo\":\"TY00221\",\"employeeName\":\"周铭\"},\n" +
                "{\"employeeNo\":\"TY00225\",\"employeeName\":\"张集成\"},\n" +
                "{\"employeeNo\":\"TY00226\",\"employeeName\":\"方敏\"},\n" +
                "{\"employeeNo\":\"TY00228\",\"employeeName\":\"刘敏\"},\n" +
                "{\"employeeNo\":\"TY00229\",\"employeeName\":\"杨洁\"},\n" +
                "{\"employeeNo\":\"TY00232\",\"employeeName\":\"何晓晴\"},\n" +
                "{\"employeeNo\":\"TY00233\",\"employeeName\":\"彭小英\"},\n" +
                "{\"employeeNo\":\"TY00235\",\"employeeName\":\"毋保保\"},\n" +
                "{\"employeeNo\":\"TY00236\",\"employeeName\":\"李柳柳\"},\n" +
                "{\"employeeNo\":\"TY00237\",\"employeeName\":\"刘良红\"},\n" +
                "{\"employeeNo\":\"TY00238\",\"employeeName\":\"程建东\"},\n" +
                "{\"employeeNo\":\"TY00239\",\"employeeName\":\"郑凯琦\"},\n" +
                "{\"employeeNo\":\"TY00240\",\"employeeName\":\"周春\"},\n" +
                "{\"employeeNo\":\"TY00241\",\"employeeName\":\"周柯详\"},\n" +
                "{\"employeeNo\":\"TY00242\",\"employeeName\":\"万文才\"},\n" +
                "{\"employeeNo\":\"TY00243\",\"employeeName\":\"史龙飞\"},\n" +
                "{\"employeeNo\":\"TY00300\",\"employeeName\":\"江爱平\"},\n" +
                "{\"employeeNo\":\"TY003001\",\"employeeName\":\"卢洪波\"},\n" +
                "{\"employeeNo\":\"TY003002\",\"employeeName\":\"朱明彪\"},\n" +
                "{\"employeeNo\":\"TY00301\",\"employeeName\":\"陈卫业\"}" +
                "]";

        private String employeeNo;
        private String employeeName;

        public String getEmployeeNo() {
            return employeeNo;
        }

        public void setEmployeeNo(String employeeNo) {
            this.employeeNo = employeeNo;
        }

        public String getEmployeeName() {
            return employeeName;
        }

        public void setEmployeeName(String employeeName) {
            this.employeeName = employeeName;
        }

        @Override
        public String toString() {
            return "NameData{" +
                    "employeeNo='" + employeeNo + '\'' +
                    ", employeeName='" + employeeName + '\'' +
                    '}';
        }

        public static List<NameData> getNameList() {
            List<NameData> arrayList = new Gson().fromJson(NAME, new TypeToken<ArrayList<NameData>>() {
            }.getType());
            arrayList = arrayList.stream().collect(collectingAndThen(toCollection(() ->
                    new TreeSet<>(comparing(NameData::getEmployeeNo))), ArrayList::new));
            return arrayList;
        }
    }

}
