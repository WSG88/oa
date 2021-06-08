import cn.hutool.core.date.DateUnit;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.text.csv.CsvData;
import cn.hutool.core.text.csv.CsvReader;
import cn.hutool.core.text.csv.CsvRow;
import cn.hutool.core.text.csv.CsvUtil;
import cn.hutool.core.util.CharsetUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import data.Data;
import data.DataBean;
import data.DateTimeBean;
import data.NameData;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

import static java.util.Comparator.comparing;
import static java.util.stream.Collectors.collectingAndThen;
import static java.util.stream.Collectors.toCollection;

public class RoomOne3 {
    static String FILE_PATH = "F:\\WORK\\oa\\file\\";
    static String YEAR_MONTH = "2021-05";
    static int ROOM = 3;
    static int SIX = 6;
    static int SIX_T = 36;
    static String FIRST_TIME = "08:00";
    static String FIRST_TIME_PRE = "08:0" + SIX;
    static String QUE_QING = "--";
    static List<NameData> userList;

    public static void main(String[] args) throws Exception {

        userList = new Gson().fromJson(NameData.NAME, new TypeToken<ArrayList<NameData>>() {
        }.getType());
        userList = userList.stream().collect(collectingAndThen(toCollection(() ->
                new TreeSet<NameData>(comparing(NameData::getEmployeeNo))), ArrayList::new));


        List<String> arrayNamesList = new ArrayList<>();
        List<DateTimeBean> listDateTimeBean = new ArrayList<>();

        List<File> files = FileUtil.loopFiles("F:\\WORK\\oa\\csv\\");
        for (File file : files) {
            String path = file.getAbsolutePath();
            if (path.endsWith(".xlsx")) {
                listDateTimeBean.addAll(getTestBean4xlsx(path));
            }
            if (path.endsWith(".csv")) {
                listDateTimeBean.addAll(getTestBean4CSV(path));
            }
        }
        Map<String, List<DateTimeBean>> mm = listDateTimeBean.stream().collect(Collectors.groupingBy(DateTimeBean::getName));

        if (ROOM == 1) {
            arrayNamesList = getNameList1();
        } else if (ROOM == 2) {
            arrayNamesList = getNameList2();
        } else if (ROOM == 3) {
            arrayNamesList = getNameList3();
        } else if (ROOM == 0) {
            arrayNamesList = getNameList();
        }

        List<Data> dataArrayList = new ArrayList<>();

        for (String name : arrayNamesList) {
            List<DateTimeBean> lists = mm.get(name);
            if (lists == null) {
                continue;
            }
            Map<String, List<DateTimeBean>> map = lists.stream().collect(Collectors.groupingBy(DateTimeBean::getDate));
            Iterator iterator = map.keySet().iterator();
            while (iterator.hasNext()) {
                String date = (String) iterator.next();
                if (!date.contains(YEAR_MONTH)) {
                    continue;
                }
                List<DateTimeBean> list = map.get(date);
                list.sort(new Comparator<DateTimeBean>() {
                    @Override
                    public int compare(DateTimeBean o1, DateTimeBean o2) {
                        return o1.getTime().compareTo(o2.getTime());
                    }
                });
                ArrayList<String> arrayList = new ArrayList<>();
                for (DateTimeBean bean : list) {
                    arrayList.add(bean.getTime());
                }
                if (!arrayList.isEmpty()) {
                    arrayList.sort(new Comparator<String>() {
                        @Override
                        public int compare(String o1, String o2) {
                            return o1.compareTo(o2);
                        }
                    });
                    //System.out.println(arrayList);
                    //--------------------------------------------------
                    HashSet<String> hashSet = new HashSet<>();
                    String tmp = null;
                    for (String s : arrayList) {
                        String string = date + " " + s;
                        if (tmp == null) {
                            tmp = string;
                            hashSet.add(s);
                        } else {
                            if (DateUtil.between(DateUtil.parse(tmp, "yyyy-MM-dd HH:mm"),
                                    DateUtil.parse(string, "yyyy-MM-dd HH:mm"), DateUnit.MINUTE) > 6) {
                                tmp = string;
                                hashSet.add(s);
                            }
                        }
                    }
                    List<String> result = new ArrayList<>(hashSet);
                    result.sort(new Comparator<String>() {
                        @Override
                        public int compare(String o1, String o2) {
                            return o1.compareTo(o2);
                        }
                    });
                    arrayList.clear();
                    arrayList.addAll(result);

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
//                    System.out.println(name + date + arrayList);

                    for (String s : arrayList) {
                        for (NameData nameData : userList) {
                            if (nameData.getEmployeeName().equals(name)) {
                                String employeeNo = nameData.getEmployeeNo();
                                String timeOrg = date + " " + s + ":00";
                                try {
                                    Db.use().insert(
                                            Entity.create("employee_time")
                                                    .set("employee", nameData.getEmployeeNo())
                                                    .set("name", nameData.getEmployeeName())
                                                    .set("time", timeOrg)
                                                    .set("date_time", DateUtil.parseDateTime(timeOrg))
                                                    .set("md_id", SecureUtil.md5(employeeNo + timeOrg))
                                                    .set("time_org", timeOrg)
                                    );
                                } catch (SQLException e) {
//                                    e.printStackTrace();
                                }

                            }
                        }
                    }

                    completeQueQing(arrayList);
                    dataArrayList.add(new Data(name, date, arrayList));
                }
            }

        }

        //计算并保存
        getData(arrayNamesList, dataArrayList);
    }


    //---------------------------------------------------------


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


    static String readExcelData(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    return "";
                } else {
                    String s = cell.getStringCellValue();
                    if (s != null && s.length() > 0) {
//                        System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + s);
                    }
                    return s;
                }
            }
        }
        return "";
    }

    static String readExcel(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        String s = "";
        Row row = childSheet.getRow(rowNum);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); // 日期格式化
        DecimalFormat df = new DecimalFormat("0"); // 格式化为整数
        DecimalFormat df1 = new DecimalFormat("0.00");
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                switch (cell.getCellTypeEnum()) {
                    case _NONE:
                        break;
                    case NUMERIC:
                        String dataFormat = cell.getCellStyle().getDataFormatString();    // 单元格格式
                        boolean isDate = isCellDateFormatted(cell);
                        if ("General".equals(dataFormat)) {
                            s = df.format(cell.getNumericCellValue());
                        } else if (isDate) {
                            s = sdf.format(cell.getDateCellValue());
                        } else {
                            s = String.valueOf(cell.getNumericCellValue());
                        }
                        break;
                    case STRING:
                        s = cell.getStringCellValue();
                        break;
                    case FORMULA:
                        s = df1.format(cell.getNumericCellValue());
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


    public static boolean isCellDateFormatted(Cell cell) {
        if (cell == null) return false;
        boolean bDate = false;

        double d = cell.getNumericCellValue();
        if (isValidExcelDate(d)) {
            CellStyle style = cell.getCellStyle();
            if (style == null) return false;
            int i = style.getDataFormat();
            String f = style.getDataFormatString();
            bDate = isADateFormat(i, f);
        }
        return bDate;
    }

    public static boolean isADateFormat(int formatIndex, String formatString) {
        if (isInternalDateFormat(formatIndex)) {
            return true;
        }

        if ((formatString == null) || (formatString.length() == 0)) {
            return false;
        }

        String fs = formatString;
        //下面这一行是自己手动添加的 以支持汉字格式wingzing
        fs = fs.replaceAll("[\"|\']", "").replaceAll("[年|月|日|时|分|秒|毫秒|微秒]", "");

        fs = fs.replaceAll("\\\\-", "-");

        fs = fs.replaceAll("\\\\,", ",");

        fs = fs.replaceAll("\\\\.", ".");

        fs = fs.replaceAll("\\\\ ", " ");

        fs = fs.replaceAll(";@", "");

        fs = fs.replaceAll("^\\[\\$\\-.*?\\]", "");

        fs = fs.replaceAll("^\\[[a-zA-Z]+\\]", "");

        return (fs.matches("^[yYmMdDhHsS\\-/,. :]+[ampAMP/]*$"));
    }

    public static boolean isInternalDateFormat(int format) {
        switch (format) {
            case 14:
            case 15:
            case 16:
            case 17:
            case 18:
            case 19:
            case 20:
            case 21:
            case 22:
            case 45:
            case 46:
            case 47:
                return true;
            case 23:
            case 24:
            case 25:
            case 26:
            case 27:
            case 28:
            case 29:
            case 30:
            case 31:
            case 32:
            case 33:
            case 34:
            case 35:
            case 36:
            case 37:
            case 38:
            case 39:
            case 40:
            case 41:
            case 42:
            case 43:
            case 44:
        }
        return false;
    }

    public static boolean isValidExcelDate(double value) {
        return (value > -4.940656458412465E-324D);
    }


    static String getFirstTime(String day, String time1, String time2, String time3, String time4, Data data) {
        if (QUE_QING.equals(time1)) {
            return time1;
        }
        String dd = day + time1;
        long l1 = DateUtil.parse(dd, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(day + FIRST_TIME_PRE, "yyyy-MM-ddHH:mm").toCalendar().getTimeInMillis();

        if (l1 > l2) {
            return getCompleteTime(time1);
        } else {
            if (QUE_QING.equals(time2) && QUE_QING.equals(time3) && QUE_QING.equals(time4)) {
//                System.out.println(data);//夜班数据
                return time1;
            }
            return FIRST_TIME;
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

    static void clearList() {
        listList0.clear();
        listList1.clear();
        listList2.clear();
    }

    static void printList() {
        ExcelWriter writer0 = ExcelUtil.getWriter(FILE_PATH + YEAR_MONTH + "_" + ROOM +
                "_异常.xlsx");
        writer0.write(listList0, true);
        writer0.close();

        ExcelWriter writer1 = ExcelUtil.getWriter(FILE_PATH + YEAR_MONTH + "_" + ROOM +
                "_迟到.xlsx");
        writer1.write(listList1, true);
        writer1.close();

        ExcelWriter writer2 = ExcelUtil.getWriter(FILE_PATH + YEAR_MONTH + "_" + ROOM +
                "_缺勤.xlsx");
        writer2.write(listList2, true);
        writer2.close();
    }

    //type 0异常1迟到2早退3缺勤
    static void setListData(Data data, int room, int type) {
        if (data.date.startsWith("20210208")) {
            return;
        }

        List<String> list = new ArrayList<>();
        list.add(room + "车间 ");
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
    static void checkData(Data data, int room) {
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
                    checkData(data, ROOM);

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
                    if (f1 > 0 && f2 > 5.5 && (Double.parseDouble(d4.substring(0, 2)) > 17)) {//晚上加班不打卡扣0.5
                        f2 = f2 - 0.5f;
                    }
                    f1=0;
                    f2=0;
                    f3=0;

                    DataBean dataBean = new DataBean(name, date, dd1, dd2, dd3, dd4, dd5, dd6, dd7, dd8, f1, f2, f3, ROOM);
                    dataBean.error = data.error;
                    arrayDataBean.add(dataBean);
                }
            }
            if (arrayDataBean.isEmpty()) {
                continue;
            }
            mapListDataBean.put(name, arrayDataBean);
        }

        saveToExcel(mapListDataBean, "");

        printList();
    }

    static void saveToExcel(Map<String, List<DataBean>> mapListDataBean, String fileName) {
        List<List<String>> rowsList = new ArrayList<>();
        List<List<String>> rowsList1 = new ArrayList<>();
        List<String> lll0 = new ArrayList<>();
        lll0.add("姓名");
        lll0.add("天数");
        lll0.add("加班时长");
        lll0.add("加班天数");
        lll0.add("总计");
        rowsList1.add(lll0);

        List<String> dayList = getDaysOfMonth(YEAR_MONTH);

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
                            timeList1.add(dataBean.d1);
                            timeList2.add(dataBean.d2);
                            timeList3.add(dataBean.d3);
                            timeList4.add(dataBean.d4);
                            timeList5.add(dataBean.d5);
                            timeList6.add(dataBean.d6);
                            timeList7.add(dataBean.d7);
                            timeList8.add(dataBean.d8);
                            timeListAm.add(dataBean.am > 0 ? dataBean.am + "" : "");
                            timeListPm.add(dataBean.pm > 0 ? dataBean.pm + "" : "");
                            timeListNm.add(dataBean.nm > 0 ? dataBean.nm + "" : "");
                            timeListC.add(dataBean.getDay() == 0 ? " " : dataBean.getDay() + "");
                            String X = dataBean.error == 1 ? "X" : "";
                            timeListN.add(dataBean.getTimes() == 0 ? " " + X : dataBean.getTimes() + "" + X);
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

        if (fileName == null || fileName.length() == 0) {
            toExcel(rowsList, FILE_PATH + YEAR_MONTH + "_" + ROOM + "车间.xlsx");
            toExcel(rowsList1, FILE_PATH + YEAR_MONTH + "_" + ROOM + "车间汇总.xlsx");
        } else {
            toExcel(rowsList, FILE_PATH + YEAR_MONTH + "_" + fileName + "_" + ROOM + "车间.xlsx");
            toExcel(rowsList1, FILE_PATH + YEAR_MONTH + "_" + fileName + "_" + ROOM + "车间汇总.xlsx");
        }
    }


    public static List<String> getNameList1() {
        List<String> list = new ArrayList<>();
        list.addAll(Arrays.asList(("何巧珍\n" +
                "徐慧林\n" +

                "王振宇\n" +
                "谭强\n" +
                "徐柳根\n" +
                "许凌玉\n" +
                "蒋婵\n" +
                "刘翔\n" +
                "郑彦俊\n" +
                "杨洁\n" +
                "何晓晴\n" +

                "李安琦\n" +
                "陈妍\n" +
                "危欢\n" +
                "张集成\n" +
                "刘良红\n" +

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
                "周春\n"

        ).split("\n")));
        list.addAll(Arrays.asList(("刘三波\n" +
                "冯金平\n").split("\n")));
        return list;
    }

    public static List<String> getNameList2() {
        List<String> list = new ArrayList<>();
        list.addAll(Arrays.asList((
                "吴兰\n" +
                        "徐育林\n" +
                        "王章美\n" +
                        "张新丽\n" +

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
                        "李柳柳\n").split("\n")));
        return list;
    }

    public static List<String> getNameList3() {
        List<String> list = new ArrayList<>();
        list.addAll(Arrays.asList((
                "方卫华\n" +
                        "陈鹏\n" +
                "张光宗\n" +
                        "张仁爱\n" +
                        "张金娥\n" +
                        "徐琳\n" +
                        "刘明巧\n" +
                        "樊玉明\n" +
                        "陈恋\n").split("\n")));
        return list;
    }

    public static List<String> getNameList() {
        List<String> list = new ArrayList<>();
        list.addAll(Arrays.asList((
                "陈伟\n" +
                        "何仁易\n" +
                        "糜火锋\n" +
                        "王扬威\n" +
                        "陈卫平\n" +
                        "何炳辉\n" +
                        "苏阳阳\n" +
                        "汪俊锋\n" +
                        "陈鹏\n" +
                        "王思刚\n" +
                        "苏金丽\n" +
                        "陈彬\n").split("\n")));
        return list;
    }


    private static List<DateTimeBean> getTestBean4xlsx(String path) throws Exception {
        List<DateTimeBean> listDateTimeBean = new ArrayList<>();
        InputStream in = new FileInputStream(path);
        Workbook wbs = WorkbookFactory.create(in);
        Sheet childSheet = wbs.getSheetAt(0);
        for (int line = 0; line < childSheet.getLastRowNum() + 1; line++) {
            String name = readExcelData(childSheet, line, 1).replace("'", "");
            String datetime = readExcel(childSheet, line, 3);
            if (datetime != null && datetime.length() > 10) {
                DateTimeBean dateTimeBean = new DateTimeBean(name, datetime.split(" ")[0], datetime.split(" ")[1].substring(0, 5));
                dateTimeBean.setDatetime(datetime);
                listDateTimeBean.add(dateTimeBean);
            }
        }
        return listDateTimeBean;
    }

    private static List<DateTimeBean> getTestBean4CSV(String path) {
        List<DateTimeBean> listDateTimeBean = new ArrayList<>();
        CsvReader reader = CsvUtil.getReader();
        CsvData data = reader.read(FileUtil.file(path), CharsetUtil.CHARSET_GBK);
        List<CsvRow> rows = data.getRows();
        for (CsvRow csvRow : rows) {
            List<String> list = csvRow.getRawList();
            String name = list.get(1);
            String datetime = list.get(3);
            if (datetime != null && datetime.length() > 10) {
                datetime = datetime.replace("/", "-");
                DateTimeBean dateTimeBean = new DateTimeBean(name, datetime.split(" ")[0], datetime.split(" ")[1].substring(0, 5));
                dateTimeBean.setDatetime(datetime);
                listDateTimeBean.add(dateTimeBean);
            }
        }
        return listDateTimeBean;

    }
}
