import cn.hutool.core.date.DateUnit;
import cn.hutool.core.date.DateUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import data.Data;
import data.DataBean;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class RoomOne2 {
    static int DAY = 32;
    static String FILE_PATH = "F:\\WORK\\oa\\file\\";
    static String YEAR_MONTH = "202103";
    static String FILE_NAME = "20210301.xlsx";
    static int ROOM = 1;
    static int SIX = 6;
    static int SIX_T = 36;
    static String FIRST_TIME = "08:00";
    static String FIRST_TIME_PRE = "08:0" + SIX;
    static String QUE_QING = "--";

    static List<String> arrayNamesList = new ArrayList<>();

    public static void main(String[] args) throws Exception {
        clearList();

        arrayNamesList.clear();

        Workbook wbs = getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);

        arrayNamesList.clear();

        List<Data> dataArrayList = new ArrayList<>();
        Map<String, ArrayList<String>> map = new HashMap<>();
        for (int line = 1; line < childSheet.getLastRowNum() + 1; line++) {
            String name = readExcelData(childSheet, line, 1).replace("'", "");
            String date = readExcelData(childSheet, line, 3);
            String s = readExcelData(childSheet, line, 4);
            //
            if (!isAdd(name)) {
//            if (!"樊玉明".equals(name)) {
                continue;
            }
            add(name);

            String key = name + "," + date.replace("-", "");
            ArrayList<String> arrayList = map.get(key);
            if (arrayList == null) {
                arrayList = new ArrayList<>();
                map.put(key, arrayList);
            }

            if (s != null && s.trim().length() > 0) {
                arrayList.add(s.trim());
            }
            map.put(key, arrayList);
        }
//        System.out.println(map);

        List<String> newList = arrayNamesList.stream().distinct().collect(Collectors.toList());
        arrayNamesList = newList;

//        System.out.println(arrayNamesList);

        for (String name : arrayNamesList) {

            for (int i1 = 1; i1 < DAY; i1++) {
                String date = YEAR_MONTH + String.format("%02d", i1);
                String key = name + "," + date;
//                System.out.println(key);
                ArrayList<String> arrayList = map.get(key);
                if (arrayList != null && !arrayList.isEmpty()) {
                    //--------------------------------------------------
                    HashSet<String> hashSet = new HashSet<>();
                    String tmp = null;
                    for (String s : arrayList) {
                        if ("11:29".equals(s)) {
                            s = "11:30";
                        }
                        if ("11:59".equals(s)) {
                            s = "12:00";
                        }
                        if ("16:29".equals(s)) {
                            s = "16:30";
                        }
                        if ("17:29".equals(s)) {
                            s = "17:30";
                        }
                        if ("17:59".equals(s)) {
                            s = "18:00";
                        }
                        if ("18:29".equals(s)) {
                            s = "18:30";
                        }
                        if ("19:59".equals(s)) {
                            s = "20:00";
                        }
                        if ("20:59".equals(s)) {
                            s = "21:00";
                        }

                        String string = date + " " + s;
                        if (tmp == null) {
                            tmp = string;
                            hashSet.add(s);
                        } else {
                            long ll = DateUtil.between(DateUtil.parse(tmp, "yyyyMMdd HH:mm"), DateUtil.parse(string, "yyyyMMdd HH:mm"), DateUnit.MINUTE);
                            if (ll < 6) {
//                                    System.out.println("时间差为" + ll + "分钟" + name + date + arrayList);
                            } else {
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

                    if (arrayList.size() == 7) {

                        if (arrayList.get(0).startsWith("07") && arrayList.get(1).startsWith("08")) {
                            arrayList.remove(0);
                        }

                        if (arrayList.get(3).startsWith("16:36") && arrayList.get(4).startsWith("17:31")) {
                            arrayList.remove("16:36");
                        }
                    }
                    if (arrayList.size() == 7) {
                        ArrayList<String> arrayList7 = new ArrayList<>();
                        ArrayList<String> arrayList12 = new ArrayList<>();
                        ArrayList<String> arrayList11 = new ArrayList<>();
                        ArrayList<String> arrayList17 = new ArrayList<>();
                        ArrayList<String> arrayList21 = new ArrayList<>();
                        for (String s : arrayList) {
                            if (s.startsWith("07")) {
                                arrayList7.add(s);
                            }
                            if (s.startsWith("11")) {
                                arrayList11.add(s);
                            }
                            if (s.startsWith("12")) {
                                arrayList12.add(s);
                            }
                            if (s.startsWith("17")) {
                                arrayList17.add(s);
                            }
                            if (s.startsWith("21")) {
                                arrayList21.add(s);
                            }
                        }
                        if (arrayList7.size() > 1) {
                            arrayList.removeAll(arrayList7);
                            arrayList.add(0, arrayList7.get(0));
                        }
                        if (arrayList11.size() == 3) {
                            arrayList.remove(arrayList11.get(1));
                        }
                        if (arrayList12.size() > 2) {
                            arrayList.remove(arrayList12.get(arrayList12.size() - 1));
                        }
                        if (arrayList17.size() > 2) {
                            arrayList.remove(arrayList17.get(arrayList17.size() - 1));
                        }
                        if (arrayList21.size() > 1) {
                            arrayList.remove(arrayList21.get(0));
                        }
                    }

                    if (arrayList.size() > 6) {
                        System.out.println(name + date + arrayList);
                    }
//                    if ("曹明光".equals(name)) {
//                        System.out.println(name + date + arrayList);
//                    }



                    completeQueQing(arrayList);
                    dataArrayList.add(new Data(name, date, arrayList));
                }
            }

        }
//        System.out.println(dataArrayList);

        //计算并保存
        getData(arrayNamesList, dataArrayList);

        printList();
    }


    //---------------------------------------------------------


    static void add(String name) {
        if (isAdd(name)) {
            arrayNamesList.add(name);
        }
    }

    static boolean isAdd(String name) {
        if (name == null
                || "".equals(name)
                || "正常".equals(name)
                || "陈卫平".equals(name)
                || "王思刚".equals(name)
                || "苏金丽".equals(name)
                || "何仁易".equals(name)
                || "王扬威".equals(name)
                || "陈鹏".equals(name)
                || "张冬生".equals(name)
                || "糜火锋".equals(name)
                || "周华栋".equals(name)
                || "汪有国".equals(name)
                || "洪志超".equals(name)
                || "程泉华".equals(name)
                || "聂玉光".equals(name)
                || "陈秋水".equals(name)
                || "陈卫锋".equals(name)
                || "方勇".equals(name)
//                || "程后盛".equals(name)
//                || "张世玉".equals(name)

//                || "冯金平".equals(name)
//                || "苏俊潇".equals(name)
//                || "夏宇恒".equals(name)
//                || "许凌玉".equals(name)
//                || "郑彦俊".equals(name)
//                || "汪胜利".equals(name)
//                || "周谟林".equals(name)
//                || "王之检".equals(name)
//                || "周左文".equals(name)
//                || "谭强".equals(name)
//                || "徐柳根".equals(name)
//                || "何巧珍".equals(name)
//                || "臧世凯".equals(name)
//                || "刘翔".equals(name)
//                || "蒋婵".equals(name)
//                || "程后平".equals(name)
//                || "危发强".equals(name)
//                || "李星星".equals(name)
//                || "王振宇".equals(name)
//                || "刘三波".equals(name)
//                || "严命坤".equals(name)
//                || "虞涛".equals(name)
//                || "金绍雷".equals(name)
                || name.length() < 2) {
            return false;
        }
        return true;
    }

    //每月天数
    static List<String> getDaysOfMonth(String year) {
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

    static void toExcel(List<List<String>> rows, String path) {
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(path);
        // 合并单元格后的标题行，使用默认标题样式
        //writer.merge(rows.size() - 1, title);
        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);

//        Sheet sheet = writer.getSheet();
//        Font font = sheet.getWorkbook().createFont();
//        font.setFontName("宋体");
//        font.setFontHeight((short) 4);
//        font.setFontHeightInPoints((short) 4);
//        font.setItalic(true);
//        font.setStrikeout(true);
//        CellStyle style = sheet.getWorkbook().createCellStyle();
//        for (int i = 0; i < writer.getRowCount(); i++) {
//            style.setFont(font);
//            sheet.getRow(i).setRowStyle(style);
//        }
//        for (int i = 0; i < 32; i++) {
//            sheet.setColumnWidth(i, 1130);
//        }

        // 关闭writer，释放内存
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
        long l1 = DateUtil.parse(d1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis() / 1000;
        long l2 = DateUtil.parse(d2, "yyyyMMddHH:mm").toCalendar().getTimeInMillis() / 1000;
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


    static Workbook getWorkbook() {
        try {
//            File f = new File(FILE_PATH + FILE_NAME);
//            POIFSFileSystem in = new POIFSFileSystem(new FileInputStream(f));
            InputStream in = new FileInputStream(FILE_PATH + FILE_NAME);
            return WorkbookFactory.create(in);
        } catch (Exception e) {
            System.out.println(FILE_PATH + FILE_NAME + "  错误");
            e.printStackTrace();
        }
        return null;
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


    static String getFirstTime(String day, String time1, String time2, String time3, String time4, Data data) {
        if (QUE_QING.equals(time1)) {
            return time1;
        }
        String dd = day + time1;
        long l1 = DateUtil.parse(dd, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
        long l2 = DateUtil.parse(day + FIRST_TIME_PRE, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();

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
            long l1 = DateUtil.parse(dd, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
            long l2 = DateUtil.parse(data.date + "08:06", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
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
                dayList1.add(s.replace(YEAR_MONTH, ""));
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
}
