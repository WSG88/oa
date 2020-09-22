import cn.hutool.core.date.DateTime;
import cn.hutool.core.date.DateUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import data.TTT;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test2 {
    public static String DATABASE_NAME_2 = "oatime";
    public static String QUE_QING = "缺勤";
    public static int S_Y_P = 0;

    public static void main(String[] args) throws Exception {
        main2();
    }

    public static void main2() throws Exception {
        List<Entity> find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d1", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d2"))) {
                System.out.println(entity);
            }
        }
        find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d2", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d1"))) {
                System.out.println(entity);
            }
        }
        find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d3", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d4"))) {
                System.out.println(entity);
            }
        }
        find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d4", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d3"))) {
                System.out.println(entity);
            }
        }
        find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d5", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d6"))) {
                System.out.println(entity);
            }
        }
        find = Db.use().find(Entity.create(DATABASE_NAME_2).set("d6", "like 缺勤%"));
        for (Entity entity : find) {
            if (!"缺勤".equals(entity.getStr("d5"))) {
                System.out.println(entity);
            }
        }


        List<String> arrayList = new ArrayList<>();

        List<Entity> entities = Db.use().findAll(DATABASE_NAME_2);
        for (Entity entity : entities) {
            String name = entity.getStr("name");
            if (!arrayList.contains(name)) {
                arrayList.add(name);
            }
        }
        Map<String, List<String>> map = new HashMap<>();
        Map<String, String> map1 = new HashMap<>();

        for (String name : arrayList) {
            int dayChuq = 0;
            int dayNumber = 0;
            float dayF = 0F;
            float f1 = 0F;
            float f2 = 0F;
            float f3 = 0F;
            List<String> array = new ArrayList<>();
            if (name.length() == 0) continue;
            List<Entity> listEntity = Db.use().findAll(Entity.create(DATABASE_NAME_2).set("name", name));
            for (Entity e : listEntity) {
                String day = e.getStr("day");
                String d1 = e.getStr("d1");
                String d2 = e.getStr("d2");
                String d3 = e.getStr("d3");
                String d4 = e.getStr("d4");
                String d5 = e.getStr("d5");
                String d6 = e.getStr("d6");
                f1 = 0.0F;
                f2 = 0.0F;
                f3 = 0.0F;
                if (!"缺勤".equals(d1) && !"缺勤".equals(d2)) {
                    f1 = Float.parseFloat(new DecimalFormat(".0").format((float) getjian(day, d1, d2) / 60));
                    //System.out.println("上午间隔 " + f1);
                }
                if (!"缺勤".equals(d3) && !"缺勤".equals(d4)) {
                    f2 = Float.parseFloat(new DecimalFormat(".0").format((float) getjian(day, d3, d4) / 60));
                    //System.out.println("下午间隔 " + f2);
                }
                if (!"缺勤".equals(d5) && !"缺勤".equals(d6)) {
                    f3 = Float.parseFloat(new DecimalFormat(".0").format((float) getjian(day, d5, d6) / 60));
                    //System.out.println("晚班间隔 " + f3);
                }

                if (f1 > 3.5) {
                    f1 = 3.5F;
                }
                if (f2 > 4 && f2 < 5) {
                    f2 = 4F;
                } else if (f2 > 5) {
                    f2 = (float) Math.floor(f2);
                }
                f3 = (float) Math.floor(f3);

                array.add(day + " " + f1 + " " + f2 + " " + f3);
//                System.out.format("%-15s %-10.1f %-10.1f %-10.1f\n",  day, f1, f2, f3);

                float ff = f1 + f2 + f3 - 7.5F;
                if (ff >= 0) {//如果当天时长大于则算考勤及加班
                    dayF = dayF + ff;
                    dayNumber++;
                } else {//不满一天则换算为加班时间
                    float saf = Float.parseFloat(new DecimalFormat(".0").format((float) ((f1 + f2) / 1.2)));
                    saf = (float) Math.floor(saf);

                    dayF = dayF + saf + f3;
                }
                if (f1 > 0 || f2 > 0 || f3 > 0) {
                    dayChuq++;
                }
            }
            map.put(name, array);
            map1.put(name, Test1.YEAR + " " + name + " " + dayChuq + " " + dayNumber + " " + dayF);
//            System.out.println(name + " 出勤天数 " + dayChuq + " 计算时间 " + dayNumber + " 加班时间 " + dayF);
//            System.out.format("%-15s %-15s %-10d %-10d %-10.1f\n", Test1.YEAR, name, dayChuq, dayNumber, dayF);
        }
        System.out.println(map);
        for (String name : arrayList) {
            String va = map1.get(name);
            System.out.println(va);
            List<String> arrayL = map.get(name);
            for (String s : arrayL) {
                System.out.println(s);
            }
        }
    }

    private static int getjian(String day, String d1, String d2) {

        String dd1 = Test1.YEAR + day + d1;
        DateTime dateTime1 = DateUtil.parse(dd1, "yyyyMMddHH:mm");
        long l1 = dateTime1.toCalendar().getTimeInMillis();

        String dd2 = Test1.YEAR + day + d2;
        DateTime dateTime2 = DateUtil.parse(dd2, "yyyyMMddHH:mm");
        long l2 = dateTime2.toCalendar().getTimeInMillis();

        long l = l2 - l1;
        int i = (int) (l / 1000 / 60);
        return i;
    }

    public static void main1() throws Exception {
        S_Y_P = 0;
        InputStream fileInputStream = new FileInputStream("C:/Work/oa/file/2.81.xls");
        Workbook workbook = WorkbookFactory.create(fileInputStream);

        Sheet childSheet = workbook.getSheetAt(0);
        String name;
        List<String> list;
        for (int index = 7; index < childSheet.getLastRowNum() + 1; index = index + 2) {
            //姓名
            name = Test1.readExcelData(childSheet, index - 1, 10);
//            System.out.println(name);

            Row row = childSheet.getRow(index);
            if (row != null) {
                int kk = row.getLastCellNum();
                for (int i = 0; i < kk; i++) {
                    Cell cell = row.getCell(i);
                    //日期
                    int dayday = i + 1;

                    String day;
                    if (dayday < 10) {
                        day = "0" + dayday;
                    } else {
                        day = "" + dayday;
                    }
//                    System.out.println(day);

                    list = new ArrayList<>();
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING.getCode()) {
                            String string = cell.getStringCellValue();
                            int len = string.length();
                            int lim = len / 5;
                            for (int ii = 0; ii < lim; ii++) {
                                String sss = string.substring(ii * 5, ii * 5 + 5);
                                list.add(sss);
                            }
                            //考勤
//                            System.out.println(list);
                            paList(list);


                            if (name.equals("王扬威")
                                    || name.equals("何仁易")
                                    || name.equals("陈鹏")
                                    || name.equals("陈卫平")) {
                            } else {
                                if ("沈永平".equals(name) && S_Y_P <= kk) {
                                    S_Y_P++;
                                    continue;
                                }

                                String d1 = list.get(0);
                                String d2 = list.get(1);
                                String d3 = list.get(2);
                                String d4 = list.get(3);
                                String d5 = list.get(4);
                                String d6 = list.get(5);

                                if (!"缺勤".equals(d1)
                                        && "缺勤".equals(d2)
                                        && "缺勤".equals(d3)
                                        && "缺勤".equals(d4)
                                        && "缺勤".equals(d5)
                                        && "缺勤".equals(d6)
                                        ) {
                                    d1 = "缺勤";
                                }
                                if (!"缺勤".equals(d1)
                                        && !"缺勤".equals(d2)
                                        && !"缺勤".equals(d3)
                                        && "缺勤".equals(d4)
                                        && "缺勤".equals(d5)
                                        && "缺勤".equals(d6)
                                        ) {
                                    d3 = "缺勤";
                                }

                                try {
                                    Db.use().insert(
                                            Entity.create(DATABASE_NAME_2)
                                                    .set("id", SecureUtil.md5(name + Test1.YEAR + day))
                                                    .set("name", name)
                                                    .set("day", Test1.YEAR + day)
                                                    .set("d1", d1)
                                                    .set("d2", d2)
                                                    .set("d3", d3)
                                                    .set("d4", d4)
                                                    .set("d5", d5)
                                                    .set("d6", d6)
                                                    .set("room", 2));
                                } catch (SQLException e) {
                                    Db.use().update(
                                            Entity.create().set("d1", d1)
                                                    .set("d2", d2)
                                                    .set("d3", d3)
                                                    .set("d4", d4)
                                                    .set("d5", d5)
                                                    .set("d6", d6), //修改的数据
                                            Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                    );
                                }

                                String dd1 = "";
                                if (!"缺勤".equals(d1)) {
                                    dd1 = Test1.YEAR + day + d1;
                                    DateTime dateTime1 = DateUtil.parse(dd1, "yyyyMMddHH:mm");
                                    long l1 = dateTime1.toCalendar().getTimeInMillis();
                                    String dd = Test1.YEAR + day + "08:05";
                                    DateTime dateTime2 = DateUtil.parse(dd, "yyyyMMddHH:mm");
                                    long l2 = dateTime2.toCalendar().getTimeInMillis();
                                    if ((l2 - l1) < 0) {
                                        //缺勤 检查是否为11:30之前的数据
                                        dd = Test1.YEAR + day + "11:30";
                                        dateTime2 = DateUtil.parse(dd, "yyyyMMddHH:mm");
                                        l2 = dateTime2.toCalendar().getTimeInMillis();
                                        if ((l2 - l1) < 0) {
                                            dd1 = "缺勤";
                                            if (!"缺勤".equals(d1)
                                                    && !"缺勤".equals(d2)
                                                    && !"缺勤".equals(d3)
                                                    && !"缺勤".equals(d4)
                                                    && !"缺勤".equals(d5)
                                                    && "缺勤".equals(d6)
                                                    ) {
                                                try {
                                                    Db.use().update(
                                                            Entity.create().set("d1", "11:30").set("d2", d1).set("d3", d2).set("d4", d3).set("d5", d4).set("d6", d5), //修改的数据
                                                            Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                                    );
                                                } catch (SQLException e) {
                                                    e.printStackTrace();
                                                }
                                            }
                                            if (!"缺勤".equals(d1)
                                                    && !"缺勤".equals(d2)
                                                    && !"缺勤".equals(d3)
                                                    && !"缺勤".equals(d4)
                                                    && "缺勤".equals(d5)
                                                    && "缺勤".equals(d6)
                                                    ) {
                                                try {
                                                    Db.use().update(
                                                            Entity.create().set("d1", dd1).set("d2", dd1).set("d3", d1).set("d4", d2).set("d5", d3).set("d6", d4), //修改的数据
                                                            Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                                    );
                                                } catch (SQLException e) {
                                                    e.printStackTrace();
                                                }
                                            }
                                            if (!"缺勤".equals(d1)
                                                    && !"缺勤".equals(d2)
                                                    && "缺勤".equals(d3)
                                                    && "缺勤".equals(d4)
                                                    && !"缺勤".equals(d5)
                                                    && !"缺勤".equals(d6)
                                                    ) {
                                                try {
                                                    Db.use().update(
                                                            Entity.create().set("d1", dd1).set("d2", dd1).set("d3", d1).set("d4", d2), //修改的数据
                                                            Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                                    );
                                                } catch (SQLException e) {
                                                    e.printStackTrace();
                                                }
                                            }

                                            //缺勤 检查是否为17:30之前的数据
                                            dd = Test1.YEAR + day + "17:30";
                                            dateTime2 = DateUtil.parse(dd, "yyyyMMddHH:mm");
                                            l2 = dateTime2.toCalendar().getTimeInMillis();
                                            if ((l2 - l1) < 0) {
                                                if (!"缺勤".equals(d1)
                                                        && !"缺勤".equals(d2)
                                                        && "缺勤".equals(d3)
                                                        && "缺勤".equals(d4)
                                                        && "缺勤".equals(d5)
                                                        && "缺勤".equals(d6)
                                                        ) {
                                                    try {
                                                        Db.use().update(
                                                                Entity.create().set("d1", dd1).set("d2", dd1).set("d3", dd1).set("d4", dd1).set("d5", d1).set("d6", d2), //修改的数据
                                                                Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                                        );
                                                    } catch (SQLException e) {
                                                        e.printStackTrace();
                                                    }
                                                }
                                            }
                                        } else {
                                            //迟到
                                            dd1 = TTT.getCompleteTime(d1);
                                            try {
                                                Db.use().update(
                                                        Entity.create().set("d1", dd1), //修改的数据
                                                        Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                                );
                                            } catch (SQLException e) {
                                                e.printStackTrace();
                                            }
                                        }
                                    } else {
                                        //统一改为八点
                                        dd1 = "08:00";

                                        try {
                                            Db.use().update(
                                                    Entity.create().set("d1", dd1), //修改的数据
                                                    Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                            );
                                        } catch (SQLException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                }

                                if (!"缺勤".equals(d3)) {
                                    long l1 = DateUtil.parse(Test1.YEAR + day + d3, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    long l2 = DateUtil.parse(Test1.YEAR + day + "11:30", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    long l3 = DateUtil.parse(Test1.YEAR + day + "12:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    if (l1 > l2 && l1 < l3) {
                                        d3 = "12:00";
                                        try {
                                            Db.use().update(
                                                    Entity.create().set("d3", d3), //修改的数据
                                                    Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                            );
                                        } catch (SQLException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                }


                                if (!"缺勤".equals(d5)) {
                                    long l1 = DateUtil.parse(Test1.YEAR + day + d5, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    long l2 = DateUtil.parse(Test1.YEAR + day + "17:30", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    long l3 = DateUtil.parse(Test1.YEAR + day + "18:00", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
                                    if (l1 > l2 && l1 < l3) {
                                        d5 = "18:00";
                                        try {
                                            Db.use().update(
                                                    Entity.create().set("d5", d5), //修改的数据
                                                    Entity.create(DATABASE_NAME_2).set("id", SecureUtil.md5(name + Test1.YEAR + day)) //where条件
                                            );
                                        } catch (SQLException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }

        }


    }

    private static void paList(List<String> list) {
        int length = list.size();
        switch (length) {
            case 0:
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                break;
            case 1:
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                break;
            case 2:
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                break;
            case 3:
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                list.add("缺勤");
                break;
            case 4:
                list.add("缺勤");
                list.add("缺勤");
                break;
            case 5:
                list.add("缺勤");
                break;
        }
    }
}
