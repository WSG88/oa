import cn.hutool.core.date.DateTime;
import cn.hutool.core.date.DateUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import data.DataBean;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.*;

public class Test2 {
    public static String DATABASE_NAME_2 = "oatime";
    public static String QUE_QING = "缺勤";

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
        List<String> namesList = getNameList(DATABASE_NAME_2);

        Map<String, List<DataBean>> mapListDataBean = new HashMap<>();

        for (String name : namesList) {
            int checkWorkDay = 0;
            int attendanceDay = 0;
            float overtime = 0F;
            float f1 = 0F;
            float f2 = 0F;
            float f3 = 0F;
            float fff = 0F;
            List<DataBean> arrayDataBean = new ArrayList<>();
            if (name == null) continue;
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
                fff = 0F;
                if (!"缺勤".equals(d1) && !"缺勤".equals(d2)) {
                    f1 = Utils.timeDifference(day + d1, day + d2);
                }
                if (!"缺勤".equals(d3) && !"缺勤".equals(d4)) {
                    f2 = Utils.timeDifference(day + d3, day + d4);
                }
                if (!"缺勤".equals(d5) && !"缺勤".equals(d6)) {
                    f3 = Utils.timeDifference(day + d5, day + d6);
                }

                //晚上加班时间
                if (f3 % 1 >= 0.45) {
                    f3 = (float) Math.floor(f3) + 0.5F;
                } else {
                    f3 = (float) Math.floor(f3);
                }

                float ff = f1 + f2 + f3 - 7.9F;
                if (ff >= 0) {
                    //如果当天时长大于则算考勤及加班
                    if (f1 + f2 - 7.9F > 0) {
                        float dds = f1 + f2 - 7.9F;
                        if ((dds % 1) >= 0.85f) {
                            fff = (float) Math.floor(dds) + 1;
//                        } else if (dds >= 0.4) {
//                            fff = (float) Math.floor(dds) + 0.5F;
                        } else {
                            fff = (float) Math.floor(dds);
                        }
                        overtime = overtime + fff + f3;
                    } else {
                        overtime = overtime + f3 + (f1 + f2 - 8F);
                    }
                    attendanceDay++;
                } else {
                    overtime = overtime + f1 + f2 + f3;
                }
                if (f1 > 0 || f2 > 0 || f3 > 0) {
                    checkWorkDay++;
                }
                DataBean dataBean = new DataBean(name, day, d1, d2, d3, d4, d5, d6, f1, f2, f3);
                arrayDataBean.add(dataBean);
            }
            mapListDataBean.put(name, arrayDataBean);
        }

        saveToExcel(mapListDataBean);
    }

    private static void saveToExcel(Map<String, List<DataBean>> mapListDataBean) {
        List<List<String>> rowsList = new ArrayList<>();
        List<String> dayList = new ArrayList<>();
        List<String> blankList = new ArrayList<>();
        for (int i = 1; i < 32; i++) {
            String dd;
            if (i < 10) {
                dd = "2020080" + i;
            } else {
                dd = "202008" + i;
            }
            dayList.add(dd);
            blankList.add(" ");
        }
        rowsList.add(dayList);

        Iterator iterator = mapListDataBean.keySet().iterator();
        while (iterator.hasNext()) {
            String name = (String) iterator.next();
            List<String> nameList = new ArrayList<>();
            nameList.add(name + "  ");
            List<DataBean> arrayDataBean1 = mapListDataBean.get(name);
            float n = 0;
            float c = 0;
            for (DataBean dataBean : arrayDataBean1) {
                c += dataBean.c();
                n += dataBean.n();
            }
            nameList.add(c + "天 ");
            nameList.add(Utils.getDecimals(n) + "小时  ");
            rowsList.add(nameList);

            List<String> timeList1 = new ArrayList<>();
            List<String> timeList2 = new ArrayList<>();
            List<String> timeList3 = new ArrayList<>();
            List<String> timeList4 = new ArrayList<>();
            List<String> timeList5 = new ArrayList<>();
            List<String> timeList6 = new ArrayList<>();
            List<String> timeListAm = new ArrayList<>();
            List<String> timeListPm = new ArrayList<>();
            List<String> timeListNm = new ArrayList<>();
            List<String> timeListA = new ArrayList<>();
            List<String> timeListP = new ArrayList<>();
            List<String> timeListN = new ArrayList<>();
            List<String> timeListC = new ArrayList<>();

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

                            timeListA.add(dataBean.a() == 0 ? " " : dataBean.a() + "");
                            timeListP.add(dataBean.p() == 0 ? " " : dataBean.p() + "");
                            timeListN.add(dataBean.n() == 0 ? " " : dataBean.n() + "");
                            timeListC.add(dataBean.c() == 0 ? " " : dataBean.c() + "");
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

                    timeListA.add("");
                    timeListP.add("");
                    timeListN.add("");
                    timeListC.add("");
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
            rowsList.add(timeListA);
            rowsList.add(timeListP);
            rowsList.add(timeListN);
            rowsList.add(timeListC);
            rowsList.add(blankList);
            rowsList.add(blankList);
        }


        Utils.toExcel(rowsList, "2车间", "C:\\Work\\2车间" + new Date().getTime() + ".xlsx");
    }

    private static List<String> getNameList(String db) throws SQLException {
        //        List<Entity> entities = Db.use().findAll(db);
        List<Entity> entities = Db.use().query("select name from " + db + " group by name");
        List<String> arrayList = new ArrayList<>();
        for (Entity entity : entities) {
            String name = entity.getStr("name");
            if (name != null && name.length() > 0) {
                arrayList.add(name);
            }
        }
        return arrayList;
    }


    public static void main1() throws Exception {
        InputStream fileInputStream = new FileInputStream("C:\\Work\\oa\\file\\2.81.xls");
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet childSheet = workbook.getSheetAt(0);
        String name;
        List<String> list;
        for (int index = 7; index < childSheet.getLastRowNum() + 1; index = index + 2) {
            //姓名
            name = Test1.readExcelData(childSheet, index - 1, 10);

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

                    list = new ArrayList<>();
                    if (cell != null) {
                        if (cell.getCellTypeEnum() == CellType.STRING) {
                            String string = cell.getStringCellValue();
                            int len = string.length();
                            int lim = len / 5;
                            for (int ii = 0; ii < lim; ii++) {
                                String sss = string.substring(ii * 5, ii * 5 + 5);
                                list.add(sss);
                            }
                            paList(list);

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
                                        dd1 = Utils.getCompleteTime(d1);
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
