import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import data.Data;
import data.DataBean;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.*;

public class Test2 {
    public static String DATABASE_NAME_2 = "oatime";
    public static String FILE_PATH = "C:\\Work\\oa\\file\\";
    public static String YEAR_MONTH = "202008";
    public static String FILE_NAME = "2.81.xls";
    public static int ROOM = 2;

    public static List<String> arrayNamesList = new ArrayList<>();

    public static void main(String[] args) throws Exception {
        setData();
        getData();
    }

    public static void getData() throws Exception {
        LinkedHashMap<String, List<DataBean>> mapListDataBean = new LinkedHashMap<>();
        for (String name : arrayNamesList) {
            List<DataBean> arrayDataBean = new ArrayList<>();
            List<Entity> listEntity = Db.use().findAll(Entity.create(DATABASE_NAME_2).set("name", name));
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
            mapListDataBean.put(name, arrayDataBean);
        }

        saveToExcel(mapListDataBean);
    }

    private static void saveToExcel(Map<String, List<DataBean>> mapListDataBean) {
        List<List<String>> rowsList = new ArrayList<>();

        List<String> dayList = Utils.getDaysOfMonth(YEAR_MONTH);
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

        Utils.toExcel(rowsList, ROOM + "车间", FILE_PATH +
                ROOM + "车间" + "_" + YEAR_MONTH + "_" + new Date().getTime() + ".xlsx");
    }


    public static void setData() throws Exception {
        arrayNamesList.clear();
        InputStream fileInputStream = new FileInputStream(FILE_PATH + FILE_NAME);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet childSheet = workbook.getSheetAt(0);
        List<String> list;
        for (int index = 7; index < childSheet.getLastRowNum() + 1; index = index + 2) {
            //姓名
            String name = Test1.readExcelData(childSheet, index - 1, 10);
            if (name == null || name.length() == 0) {
                continue;
            }
            arrayNamesList.add(name);
            Row row = childSheet.getRow(index);
            if (row != null) {
                int kk = row.getLastCellNum();
                for (int i = 0; i < kk; i++) {
                    Cell cell = row.getCell(i);
                    //日期
                    String day;
                    int j = i + 1;
                    if (j < 10) {
                        day = "0" + j;
                    } else {
                        day = "" + j;
                    }
                    String date = YEAR_MONTH + day;

                    //考勤记录
                    list = new ArrayList<>();
                    if (cell != null) {
                        if (cell.getCellTypeEnum() == CellType.STRING) {
                            String string = cell.getStringCellValue();
                            int len = string.length();
                            int ll = len / 5;
                            for (int ii = 0; ii < ll; ii++) {
                                String sss = string.substring(ii * 5, ii * 5 + 5);
                                list.add(sss);
                            }
                            //检查考勤数据是否完整

//                            //早上多打卡
//                            if (list.size() > 1) {
//                                int l = 0;
//                                for (String s : list) {
//                                    if (s.startsWith("07")) {
//                                        l++;
//                                    }
//                                }
//                                if (l > 1) {
//                                    System.out.println(name + "   " + date + " " + list);
//                            return;
//                                }
//                            }
//                            //打卡次数缺失
//                            if (list.size() == 1 || list.size() == 3 || list.size() == 5) {
//                                System.out.println(name + "   " + date + " " + list);
//                            return;
//                            }

//                            //是否请假
//                            if (list.size() == 2) {
//                                System.out.println(name + "   " + date + " " + list);
//                            } else if (list.size() == 4) {
//                                if (!(list.get(0).startsWith("07") || list.get(0).startsWith("08"))) {
//                                    System.out.println(name + "   " + date + " " + list);
//                                }
//                            }

                            //清空一次打卡
                            if (list.size() == 1) {
                                list.clear();
                                continue;
                            }
                            //补全考勤数据
                            Utils.completeQueQing(list);
                            //检查各区间值是否正常

                            Data data = new Data(name, date, list);
//                            System.out.println(data.toString());

                            saveToDatabase(data, ROOM);

                        }

                    }

                }
            }

        }


    }

    private static void saveToDatabase(Data data, int room) throws SQLException {
        String name = data.name;
        String date = data.date;
        List<String> list = data.list;

        String d1 = list.get(0);
        String d2 = list.get(1);
        String d3 = list.get(2);
        String d4 = list.get(3);
        String d5 = list.get(4);
        String d6 = list.get(5);

        if (!"缺勤".equals(d1)) {
            d1 = Utils.getFirstTime(date, d1);
        }
        if (!"缺勤".equals(d3)) {
            d3 = Utils.getFirstTime(date, d3);
        }
        if (!"缺勤".equals(d5)) {
            d5 = Utils.getFirstTime(date, d5);
        }
        if (!"缺勤".equals(d6)) {
            d6 = Utils.getLastCompleteTime(d6);
        }
        String id = SecureUtil.md5(name + date);

        float f1 = 0F;
        float f2 = 0F;
        float f3 = 0F;
        if (!"缺勤".equals(d1) && !"缺勤".equals(d2)) {
            f1 = Utils.timeDifference(date + d1, date + d2);
        }
        if (!"缺勤".equals(d3) && !"缺勤".equals(d4)) {
            f2 = Utils.timeDifference(date + d3, date + d4);
        }
        if (!"缺勤".equals(d5) && !"缺勤".equals(d6)) {
            f3 = Utils.timeDifference(date + d5, date + d6);
        }

        try {
            //插入数据
            Db.use().insert(
                    Entity.create(DATABASE_NAME_2)
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
                    Entity.create(DATABASE_NAME_2).set("id", id)
            );
        }
    }


}
