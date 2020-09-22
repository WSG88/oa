import cn.hutool.core.date.DateTime;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;

import java.sql.SQLException;
import java.util.List;

public class Fs {
    public static String DB_NAME="oatime";

    //不参加考勤
    //正常上下班
    public static void main(String[] args) {
        try {
            List<Entity> list = Db.use().findAll(Entity.create(DB_NAME).set("name", "程泉华"));
            list = Db.use().find(Entity.create(DB_NAME).set("d1", "like 11%"));
//            list = Db.use().findAll(DB_NAME);


            for (int i = 0; i < list.size(); i++) {
                Entity entity = list.get(i);
                System.out.println(entity);

                String name = entity.getStr("name");
                String day = entity.getStr("day");
                String d1 = entity.getStr("d1");
                String d2 = entity.getStr("d2");
                String d3 = entity.getStr("d3");
                String d4 = entity.getStr("d4");
                String d5 = entity.getStr("d5");
                String d6 = entity.getStr("d6");


                if (!"缺勤".equals(d1)) {
                    DateTime dt1 = new DateTime(day + d1, "yyyyMMddHH:mm");
                    long l1 = dt1.toCalendar().getTimeInMillis();
                    System.out.println(l1);

                    DateTime dt2 = new DateTime(day + "08:05", "yyyyMMddHH:mm");
                    long l2 = dt2.toCalendar().getTimeInMillis();
                    System.out.println(l2);
                    long l3 = (l2 - l1) / 1000 / 60;
                    System.out.println("l2-l1=" + l3);
                    if (l3 < 0) {
                        System.out.println(name + day + "上午迟到" + Math.abs(l3) + "分钟");
                    }
                }

                if (!"缺勤".equals(d1) && !"缺勤".equals(d2)) {
                    DateTime dt1 = new DateTime(day + d1, "yyyyMMddHH:mm");
                    long l1 = dt1.toCalendar().getTimeInMillis();
                    System.out.println(l1);
                    DateTime dt2 = new DateTime(day + d2, "yyyyMMddHH:mm");
                    long l2 = dt2.toCalendar().getTimeInMillis();
                    System.out.println(l2);
                    long l3 = (l2 - l1) / 1000 / 60;
                    System.out.println("l2-l1=" + l3);
                    System.out.println(name + day + "上午时长" + Math.abs(l3) + "分钟");
                }

            }


        } catch (SQLException e) {
            e.printStackTrace();
        }

    }
}
