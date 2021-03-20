package kq;

import cn.hutool.db.Db;
import cn.hutool.db.Entity;

import java.sql.SQLException;
import java.util.List;

public class Kq {
    public static void main(String[] args) throws SQLException {
        List<Entity> list= Db.use().findAll("sys_user");
        System.out.println(list);

        Db.use().findAll(Entity.create("sys_user").set("name", "unitTestUser"));
    }
}
