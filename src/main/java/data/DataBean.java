package data;

import cn.hutool.core.date.DateUtil;

public class DataBean {

    public String name;
    public String day;
    public String d1;
    public String d2;
    public String d3;
    public String d4;
    public String d5;
    public String d6;

    public float am;
    public float pm;
    public float nm;
    public int isLate;
    public int isLeaveEarly;
    public int isAbsence;

    public DataBean(String name, String day, String d1, String d2, String d3, String d4, String d5, String d6, float am, float pm, float nm) {
        this.name = name;
        this.day = day;
        this.d1 = d1;
        this.d2 = d2;
        this.d3 = d3;
        this.d4 = d4;
        this.d5 = d5;
        this.d6 = d6;
        this.am = am;
        this.pm = pm;
        this.nm = nm;
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
                ", am=" + am +
                ", pm=" + pm +
                ", nm=" + nm +
                ", isLate=" + isLate +
                ", isLeaveEarly=" + isLeaveEarly +
                ", isAbsence=" + isAbsence +
                '}';
    }

    public float a() {
        if (!"缺勤".equals(d1)) {
            long l1 = DateUtil.parse(day + d1, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
            long l2 = DateUtil.parse(day + "08:05", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
            float l = (l1 - l2) / 1000 / 60 / 60;
            if (l > 0) {
                if (l % 1 <= 0.5F) {
                    l = (float) Math.floor(l) + 0.5F;
                } else {
                    l = (float) Math.floor(l) + 1F;
                }
                return l;
            }
        }
        return 0;
    }

    public float p() {
        if (!"缺勤".equals(d3)) {
            long l1 = DateUtil.parse(day + d3, "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
            long l2 = DateUtil.parse(day + "13:05", "yyyyMMddHH:mm").toCalendar().getTimeInMillis();
            float l = (l1 - l2) / 1000 / 60 / 60;
            if (l > 0) {
                if (l % 1 <= 0.5F) {
                    l = (float) Math.floor(l) + 0.5F;
                } else {
                    l = (float) Math.floor(l) + 1F;
                }
                return l;
            }
        }
        return 0;
    }

    public float m(float nm) {
        if (nm % 1 >= 0.5F) {
            return (float) Math.floor(nm) + 0.5F;
        } else {
            return (float) Math.floor(nm);
        }
    }

    float NIGHT = 7.95F;

    public float n() {
        if (am + pm > NIGHT && pm > 4.5F) {
            return (float) Math.floor(pm - 4.5F) + m(nm);
        } else {
            if (nm + am + pm - NIGHT > 0) {
                return m(nm + am + pm - NIGHT);
            } else {
                return nm + am + pm;
            }
        }
    }

    public int c() {
        if (am + pm > NIGHT) {
            return 1;
        }
        if (nm + am + pm - NIGHT > 0) {
            return 1;
        }
        return 0;
    }
}
