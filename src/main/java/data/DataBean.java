package data;

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
    public int room = 2;
    public int error;

    public DataBean(String name, String day, String d1, String d2, String d3, String d4, String d5, String d6,
                    float am, float pm, float nm, int room) {
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
