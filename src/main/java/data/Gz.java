package data;

import java.util.Objects;

public class Gz {
    String name;
    String sale;
    String month;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSale() {
        return sale;
    }

    public void setSale(String sale) {
        this.sale = sale;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    @Override
    public String toString() {
        return name + ',' + month + ',' + sale;
    }

    public String getString() {
        return "{" +
                "" + name +
                "," + sale +
                '}';
    }


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Gz gz = (Gz) o;
        return Objects.equals(name, gz.name) && Objects.equals(sale, gz.sale) && Objects.equals(month, gz.month);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, sale, month);
    }
}
