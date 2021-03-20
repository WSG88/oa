package data;

import java.util.List;

public class Data {
    public String name;
    public String date;
    public List<String> list;
    public int error;

    public Data(String name, String date, List<String> list) {
        this.name = name;
        this.date = date;
        this.list = list;
    }


    @Override
    public String toString() {
        return "Data{" +
                "name='" + name + '\'' +
                ", date='" + date + '\'' +
                ", list=" + list +
                '}';
    }

}
