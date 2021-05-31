package data;

public class DateTimeBean {
    private String name;
    private String datetime;
    private String date;
    private String time;

    public DateTimeBean() {
    }

    public DateTimeBean(String name, String date, String time) {
        this.name = name;
        this.date = date;
        this.time = time;
    }

    public String getGroupBy() {
        return name + date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDatetime() {
        return datetime;
    }

    public void setDatetime(String datetime) {
        this.datetime = datetime;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }

    @Override
    public String toString() {
        return "DateTimeBean{" +
                "name='" + name + '\'' +
                ", datetime='" + datetime + '\'' +
                ", date='" + date + '\'' +
                ", time='" + time + '\'' +
                '}';
    }
}