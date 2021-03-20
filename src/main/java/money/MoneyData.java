package money;

public class MoneyData {

    public String 序号;
    public String 姓名;
    public double 应出勤天数;
    public double 实出勤天数;
    public double 基本工资;
    public double 技术工资;
    public double 满勤工资;
    public double 岗位津贴;
    public double 工龄工资;
    public double 应付工资小计;
    //补贴款项;
    public double 加班补贴;
    public double 夜班补贴;
    public double 操机补贴;
    public double 绩效奖励;
    public double 住宿补贴;
    public double 差旅补贴;
    public double 高温补贴;
    public double 劳保补贴;
    public double 补贴小计;
    //代扣款项;
    public double 缺勤工资;
    public double 养老险8;
    public double 医疗险2;
    public double 失业险;
    public double 个人所得税;
    public double 质量扣款;
    public double 绩效处罚;
    public double 电费;
    public double 代扣款小计;

    public double 实付工资;
    public String 备注;

    @Override
    public String toString() {
        return "MoneyData{" +
                "序号='" + 序号 + '\'' +
                ", 姓名='" + 姓名 + '\'' +
                ", 应出勤天数=" + 应出勤天数 +
                ", 实出勤天数=" + 实出勤天数 +
                ", 基本工资=" + 基本工资 +
                ", 技术工资=" + 技术工资 +
                ", 满勤工资=" + 满勤工资 +
                ", 岗位津贴=" + 岗位津贴 +
                ", 工龄工资=" + 工龄工资 +
                ", 应付工资小计=" + 应付工资小计 +
                ", 加班补贴=" + 加班补贴 +
                ", 夜班补贴=" + 夜班补贴 +
                ", 操机补贴=" + 操机补贴 +
                ", 绩效奖励=" + 绩效奖励 +
                ", 住宿补贴=" + 住宿补贴 +
                ", 差旅补贴=" + 差旅补贴 +
                ", 高温补贴=" + 高温补贴 +
                ", 劳保补贴=" + 劳保补贴 +
                ", 补贴小计=" + 补贴小计 +
                ", 缺勤工资=" + 缺勤工资 +
                ", 养老险8=" + 养老险8 +
                ", 医疗险2=" + 医疗险2 +
                ", 失业险=" + 失业险 +
                ", 个人所得税=" + 个人所得税 +
                ", 质量扣款=" + 质量扣款 +
                ", 绩效处罚=" + 绩效处罚 +
                ", 电费=" + 电费 +
                ", 代扣款小计=" + 代扣款小计 +
                ", 实付工资=" + 实付工资 +
                ", 备注='" + 备注 + '\'' +
                '}';
    }

    public static double parseDouble(String str) {
        try {
            return Double.parseDouble(str);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
}
