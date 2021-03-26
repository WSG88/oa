package kq;

import cn.hutool.core.date.DateUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.db.Db;
import cn.hutool.db.Entity;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//处理人脸考勤机数据插入数据库
public class Kq {
    public static void main(String[] args) throws Exception {
        InputStream in = new FileInputStream("F:\\WORK\\oa\\file\\Allevent.xlsx");
        Workbook wbs = WorkbookFactory.create(in);
        Sheet childSheet = wbs.getSheetAt(0);
        List<Entity> list = new ArrayList<>();

        Map<String, TimeData> map = new HashMap<>();
        for (int index = 1; index < childSheet.getLastRowNum() + 1; index++) {
            TimeData timeData = new TimeData();
            timeData.setEmployee(readExcel(childSheet, index, 0).replace("'", "").trim());
            timeData.setName(readExcel(childSheet, index, 2).replace("'", "").trim());
            timeData.setTime(readExcel(childSheet, index, 3).replace("'", "").trim() + ":00");
            timeData.setMd(SecureUtil.md5(timeData.getEmployee() + timeData.getTime()));
            map.put(timeData.getMd(), timeData);
        }
        for (TimeData timeData : map.values()) {
            Entity entity = Entity.create("employee_time")
                    .set("employee", timeData.getEmployee())
                    .set("name", timeData.getName())
                    .set("time", timeData.getTime())
                    .set("date_time", DateUtil.parseDateTime(timeData.getTime()))
                    .set("md_id", timeData.getMd());
            list.add(entity);
        }
        System.out.println(list);
        try {
            //Db.use().insert(entity);
            Db.use().insert(list);
        } catch (SQLException e) {
        }

        System.out.println("导入数据结束");

    }

    public static String readExcel(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        String s = "";
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                switch (cell.getCellTypeEnum()) {
                    case _NONE:
                        break;
                    case NUMERIC:
                        s = String.valueOf(cell.getNumericCellValue());
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            //s = NumberToTextConverter.toText(cell.getNumericCellValue());
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
                            s = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                        }
                        break;
                    case STRING:
                        s = cell.getStringCellValue();
                        break;
                    case FORMULA:
                        DecimalFormat df = new DecimalFormat("0.00");
                        s = df.format(cell.getNumericCellValue());
                        break;
                    case BLANK:
                        break;
                    case BOOLEAN:
                        s = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case ERROR:
                        break;

                }
                //System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + s);
            }
        }
        return s;
    }
}
