import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class RoomTwo {

    public static void main(String[] args) throws Exception {
        Utils.clearList();
        Utils.ROOM = 2;
        Utils.YEAR_MONTH = "202010";
        Utils.FILE_PATH = "d:\\Work\\oa\\file\\";
        Utils.FILE_NAME = "2.10.xls";
        Utils.clear();

        Workbook wbs = Utils.getWorkbook();
        Sheet childSheet = wbs.getSheetAt(0);

        List<Data> dataArrayList = new ArrayList<>();
        for (int index = 5; index < childSheet.getLastRowNum() + 1; index = index + 2) {
            //姓名
            String name = Utils.readExcelData(childSheet, index - 1, 10);
            if (!Utils.isAdd(name)) {
                continue;
            }
            Utils.add(name);
            Row row = childSheet.getRow(index);
            if (row != null) {
                int kk = row.getLastCellNum();
                for (int i = 0; i < kk; i++) {
                    Cell cell = row.getCell(i);
                    //日期
                    String date = Utils.YEAR_MONTH + String.format("%02d", i + 1);
                    //考勤记录
                    List<String> list = new ArrayList<>();
                    if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
                        String string = cell.getStringCellValue();
                        int len = string.length();
                        int ll = len / 5;
                        for (int ii = 0; ii < ll; ii++) {
                            String sss = string.substring(ii * 5, ii * 5 + 5);
                            list.add(sss);
                        }
                    }
                    if (!list.isEmpty()) {
                        //补全考勤数据
                        Utils.completeQueQing(list);
                        dataArrayList.add(new Data(name, date, list));
                    }
                }
            }
        }

        //计算并保存
        Utils.getData(Utils.arrayNamesList, dataArrayList);

        Utils.printList();
    }
}
