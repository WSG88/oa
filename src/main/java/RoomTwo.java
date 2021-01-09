import data.Data;
import data.Utils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class RoomTwo {

    public static void main(String[] args) throws Exception {
        Utils.clearList();
        Utils.ROOM = 2;
        Utils.YEAR_MONTH = "202012";
        Utils.FILE_PATH = "d:\\Work\\oa\\file\\";
        Utils.FILE_NAME = "2.12.xls";
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
                        int ll = 0;
                        if (string.length() > 30) {
                            ll = len / 7;
                        } else {
                            ll = len / 5;
                        }
                        for (int ii = 0; ii < ll; ii++) {
                            String sss = string.substring(ii * 5, ii * 5 + 5);
                            list.add(sss);
                        }
                    }
                    //培训
                    if ("20201226".equals(date)) {
                        if (
                                "	蓝天航	".trim().equals(name) ||
                                        "	肖立军	".trim().equals(name) ||
                                        "	程宇文	".trim().equals(name) ||
                                        "	危发强	".trim().equals(name) ||
                                        "	陈震宇	".trim().equals(name) ||
                                        "	刘小龙	".trim().equals(name) ||
                                        "	万立妹	".trim().equals(name) ||
                                        "	万运来	".trim().equals(name) ||
                                        "	曹明光	".trim().equals(name) ||
                                        "	张康文	".trim().equals(name) ||
                                        "	葛银保	".trim().equals(name) ||
                                        "	张新丽	".trim().equals(name) ||
                                        "	王思刚	".trim().equals(name) ||
                                        "	苏金丽	".trim().equals(name) ||
                                        "	程厚阳	".trim().equals(name) ||
                                        "	严命坤	".trim().equals(name) ||
                                        "	黄亦龙	".trim().equals(name) ||
                                        "	王章美	".trim().equals(name) ||
                                        "	江鹏	".trim().equals(name) ||
                                        "	虞涛	".trim().equals(name) ||
                                        "	陈亚军	".trim().equals(name) ||
                                        "	张志旗	".trim().equals(name) ||
                                        "	陈妍	".trim().equals(name) ||
                                        "	徐靖	".trim().equals(name) ||
                                        "	苏芳华	".trim().equals(name) ||
                                        "	刘镇	".trim().equals(name) ||
                                        "	张学成	".trim().equals(name) ||
                                        "	方兴兴	".trim().equals(name) ||
                                        "	沈长征	".trim().equals(name) ||
                                        "	侯木财	".trim().equals(name) ||
                                        "	王金宝	".trim().equals(name) ||
                                        "	蒋婵	".trim().equals(name) ||
                                        "	罗飞	".trim().equals(name) ||
                                        "	邵泽球	".trim().equals(name) ||
                                        "	张涛	".trim().equals(name) ||
                                        "	张祖胜	".trim().equals(name) ||
                                        "	张欢	".trim().equals(name) ||
                                        "	李开立	".trim().equals(name) ||
                                        "	周谟林	".trim().equals(name) ||
                                        "	杜传国	".trim().equals(name) ||
                                        "	刘太华	".trim().equals(name) ||
                                        "	许凌玉	".trim().equals(name) ||
                                        "	徐慧林	".trim().equals(name) ||
                                        "	苏俊潇	".trim().equals(name) ||
                                        "	章水根	".trim().equals(name) ||
                                        "	方智鑫	".trim().equals(name) ||
                                        "	郑彦俊	".trim().equals(name) ||
                                        "	张志成	".trim().equals(name)
                        ) {
                            list.clear();
                            list.add("07:59");
                            list.add("11:31");
                            list.add("11:59");
                            list.add("16:31");
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
