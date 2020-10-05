import data.Utils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WWW {
    public static void main(String[] args) throws Exception {

        Utils.FILE_PATH = "C:\\Work\\oa\\file\\";
        Utils.FILE_NAME = "在制品清单.xls";

        Workbook wbs = Utils.getWorkbook();
        int cnt = wbs.getNumberOfSheets();
        for (int jj = 0; jj < cnt; jj++) {
            Sheet childSheet = wbs.getSheetAt(jj);

            Font font = wbs.getFontAt(row.getCell(k).getCellStyle().getFontIndex());

            System.out.println(childSheet.getSheetName());
            for (int index = 1; index < childSheet.getLastRowNum() + 1; index++) {
                for (int i = 0; i < 10; i++) {
                    try {
                        Utils.readExcelData(childSheet, index, i);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    try {
                        Utils.readExcelDataGetNumericCellValue(childSheet, index, i);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }

            }
        }


    }
}
