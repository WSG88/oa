import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelUtil {
    public static String readExcelData(Sheet childSheet, int rowNum, int cellNum) throws Exception {
        Row row = childSheet.getRow(rowNum);
        if (row != null) {
            Cell cell = row.getCell(cellNum);
            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    return "";
                } else {
//                    System.out.println("第" + (rowNum + 1) + "行" + "第" + (cellNum + 1) + "列的值： " + cell.getStringCellValue());
                    return cell.getStringCellValue();
                }
            }
        }
        return "";
    }
}
