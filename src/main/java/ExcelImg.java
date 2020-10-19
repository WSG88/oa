import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelImg {

    public static void getDataFromExcel(String filePath) throws Exception {
        InputStream in = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(in);
        Sheet sheet = workbook.getSheetAt(7);
        Map<String, PictureData> pictureDataMap = getPictureDataMap(sheet);
        setPictureDataMap(pictureDataMap);
    }

    private static Map<String, PictureData> getPictureDataMap(Sheet sheet) throws IOException {
        Map<String, PictureData> pictureDataMap = null;
        if (sheet instanceof HSSFSheet) {
            pictureDataMap = getHSSFSheetPictures((HSSFSheet) sheet);
        }
        if (sheet instanceof XSSFSheet) {
            pictureDataMap = getXSSFSheetPictures((XSSFSheet) sheet);
        }
        return pictureDataMap;
    }

    public static Map<String, PictureData> getHSSFSheetPictures(HSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<>();
        List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = (HSSFClientAnchor) picture.getAnchor();
                PictureData pdata = picture.getPictureData();
                String key = cAnchor.getRow1() + "-" + cAnchor.getCol1(); // 行号-列号
                map.put(key, pdata);
            }
        }
        return map;
    }

    public static Map<String, PictureData> getXSSFSheetPictures(XSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<>();
        List<POIXMLDocumentPart> list = sheet.getRelations();
        for (POIXMLDocumentPart part : list) {
            if (part instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) part;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker marker = anchor.getFrom();
                    String key = marker.getRow() + "-" + marker.getCol();// 行号-列号
                    map.put(key, picture.getPictureData());
                }
            }
        }
        return map;
    }

    public static void setPictureDataMap(Map<String, PictureData> pictureDataMap) throws IOException {
        Object key[] = pictureDataMap.keySet().toArray();
        for (int i = 0; i < pictureDataMap.size(); i++) {
            // 获取图片流
            PictureData pic = pictureDataMap.get(key[i]);
            // 获取图片索引
            String picName = key[i].toString();
            // 获取图片格式
            String ext = pic.suggestFileExtension();
            byte[] data = pic.getData();
            //图片保存路径
            FileOutputStream out = new FileOutputStream("D:\\pic" + picName + "." + ext);
            out.write(data);
            out.close();
        }
    }

    public static void main(String[] args) throws Exception {
        getDataFromExcel("D:\\ERP&MES\\天一指令\\239指令\\001-0500-06 盖板.xls");
    }
}
