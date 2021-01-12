package cn.iscgm.poi.utils;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author cgm
 */
public class ExcelUtils {
    /**
     * 获取包含指定元格内容的行号集合
     *
     * @param sheet
     * @param cellContent
     * @return 行号集合
     */
    public static List<Integer> findRowNum(XSSFSheet sheet, String cellContent) {
        List<Integer> list = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                        Integer num = row.getRowNum();
                        list.add(num);
                    }
                }
            }
        }
        return list;
    }

    /**
     * 获取指定单元格的字符内容
     *
     * @param fis
     * @param sheetNum
     * @param fileName
     * @param rowNum
     * @param cellNum
     * @return 单元格内容字符串
     * @throws EncryptedDocumentException
     * @throws IOException
     */
    public static String getStringCellContent(FileInputStream fis, int
            sheetNum, String fileName, Integer rowNum, Integer cellNum) throws EncryptedDocumentException, IOException {
        DataFormatter formatter = new DataFormatter();
        String cellContent;
        if (fileName.endsWith(".xlsx")) {
            //获取XSSFWorkbook对象
            XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis);
            XSSFSheet sheet = workbook.getSheetAt(sheetNum - 1);
            XSSFRow row = sheet.getRow(rowNum);
            XSSFCell cell = row.getCell(cellNum);
            cellContent = formatter.formatCellValue(cell);
        } else {
            //获取HSSFWorkbook对象
            HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(fis);
            HSSFSheet sheet = workbook.getSheetAt(sheetNum - 1);
            HSSFRow row = sheet.getRow(rowNum);
            HSSFCell cell = row.getCell(cellNum);
            cellContent = formatter.formatCellValue(cell);
        }

        if (cellContent == null) {
            System.err.println(fileName + rowNum + "-" + cellNum + "单元格内容为空！");
        }
        return cellContent;
    }

    /**
     * 获取excel中的图片存入map
     *
     * @param fis      excel文件输入流
     * @param sheetNum 表格编号
     * @param fileName excel文件名
     * @return Map<String, PictureData> key:行号-列号 value:PictureData
     * @throws EncryptedDocumentException
     */
    public static Map<String, PictureData> getPicture(FileInputStream fis, int
            sheetNum, String fileName) throws EncryptedDocumentException, IOException {
        Map<String, PictureData> map = new HashMap<>();
        if (fileName.endsWith(".xlsx")) {
            XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis);
            XSSFSheet sheet = workbook.getSheetAt(sheetNum - 1);
            if (sheet != null) {
                List<POIXMLDocumentPart> list = sheet.getRelations();
                for (POIXMLDocumentPart part : list) {
                    if (part instanceof XSSFDrawing) {
                        XSSFDrawing drawing = (XSSFDrawing) part;
                        List<XSSFShape> shapes = drawing.getShapes();
                        for (XSSFShape shape : shapes) {
                            XSSFPicture picture = (XSSFPicture) shape;
                            XSSFClientAnchor anchor = picture.getPreferredSize();
                            CTMarker marker = anchor.getFrom();
                            String key = marker.getRow() + "-" + marker.getCol();
                            map.put(key, picture.getPictureData());
                        }
                    }
                }
            }
        } else {
            //获取HSSFWorkbook对象
            HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(fis);
            //获取图片HSSFPictureData集合
            List<HSSFPictureData> pictures = workbook.getAllPictures();
            //获取当前表编码所对应的表
            HSSFSheet sheet = workbook.getSheetAt(sheetNum - 1);
            if (sheet != null) {
                //对表格进行操作
                for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                    HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                    if (shape instanceof HSSFPicture) {
                        HSSFPicture pic = (HSSFPicture) shape;
                        //获取行编号
                        int row = anchor.getRow2();
                        //获取列编号
                        int col = anchor.getCol2();
                        int pictureIndex = pic.getPictureIndex() - 1;
                        HSSFPictureData picData = pictures.get(pictureIndex);
                        map.put(row + "-" + col, picData);
                    }
                }
            } else {
                System.err.println(fileName + "图片为空！");
            }
        }
        return map;
    }

    /**
     * 查找包含指定单元格为空的一个空行行号
     *
     * @param sheet   表
     * @param cellNum 单元格（列）号
     * @return rowNum 行号；如果返回-1则代表未找到
     */
    public static Integer findNullRow(XSSFSheet sheet, Integer cellNum) {
        Integer rowNum = -1;
        for (Row row : sheet) {
            Cell cell = row.getCell(cellNum);
            if (cell.getCellType() == CellType.BLANK) {
                rowNum = cell.getRow().getRowNum();
            }

        }
        return rowNum;

    }

    /**
     * 查找包含指定单元格内容且包含另一指定单元格号为空的行的行号集合
     *
     * @param sheet       表
     * @param cellNum     指定单元格行号
     * @param cellContent 指定单元格内容
     * @return 行号集合
     */
    public static List<Integer> findNullRows(XSSFSheet sheet, Integer cellNum, String cellContent) {
        List<Integer> list = new ArrayList<>();
        List<Integer> rowNums = findRowNum(sheet, cellContent);
        for (Integer rowNum : rowNums
        ) {
            Row row = sheet.getRow(rowNum);
            Cell emptyCell = row.getCell(cellNum);
            if (emptyCell.getCellType() == CellType.BLANK) {
                list.add(row.getRowNum());
            }
        }

        return list;
    }
}