package javaExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;

/**
 * Created by USER on 07.02.2017.
 */
public class comparison {
    public static SimpleDateFormat sdf = new SimpleDateFormat("YYYY.MM.DD");
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream(("D:\\example.xls"));
        Workbook wb = new HSSFWorkbook(fis);
        for (Row row:wb.getSheetAt(0)){
            for (Cell cell:row){

                System.out.println(getCellText(cell));
            }
        }
        for (Row row:wb.getSheetAt(1)){
            for (Cell cell:row){

                System.out.println(getCellText(cell));
            }
        }
        fis.close();
    }

    public static String getCellText(Cell cell) {
        String result = "";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = sdf.format(cell.getDateCellValue().toString());
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_BLANK:
                System.out.println();
                break;
            default:
                break;
        }
        return result;
    }
}


