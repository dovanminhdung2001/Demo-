package excel1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Demo_function {
    public static void main(String[] args) throws IOException {
        String excelPath = String.valueOf(new StringBuffer("E://Demo-function-").append(DateUtils.dateUpFile()).append(".xlsx"));
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("sheet1");
        int rowIndex = 0;
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("SOLID_FOREGROUND");

        rowIndex = 2;
        cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        row = sheet.createRow(rowIndex);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("SOLID_FOREGROUND row");
        row.setRowStyle(cellStyle);

        workbook.write(new FileOutputStream(excelPath));
        workbook.close();
    }
}
