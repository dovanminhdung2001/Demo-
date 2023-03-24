package excel1;

import excel1.Entity.BookEntity;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class Demo1 {
    public static final int COL_ID = 0;
    public static final int COL_NAME = 1;
    public static final int COL_PRICE = 2;
    public static final int COL_TOTAL = 3;
    private static CellStyle cellStyleFormatNumber = null;

    public static void main(String[] args) {
        List<BookEntity> list = List.of(new BookEntity[]{
                new BookEntity(1L, "b1", 1.0),
                new BookEntity(2L, "b2asdfgfhgj.,waesdfhgjasdfgsadf à à à", 2.0),
                new BookEntity(3L, "b3", 3.0)
        });
        String excelPath = String.valueOf(new StringBuffer("E://dsBanHangTheoSanPham-").append(DateUtils.dateUpFile()).append(".xlsx"));

        try {
            Workbook workbook = getWorkbook(excelPath);

            // Create sheet
            Sheet sheet = workbook.createSheet("Books"); // Create sheet with sheet name

            int rowIndex = 0;
            // write  header
            writeHeader(sheet, rowIndex);

            rowIndex = 12;
            // Write table header
            writeTable(sheet, rowIndex);

            // Write data
            rowIndex++;
            for (BookEntity book : list) {
                // Create row
                Row row = sheet.createRow(rowIndex);
                // Write data on row
                writeBook(book, row);
                rowIndex++;
            }

            // Write footer
            writeFooter(sheet, rowIndex);

            // Auto resize column witdth
            int numberOfColumn = sheet.getRow(1).getPhysicalNumberOfCells();
            autosizeColumn(sheet, numberOfColumn);

            // Create file excel
            createOutputFile(workbook, excelPath);
            System.out.println("Export Done!!!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Workbook getWorkbook(String excelFilePath) throws IOException {
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }
    // Write header
    private static void writeHeader(Sheet sheet, int rowIndex) {
        Font font1 = sheet.getWorkbook().createFont();
        font1.setBold(true);
        font1.setFontHeightInPoints((short) 16); // font size

        // Create CellStyle
        CellStyle cellStyle1 = sheet.getWorkbook().createCellStyle();
        cellStyle1.setFont(font1);
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);

        Row row = sheet.createRow(++rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellStyle(cellStyle1);
        cell.setCellValue("Báo cáo bán hàng theo sản phẩm");

        int firstCol = 0;
        int lastCol = 7;
        for (int i = rowIndex; i < rowIndex + 4; i++) {
            int firstRow = i;
            int lastRow = i;
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }

        Font font2 = sheet.getWorkbook().createFont();
        font2.setBold(true);
        font2.setFontHeightInPoints((short) 12);

        CellStyle cellStyle2 = sheet.getWorkbook().createCellStyle();
        cellStyle2.setFont(font2);
        cellStyle2.setAlignment(HorizontalAlignment.CENTER);

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("Cửa hàng: Shop KhangBaby");

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("Địa chỉ: Đâu đó ở Bắc Giang");
    }


    // Write table with format
    private static void writeTable(Sheet sheet, int rowIndex) {
        // create CellStyle
        CellStyle cellStyle = createStyleForTable(sheet);

        // Create row
        Row row = sheet.createRow(rowIndex);

        // Create cells
        Cell cell = row.createCell(COL_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Id");

        cell = row.createCell(COL_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Name");

        cell = row.createCell(COL_PRICE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Price");

        cell = row.createCell(COL_TOTAL);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Total");
    }

    // Write data
    private static void writeBook(BookEntity book, Row row) {
        if (cellStyleFormatNumber == null) {
            // Format number
            short format = (short)BuiltinFormats.getBuiltinFormat("#,##0");
            // DataFormat df = workbook.createDataFormat();
            // short format = df.getFormat("#,##0");

            //Create CellStyle
            Workbook workbook = row.getSheet().getWorkbook();
            cellStyleFormatNumber = workbook.createCellStyle();
            cellStyleFormatNumber.setDataFormat(format);
        }

        Cell cell = row.createCell(COL_ID);
        cell.setCellValue(book.getId());

        cell = row.createCell(COL_NAME);
        cell.setCellValue(book.getName());

        cell = row.createCell(COL_PRICE);
        cell.setCellValue(book.getPrice());
        cell.setCellStyle(cellStyleFormatNumber);

        // Create cell formula
        // totalMoney = price * quantity
        cell = row.createCell(COL_TOTAL, CellType.FORMULA);
        cell.setCellStyle(cellStyleFormatNumber);
        int currentRow = row.getRowNum() + 1;
        String columnPrice = CellReference.convertNumToColString(COL_PRICE);
        String columnId = CellReference.convertNumToColString(COL_ID);
        cell.setCellFormula(columnPrice + currentRow + "*" + columnId + currentRow);
    }

    // Create CellStyle for header
    private static CellStyle createStyleForTable(Sheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setColor(IndexedColors.BLACK.getIndex()); // text color

        // Create CellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        return cellStyle;
    }

    // Write footer
    private static void writeFooter(Sheet sheet, int rowIndex) {
        // Create row
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(COL_TOTAL, CellType.FORMULA);
        cell.setCellFormula("SUM(E2:E6)");
    }

    // Auto resize column width
    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    // Create output file
    private static void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
        try (OutputStream os = new FileOutputStream(excelFilePath)) {
            workbook.write(os);
        }
    }
}