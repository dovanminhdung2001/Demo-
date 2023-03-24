package excel1.pmud;

import excel1.DateUtils;
import excel1.pmud.model.dto.CheckInDTO;
import excel1.pmud.model.req.FilterCheckinReq;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Nv_s_Day {
    public static final int COL_USER_ID = 0;
    public static final int COL_USERNAME = 1;
    public static final int COL_CHECKIN = 2;
    public static final int COL_OUT1 = 3;
    public static final int COL_IN1 = 4;
    public static final int COL_OUT2 = 5;
    public static final int COL_IN2 = 6;
    public static final int COL_OUT3 = 7;
    public static final int COL_IN3 = 8;
    public static final int COL_CHECKOUT = 9;
    public static final int COL_CHECKIN_LATE = 10;
    public static final int COL_CHECKOUT_EARLY = 11;
    public static final int COL_GO_OUT_AMOUNT = 12;
    public static final int COL_GO_OUT_TIME = 13;
    public static final int COL_WORK_TIME = 14;
    private static CellStyle cellStyleFormatNumber = null;
    public static int rowIndex =0;
    public static void main(String[] args) {
        FilterCheckinReq req = new FilterCheckinReq("2022-01-01", "2022-01-01", "1");
        List<CheckInDTO>  listDto = List.of(
                new CheckInDTO(1, "user 1", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09"),
                new CheckInDTO(2, "user 2", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09"),
                new CheckInDTO(3, "user 3", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09")
        );
        CheckInDTO checkInDTO = new CheckInDTO(1, "user 1", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09");
        String excelPath = String.valueOf(new StringBuffer("D://BangChamCong_DS_").append(req.getFrom()).append(".xlsx"));

        try {
            Workbook workbook = getWorkbook(excelPath);
            Sheet sheet = workbook.createSheet("Sheet1");

            writeHeader(sheet, rowIndex, req);

            writeTable(sheet);

            writeContent(sheet, listDto);

            int numberOfColumn = sheet.getRow(1).getPhysicalNumberOfCells();
            autosizeColumn(sheet, numberOfColumn);

            workbook.write(new FileOutputStream(excelPath));
            workbook.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Workbook getWorkbook(String excelFilePath) throws IOException {
        if (excelFilePath.endsWith("xlsx"))
            return new XSSFWorkbook();

        if (excelFilePath.endsWith("xls"))
            return new HSSFWorkbook();

        throw new IllegalArgumentException("The specified file is not Excel file");
    }

    private static void writeHeader(Sheet sheet, int rowIndex, FilterCheckinReq req) {
        Font font1 = sheet.getWorkbook().createFont();
        font1.setBold(true);
        font1.setFontHeightInPoints((short) 16); // font size

        // Create CellStyle
        CellStyle cellStyle1 = sheet.getWorkbook().createCellStyle();
        cellStyle1.setFont(font1);

        Row row = sheet.createRow(++rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellStyle(cellStyle1);
        cell.setCellValue("Công ty TNHH Dịch vụ Đào tạo Thiên Ưng");

        int firstCol = 0;
        int lastCol = 15;
        for (int i = rowIndex; i < rowIndex + 5; i++) {
            int firstRow = i;
            int lastRow = i;
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }

        Font font2 = sheet.getWorkbook().createFont();
        font2.setFontHeightInPoints((short) 12);

        CellStyle cellStyle2 = sheet.getWorkbook().createCellStyle();
        cellStyle2.setFont(font2);

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("Đ/C: Lô B11 - Ngõ 233 - Xuân Thủy - Cầu giấy");

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("ĐT: Mr Nam 0984.322.539  | Web: Ketoanthienung.net");

        row = sheet.createRow(rowIndex + 2);
        cell = row.createCell(0);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("Ngày " + req.getFrom());
    }

    private static void writeTable(Sheet sheet) {
        rowIndex = 7;

        CellStyle cellStyle = createStyleForTable(sheet);
        Row row = sheet.createRow(rowIndex);
        String[] tblHeader = {"User id", "Name", "Checkin", "Go out 1", "Go in 1", "Go out 2", "Go in 2", "Go out 3", "Go in 3", "Checkout", "Checkin late", "Checkout early", "Go out turns", "Go out time", "Work time"};

        for (int i = 0; i < tblHeader.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(tblHeader[i]);
        }
    }

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


    private static void writeContent(Sheet sheet, List<CheckInDTO> list) {
        rowIndex = 8;

        for (CheckInDTO dto : list) {
            Row row = sheet.createRow(rowIndex);

            Cell cell = row.createCell(COL_USER_ID);
            cell.setCellValue(dto.getUserId());

            cell = row.createCell(COL_USERNAME);
            cell.setCellValue(dto.getUsername());

            cell = row.createCell(COL_OUT1);
            cell.setCellValue(dto.getGoOut1());

            cell = row.createCell(COL_IN1);
            cell.setCellValue(dto.getGoIn1());

            cell = row.createCell(COL_OUT2);
            cell.setCellValue(dto.getGoOut2());

            cell = row.createCell(COL_IN2);
            cell.setCellValue(dto.getGoIn2());

            cell = row.createCell(COL_OUT3);
            cell.setCellValue(dto.getGoOut3());

            cell = row.createCell(COL_IN3);
            cell.setCellValue(dto.getGoIn3());

            cell = row.createCell(COL_CHECKOUT);
            cell.setCellValue(dto.getCheckout());

            cell = row.createCell(COL_CHECKIN_LATE);
            cell.setCellValue(dto.getCheckinLate());

            cell = row.createCell(COL_CHECKOUT_EARLY);
            cell.setCellValue(dto.getCheckoutEarly());

            cell = row.createCell(COL_GO_OUT_AMOUNT);
            cell.setCellValue(dto.getGoOutAmount());

            cell = row.createCell(COL_GO_OUT_TIME);
            cell.setCellValue(dto.getGoOutTime());

            cell = row.createCell(COL_WORK_TIME);
            cell.setCellValue(dto.getWorkTime());

            rowIndex ++;
        }
    }

    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }
}
