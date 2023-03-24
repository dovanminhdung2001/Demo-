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

public class Nv_Day {
    private static CellStyle cellStyleFormatNumber = null;
    public static int rowIndex =0;
    public static void main(String[] args) {
        FilterCheckinReq req = new FilterCheckinReq("2022-01-01", "2022-01-01", "1");
        CheckInDTO checkInDTO = new CheckInDTO(1, "user 1", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09");
        String excelPath = String.valueOf(new StringBuffer("D://BangChamCong_").append(DateUtils.dateUpFile()).append(".xlsx"));

        try {
            Workbook workbook = getWorkbook(excelPath);
            Sheet sheet = workbook.createSheet("Sheet1");

            writeHeader(sheet, rowIndex, req);
            writeContent(sheet, checkInDTO);

            int numberOfColumn = sheet.getRow(1).getPhysicalNumberOfCells();
            autosizeColumn(sheet, numberOfColumn);

            workbook.write(new FileOutputStream(excelPath));
            System.out.println("Export Done!!!");
        } catch (Exception e) {
            e.printStackTrace();
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
        cell.setCellValue("Ngày" + req.getFrom());
    }

    private static void writeContent(Sheet sheet, CheckInDTO dto) {
        rowIndex = 6;
        Row row = sheet.createRow(++rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue("User id");
        cell = row.createCell(1);
        cell.setCellValue(dto.getUserId());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("User name");
        cell = row.createCell(1);
        cell.setCellValue(dto.getUsername());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Checkin");
        cell = row.createCell(1);
        cell.setCellValue(dto.getCheckin());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go out 1");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoOut1());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go in 1");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoIn1());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go out 2");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoOut2());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go in 2");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoIn2());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go out 3");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoOut3());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go in 3");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoIn3());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Check out");
        cell = row.createCell(1);
        cell.setCellValue(dto.getCheckout());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Checkin late");
        cell = row.createCell(1);
        cell.setCellValue(dto.getCheckinLate().equals("     00:-1") ? "" : dto.getCheckinLate());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Checkout early");
        cell = row.createCell(1);
        cell.setCellValue(dto.getCheckoutEarly().equals("     00:-1") ? "" : dto.getCheckoutEarly());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go out amount");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoOutAmount());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Go out time");
        cell = row.createCell(1);
        cell.setCellValue(dto.getGoOutTime());

        row = sheet.createRow(++rowIndex);
        cell = row.createCell(0);
        cell.setCellValue("Work time");
        cell = row.createCell(1);
        cell.setCellValue(dto.getWorkTime());
    }

    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }
}