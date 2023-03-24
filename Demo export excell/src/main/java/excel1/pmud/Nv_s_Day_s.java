package excel1.pmud;

import excel1.DateUtils;
import excel1.pmud.model.req.FilterCheckinReq;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.text.ParseException;
import java.util.Date;

public class Nv_s_Day_s {
    private static CellStyle cellStyleFormatNumber = null;
    public static int rowIndex =0;
    public static void main(String[] args) throws ParseException, IOException {
        FilterCheckinReq req = new FilterCheckinReq("2022-12-25", "2023-01-09", "1");
        Date fromDate = DateUtils.sdf.parse(req.getFrom());
        Date toDate = DateUtils.sdf.parse(req.getTo());
        String excelPath = String.valueOf(new StringBuffer("D://BangChamCong_DS_").append(req.getFrom()).append(".xlsx"));

        Long totalDay = (toDate.getTime() / 1000 - fromDate.getTime() / 1000) / 3600 / 24;
//        Date[totalDay] arrDa
        Workbook workbook = getWorkbook(excelPath);
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

}
