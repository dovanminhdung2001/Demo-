package excel1.pmud;

import excel1.DateUtils;
import excel1.pmud.model.dto.CheckInDTO;
import excel1.pmud.model.req.FilterCheckinReq;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.TreeSet;

public class Nv_Day {
    private static CellStyle cellStyleFormatNumber = null;
    public static int rowIndex =0;
    public static void main(String[] args) {
        test();
//        FilterCheckinReq req = new FilterCheckinReq("2022-01-01", "2022-01-01", "1");
//        CheckInDTO checkInDTO = new CheckInDTO(1, "user 1", "2023-01-09 01:39:32", "", "", "", "", "", "", "2023-01-09 01:49:41", "     00:-1", "16:10:18", 0, "     00:00", "     10:09");
//        String excelPath = String.valueOf(new StringBuffer("E://BangChamCong_").append(DateUtils.dateUpFile()).append(".xlsx"));
//
//        try {
//            Workbook workbook = getWorkbook(excelPath);
//            Sheet sheet = workbook.createSheet("Sheet1");
//
//            writeHeader(sheet, rowIndex, req);
//            writeContent(sheet, checkInDTO);
//
//            String[] positionArr = { "0", "1", "2", "3", "4"} ;
//
//            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
//            DataValidationConstraint positionConstraint = validationHelper.createExplicitListConstraint(positionArr);
//            CellRangeAddressList positionColRange = new CellRangeAddressList(3, 10, 5, 5);
//            DataValidation positionValidation = validationHelper.createValidation(positionConstraint, positionColRange);
//            sheet.addValidationData(positionValidation);
//
//            int numberOfColumn = sheet.getRow(1).getPhysicalNumberOfCells();
//            autosizeColumn(sheet, numberOfColumn);
//
//            workbook.write(new FileOutputStream(excelPath));
//            System.out.println("Export Done!!!");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
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

    private static void test() {

        String[] countryName = {
                "DVDT - Đơn vị đào tạo",
                "DVCN - Đơn vị chức năng",
                "DVHTDT - Đơn vị hỗ trợ đào tạo",
                "BGH - Ban Giám Hiệu",
                "HDT - Hội đồng trường",
                "KTTTTT - Khoa kinh tếTTT",
                "KNN - Khoa ngoại ngữ",
                "KKTTC - Khoa Kinh tế - Tài chính",
                "KQTKD - Khoa Quản trị kinh doanh",
                "BMQTCL - Tổ Quản trị chiến lược",
                "BMQTNL - Tổ bộ môn Quản trị nhân lực",
                "BMQTDA - Tổ bộ môn Quản trị dự án",
                "KDLKS@ - Khoa Du lịch khách sạn",
                "TIEN - Phòng của tiến",
                "BXH - Ban xã hội",
                "PCT# - Phòng Lap",
                "TKDH - Khoa thiết kế đồ hoạ",
                "PCT1 - Khoa kinh tế",
                "TV - Thư viện",
                "KL - Khoa Luật",
                "PCTSV - Phòng công tác sinh viên",
                "PBTS - Phòng ban tuyển sinh",
                "PTT - Phòng truyền thông",
                "HT - Hiệu trưởng",
                "PHT - Phó Hiệu trưởng",
                "PQHDN - Phòng Quan hệ Doanh nghiệp",
                "Laosedu - lao edu",
                "May - MayLao",
                "BT - Bí thư",
                "Unit_test - Unit_test",
                "KKKK - KKKK",
                "DV_con1 - DV_con1",
                "BTN - Ban Tự Nhiên",
                "SAO - Student Affair Office",
                "ddz1 - dungdemo1",
                "ddz1 - dungdemo1",
                "ddz1 - dungdemo1",
                "TIEN2 - Phòng của tiến 2",
                "PKHTC - Phòng kế hoạch tài chính",
                "QTKD - Khoa QTKD",
                "KNN1 - Khoa ngoại ngữ",
                "KHOANN - Khoa ngôn ngữ Anh",
                "KTT1 - Khoa kinh tế",
                "TVC - Thư viện con"
        };

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet realSheet = workbook.createSheet("realSheet");

        XSSFSheet hidden = workbook.createSheet("hidden");

        for (int i = 0, length= countryName.length; i < length; i++) {
            String name = countryName[i];
            XSSFRow row = hidden.createRow(i);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(name);
        }

        DataValidation dataValidation = null;
        DataValidationConstraint constraint = null;
        DataValidationHelper validationHelper = null;
        validationHelper=new XSSFDataValidationHelper(realSheet);
        CellRangeAddressList addressList = new  CellRangeAddressList(0,0,0,0);
        constraint =validationHelper.createFormulaListConstraint("hidden!$A$1:$A$" + countryName.length);
        dataValidation = validationHelper.createValidation(constraint, addressList);
        dataValidation.setSuppressDropDownArrow(true);
        workbook.setSheetHidden(1, true);
        realSheet.addValidationData(dataValidation);
        FileOutputStream stream = null;
        try {
            stream = new FileOutputStream("E:\\testExcel.xlsx");
            workbook.write(stream) ;
            stream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}