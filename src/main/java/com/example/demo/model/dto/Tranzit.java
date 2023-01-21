package com.example.demo.model.dto;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Tranzit {

    public static void createRows(int cellCount, XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, List<ElementRegion> regions, List<NamePoints> namePoints) {

        XSSFCellStyle cellStylobl = xssfWorkbook.createCellStyle();
        cellStylobl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//        cellStylobl.setAlignment(CENTER); // стиль выравнивания по горизонтали
//        cellStylobl.setVerticalAlignment(CENTER); // стиль выравивание по вертикали
        cellStylobl.setAlignment(CellStyle.ALIGN_CENTER);
        cellStylobl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
//        font.
        cellStylobl.setFont(font);

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast2 = sheet.getLastRowNum();

        XSSFCell cellName_obl = row.createCell(0);
        cellName_obl.setCellStyle(cellStylobl);

        XSSFCell cellName_obl1 = row.createCell(1);
        cellName_obl1.setCellStyle(cellStyle);
        XSSFCell cellName_obl2 = row.createCell(2);
        cellName_obl2.setCellStyle(cellStyle);
        XSSFCell cellName_obl3 = row.createCell(3);
        cellName_obl3.setCellStyle(cellStyle);
        XSSFCell cellName_obl4 = row.createCell(4);
        cellName_obl4.setCellStyle(cellStyle);
        XSSFCell cellName_obl5 = row.createCell(5);
        cellName_obl5.setCellStyle(cellStyle);
        XSSFCell cellName_obl6 = row.createCell(6);
        cellName_obl6.setCellStyle(cellStyle);
        XSSFCell cellName_obl7 = row.createCell(7);
        cellName_obl7.setCellStyle(cellStyle);
        XSSFCell cellName_obl8 = row.createCell(8);
        cellName_obl8.setCellStyle(cellStyle);
        XSSFCell cellName_obl9 = row.createCell(9);
        cellName_obl9.setCellStyle(cellStyle);
        XSSFCell cellName_obl10 = row.createCell(10);
        cellName_obl10.setCellStyle(cellStyle);
        XSSFCell cellName_obl11 = row.createCell(11);
        cellName_obl11.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast2, rowLast2, 0, 11));

        // for (int k = cellCount; regions.size() > k; k++) {
//            for (int k = 0; namePoints.size() > k; k++) {
        ElementRegion elementRegion = regions.get(cellCount);
        switch (elementRegion.getRegion()) {
            case 1:
                cellName_obl.setCellValue("БРЕСТСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    cicle(sheet, namePoints, cellCount, cellStyle, cellStyleRow);
                }
                plus(xssfWorkbook, sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                break;
            case 2:
                cellName_obl.setCellValue("ВИТЕБСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    cicle(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                break;
            case 3:
                cellName_obl.setCellValue("ГОМЕЛЬСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    cicle(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                break;

            case 4:
                cellName_obl.setCellValue("ГРОДНЕНСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    cicle(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                break;
            case 5:
                cellName_obl.setCellValue("МИНСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    cicle(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                break;

//            }
        }
    }

    private static void plus(XSSFWorkbook xssfWorkbook, XSSFSheet sheet, List<NamePoints> namePoints, int k, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow) {

        XSSFCellStyle cellStylobl = xssfWorkbook.createCellStyle();
        cellStylobl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStylobl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
//        font.
        cellStylobl.setFont(font);

        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName = row2.createCell(0);
        cellName.setCellStyle(cellStylobl);
        cellName.setCellValue("ИТОГО: ");

        XSSFCell cellThousand_t = row2.createCell(1);
        cellThousand_t.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_pos_ed = row2.createCell(2);
        cellThousand_pos_ed.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_sht = row2.createCell(3);
        cellThousand_sht.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_part = row2.createCell(4);
        cellThousand_part.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_m2 = row2.createCell(5);
        cellThousand_m2.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_m3 = row2.createCell(6);
        cellThousand_m3.setCellStyle(cellStyleRow);
        XSSFCell cellRailway_wagons = row2.createCell(7);
        cellRailway_wagons.setCellStyle(cellStyleRow);
        XSSFCell cellMotor_transport = row2.createCell(8);
        cellMotor_transport.setCellStyle(cellStyleRow);
        XSSFCell cellContainer = row2.createCell(9);
        cellContainer.setCellStyle(cellStyleRow);
        XSSFCell cellBaggage = row2.createCell(10);
        cellBaggage.setCellStyle(cellStyleRow);
        XSSFCell cellAirplane = row2.createCell(11);
        cellAirplane.setCellStyle(cellStyleRow);

        // cellThousand_t.setCellType(XSSFCell.CELL_TYPE_FORMULA);
        int summa_tonn = 0;
        int summa_pos_ed = 0;
        int summa_sht = 0;
        int summa_part = 0;
        int summa_m2 = 0;
        int summa_m3 = 0;
        int summa_wagons = 0;
        int summa_transport = 0;
        int summa_container = 0;
        int summa_baggage = 0;
        int summa_airplane = 0;

        for (k = 0; namePoints.size() > k; k++) {

            int tonn = namePoints.get(k).getTonn();
            int pos_ed = namePoints.get(k).getUnits();
            int sht = namePoints.get(k).getPieces();
            int part = namePoints.get(k).getParties();
            int m2 = namePoints.get(k).getM2();
            int m3 = namePoints.get(k).getM3();
            int wagons = namePoints.get(k).getWagons();
            int transport = namePoints.get(k).getMotor_transport();
            int container = namePoints.get(k).getContainers();
            int baggage = namePoints.get(k).getBaggage();
            int airplane = namePoints.get(k).getAirplane();

            summa_tonn = summa_tonn + tonn;
            summa_pos_ed = summa_pos_ed + pos_ed;
            summa_sht = summa_sht + sht;
            summa_part = summa_part + part;
            summa_m2 = summa_m2 + m2;
            summa_m3 = summa_m3 + m3;
            summa_wagons = summa_wagons + wagons;
            summa_transport = summa_transport + transport;
            summa_container = summa_container + container;
            summa_baggage = summa_baggage + baggage;
            summa_airplane = summa_airplane + airplane;

            cellThousand_t.setCellValue(summa_tonn);
            cellThousand_pos_ed.setCellValue(summa_pos_ed);
            cellThousand_sht.setCellValue(summa_sht);
            cellThousand_part.setCellValue(summa_part);
            cellThousand_m2.setCellValue(summa_m2);
            cellThousand_m3.setCellValue(summa_m3);
            cellRailway_wagons.setCellValue(summa_wagons);
            cellMotor_transport.setCellValue(summa_transport);
            cellContainer.setCellValue(summa_container);
            cellBaggage.setCellValue(summa_baggage);
            cellAirplane.setCellValue(summa_airplane);


        }


    }

    public static void cicle(XSSFSheet sheet, List<NamePoints> namePoints, int k, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow) {
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast2 = sheet.getLastRowNum();
        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName_points = row2.createCell(0);
        cellName_points.setCellStyle(cellStyle);

        XSSFCell cellThousand_t = row2.createCell(1);
        cellThousand_t.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_pos_ed = row2.createCell(2);
        cellThousand_pos_ed.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_sht = row2.createCell(3);
        cellThousand_sht.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_part = row2.createCell(4);
        cellThousand_part.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_m2 = row2.createCell(5);
        cellThousand_m2.setCellStyle(cellStyleRow);
        XSSFCell cellThousand_m3 = row2.createCell(6);
        cellThousand_m3.setCellStyle(cellStyleRow);
        XSSFCell cellRailway_wagons = row2.createCell(7);
        cellRailway_wagons.setCellStyle(cellStyleRow);
        XSSFCell cellMotor_transport = row2.createCell(8);
        cellMotor_transport.setCellStyle(cellStyleRow);
        XSSFCell cellContainer = row2.createCell(9);
        cellContainer.setCellStyle(cellStyleRow);
        XSSFCell cellBaggage = row2.createCell(10);
        cellBaggage.setCellStyle(cellStyleRow);
        XSSFCell cellAirplane = row2.createCell(11);
        cellAirplane.setCellStyle(cellStyleRow);

        cellName_points.setCellValue(namePoints.get(k).getName());
        cellThousand_t.setCellValue(namePoints.get(k).getTonn());
        cellThousand_pos_ed.setCellValue(namePoints.get(k).getUnits());
        cellThousand_sht.setCellValue(namePoints.get(k).getPieces());
        cellThousand_part.setCellValue(namePoints.get(k).getParties());
        cellThousand_m2.setCellValue(namePoints.get(k).getM2());
        cellThousand_m3.setCellValue(namePoints.get(k).getM3());
        cellRailway_wagons.setCellValue(namePoints.get(k).getWagons());
        cellMotor_transport.setCellValue(namePoints.get(k).getMotor_transport());
        cellContainer.setCellValue(namePoints.get(k).getContainers());
        cellBaggage.setCellValue(namePoints.get(k).getBaggage());
        cellAirplane.setCellValue(namePoints.get(k).getAirplane());

    }
}
