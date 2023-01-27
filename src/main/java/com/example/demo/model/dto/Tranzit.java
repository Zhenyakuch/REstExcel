package com.example.demo.model.dto;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Tranzit {

    public static int summa_tonn2;
    public static int summa_pos_ed2;
    public static int summa_sht2;
    public static int summa_part2;
    public static int summa_m22;
    public static int summa_m32;
    public static int summa_wagons2;
    public static int summa_transport2;
    public static int summa_container2;
    public static int summa_baggage2;
    public static int summa_airplane2;


    public static void create_obl(int cellCount, XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle,
                                  XSSFCellStyle cellStyleRow, List<ElementRegion> regions, List<NamePoints> namePoints) {

        XSSFCellStyle cell_styl_obl = xssfWorkbook.createCellStyle();
        cell_styl_obl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setAlignment(CellStyle.ALIGN_CENTER);
        cell_styl_obl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
        cell_styl_obl.setFont(font);

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast2 = sheet.getLastRowNum();

        XSSFCell cellName_obl = row.createCell(0);
        cellName_obl.setCellStyle(cell_styl_obl);

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

        ElementRegion elementRegion = regions.get(cellCount);
        switch (elementRegion.getRegion()) {
            case 1:
                cellName_obl.setCellValue("БРЕСТСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    create_rows(sheet, namePoints, cellCount, cellStyle, cellStyleRow);
                }
                plus(xssfWorkbook, sheet, namePoints, cellStyleRow);

                break;
            case 2:
                cellName_obl.setCellValue("ВИТЕБСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    create_rows(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellStyleRow);

                break;
            case 3:
                cellName_obl.setCellValue("ГОМЕЛЬСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    create_rows(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellStyleRow);

                break;

            case 4:
                cellName_obl.setCellValue("ГРОДНЕНСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    create_rows(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellStyleRow);

                break;
            case 5:
                cellName_obl.setCellValue("МИНСКАЯ ОБЛАСТЬ");
                for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                    create_rows(sheet, namePoints, cellCount, cellStyle, cellStyleRow);

                }
                plus(xssfWorkbook, sheet, namePoints, cellStyleRow);

                break;
        }

    }


    public static void plus(XSSFWorkbook xssfWorkbook, XSSFSheet sheet, List<NamePoints> namePoints, XSSFCellStyle cellStyleRow) {

        XSSFCellStyle cell_styl_obl = xssfWorkbook.createCellStyle();
        cell_styl_obl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
        cell_styl_obl.setFont(font);

        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName = row2.createCell(0);
        cellName.setCellStyle(cell_styl_obl);
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


        for (NamePoints namePoint : namePoints) {
            int tonn = namePoint.getTonn();
            int pos_ed = namePoint.getUnits();
            int sht = namePoint.getPieces();
            int part = namePoint.getParties();
            int m2 = namePoint.getM2_packages();
            int m3 = namePoint.getM3();
            int wagons = namePoint.getWagons();
            int transport = namePoint.getMotor_transport();
            int container = namePoint.getContainers();
            int baggage = namePoint.getBaggage();
            int airplane = namePoint.getAirplane();

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

        summa_tonn2 = summa_tonn2 + summa_tonn;
        summa_pos_ed2 = summa_pos_ed2 + summa_pos_ed;
        summa_sht2 = summa_sht2 + summa_sht;
        summa_part2 = summa_part2 + summa_part;
        summa_m22 = summa_m22 + summa_m2;
        summa_m32 = summa_m32 + summa_m3;
        summa_wagons2 = summa_wagons2 + summa_wagons;
        summa_transport2 = summa_transport2 + summa_transport;
        summa_container2 = summa_container2 + summa_container;
        summa_baggage2 = summa_baggage2 + summa_baggage;
        summa_airplane2 = summa_airplane2 + summa_airplane;

    }

    public static void create_rows(XSSFSheet sheet, List<NamePoints> namePoints, int k, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow) {

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
        cellThousand_m2.setCellValue(namePoints.get(k).getM2_packages());
        cellThousand_m3.setCellValue(namePoints.get(k).getM3());
        cellRailway_wagons.setCellValue(namePoints.get(k).getWagons());
        cellMotor_transport.setCellValue(namePoints.get(k).getMotor_transport());
        cellContainer.setCellValue(namePoints.get(k).getContainers());
        cellBaggage.setCellValue(namePoints.get(k).getBaggage());
        cellAirplane.setCellValue(namePoints.get(k).getAirplane());

    }

    public static void plusAll(XSSFWorkbook xssfWorkbook, XSSFSheet sheet, XSSFCellStyle cellStyleRow, int summa_tonn,
                               int summa_pos_ed, int summa_sht, int summa_part, int summa_m2, int summa_m3,
                               int summa_wagons, int summa_transport, int summa_container, int summa_baggage,
                               int summa_airplane) {

        XSSFCellStyle cell_styl_obl = xssfWorkbook.createCellStyle();
        cell_styl_obl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
        cell_styl_obl.setFont(font);

        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName = row2.createCell(0);
        cellName.setCellStyle(cell_styl_obl);
        cellName.setCellValue("ВСЕГО ПО РБ: ");

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

    public static void plus_eaeu(CountryReport countryRequest, CountryRow countryRow, XSSFWorkbook xssfWorkbook,
                                 XSSFSheet sheet, XSSFCellStyle cellStyleRow) {

        XSSFCellStyle cell_styl_obl = xssfWorkbook.createCellStyle();
        cell_styl_obl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setWrapText(true);//перенос слов

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        font.setBold(true);
        cell_styl_obl.setFont(font);

        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName = row2.createCell(0);
        cellName.setCellStyle(cell_styl_obl);
        cellName.setCellValue("В том числе в страны ЕАЭС: ");

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

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().get(0).getRegions().size(); cellCount++) {

            for (int k = 0; k < countryRow.getRegions().get(cellCount).getNamePoints().size(); k++) {

                if (countryRow.getRegions().get(cellCount).getNamePoints().get(k).getName().equals("Россия") |
                        countryRow.getRegions().get(cellCount).getNamePoints().get(k).getName().equals("Казахстан") |
                        countryRow.getRegions().get(cellCount).getNamePoints().get(k).getName().equals("Кыргызстан") |
                        countryRow.getRegions().get(cellCount).getNamePoints().get(k).getName().equals("Армения")) {

                    summa_tonn = summa_tonn + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getTonn();
                    summa_pos_ed = summa_pos_ed + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getUnits();
                    summa_sht = summa_sht + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getPieces();
                    summa_part = summa_part + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getParties();
                    summa_m2 = summa_m2 + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getM2_packages();
                    summa_m3 = summa_m3 + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getM3();
                    summa_wagons = summa_wagons + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getWagons();
                    summa_transport = summa_transport + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getMotor_transport();
                    summa_container = summa_container + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getContainers();
                    summa_baggage = summa_baggage + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getBaggage();
                    summa_airplane = summa_airplane + countryRow.getRegions().get(cellCount).getNamePoints().get(k).getAirplane();
                }
            }
        }
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

    public static void nullable() {
        summa_tonn2=0;
        summa_pos_ed2=0;
        summa_sht2=0;
        summa_part2=0;
        summa_m22=0;
        summa_m32=0;
        summa_wagons2=0;
        summa_transport2=0;
        summa_container2=0;
        summa_baggage2=0;
        summa_airplane2=0;
    }
}

