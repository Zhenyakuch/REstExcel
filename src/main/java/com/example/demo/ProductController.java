package com.example.demo;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.CountryReport;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.OutputStream;

@RestController
@Slf4j
public class ProductController {

    @PostMapping("/product")
    public String printData(@RequestBody CountryReport countryRow) throws Exception {
        log.debug("CountryReport " + countryRow);

        System.out.println(countryRow.getCountryRows().get(0).getCountry());
        System.out.println(countryRow.getCountryRows().get(0).getMassProduct().getDateWeight());
        System.out.println(countryRow.getCountryRows().get(0).getMassProduct().getWeekWeight());
        System.out.println(countryRow.getCountryRows().get(0).getRegions().get(0).getDateWeight());
        System.out.println(countryRow.getCountryRows().get(0).getRegions().get(0).getWeekWeight());
        System.out.println();
        TotalMass totalMass = new TotalMass(countryRow);
        log.debug("TotalMass " + totalMass);


        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/Ввоз.xlsx");
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        String fromTitle = sheet.getRow(0).getCell(0).toString();
        fromTitle.replace("product", countryRow.getProduct());
        fromTitle.replace("startDate", countryRow.getStarDate().toString());
        fromTitle.replace("endDate", countryRow.getEndDate().toString());
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

//        int rowLast = sheet.getLastRowNum();
//        XSSFRow row = sheet.createRow(rowLast+1);
//        XSSFRow row2 = sheet.getRow(rowLast+1);// возможно это лишнее, но я просто пробовала думала вдруг поможет.
//        // посмотри плиз своим взглядом что тут в логике не так, что он не заполняет строку. но ошибок не выдает
//        XSSFCell cellCountry = row2.createCell(0);
//        XSSFCell cellMassProductDateWeight = row2.createCell(1);
//        XSSFCell cellMassProductWeekWeight = row2.createCell(2);
//        XSSFCell cellRegionsBrestDateWeight = row2.createCell(3);
//        XSSFCell cellRegionsBrestWeekWeight = row2.createCell(4);
//        XSSFCell cellRegionsVitebskDateWeight = row2.createCell(5);
//        XSSFCell cellRegionsVitebskWeekWeight = row2.createCell(6);
//        XSSFCell cellRegionsGomelDateWeight = row2.createCell(7);
//        XSSFCell cellRegionsGomelWeekWeight = row2.createCell(8);
//        XSSFCell cellRegionsGrodnoDateWeight = row2.createCell(9);
//        XSSFCell cellRegionsGrodnoWeekWeight = row2.createCell(10);
//        XSSFCell cellRegionsMinskDateWeight = row2.createCell(11);
//        XSSFCell cellRegionsMiskWeekWeight = row2.createCell(12);
//        XSSFCell cellRegionsMogilevDateWeight = row2.createCell(13);
//        XSSFCell cellRegionsMogilexWeekWeight = row2.createCell(14);
        int rowLast;
        for (int cellcountry = 0; cellcountry < countryRow.getCountryRows().size(); cellcountry++)
            {
                rowLast = sheet.getLastRowNum();
                XSSFRow row = sheet.createRow(rowLast+1);
//                XSSFRow row2 = sheet.getRow(rowLast+1);
                XSSFCell cellCountry = row.createCell(0);
                XSSFCell cellMassProductDateWeight = row.createCell(1);
                XSSFCell cellMassProductWeekWeight = row.createCell(2);
                XSSFCell cellRegionsBrestDateWeight = row.createCell(3);
                XSSFCell cellRegionsBrestWeekWeight = row.createCell(4);
                XSSFCell cellRegionsVitebskDateWeight = row.createCell(5);
                XSSFCell cellRegionsVitebskWeekWeight = row.createCell(6);
                XSSFCell cellRegionsGomelDateWeight = row.createCell(7);
                XSSFCell cellRegionsGomelWeekWeight = row.createCell(8);
                XSSFCell cellRegionsGrodnoDateWeight = row.createCell(9);
                XSSFCell cellRegionsGrodnoWeekWeight = row.createCell(10);
                XSSFCell cellRegionsMinskDateWeight = row.createCell(11);
                XSSFCell cellRegionsMinskWeekWeight = row.createCell(12);
                XSSFCell cellRegionsMogilevDateWeight = row.createCell(13);
                XSSFCell cellRegionsMogilexWeekWeight = row.createCell(14);

        cellCountry.setCellValue(countryRow.getCountryRows().get(cellcountry).getCountry());
        cellMassProductDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getMassProduct().getDateWeight());
        cellMassProductWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getMassProduct().getWeekWeight());

        switch (countryRow.getCountryRows().get(cellcountry).getRegions().get(0).getRegion()) {
            case 1:
            cellRegionsBrestDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(0).getDateWeight());
            cellRegionsBrestWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(0).getWeekWeight());
            case 2:
            cellRegionsVitebskDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(1).getDateWeight());
            cellRegionsVitebskWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(1).getWeekWeight());
            case 3:
            cellRegionsGomelDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(2).getDateWeight());
            cellRegionsGomelWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(2).getWeekWeight());
            case 4:
            cellRegionsGrodnoDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(3).getDateWeight());
            cellRegionsGrodnoWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(3).getWeekWeight());
            case 5:
            cellRegionsMinskDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(4).getDateWeight());
            cellRegionsMinskWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(4).getWeekWeight());
            case 6:
            cellRegionsMogilevDateWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(5).getDateWeight());
            cellRegionsMogilexWeekWeight.setCellValue(countryRow.getCountryRows().get(cellcountry).getRegions().get(5).getWeekWeight());
            break;
        }
        }

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/updated.xlsx")) {

            xssfWorkbook.write(fileOut);
        }


            //  }

      //  }

        //workbook.getWorksheets().removeAt("Ввоз");





        return "";

    }
}