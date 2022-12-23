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

        Workbook workbook = new Workbook("src/main/resources/Ввоз.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.replace("product", countryRow.getProduct());
        worksheet.replace("startDate", countryRow.getStarDate().toString());
        worksheet.replace("endDate", countryRow.getEndDate().toString());

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/Ввоз.xlsx");

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();
        XSSFRow row = sheet.createRow(rowLast+1);
        XSSFRow row2 = sheet.getRow(rowLast+1);// возможно это лишнее, но я просто пробовала думала вдруг поможет.
        // посмотри плиз своим взглядом что тут в логике не так, что он не заполняет строку. но ошибок не выдает
        XSSFCell cellCountry = row2.createCell(0);
        XSSFCell cellMassProductDateWeight = row2.createCell(1);
        XSSFCell cellMassProductWeekWeight = row2.createCell(2);
        XSSFCell cellRegionsDateWeight = row2.createCell(3);
        XSSFCell cellRegionsWeekWeight = row2.createCell(4);


       // for (int i = rowLast; i <= 10; i++) {
            // for (int cellLast = 0; cellLast<=14; cellLast++) {
           // cellCountry = row.createCell(0);
            cellCountry.setCellValue(countryRow.getCountryRows().get(0).getCountry());
            cellMassProductDateWeight.setCellValue(countryRow.getCountryRows().get(0).getMassProduct().getDateWeight());
            cellMassProductWeekWeight.setCellValue(countryRow.getCountryRows().get(0).getMassProduct().getWeekWeight());
            cellRegionsDateWeight.setCellValue(countryRow.getCountryRows().get(0).getRegions().get(0).getDateWeight());
            cellRegionsWeekWeight.setCellValue(countryRow.getCountryRows().get(0).getRegions().get(0).getWeekWeight());

            //  }

      //  }


        //workbook.getWorksheets().removeAt("Ввоз");


        workbook.save("src/main/resources/updated.xlsx");


        return "";

    }
}