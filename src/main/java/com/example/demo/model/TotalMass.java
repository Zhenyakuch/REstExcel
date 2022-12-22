package com.example.demo.model;

import com.example.demo.model.dto.CountryReport;
import com.example.demo.model.dto.CountryRow;
import com.example.demo.model.dto.ElementMass;
import com.example.demo.model.dto.ElementRegion;
import lombok.ToString;

import java.util.ArrayList;
import java.util.List;

@ToString
public class TotalMass {
    private ElementMass massProduct;
    private List<ElementRegion> regions;

    public TotalMass(CountryReport report) {
        massProduct = new ElementMass();
        regions = new ArrayList<>();

        List<CountryRow> countryRows = report.getCountryRows();
        for (CountryRow row : countryRows) {
            plus(massProduct, row.getMassProduct());
            plusList(regions, row.getRegions());
        }
    }

    private void plusList(List<ElementRegion> first, List<ElementRegion> second) {
        for (ElementRegion reg : second) {
            boolean find = false;
            for (ElementRegion total : first) {
                if (reg.getRegion() == total.getRegion()) {
                    plus(total, reg);
                    find = true;
                    break;
                }
            }
            if (!find)
                first.add(reg);
        }
    }

    private ElementMass plus(ElementMass first, ElementMass second) {
        first.setDateWeight(first.getDateWeight() + second.getDateWeight());
        first.setWeekWeight(first.getWeekWeight() + second.getWeekWeight());
        return first;
    }
}
