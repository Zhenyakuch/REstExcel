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
        this.massProduct = new ElementMass();
        this.regions = new ArrayList<>();

        List<CountryRow> countryRows = report.getCountryRows();
        for (CountryRow row : countryRows) {
            plus(this.massProduct, row.getMassProduct());
            plusList(this.regions, row.getRegions());
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
            if (!find) {
                ElementRegion newReg = new ElementRegion();
                newReg.setRegion(reg.getRegion());
                newReg.setDateWeight(reg.getDateWeight());
                newReg.setWeekWeight(reg.getWeekWeight());
                first.add(newReg);
            }
        }
    }

    private ElementMass plus(ElementMass first, ElementMass second) {
        first.setDateWeight(first.getDateWeight().add(second.getDateWeight()));
        first.setWeekWeight(first.getWeekWeight().add(second.getWeekWeight()));
        return first;
    }
}
