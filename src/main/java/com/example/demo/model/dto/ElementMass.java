package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.math.BigDecimal;

@Setter
@Getter
@ToString
public class ElementMass {

    //переменные json импорта экспорта и реэк
    private BigDecimal dateWeight = new BigDecimal(0);
    private BigDecimal weekWeight = new BigDecimal(0);
    private BigDecimal dateFromTo = new BigDecimal(0);
    private BigDecimal date2021 = new BigDecimal(0);
    private BigDecimal date2022 = new BigDecimal(0);

    public Double date2021Double() {
        return date2021.doubleValue();
    }

    public double date2022Double() {
        return date2022.doubleValue();
    }

    public Double dateFromToDouble() {
        return dateFromTo.doubleValue();
    }

    public Double dateWeightDouble() {
        return dateWeight.doubleValue();
    }

    public Double weekWeightDouble() {
        return weekWeight.doubleValue();
    }
}
