package com.example.demo.model.dto;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.math.BigDecimal;

@Setter
@Getter
@ToString
public class ElementMass {
    private BigDecimal dateWeight = new BigDecimal(0);
    private BigDecimal weekWeight = new BigDecimal(0);
    private BigDecimal dateFromTo= new BigDecimal(0);

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
