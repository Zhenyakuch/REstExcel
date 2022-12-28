package com.example.demo.model.dto;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.List;

@Getter
@Setter
@ToString
public class CountryRow {
    private String resCountryOrProduct;
    private ElementMass massProduct;
    private List<ElementRegion> regions;

}