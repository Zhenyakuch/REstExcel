package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;
import java.util.List;

@Getter
@Setter
@ToString
public class CountryReport {
    private Date starDate;
    private Date endDate;
    private String product;
    private boolean isImport;
    private List<CountryRow> countryRows;
}
