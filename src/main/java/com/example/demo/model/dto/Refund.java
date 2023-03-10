package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;

@Getter
@Setter
@ToString
public class Refund {

    private int number;
    private LocalDate date1;
    private LocalDate date2;
    private String  place;
    private String name;
    private double quantity;
    private String recipient;
    private String place_sender;
    private String number_TS;
    private String  numberFSS;
    private String return_reasons;
    private String  organizationFSS;
    private String fio1;
    private String fio2;
    private String  fio3;
}
