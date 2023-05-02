package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.util.Date;

@Getter
@Setter
@ToString
public class Refund {

    private String number;
    private Date date1;
    private Date date2;
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
