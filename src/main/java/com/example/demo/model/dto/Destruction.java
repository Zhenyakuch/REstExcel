package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;

@Getter
@Setter
@ToString
public class Destruction {

    private String number;
    private String date1;
    private String date2;
    private String method_destruction;
    private String name;
    private double quantity;
    private String units;

    private String place;
    private String position1;
    private String position2;
    private String position3;
    private String fio1;
    private String fio2;
    private String fio3;
}
