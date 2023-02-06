package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;

@Getter
@Setter
@ToString
public class Sticker {

    private int number;
    private String name;
    private double weight;
    private String origin;
    private String place;
    private double net_weight;
    private String recipient;
    private String appointment;
    private double area;
    private String external_sings;
    private String provisional_definition;
    private String additional_info;
    private String seal_number;
    private String position;
    private LocalDate date;
    private String fio1;
    private String fio2;

}
