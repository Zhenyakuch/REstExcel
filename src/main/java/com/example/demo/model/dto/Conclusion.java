package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;

@Getter
@Setter
@ToString
public class Conclusion {

    private int number1;
    private int number2;
    private int number3;
    private String name_legal;
    private String issued;
    private String name;
    private double weight;
    private String origin;
    private String place;
    private String from_whos;
    private String recipient;
    private String result;
    private String events;
    private LocalDate date1;
    private LocalDate date2;
    private LocalDate date3;
    private LocalDate date4;
    private String fio;
}
