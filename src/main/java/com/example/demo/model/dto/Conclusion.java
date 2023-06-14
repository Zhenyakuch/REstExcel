package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.util.List;

@Getter
@Setter
@ToString
public class Conclusion {

    private List<ConclusionProduct> conclusionProducts;
    private String number1;
    private String number3;
    private String name_legal;
    private String issued;
    private String place;
    private String recipient;
    private String events;
    private String date1;
    private String date3;
    private String date4;
    private String fio;
}
