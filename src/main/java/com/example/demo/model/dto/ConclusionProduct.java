package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString(callSuper = true)
public class ConclusionProduct {

    private String name;
    private String weight;
    private String origin;
    private String result;
    private String number2;
    private String date2;
    private String from_whos;


}
