package com.example.demo;

public class Product {
    private String product;
    private String country;
    private int Brest;
    private int Vitebsk;
    private int Gomel;
    private int Grodno;
    private int Minsk;
    private int Mogilev;
    private int seven_day;
    private int period;

    public String getProduct() {
        return product;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public int getBrest() {
        return Brest;
    }

    public void setBrest(int brest) {
        Brest = brest;
    }

    public int getVitebsk() {
        return Vitebsk;
    }

    public void setVitebsk(int vitebsk) {
        Vitebsk = vitebsk;
    }

    public int getGomel() {
        return Gomel;
    }

    public void setGomel(int gomel) {
        Gomel = gomel;
    }

    public int getGrodno() {
        return Grodno;
    }

    public void setGrodno(int grodno) {
        Grodno = grodno;
    }

    public int getMinsk() {
        return Minsk;
    }

    public void setMinsk(int minsk) {
        Minsk = minsk;
    }

    public int getMogilev() {
        return Mogilev;
    }

    public void setMogilev(int mogilev) {
        Mogilev = mogilev;
    }

    public int getSeven_day() {
        return seven_day;
    }

    public void setSeven_day(int seven_day) {
        this.seven_day = seven_day;
    }

    public int getPeriod() {
        return period;
    }

    public void setPeriod(int period) {
        this.period = period;
    }

    public int getAll_ton() {
        return all_ton;
    }

    public void setAll_ton(int all_ton) {
        this.all_ton = all_ton;
    }

    private int all_ton;


@Override
    public String toString (){
    return "Product [" +
            "product=" + product + ", " +
            "country=" + country + ", " +
            "Brest=" + Brest+ ", " +
            "Vitebsk=" + Vitebsk+ ", " +
            "Gomel=" + Gomel+ ", " +
            "Grodno=" + Grodno+ ", " +
            "Minsk=" + Minsk+ ", " +
            "Mogilev=" + Mogilev+ ", " +
            "seven_day=" + seven_day + ", " +
            "period=" + period+
            "]";


}




}