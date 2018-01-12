package com.tongguan.main;

import java.math.BigDecimal;

public class DataBean {
    private String customer;
    private String titleName;
    private String year;
    private Double value;

    public DataBean(String customer, String titleName, String year, Double value) {
        this.customer = customer;
        this.titleName = titleName;
        this.year = year;
        this.value = value;
        BigDecimal b = new BigDecimal(this.value);
        this.value = b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
    }

    public DataBean() {
    }

    public String getCustomer() {
        return customer;
    }

    public String getTitleName() {
        return titleName;
    }

    public String getYear() {
        return year;
    }

    public Double getValue() {
        return value;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public void setValue(Double value) {
        this.value = value;
    }
}
