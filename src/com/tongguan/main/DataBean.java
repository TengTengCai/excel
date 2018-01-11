package com.tongguan.main;

public class DataBean {
    private String customer;
    private String titleName;
    private String year;
    private String value;

    public DataBean(String customer, String titleName, String year, String value) {
        this.customer = customer;
        this.titleName = titleName;
        this.year = year;
        this.value = value;
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

    public String getValue() {
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

    public void setValue(String value) {
        this.value = value;
    }
}
