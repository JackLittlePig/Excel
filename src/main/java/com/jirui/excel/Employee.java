package com.jirui.excel;

public class Employee {

    public String name;
    public String depart;
    public String type;


    public Employee(String name, String depart, String type) {
        this.name = name;
        this.depart = depart;
        this.type = type;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "depart='" + depart + '\'' +
                ", name='" + name + '\'' +
                ", type='" + type + '\'' +
                '}';
    }
}
