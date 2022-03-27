package com.jirui.excel;

public class AllResult {

    public String depart;
    public String allNo;
    public String signInNo;
    public String notSignInNo;

    public AllResult(String depart, String allNo, String signInNo, String notSignInNo) {
        this.depart = depart;
        this.allNo = allNo;
        this.signInNo = signInNo;
        this.notSignInNo = notSignInNo;
    }

    @Override
    public String toString() {
        return "AllResult{" +
                "depart='" + depart + '\'' +
                ", allNo='" + allNo + '\'' +
                ", signInNo='" + signInNo + '\'' +
                ", notSignInNo='" + notSignInNo + '\'' +
                '}';
    }
}
