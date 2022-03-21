package com.jirui.excel;

import java.io.IOException;

public class Main {

    public static void main(String[] args){
        try {
            new AnalyzeExcel().start();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
