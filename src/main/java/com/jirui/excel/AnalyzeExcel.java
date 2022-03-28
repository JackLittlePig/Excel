package com.jirui.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class AnalyzeExcel {

    private static final String PATH_MAC = "/Users/leizhao/leizhao/excel_data";
    private static final String PATH_WINDOWS = "E:\\未考勤人员";

    private static final String FILE_PATH = PATH_WINDOWS;

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    private static final String DEFAULT_ALL_EMP_FILE_PATH = FILE_PATH + "/总表.xls";
    private static final String DEFAULT_ALL_SIGN_IN_FILE_PATH = FILE_PATH + "/刷卡表.xlsx";
    private static final String DEFAULT_RESULT_EXPORT_NOT_SIGN_IN_PATH = FILE_PATH + "/未刷卡表.xls";
    private static final String DEFAULT_RESULT_EXPORT_SIGN_IN_PATH = FILE_PATH + "/已刷卡表.xls";
    private static final String DEFAULT_RESULT_EXPORT_ALL_PATH = FILE_PATH + "/签到未签到总表.xls";


    private static final String FILTER_EMP_TYPE = "正式员工";

    private static final int CELL_COUNT = 3;

    private static final int CELL_ALL_SIGN_IN_EMP_NAME = 5;
    private static final int CELL_ALL_SIGN_IN_DEPART_NO = 6;


    private static final int CELL_ALL_EMP_NAME = 4;
    private static final int CELL_ALL_EMP_DEPART_NO = 6;
    private static final int CELL_ALL_EMP_TYPE = 17;

    private final List<Employee> noRegEmployeeList = new ArrayList<Employee>();
    private final List<Employee> employeeList = new ArrayList<Employee>();
    private final List<Employee> signInList = new ArrayList<Employee>();

    private final Map<String, List<Employee>> signInMap = new HashMap<String, List<Employee>>();
    private final Map<String, List<Employee>> noRegEmployeeMap = new HashMap<String, List<Employee>>();
    private final Map<String, List<Employee>> employeeMap = new HashMap<String, List<Employee>>();

    private final List<Employee> exportExcelSignInList = new ArrayList<Employee>();
    private final List<Employee> exportExcelNotSignInList = new ArrayList<Employee>();
    private final List<AllResult> exportExcelAllResultList = new ArrayList<AllResult>();

    public void start() throws IOException {
        System.out.println("程序开始执行.....");
        readEmployee();
        readSignIn();
        analyze();
        writeResultToFile(exportExcelNotSignInList, DEFAULT_RESULT_EXPORT_NOT_SIGN_IN_PATH);
        writeResultToFile(exportExcelSignInList, DEFAULT_RESULT_EXPORT_SIGN_IN_PATH);
        writeAllResultToFile(exportExcelAllResultList, DEFAULT_RESULT_EXPORT_ALL_PATH);
        System.out.println("程序执行结束.....");
    }

    private void writeAllResultToFile(List<AllResult> list, String filePath) throws IOException {
        if (list == null || list.size() == 0) {
            return;
        }
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int i = 0; i < 4; i++) {
            sheet.setColumnWidth(i, 7000);
        }
        sheet.setDefaultRowHeight((short) 400);
        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(list.get(i).depart);
            row.createCell(1).setCellValue(list.get(i).allNo);
            row.createCell(2).setCellValue(list.get(i).signInNo);
            row.createCell(3).setCellValue(list.get(i).notSignInNo);
        }
        FileOutputStream fos = null;
        try {
            File notSignInFile = new File(filePath);
            if (!notSignInFile.exists()) {
                notSignInFile.createNewFile();
            }
            fos = new FileOutputStream(notSignInFile);
            workbook.write(fos);
            fos.flush();
        } finally {
            if (fos != null) {
                fos.close();
            }
        }
    }

    private void writeResultToFile(List<Employee> list, String filePath) throws IOException {
        if (list == null || list.size() == 0) {
            return;
        }
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int i = 0; i < CELL_COUNT; i++) {
            sheet.setColumnWidth(i, 7000);
        }
        sheet.setDefaultRowHeight((short) 400);
        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(list.get(i).depart);
            row.createCell(1).setCellValue(list.get(i).name);
            row.createCell(2).setCellValue(list.get(i).type);
        }
        FileOutputStream fos = null;
        try {
            File notSignInFile = new File(filePath);
            if (!notSignInFile.exists()) {
                notSignInFile.createNewFile();
            }
            fos = new FileOutputStream(notSignInFile);
            workbook.write(fos);
            fos.flush();
        } finally {
            if (fos != null) {
                fos.close();
            }
        }
    }

    private void analyze() {
        exportExcelAllResultList.add(new AllResult("部门名", "部门所有人数", "所有已考勤人数", "非正式未考勤人数"));
        exportExcelNotSignInList.add(new Employee("非正式未考勤人员名" , "部门名" , "非正式未考勤人数"));
        exportExcelSignInList.add(new Employee("所有已考勤人员名" , "部门名" , "所有已考勤人数"));
        for (Map.Entry<String, List<Employee>> entry : signInMap.entrySet()) {
            if (noRegEmployeeMap.get(entry.getKey()) == null) {
                continue;
            }
            Set<String> employeeSet = new HashSet<String>();
            for (Employee employee : noRegEmployeeMap.get(entry.getKey())) {
                employeeSet.add(employee.name);
            }

            Set<String> signInSet = new HashSet<String>();
            for (Employee signIn : entry.getValue()) {
                signInSet.add(signIn.name);

                exportExcelSignInList.add(new Employee(signIn.name, signIn.depart, String.valueOf(entry.getValue().size())));
            }

            int signInNo = signInSet.size();

            employeeSet.removeAll(signInSet);

            int notSignNo = employeeSet.size();

            if (employeeSet.size() > 0) {
                for (String empName : employeeSet) {
                    exportExcelNotSignInList.add(new Employee(empName, entry.getKey(), String.valueOf(employeeSet.size())));
                }
            }
            exportExcelAllResultList.add(new AllResult(entry.getKey()
                    , String.valueOf(employeeMap.get(entry.getKey()).size())
                    , String.valueOf(signInNo)
                    , String.valueOf(notSignNo)
            ));
        }
    }

    private void readSignIn() throws IOException {
        File allUserExcelFile = new File(DEFAULT_ALL_SIGN_IN_FILE_PATH);
        Workbook workBook = getWorkbook(allUserExcelFile);
        Sheet sheet = workBook.getSheetAt(0);
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = firstRowNum + 1; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            Cell empNameCell = row.getCell(CELL_ALL_SIGN_IN_EMP_NAME);
            Cell departNoCell = row.getCell(CELL_ALL_SIGN_IN_DEPART_NO);
            if (empNameCell != null && departNoCell != null) {
                String empName = empNameCell.toString();
                String departNo = departNoCell.toString();
                signInList.add(new Employee(empName, departNo, null));
            }
        }

        for (Employee employee : signInList) {
            if (employee == null) {
                continue;
            }
            if (signInMap.get(employee.depart) == null) {
                signInMap.put(employee.depart, new ArrayList<Employee>());
            }
            List<Employee> values = signInMap.get(employee.depart);
            values.add(employee);
        }
    }

    private void readEmployee() throws IOException {
        File allUserExcelFile = new File(DEFAULT_ALL_EMP_FILE_PATH);
        Workbook workBook = getWorkbook(allUserExcelFile);
        Sheet sheet = workBook.getSheetAt(0);
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = firstRowNum + 1; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);
            Cell empNameCell = row.getCell(CELL_ALL_EMP_NAME);
            Cell departNoCell = row.getCell(CELL_ALL_EMP_DEPART_NO);
            Cell empTypeCell = row.getCell(CELL_ALL_EMP_TYPE);
            if (empNameCell != null && departNoCell != null && empTypeCell != null) {
                String empName = empNameCell.toString();
                String departNo = departNoCell.toString();
                String empType = empTypeCell.toString();
                if (!FILTER_EMP_TYPE.equals(empType)) {
                    noRegEmployeeList.add(new Employee(empName, departNo, empType));
                }
                employeeList.add(new Employee(empName, departNo, empType));
            }
        }

        for (Employee employee : noRegEmployeeList) {
            if (employee == null) {
                continue;
            }
            if (noRegEmployeeMap.get(employee.depart) == null) {
                noRegEmployeeMap.put(employee.depart, new ArrayList<Employee>());
            }
            List<Employee> values = noRegEmployeeMap.get(employee.depart);
            values.add(employee);
        }

        for (Employee employee : employeeList) {
            if (employee == null) {
                continue;
            }
            if (signInMap.get(employee.depart) == null) {
                signInMap.put(employee.depart, new ArrayList<Employee>());
            }

            if (noRegEmployeeMap.get(employee.depart) == null) {
                noRegEmployeeMap.put(employee.depart, new ArrayList<Employee>());
            }

            if (employeeMap.get(employee.depart) == null) {
                employeeMap.put(employee.depart, new ArrayList<Employee>());
            }
            List<Employee> values = employeeMap.get(employee.depart);
            values.add(employee);
        }
    }

    /**
     * 判断Excel的版本,获取Workbook
     */
    public Workbook getWorkbook(File file) throws IOException {
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(EXCEL_XLS)) {
            //Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {
            // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }
}
