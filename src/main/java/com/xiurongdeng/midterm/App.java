package com.xiurongdeng.midterm;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class App {
  public static void main(String[] args) {
    TestLogin tryLogin = new TestLogin("http://demo.guru99.com/test/login.html");
    String cwd = new File("").getAbsolutePath();
    System.out.println(cwd);
    ExcelFileUtil objExcelFile = new ExcelFileUtil("src/main/resources/", "Lab4.xlsx");
    List<Map<String, String>> testCase = objExcelFile.readExcel("Test Case");
    // create a report sheet with title row
    List<List<String>> report = new ArrayList<>();
    List<String> title = new ArrayList<>();
    title.add("username");
    title.add("password");
    title.add("test result");
    report.add(title);
    for(Map<String, String> caseData : testCase) {
      String testResult = tryLogin.login(caseData.get("username"), caseData.get("password")) ? "True" : "False";
      List<String> row = new ArrayList<>();
      row.add(caseData.get("username"));
      row.add(caseData.get("password"));
      row.add(testResult);
      report.add(row);
    }
    objExcelFile.writeExcel("Test Report", report);
  }
}
