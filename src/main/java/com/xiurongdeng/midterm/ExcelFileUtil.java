package com.xiurongdeng.midterm;

import java.io.*;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {
  private final String filePath;
  private final String fileName;

  public ExcelFileUtil(String filePath, String fileName) {
    this.filePath = filePath;
    this.fileName = fileName;
  }
  private Workbook getExcelWorkbook() {
    Workbook book = null;
    // Create an object of File class to open xlsx file
    File file = new File(filePath + fileName);
    // Create an object of FileInputStream class to read excel file
    try (FileInputStream inputStream = new FileInputStream(file)) {
      // Find the file extension by splitting file name in substring  and getting only extension
      // name
      String fileExtensionName = fileName.substring(fileName.indexOf("."));
      // Check condition if the file is xlsx file
      if (fileExtensionName.equals(".xlsx")) {
        book = new XSSFWorkbook(inputStream);
      } else if (fileExtensionName.equals(".xls")) {
        book = new HSSFWorkbook(inputStream);
      }
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (SecurityException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
    return book;
  }

  public List<Map<String, String>> readExcel(String sheetName) {
    Workbook book = getExcelWorkbook();
    // Read sheet inside the workbook by its name
    Sheet sheet = book.getSheet(sheetName);

    // Find number of rows in excel sheet
    int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
    List<Map<String, String>> result = new ArrayList();
    if (rowCount > 0) {
      Row rowHead = sheet.getRow(0);
      // Create a loop over all the rows of excel sheet to read it
      for (int i = 1; i < rowCount + 1; i++) {
        Row row = sheet.getRow(i);
        Map<String, String> data = new HashMap<>();
        // Create a loop to print cell values in a row
        for (int j = 0; j < row.getLastCellNum(); j++) {
          data.put(rowHead.getCell(j).getStringCellValue(), row.getCell(j).getStringCellValue());
        }
        result.add(data);
      }
    }
    return result;
  }

  public void writeExcel(String sheetName, List<List<String>> dataToWrite) {
    Workbook book = getExcelWorkbook();
    // Read excel sheet by sheet name
    LocalTime now = LocalTime.now();
    Sheet sheet = book.getSheet(sheetName) == null ? book.createSheet(sheetName) : book.createSheet(sheetName + String.valueOf(now.getSecond()));
    for(int i=0; i < dataToWrite.size(); i++) {
      Row newRow = sheet.createRow(i);
      // Create a loop over the cell of newly created Row
      for (int j = 0; j < dataToWrite.get(i).size(); j++) {
        // Fill data in row
        Cell cell = newRow.createCell(j);
        cell.setCellValue(dataToWrite.get(i).get(j));
        cell.getSheet();
      }
    }
    // Create an object of FileOutputStream class to create write data in excel file
    try(FileOutputStream outputStream = new FileOutputStream(filePath + fileName)) {
      book.write(outputStream);
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
