package com.clearlove;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

/**
 * @author promise
 * @date 2022/5/18 - 22:28
 */
public class ExcelReadTest {

  String PATH = "E:\\workspace\\workspace4POI&EasyExcel\\poi\\";

  @Test
  public void testRead03() throws Exception {

    // 获取文件流
    FileInputStream fileInputStream = new FileInputStream(PATH + "Clearlove统计表03.xls");

    // 1.创建一个工作簿
    Workbook workbook = new HSSFWorkbook(fileInputStream);

    // 2.得到表
    Sheet sheet = workbook.getSheetAt(0);

    // 3.得到行
    Row row = sheet.getRow(0);

    // 4.得到列
    Cell cell1 = row.getCell(0);
    Cell cell2 = row.getCell(1);

    // 读取值的时候，一定要注意类型
    // getStringCellValue 字符串类型
    System.out.println(cell1.getStringCellValue());
    System.out.println(cell2.getNumericCellValue());
    fileInputStream.close();

  }

  @Test
  public void testRead07() throws Exception {

    // 获取文件流
    FileInputStream fileInputStream = new FileInputStream(PATH + "Clearlove统计表07.xlsx");

    // 1.创建一个工作簿
    Workbook workbook = new XSSFWorkbook(fileInputStream);

    // 2.得到表
    Sheet sheet = workbook.getSheetAt(0);

    // 3.得到行
    Row row = sheet.getRow(0);

    // 4.得到列
    Cell cell1 = row.getCell(0);
    Cell cell2 = row.getCell(1);

    // 读取值的时候，一定要注意类型
    // getStringCellValue 字符串类型
    System.out.println(cell1.getStringCellValue());
    System.out.println(cell2.getNumericCellValue());
    fileInputStream.close();

  }

}
