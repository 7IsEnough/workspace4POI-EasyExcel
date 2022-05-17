package com.clearlove;

import java.io.FileNotFoundException;
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
 * @date 2022/5/17 - 21:59
 */
public class ExcelWriteTest {

  String PATH = "E:\\workspace\\workspace4POI&EasyExcel\\poi";

  @Test
  public void testWrite03() throws Exception {
    // 1.创建一个工作簿
    Workbook workbook = new HSSFWorkbook();
    // 2.创建一个工作表
    Sheet sheet = workbook.createSheet("统计表");
    // 3.创建一个行
    Row row1 = sheet.createRow(0);
    // 4.创建一个单元格  (1,1)
    Cell cell11 = row1.createCell(0);
    cell11.setCellValue("今日新增人数");
    // (1,2)
    Cell cell12 = row1.createCell(1);
    cell12.setCellValue("666");


    // 第二行
    Row row2 = sheet.createRow(1);
    // 4.创建一个单元格  (2,1)
    Cell cell21 = row2.createCell(0);
    cell21.setCellValue("统计时间");
    // (2,2)
    Cell cell22 = row2.createCell(1);
    String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
    cell22.setCellValue(time);

    // 生成一张表 (IO流)  03版本使用xls结尾
    FileOutputStream fileOutputStream = new FileOutputStream(PATH + "\\Clearlove统计表03.xls");
    // 输出
    workbook.write(fileOutputStream);
    // 关闭流
    fileOutputStream.close();
    System.out.println("Excel生成完毕");
  }

  @Test
  public void testWrite07() throws Exception {
    // 1.创建一个工作簿  07
    Workbook workbook = new XSSFWorkbook();
    // 2.创建一个工作表
    Sheet sheet = workbook.createSheet("统计表");
    // 3.创建一个行
    Row row1 = sheet.createRow(0);
    // 4.创建一个单元格  (1,1)
    Cell cell11 = row1.createCell(0);
    cell11.setCellValue("今日新增人数");
    // (1,2)
    Cell cell12 = row1.createCell(1);
    cell12.setCellValue("666");


    // 第二行
    Row row2 = sheet.createRow(1);
    // 4.创建一个单元格  (2,1)
    Cell cell21 = row2.createCell(0);
    cell21.setCellValue("统计时间");
    // (2,2)
    Cell cell22 = row2.createCell(1);
    String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
    cell22.setCellValue(time);

    // 生成一张表 (IO流)  03版本使用xlsx结尾
    FileOutputStream fileOutputStream = new FileOutputStream(PATH + "\\Clearlove统计表07.xlsx");
    // 输出
    workbook.write(fileOutputStream);
    // 关闭流
    fileOutputStream.close();
    System.out.println("Excel生成完毕");
  }

}
