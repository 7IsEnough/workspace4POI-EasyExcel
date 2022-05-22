package com.clearlove;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
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

  @Test
  public void testCellType() throws Exception {
    // 获取文件流
    FileInputStream fileInputStream = new FileInputStream(PATH + "明细表.xls");

    // 1.创建一个工作簿 使用excel能操作的这边都可以操作
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    Sheet sheet = workbook.getSheetAt(0);

    // 2.获取标题内容
    Row rowTitle = sheet.getRow(0);
    if (rowTitle != null) {
      // 获取列数
      int cellCount = rowTitle.getPhysicalNumberOfCells();
      for (int cellNum = 0; cellNum < cellCount; cellNum++) {
        Cell cell = rowTitle.getCell(cellNum);
        if (cell != null) {
          int cellType = cell.getCellType();
          String cellValue = cell.getStringCellValue();
          System.out.print(cellValue + " | ");
        }
      }
      System.out.println();
    }

    // 获取表中的内容
    int rowCount = sheet.getPhysicalNumberOfRows();
    for (int rowNum = 1; rowNum < rowCount; rowNum++) {
      Row rowData = sheet.getRow(rowNum);
      // 读取列
      int cellCount = rowTitle.getPhysicalNumberOfCells();
      for (int cellNum = 0; cellNum < cellCount; cellNum++) {
        System.out.print("[" + (rowNum + 1) + "-" + (cellNum + 1) + "]");
        Cell cell = rowData.getCell(cellNum);

        // 匹配列的数据类型
        if (cell != null) {
          int cellType = cell.getCellType();
          String cellValue = "";
          switch (cellType) {
            // 字符串
            case XSSFCell
                .CELL_TYPE_STRING:
              System.out.print("[String]");
            cellValue = cell.getStringCellValue();
            break;
            // 布尔
            case XSSFCell
                .CELL_TYPE_BOOLEAN:
              System.out.print("[Boolean]");
              cellValue = String.valueOf(cell.getBooleanCellValue());
              break;
            // 空
            case XSSFCell
                .CELL_TYPE_BLANK:
              System.out.print("[BLANK]");
              break;
            // 数字(日期、普通数字)
            case XSSFCell
                .CELL_TYPE_NUMERIC:
              System.out.print("[Numeric]");
              // 日期
              if (DateUtil.isCellDateFormatted(cell)) {
                System.out.print("[日期]");
                cellValue = new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd");
              } else {
                // 不是日期格式，防止数字过长
                System.out.print("[普通数字，转换为字符串输出]");
                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                cellValue = cell.toString();
              }
              break;
            // 布尔
            case XSSFCell
                .CELL_TYPE_ERROR:
              System.out.print("[数据类型错误]");
              break;
          }
          System.out.println(cellValue);
        }


      }
    }
    fileInputStream.close();
  }

  @Test
  public void testFormula() throws Exception {
    FileInputStream inputStream = new FileInputStream(PATH + "公式.xls");
    Workbook workbook = new XSSFWorkbook(inputStream);

    Sheet sheet = workbook.getSheetAt(0);

    Row row = sheet.getRow(4);
    Cell cell = row.getCell(0);

    // 拿到计算公式 eval
    FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);

    // 输出单元格的内容
    int cellType = cell.getCellType();
    switch (cellType) {
      // 公式
      case Cell.CELL_TYPE_FORMULA:
        String formula = cell.getCellFormula();
        System.out.println(formula);

        // 计算
        CellValue evaluate = formulaEvaluator.evaluate(cell);
        String cellValue = evaluate.formatAsString();
        System.out.println(cellValue);
    }
  }

}
