package com.clearlove.easy;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import org.apache.commons.collections4.ListUtils;
import org.junit.Test;

/**
 * @author promise
 * @date 2022/5/22 - 23:03
 */
public class EasyTest {

  String PATH = "E:\\workspace\\workspace4POI&EasyExcel\\poi\\";

  private List<DemoData> data() {
    List<DemoData> list = new ArrayList<>();
    for (int i = 0; i < 10; i++) {
      DemoData data = new DemoData();
      data.setString("字符串" + i).setDate(new Date()).setDoubleData(0.56);
      list.add(data);
    }
    return list;
  }

  // 根据List 写入Excel

  @Test
  public void simpleWrite() {
    // 写法1
    String fileName = PATH + "EasyTest.xlsx";
    // 这里需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
    // 如果这里想使用03 则 传入excelType参数即可
    // write(filename, 实体类)
    // sheet(表名)
    // doWrite(数据)
    EasyExcel.write(fileName, DemoData.class)
        .sheet("模板")
        .doWrite(data());

  }

}
