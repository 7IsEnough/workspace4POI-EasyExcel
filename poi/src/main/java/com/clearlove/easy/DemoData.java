package com.clearlove.easy;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import java.util.Date;
import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @author promise
 * @date 2022/5/22 - 22:49
 */
@Data
@Accessors(chain = true)
public class DemoData {

  @ExcelProperty("字符串标题")
  private String string;

  @ExcelProperty("日期标题")
  private Date date;

  @ExcelProperty("数字标题")
  private Double doubleData;

  // 忽略这个字段
  @ExcelIgnore
  private String ignore;

}
