package com.cty.easyexceldemo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

/**
 * @author ：Mr.chen
 * @date ：Created in 2022/4/22 22:27
 * @description：
 * @modified By：
 * @version: $
 */
@Data
public class Person {
    @ColumnWidth(10)//单元格长度
    @ExcelProperty(value = "序号", index = 0)
    private String id;
    @ColumnWidth(20)//单元格长度
    @ExcelProperty(value = "姓名", index = 1)
    private String name;
    @ColumnWidth(20)//单元格长度
    @ExcelProperty(value = "手机号", index = 2)
    private String phone;
    @ColumnWidth(20)//单元格长度
    @ExcelProperty(value = "性别", index = 3)
    private String sex;
    @ColumnWidth(20)//单元格长度
    @ExcelProperty(value = "电话号", index = 4)
    private String mobile;

    public Person(String id, String name, String phone, String sex, String mobile) {
        this.id = id;
        this.name = name;
        this.phone = phone;
        this.sex = sex;
        this.mobile = mobile;
    }
}
