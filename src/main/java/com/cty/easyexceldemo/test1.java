package com.cty.easyexceldemo;

import com.alibaba.excel.EasyExcel;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @author ：Mr.chen
 * @date ：Created in 2022/4/22 22:27
 * @description：
 * @modified By：
 * @version: $
 */
@RestController
public class test1 {

    @GetMapping("cty")
    public void test(HttpServletResponse response) throws IOException {
        // 这里注意 使用swagger 可能会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String sheetName = "模板";
        String name = UUID.randomUUID().toString()+"测试";
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode(name, "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(),Person.class)
                .autoCloseStream(Boolean.TRUE)
                .registerWriteHandler(new FreezeAndFilter())
                .registerWriteHandler(new ExcelFillCellMergeStrategy(2))
                .sheet(sheetName).doWrite(getData());
    }

    public static List<Person> getData() {
        List<Person> list = new ArrayList<>();
        Person person1 = new Person("1","cty1", "111", "aaa", "111111");
        Person person2 = new Person("2","cty2", "222", "bbb", "222222");
        Person person3 = new Person("2",null, "333", null, "333333");
        Person person4 = new Person("2","cty4", "444", "ddd", "444444");
        Person person5 = new Person("3","cty5", "555", "eee", "555555");
        list.add(person1);
        list.add(person2);
        list.add(person3);
        list.add(person4);
        list.add(person5);
        return list;
    }
}
