package com.example.exceldemo.entity;

import com.example.exceldemo.util.ExcelExport;
import com.example.exceldemo.util.ExcelImport;
import lombok.Data;

@Data
public class UserExport {

    @ExcelExport(value = "姓名",example = "张三",sort = 10)
    private String name;
    @ExcelExport(value = "年龄",example = "12",sort = 5)
    private String age;
    @ExcelExport(value = "性别",example = "男",sort = 3,kv = "1-男;2-女")
    private String sex;
    @ExcelExport(value = "手机",example = "17600774563",sort = 4)
    private String tel;
    @ExcelExport(value = "城市",example = "杭州",sort = 10)
    private String city;
    @ExcelExport(value = "头像",example = "123123",sort = 10)
    private String avatar;

}
