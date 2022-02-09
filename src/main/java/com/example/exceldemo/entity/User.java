package com.example.exceldemo.entity;

import com.example.exceldemo.util.ExcelImport;
import lombok.Data;

@Data
public class User {

    @ExcelImport("姓名")
    private String name;
    @ExcelImport("年龄")
    private String age;
    @ExcelImport("性别")
    private String sex;
    @ExcelImport("手机")
    private String tel;
    @ExcelImport("城市")
    private String city;
    @ExcelImport("头像")
    private String avatar;

    /**
     * 错误提示
     */
    private String rowTips;
    /**
     * 原始数据
     */
    private String rowData;
    /**
     * 行号
     */
    private Integer rowNum;

}
