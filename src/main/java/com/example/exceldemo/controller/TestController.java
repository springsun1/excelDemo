package com.example.exceldemo.controller;

import cn.hutool.json.JSONArray;
import com.example.exceldemo.entity.User;
import com.example.exceldemo.entity.UserExport;
import com.example.exceldemo.util.ExcelUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@RestController
@RequestMapping("/api")
public class TestController {

    /**
     * 导出模板，并且导出示例
     * @param response
     */
    @GetMapping("/export")
    public void export(HttpServletResponse response) {
        ExcelUtils.exportTemplate(response, "用户表", UserExport.class,true);
    }


    @GetMapping("/export1")
    public void export1(HttpServletResponse response) {
        // 表头数据
        List<Object> head = Arrays.asList("姓名","年龄","性别","头像");
        // 用户1数据
        List<Object> user1 = new ArrayList<>();
        user1.add("诸葛亮");
        user1.add(60);
        user1.add("男");
        user1.add("https://profile.csdnimg.cn/A/7/3/3_sunnyzyq");
        // 用户2数据
        List<Object> user2 = new ArrayList<>();
        user2.add("大乔");
        user2.add(28);
        user2.add("女");
        user2.add("https://profile.csdnimg.cn/6/1/9/0_m0_48717371");
        // 将数据汇总
        List<List<Object>> sheetDataList = new ArrayList<>();
        sheetDataList.add(head);
        sheetDataList.add(user1);
        sheetDataList.add(user2);
        // 导出数据
        ExcelUtils.export(response,"用户表", sheetDataList);
    }


    @PostMapping("/import")
    public void importUser(@RequestPart("file") MultipartFile file) throws Exception {
        List<User> users = ExcelUtils.readMultipartFile(file, User.class);
        for (User user : users) {
            System.out.println(user.toString());
        }
    }

    @PostMapping("/import1")
    public JSONArray importUser1(@RequestPart("file")MultipartFile file) throws Exception {
        JSONArray array = ExcelUtils.readMultipartFile(file);
        System.out.println("导入数据为:" + array);
        return array;
    }

}
