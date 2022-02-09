* * *

前言
==

最近抽了两天时间，把Java实现表格的相关操作进行了封装，本次封装是基于 POI 的二次开发，最终使用只需要调用一个工具类中的方法，就能满足业务中绝大部门的导入和导出需求。

1\. 功能测试
========

1.1 测试准备
--------

在做测试前，我们需要將【2. 环境准备】中的四个文件拷贝在工程里（如：我这里均放在了com.zyq.util.excel 包下）。

![](https://img-blog.csdnimg.cn/3aa78ee37e3744c4be8ddae394dc2141.png)

1.2 数据导入
--------

### 1.2.1 导入解析为JSON

比如，我们有下面一个表格：

![](https://img-blog.csdnimg.cn/ae7e174ace8f4e59b6f2db4af1e9dceb.png)

Controller 代码：

    @PostMapping("/import")
    public JSONArray importUser(@RequestPart("file")MultipartFile file) throws Exception {
        JSONArray array = ExcelUtils.readMultipartFile(file);
        System.out.println("导入数据为:" + array);
        return array;
    }

测试效果：

![](https://img-blog.csdnimg.cn/0cc2f19d34754c5a87c4eff66b49da73.png)

### 1.2.2 导入解析为对象（基础）

首先，你需要创建一个与导入表格对应的Java实体对象，并打上对应的Excel解析的导入注解，@ExcelImport注解的value则为表头名称。

![](https://img-blog.csdnimg.cn/8828debd5d644242aa91e95f15655724.png)

Controller 代码：

    @PostMapping("/import")
    public void importUser(@RequestPart("file")MultipartFile file) throws Exception {
        List<User> users = ExcelUtils.readMultipartFile(file, User.class);
        for (User user : users) {
            System.out.println(user.toString());
        }
    }

测试效果：

![](https://img-blog.csdnimg.cn/15b12e9cc5be4176972521c740398122.png)

### 1.2.3 导入解析为对象（字段自动映射）

对于有的枚举数据，通常我们导入的时候，表格中的数据是值，而在数据保存时，往往用的是键，比如：我们用sex=1可以表示为男，sex=2表示为女，那么我们通过配置也可以达到导入时，数据的自动映射。

那么，我们只需要将Java实体中的对象sex字段的类型改为对应的数字类型Integer，然后再注解中配置好 kv 属性（属性格式为：键1-值1;键2-值2;键3-值3;.....）

![](https://img-blog.csdnimg.cn/a4ee2b8d043046e5ac1cfe40b191eb82.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：可以看到已经自动映射成功了。

![](https://img-blog.csdnimg.cn/99957f87d97d47d18624ddae7cc52966.png)

### 1.2.4 导入解析为对象（获取行号）

我们在做页面数据导入时，有时候可能需要获取行号，好追踪导入的数据。

那么，我们只需要在对应的实体中加入一个 int 类型的 rowNum 字段即可。

![](https://img-blog.csdnimg.cn/efaed1ec599a4ed7bc40de135638d856.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：

![](https://img-blog.csdnimg.cn/904642ea718e4273b70bf3b914f35c96.png)

### 1.2.5 导入解析为对象（获取原始数据）

在做页面数据导入的时候，如果某行存在错误，一般我们会将原始的数据拿出来分析，为什么会造成数据错误。那么，我们在实体类中，增加一个 String 类型的 rowData 字段即可。

![](https://img-blog.csdnimg.cn/38d8ea403dec40f2b1cfbad8e2c5b031.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：

![](https://img-blog.csdnimg.cn/d497ee6c55b244bb8b1c901861531499.png)

### 1.2.6 导入解析为对象（获取错误提示）

当我们在导入数据的时候，如果某行数据存在，字段类型不正确，长度超过最大限制（详见1.2.7），必填字段验证（1.2.8），数据唯一性验证（1.2.9）等一些错误时候，我们可以往对象中添加一个 String 类型的 rowTips 字段，则可以直接拿到对应的错误信息。

![](https://img-blog.csdnimg.cn/02c3160a862742a4bd78671a8b577244.png)

比如，我们将表格中赵子龙的性别改为F（F并不是映射数据），将大乔的性别改为二十八（不能转换为Integer类型数据）。

![](https://img-blog.csdnimg.cn/23c297fbb55e429f86bc280d2a5b7e95.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：可以看到，我们可以通过 rowTips 直接拿到对应的错误数据提示。

![](https://img-blog.csdnimg.cn/57f779a6acc147c28c206540862ac884.png)

### 1.2.7 导入解析为对象（限制字段长度）

比如，我们手机通常为11为长度，那么不妨限制电话的最大长度位数为11位。

对应的做法，就是在 @ExcelImport 注解中，设置 maxLength = 11 即可。

![](https://img-blog.csdnimg.cn/ab0f871993cb4e79a48ea58ad194d271.png)

比如，我们将诸葛孔明的电话长度设置为超过11位数的一个字符串。

![](https://img-blog.csdnimg.cn/4bc6a29fb7e7497bb85a2e31d4a88484.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：

![](https://img-blog.csdnimg.cn/38e835030e7d4995a1cbda4e1adec136.png)

### 1.2.8 导入解析为对象（必填字段验证）

我们在做数据导入的时候，往往还会有一些必填字段，比如用户的名称，电话。

那么，我们只需要在 @ExcelImport 注解属性中，加上 required = true 即可。

![](https://img-blog.csdnimg.cn/137ea8a724da4a5a9a1cab0b1e06dbf3.png)

我们将诸葛孔明的电话，以及第4行的姓名去掉，进行测试。

![](https://img-blog.csdnimg.cn/efd46057ac794376a3777afdcbf88c40.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：

![](https://img-blog.csdnimg.cn/e061e65ea10841f69efd366e8bf69bc5.png)

### 1.2.9 导入解析为对象（数据唯一性验证）

(1) 单字段唯一性验证

我们在导入数据的时候，某个字段是具有唯一性的，比如我们这里假设规定姓名不能重复，那么则可以在对应字段的 @ExcelImport 注解上加上 unique = true 属性。

![](https://img-blog.csdnimg.cn/1245c763f0f941bd83e0beaba1f99c07.png)

这里我们构建2条姓名一样的数据进行测试。

![](https://img-blog.csdnimg.cn/ecd21037d57545fbb8789284529461c9.png)

Cotroller 代码略（和 1.2.2 完全一致）。

测试效果：

![](https://img-blog.csdnimg.cn/7bf2e21438ad41e49307f460ad348b1a.png)

（2）多字段唯一性验证

如果你导入的数据存在多字段唯一性验证这种情况，只需要将每个对应字段的 @ExcelImport 注解属性中，都加上 required = true 即可。

比如：我们将姓名和电话两个字段进行联合唯一性验证（即不能存在有名称和电话都一样的数据，单个字段属性重复允许）。

![](https://img-blog.csdnimg.cn/ac3508e7dc3940e08be828b9836912f7.png)

首先，我们将刚刚（1）的数据进行导入。

![](https://img-blog.csdnimg.cn/85bfd2935a2e425b9d9d0d3990423e2e.png)

测试效果：可以看到，虽然名称有相同，但电话不相同，所以这里并没有提示唯一性验证错误。

![](https://img-blog.csdnimg.cn/40086da03cd144e793dda6cfdd792e7e.png)

现在，我们将最后一行的电话也改为和第1行一样的，于是，现在就存在了违背唯一性的两条数据。

![](https://img-blog.csdnimg.cn/e1881bd83fb04a75805b623a5beae91d.png)

测试效果：可以看到，我们的联合唯一性验证生效了。

![](https://img-blog.csdnimg.cn/1d9de040614740679087236191275832.png)

1.3 数据导出
--------

### 1.3.1 动态导出（基础）

这种方式十分灵活，表中的数据，完全自定义设置。

Controller 代码：

    @GetMapping("/export")
    public void export(HttpServletResponse response) {
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

代码截图：

![](https://img-blog.csdnimg.cn/bc67613f57394018824f8fa37ad2132e.png)

由于是 get 请求，我们直接在浏览器上输入请求地址即可触发下载。

![](https://img-blog.csdnimg.cn/bb7f9faebb074b77b9248345d085ce8d.png)

打开下载表格，我们可以看到，表中的数据和我们代码组装的顺序一致。

![](https://img-blog.csdnimg.cn/101ed6ed191e4691987768c2247ba037.png)

### 1.3.2 动态导出（导出图片）

如果你的导出中，需要将对应图片链接直接显示为图片的话，那么，这里也是可以的，只需要将对应的类型转为 java.net.URL 类型即可（注意：转的时候有异常处理，为了方便演示，我这里直接抛出）

![](https://img-blog.csdnimg.cn/ffe697f7618f4b79aafc7f76185e448f.png)

测试效果：

![](https://img-blog.csdnimg.cn/5cc3c71913ab4ee1971bb99b680b8451.png)

### 1.3.3 动态导出（实现下拉列表）

我们在做一些数据导出的时候，可能要对某一行的下拉数据进行约束限制。

比如，当我们下载一个导入模版的时候，我们可以将性别，城市对应的列设置为下拉选择。

![](https://img-blog.csdnimg.cn/8cccb5da12c04c0a972c1a3609a4b8ab.png)

测试效果：

![](https://img-blog.csdnimg.cn/9616a829ee8043f8af859de39710bca9.png)

![](https://img-blog.csdnimg.cn/fedda83f4c634fcbb8909591f9604b2a.png)

### 1.3.4 动态导出（横向合并）

比如，我们将表头横向合并，只需要将合并的单元格设置为 ExcelUtils.COLUMN\_MERGE 即可。

![](https://img-blog.csdnimg.cn/142a2a77f25d49e7b5367e7e40165304.png)

测试效果：可以看到表头的地址已经被合并了。

![](https://img-blog.csdnimg.cn/a9288eb85260457db1b3babe88b349cc.png)

### 1.3.5 动态导出（纵向合并）

除了横向合并，我们还可以进行纵向合并，只需要将合并的单元格设置为 ExcelUtils.ROW\_MERGE 即可。

![](https://img-blog.csdnimg.cn/437cb122f3bb479eb00adca4af07c05f.png)

测试效果：

![](https://img-blog.csdnimg.cn/854e627ebc4448aa8c09ef8d5e7dd402.png)

### 1.3.6 导出模板（基础）

我们在做数据导入的时候，往往首先会提供一个模版供其下载，这样用户在导入的时候才知道如何去填写数据。导出模板除了可以用上面的动态导出，这里还提供了一种更加便捷的写法。只需要创建一个类，然后再对应字段上打上 @ExcelExport 注解类即可。

![](https://img-blog.csdnimg.cn/70d5e17544754ed795e58b3f74a719f2.png)

Controller 代码：

    @GetMapping("/export")
    public void export(HttpServletResponse response) {
        ExcelUtils.exportTemplate(response, "用户表", User.class);
    }

代码截图：

![](https://img-blog.csdnimg.cn/10c07be81be4410b86ebd0020325f7ff.png)

测试效果：

![](https://img-blog.csdnimg.cn/073f4a18bc2f45b3801a58fb52123105.png)

### 1.3.7 导出模板（附示例数据）

我们在做模版下载时候，有时往往会携带一条样本数据，好提示用户数据格式是什么，那么我们只需要在对应字段上进行配置即可。

![](https://img-blog.csdnimg.cn/4209f2fde7884157aa7fd24fa1bf53d2.png)

Controller代码：

![](https://img-blog.csdnimg.cn/83048c60e75d42f7b1a0e72a59f29da9.png)

测试效果：

![](https://img-blog.csdnimg.cn/f5c74379bf6f4b7d9c09d176abcc52cd.png)

### 1.3.8 按对象导出（基础）

我们还可以通过 List 对象，对数据直接进行导出。首先，同样需要在对应类的字段上，设置导出名称。

![](https://img-blog.csdnimg.cn/a45a495287f64b6b8d5d300c060974de.png)

Controller 代码：

![](https://img-blog.csdnimg.cn/64d8df2c39e2435493d67fb9c37d4c1a.png)

测试效果：

![](https://img-blog.csdnimg.cn/be2797f4c0a046e098aedf962c055cfe.png)

### 1.3.9 按对象导出（数据映射）

在上面 1.3.8 的导出中，我们可以看到，性别数据导出来是1和2，这个不利于用户体验，应该需要转换为对应的中文，我们可以在字段注解上进行对应的配置。

![](https://img-blog.csdnimg.cn/823809bcb7e741b2a44ebd60ca7baf44.png)

Controller 代码略（和1.3.8完全一致）

测试效果：可以看到1和2显示为了对应的男和女

### ![](https://img-blog.csdnimg.cn/28d4044bd181410b905da26cf241d0b6.png)

### 1.3.10 按对象导出（调整表头顺序）

如果你需要对表头字段进行排序，有两种方式：

第一种：按照表格的顺序，排列Java类中的字段；

第二种：在 @ExcelExport 注解中，指定 sort 属性，其值越少，排名越靠前。

![](https://img-blog.csdnimg.cn/238d7b17833545588b5b49273b6d0595.png)

Controller 代码略（和1.3.8完全一致）

测试效果：可以看到，此时导出数据的表头顺序，和我们指定的顺序完全一致。![](https://img-blog.csdnimg.cn/00cb4a2ef6c348ee99538cfefa8c6827.png)

2\. 环境准备
========

2.1 Maven 依赖
------------

本次工具类的封装主要依赖于阿里巴巴的JSON包，以及表格处理的POI包，所以我们需要导入这两个库的依赖包，另外，我们还需要文件上传的相关包，毕竟我们在浏览器页面，做Excel导入时，是上传的Excel文件。

    <!-- 文件上传 -->
    <dependency>
        <groupId>org.apache.httpcomponents</groupId>
        <artifactId>httpmime</artifactId>
    	<artifactId>4.5.7</artifactId>
    </dependency>
    <!-- JSON -->
    <dependency>
        <groupId>com.alibaba</groupId>
        <artifactId>fastjson</artifactId>
        <version>1.2.41</version>
    </dependency>
    <!-- POI -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>3.16</version>
    </dependency>

2.2 类文件
-------

### ExcelUtils

    package com.zyq.util.excel;
    
    import java.io.ByteArrayOutputStream;
    import java.io.File;
    import java.io.FileInputStream;
    import java.io.FileOutputStream;
    import java.io.IOException;
    import java.io.InputStream;
    import java.lang.reflect.Field;
    import java.math.BigDecimal;
    import java.math.RoundingMode;
    import java.net.URL;
    import java.text.NumberFormat;
    import java.text.SimpleDateFormat;
    import java.util.*;
    import java.util.Map.Entry;
    
    import javax.servlet.ServletOutputStream;
    import javax.servlet.http.HttpServletResponse;
    
    import com.zyq.entity.User;
    import org.apache.poi.hssf.usermodel.HSSFDataValidation;
    import org.apache.poi.hssf.usermodel.HSSFWorkbook;
    import org.apache.poi.poifs.filesystem.POIFSFileSystem;
    import org.apache.poi.ss.usermodel.Cell;
    import org.apache.poi.ss.usermodel.CellStyle;
    import org.apache.poi.ss.usermodel.CellType;
    import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
    import org.apache.poi.ss.usermodel.DataValidation;
    import org.apache.poi.ss.usermodel.DataValidationConstraint;
    import org.apache.poi.ss.usermodel.DataValidationHelper;
    import org.apache.poi.ss.usermodel.Drawing;
    import org.apache.poi.ss.usermodel.FillPatternType;
    import org.apache.poi.ss.usermodel.HorizontalAlignment;
    import org.apache.poi.ss.usermodel.IndexedColors;
    import org.apache.poi.ss.usermodel.Row;
    import org.apache.poi.ss.usermodel.Sheet;
    import org.apache.poi.ss.usermodel.VerticalAlignment;
    import org.apache.poi.ss.usermodel.Workbook;
    import org.apache.poi.ss.util.CellRangeAddress;
    import org.apache.poi.ss.util.CellRangeAddressList;
    import org.apache.poi.xssf.streaming.SXSSFWorkbook;
    import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;
    import org.springframework.web.multipart.MultipartFile;
    
    import com.alibaba.fastjson.JSONArray;
    import com.alibaba.fastjson.JSONObject;
    
    /**
     * Excel导入导出工具类
     *
     * @author sunnyzyq
     * @date 2021/12/17
     */
    public class ExcelUtils {
    
        private static final String XLSX = ".xlsx";
        private static final String XLS = ".xls";
        public static final String ROW_MERGE = "row_merge";
        public static final String COLUMN_MERGE = "column_merge";
        private static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";
        private static final String ROW_NUM = "rowNum";
        private static final String ROW_DATA = "rowData";
        private static final String ROW_TIPS = "rowTips";
        private static final int CELL_OTHER = 0;
        private static final int CELL_ROW_MERGE = 1;
        private static final int CELL_COLUMN_MERGE = 2;
        private static final int IMG_HEIGHT = 30;
        private static final int IMG_WIDTH = 30;
        private static final char LEAN_LINE = '/';
        private static final int BYTES_DEFAULT_LENGTH = 10240;
        private static final NumberFormat NUMBER_FORMAT = NumberFormat.getNumberInstance();
    
    
        public static <T> List<T> readFile(File file, Class<T> clazz) throws Exception {
            JSONArray array = readFile(file);
            return getBeanList(array, clazz);
        }
    
        public static <T> List<T> readMultipartFile(MultipartFile mFile, Class<T> clazz) throws Exception {
            JSONArray array = readMultipartFile(mFile);
            return getBeanList(array, clazz);
        }
    
        public static JSONArray readFile(File file) throws Exception {
            return readExcel(null, file);
        }
    
        public static JSONArray readMultipartFile(MultipartFile mFile) throws Exception {
            return readExcel(mFile, null);
        }
    
        private static <T> List<T> getBeanList(JSONArray array, Class<T> clazz) throws Exception {
            List<T> list = new ArrayList<>();
            Map<Integer, String> uniqueMap = new HashMap<>(16);
            for (int i = 0; i < array.size(); i++) {
                list.add(getBean(clazz, array.getJSONObject(i), uniqueMap));
            }
            return list;
        }
    
        /**
         * 获取每个对象的数据
         */
        private static <T> T getBean(Class<T> c, JSONObject obj, Map<Integer, String> uniqueMap) throws Exception {
            T t = c.newInstance();
            Field[] fields = c.getDeclaredFields();
            List<String> errMsgList = new ArrayList<>();
            boolean hasRowTipsField = false;
            StringBuilder uniqueBuilder = new StringBuilder();
            int rowNum = 0;
            for (Field field : fields) {
                // 行号
                if (field.getName().equals(ROW_NUM)) {
                    rowNum = obj.getInteger(ROW_NUM);
                    field.setAccessible(true);
                    field.set(t, rowNum);
                    continue;
                }
                // 是否需要设置异常信息
                if (field.getName().equals(ROW_TIPS)) {
                    hasRowTipsField = true;
                    continue;
                }
                // 原始数据
                if (field.getName().equals(ROW_DATA)) {
                    field.setAccessible(true);
                    field.set(t, obj.toString());
                    continue;
                }
                // 设置对应属性值
                setFieldValue(t,field, obj, uniqueBuilder, errMsgList);
            }
            // 数据唯一性校验
            if (uniqueBuilder.length() > 0) {
                if (uniqueMap.containsValue(uniqueBuilder.toString())) {
                    Set<Integer> rowNumKeys = uniqueMap.keySet();
                    for (Integer num : rowNumKeys) {
                        if (uniqueMap.get(num).equals(uniqueBuilder.toString())) {
                            errMsgList.add(String.format("数据唯一性校验失败,(%s)与第%s行重复)", uniqueBuilder, num));
                        }
                    }
                } else {
                    uniqueMap.put(rowNum, uniqueBuilder.toString());
                }
            }
            // 失败处理
            if (errMsgList.isEmpty() && !hasRowTipsField) {
                return t;
            }
            StringBuilder sb = new StringBuilder();
            int size = errMsgList.size();
            for (int i = 0; i < size; i++) {
                if (i == size - 1) {
                    sb.append(errMsgList.get(i));
                } else {
                    sb.append(errMsgList.get(i)).append(";");
                }
            }
            // 设置错误信息
            for (Field field : fields) {
                if (field.getName().equals(ROW_TIPS)) {
                    field.setAccessible(true);
                    field.set(t, sb.toString());
                }
            }
            return t;
        }
    
        private static <T> void setFieldValue(T t, Field field, JSONObject obj, StringBuilder uniqueBuilder, List<String> errMsgList) {
            // 获取 ExcelImport 注解属性
            ExcelImport annotation = field.getAnnotation(ExcelImport.class);
            if (annotation == null) {
                return;
            }
            String cname = annotation.value();
            if (cname.trim().length() == 0) {
                return;
            }
            // 获取具体值
            String val = null;
            if (obj.containsKey(cname)) {
                val = getString(obj.getString(cname));
            }
            if (val == null) {
                return;
            }
            field.setAccessible(true);
            // 判断是否必填
            boolean require = annotation.required();
            if (require && val.isEmpty()) {
                errMsgList.add(String.format("[%s]不能为空", cname));
                return;
            }
            // 数据唯一性获取
            boolean unique = annotation.unique();
            if (unique) {
                if (uniqueBuilder.length() > 0) {
                    uniqueBuilder.append("--").append(val);
                } else {
                    uniqueBuilder.append(val);
                }
            }
            // 判断是否超过最大长度
            int maxLength = annotation.maxLength();
            if (maxLength > 0 && val.length() > maxLength) {
                errMsgList.add(String.format("[%s]长度不能超过%s个字符(当前%s个字符)", cname, maxLength, val.length()));
            }
            // 判断当前属性是否有映射关系
            LinkedHashMap<String, String> kvMap = getKvMap(annotation.kv());
            if (!kvMap.isEmpty()) {
                boolean isMatch = false;
                for (String key : kvMap.keySet()) {
                    if (kvMap.get(key).equals(val)) {
                        val = key;
                        isMatch = true;
                        break;
                    }
                }
                if (!isMatch) {
                    errMsgList.add(String.format("[%s]的值不正确(当前值为%s)", cname, val));
                    return;
                }
            }
            // 其余情况根据类型赋值
            String fieldClassName = field.getType().getSimpleName();
            try {
                if ("String".equalsIgnoreCase(fieldClassName)) {
                    field.set(t, val);
                } else if ("boolean".equalsIgnoreCase(fieldClassName)) {
                    field.set(t, Boolean.valueOf(val));
                } else if ("int".equalsIgnoreCase(fieldClassName) || "Integer".equals(fieldClassName)) {
                    try {
                        field.set(t, Integer.valueOf(val));
                    } catch (NumberFormatException e) {
                        errMsgList.add(String.format("[%s]的值格式不正确(当前值为%s)", cname, val));
                    }
                } else if ("double".equalsIgnoreCase(fieldClassName)) {
                    field.set(t, Double.valueOf(val));
                } else if ("long".equalsIgnoreCase(fieldClassName)) {
                    field.set(t, Long.valueOf(val));
                } else if ("BigDecimal".equalsIgnoreCase(fieldClassName)) {
                    field.set(t, new BigDecimal(val));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    
        private static JSONArray readExcel(MultipartFile mFile, File file) throws IOException {
            boolean fileNotExist = (file == null || !file.exists());
            if (mFile == null && fileNotExist) {
                return new JSONArray();
            }
            // 解析表格数据
            InputStream in;
            String fileName;
            if (mFile != null) {
                // 上传文件解析
                in = mFile.getInputStream();
                fileName = getString(mFile.getOriginalFilename()).toLowerCase();
            } else {
                // 本地文件解析
                in = new FileInputStream(file);
                fileName = file.getName().toLowerCase();
            }
            Workbook book;
            if (fileName.endsWith(XLSX)) {
                book = new XSSFWorkbook(in);
            } else if (fileName.endsWith(XLS)) {
                POIFSFileSystem poifsFileSystem = new POIFSFileSystem(in);
                book = new HSSFWorkbook(poifsFileSystem);
            } else {
                return new JSONArray();
            }
            JSONArray array = read(book);
            book.close();
            in.close();
            return array;
        }
    
        private static JSONArray read(Workbook book) {
            // 获取 Excel 文件第一个 Sheet 页面
            Sheet sheet = book.getSheetAt(0);
            return readSheet(sheet);
        }
    
        private static JSONArray readSheet(Sheet sheet) {
            // 首行下标
            int rowStart = sheet.getFirstRowNum();
            // 尾行下标
            int rowEnd = sheet.getLastRowNum();
            // 获取表头行
            Row headRow = sheet.getRow(rowStart);
            if (headRow == null) {
                return new JSONArray();
            }
            int cellStart = headRow.getFirstCellNum();
            int cellEnd = headRow.getLastCellNum();
            Map<Integer, String> keyMap = new HashMap<>();
            for (int j = cellStart; j < cellEnd; j++) {
                // 获取表头数据
                String val = getCellValue(headRow.getCell(j));
                if (val != null && val.trim().length() != 0) {
                    keyMap.put(j, val);
                }
            }
            // 如果表头没有数据则不进行解析
            if (keyMap.isEmpty()) {
                return (JSONArray) Collections.emptyList();
            }
            // 获取每行JSON对象的值
            JSONArray array = new JSONArray();
            // 如果首行与尾行相同，表明只有一行，返回表头数据
            if (rowStart == rowEnd) {
                JSONObject obj = new JSONObject();
                // 添加行号
                obj.put(ROW_NUM, 1);
                for (int i : keyMap.keySet()) {
                    obj.put(keyMap.get(i), "");
                }
                array.add(obj);
                return array;
            }
            for (int i = rowStart + 1; i <= rowEnd; i++) {
                Row eachRow = sheet.getRow(i);
                JSONObject obj = new JSONObject();
                // 添加行号
                obj.put(ROW_NUM, i + 1);
                StringBuilder sb = new StringBuilder();
                for (int k = cellStart; k < cellEnd; k++) {
                    if (eachRow != null) {
                        String val = getCellValue(eachRow.getCell(k));
                        // 所有数据添加到里面，用于判断该行是否为空
                        sb.append(val);
                        obj.put(keyMap.get(k), val);
                    }
                }
                if (sb.length() > 0) {
                    array.add(obj);
                }
            }
            return array;
        }
    
        private static String getCellValue(Cell cell) {
            // 空白或空
            if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                return "";
            }
            // String类型
            if (cell.getCellTypeEnum() == CellType.STRING) {
                String val = cell.getStringCellValue();
                if (val == null || val.trim().length() == 0) {
                    return "";
                }
                return val.trim();
            }
            // 数字类型
            if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                // 科学计数法类型
                return NUMBER_FORMAT.format(cell.getNumericCellValue()) + "";
            }
            // 布尔值类型
            if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
                return cell.getBooleanCellValue() + "";
            }
            // 错误类型
            return cell.getCellFormula();
        }
    
        public static <T> void exportTemplate(HttpServletResponse response, String fileName, Class<T> clazz) {
            exportTemplate(response, fileName, fileName, clazz, false);
        }
    
        public static <T> void exportTemplate(HttpServletResponse response, String fileName, String sheetName,
                                              Class<T> clazz) {
            exportTemplate(response, fileName, sheetName, clazz, false);
        }
    
        public static <T> void exportTemplate(HttpServletResponse response, String fileName, Class<T> clazz,
                                              boolean isContainExample) {
            exportTemplate(response, fileName, fileName, clazz, isContainExample);
        }
    
        public static <T> void exportTemplate(HttpServletResponse response, String fileName, String sheetName,
                                              Class<T> clazz, boolean isContainExample) {
            // 获取表头字段
            List<ExcelClassField> headFieldList = getExcelClassFieldList(clazz);
            // 获取表头数据和示例数据
            List<List<Object>> sheetDataList = new ArrayList<>();
            List<Object> headList = new ArrayList<>();
            List<Object> exampleList = new ArrayList<>();
            Map<Integer, List<String>> selectMap = new LinkedHashMap<>();
            for (int i = 0; i < headFieldList.size(); i++) {
                ExcelClassField each = headFieldList.get(i);
                headList.add(each.getName());
                exampleList.add(each.getExample());
                LinkedHashMap<String, String> kvMap = each.getKvMap();
                if (kvMap != null && kvMap.size() > 0) {
                    selectMap.put(i, new ArrayList<>(kvMap.values()));
                }
            }
            sheetDataList.add(headList);
            if (isContainExample) {
                sheetDataList.add(exampleList);
            }
            // 导出数据
            export(response, fileName, sheetName, sheetDataList, selectMap);
        }
    
        private static <T> List<ExcelClassField> getExcelClassFieldList(Class<T> clazz) {
            // 解析所有字段
            Field[] fields = clazz.getDeclaredFields();
            boolean hasExportAnnotation = false;
            Map<Integer, List<ExcelClassField>> map = new LinkedHashMap<>();
            List<Integer> sortList = new ArrayList<>();
            for (Field field : fields) {
                ExcelClassField cf = getExcelClassField(field);
                if (cf.getHasAnnotation() == 1) {
                    hasExportAnnotation = true;
                }
                int sort = cf.getSort();
                if (map.containsKey(sort)) {
                    map.get(sort).add(cf);
                } else {
                    List<ExcelClassField> list = new ArrayList<>();
                    list.add(cf);
                    sortList.add(sort);
                    map.put(sort, list);
                }
            }
            Collections.sort(sortList);
            // 获取表头
            List<ExcelClassField> headFieldList = new ArrayList<>();
            if (hasExportAnnotation) {
                for (Integer sort : sortList) {
                    for (ExcelClassField cf : map.get(sort)) {
                        if (cf.getHasAnnotation() == 1) {
                            headFieldList.add(cf);
                        }
                    }
                }
            } else {
                headFieldList.addAll(map.get(0));
            }
            return headFieldList;
        }
    
        private static ExcelClassField getExcelClassField(Field field) {
            ExcelClassField cf = new ExcelClassField();
            String fieldName = field.getName();
            cf.setFieldName(fieldName);
            ExcelExport annotation = field.getAnnotation(ExcelExport.class);
            // 无 ExcelExport 注解情况
            if (annotation == null) {
                cf.setHasAnnotation(0);
                cf.setName(fieldName);
                cf.setSort(0);
                return cf;
            }
            // 有 ExcelExport 注解情况
            cf.setHasAnnotation(1);
            cf.setName(annotation.value());
            String example = getString(annotation.example());
            if (!example.isEmpty()) {
                if (isNumeric(example)) {
                    cf.setExample(Double.valueOf(example));
                } else {
                    cf.setExample(example);
                }
            } else {
                cf.setExample("");
            }
            cf.setSort(annotation.sort());
            // 解析映射
            String kv = getString(annotation.kv());
            cf.setKvMap(getKvMap(kv));
            return cf;
        }
    
        private static LinkedHashMap<String, String> getKvMap(String kv) {
            LinkedHashMap<String, String> kvMap = new LinkedHashMap<>();
            if (kv.isEmpty()) {
                return kvMap;
            }
            String[] kvs = kv.split(";");
            if (kvs.length == 0) {
                return kvMap;
            }
            for (String each : kvs) {
                String[] eachKv = getString(each).split("-");
                if (eachKv.length != 2) {
                    continue;
                }
                String k = eachKv[0];
                String v = eachKv[1];
                if (k.isEmpty() || v.isEmpty()) {
                    continue;
                }
                kvMap.put(k, v);
            }
            return kvMap;
        }
    
        /**
         * 导出表格到本地
         *
         * @param file      本地文件对象
         * @param sheetData 导出数据
         */
        public static void exportFile(File file, List<List<Object>> sheetData) {
            if (file == null) {
                System.out.println("文件创建失败");
                return;
            }
            if (sheetData == null) {
                sheetData = new ArrayList<>();
            }
            export(null, file, file.getName(), file.getName(), sheetData, null);
        }
    
        /**
         * 导出表格到本地
         *
         * @param <T>      导出数据类似，和K类型保持一致
         * @param filePath 文件父路径（如：D:/doc/excel/）
         * @param fileName 文件名称（不带尾缀，如：学生表）
         * @param list     导出数据
         * @throws IOException IO异常
         */
        public static <T> File exportFile(String filePath, String fileName, List<T> list) throws IOException {
            File file = getFile(filePath, fileName);
            List<List<Object>> sheetData = getSheetData(list);
            exportFile(file, sheetData);
            return file;
        }
    
        /**
         * 获取文件
         *
         * @param filePath filePath 文件父路径（如：D:/doc/excel/）
         * @param fileName 文件名称（不带尾缀，如：用户表）
         * @return 本地File文件对象
         */
        private static File getFile(String filePath, String fileName) throws IOException {
            String dirPath = getString(filePath);
            String fileFullPath;
            if (dirPath.isEmpty()) {
                fileFullPath = fileName;
            } else {
                // 判定文件夹是否存在，如果不存在，则级联创建
                File dirFile = new File(dirPath);
                if (!dirFile.exists()) {
                    dirFile.mkdirs();
                }
                // 获取文件夹全名
                if (dirPath.endsWith(String.valueOf(LEAN_LINE))) {
                    fileFullPath = dirPath + fileName + XLSX;
                } else {
                    fileFullPath = dirPath + LEAN_LINE + fileName + XLSX;
                }
            }
            System.out.println(fileFullPath);
            File file = new File(fileFullPath);
            if (!file.exists()) {
                file.createNewFile();
            }
            return file;
        }
    
        private static <T> List<List<Object>> getSheetData(List<T> list) {
            // 获取表头字段
            List<ExcelClassField> excelClassFieldList = getExcelClassFieldList(list.get(0).getClass());
            List<String> headFieldList = new ArrayList<>();
            List<Object> headList = new ArrayList<>();
            Map<String, ExcelClassField> headFieldMap = new HashMap<>();
            for (ExcelClassField each : excelClassFieldList) {
                String fieldName = each.getFieldName();
                headFieldList.add(fieldName);
                headFieldMap.put(fieldName, each);
                headList.add(each.getName());
            }
            // 添加表头名称
            List<List<Object>> sheetDataList = new ArrayList<>();
            sheetDataList.add(headList);
            // 获取表数据
            for (T t : list) {
                Map<String, Object> fieldDataMap = getFieldDataMap(t);
                Set<String> fieldDataKeys = fieldDataMap.keySet();
                List<Object> rowList = new ArrayList<>();
                for (String headField : headFieldList) {
                    if (!fieldDataKeys.contains(headField)) {
                        continue;
                    }
                    Object data = fieldDataMap.get(headField);
                    if (data == null) {
                        rowList.add("");
                        continue;
                    }
                    ExcelClassField cf = headFieldMap.get(headField);
                    // 判断是否有映射关系
                    LinkedHashMap<String, String> kvMap = cf.getKvMap();
                    if (kvMap == null || kvMap.isEmpty()) {
                        rowList.add(data);
                        continue;
                    }
                    String val = kvMap.get(data.toString());
                    if (isNumeric(val)) {
                        rowList.add(Double.valueOf(val));
                    } else {
                        rowList.add(val);
                    }
                }
                sheetDataList.add(rowList);
            }
            return sheetDataList;
        }
    
        private static <T> Map<String, Object> getFieldDataMap(T t) {
            Map<String, Object> map = new HashMap<>();
            Field[] fields = t.getClass().getDeclaredFields();
            try {
                for (Field field : fields) {
                    String fieldName = field.getName();
                    field.setAccessible(true);
                    Object object = field.get(t);
                    map.put(fieldName, object);
                }
            } catch (IllegalArgumentException | IllegalAccessException e) {
                e.printStackTrace();
            }
            return map;
        }
    
        public static void exportEmpty(HttpServletResponse response, String fileName) {
            List<List<Object>> sheetDataList = new ArrayList<>();
            List<Object> headList = new ArrayList<>();
            headList.add("导出无数据");
            sheetDataList.add(headList);
            export(response, fileName, sheetDataList);
        }
    
        public static void export(HttpServletResponse response, String fileName, List<List<Object>> sheetDataList) {
            export(response, fileName, fileName, sheetDataList, null);
        }
    
        public static void export(HttpServletResponse response, String fileName, String sheetName,
                                  List<List<Object>> sheetDataList) {
            export(response, fileName, sheetName, sheetDataList, null);
        }
    
        public static void export(HttpServletResponse response, String fileName, String sheetName,
                                  List<List<Object>> sheetDataList, Map<Integer, List<String>> selectMap) {
            export(response, null, fileName, sheetName, sheetDataList, selectMap);
        }
    
        public static <T, K> void export(HttpServletResponse response, String fileName, List<T> list, Class<K> template) {
            // list 是否为空
            boolean lisIsEmpty = list == null || list.isEmpty();
            // 如果模板数据为空，且导入的数据为空，则导出空文件
            if (template == null && lisIsEmpty) {
                exportEmpty(response, fileName);
                return;
            }
            // 如果 list 数据，则导出模板数据
            if (lisIsEmpty) {
                exportTemplate(response, fileName, template);
                return;
            }
            // 导出数据
            List<List<Object>> sheetDataList = getSheetData(list);
            export(response, fileName, sheetDataList);
        }
    
        public static void export(HttpServletResponse response, String fileName, List<List<Object>> sheetDataList, Map<Integer, List<String>> selectMap) {
            export(response, fileName, fileName, sheetDataList, selectMap);
        }
    
        private static void export(HttpServletResponse response, File file, String fileName, String sheetName,
                                   List<List<Object>> sheetDataList, Map<Integer, List<String>> selectMap) {
            // 整个 Excel 表格 book 对象
            SXSSFWorkbook book = new SXSSFWorkbook();
            // 每个 Sheet 页
            Sheet sheet = book.createSheet(sheetName);
            Drawing<?> patriarch = sheet.createDrawingPatriarch();
            // 设置表头背景色（灰色）
            CellStyle headStyle = book.createCellStyle();
            headStyle.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.index);
            headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headStyle.setAlignment(HorizontalAlignment.CENTER);
            headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            // 设置表身背景色（默认色）
            CellStyle rowStyle = book.createCellStyle();
            rowStyle.setAlignment(HorizontalAlignment.CENTER);
            rowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            // 设置表格列宽度（默认为15个字节）
            sheet.setDefaultColumnWidth(15);
            // 创建合并算法数组
            int rowLength = sheetDataList.size();
            int columnLength = sheetDataList.get(0).size();
            int[][] mergeArray = new int[rowLength][columnLength];
            for (int i = 0; i < sheetDataList.size(); i++) {
                // 每个 Sheet 页中的行数据
                Row row = sheet.createRow(i);
                List<Object> rowList = sheetDataList.get(i);
                for (int j = 0; j < rowList.size(); j++) {
                    // 每个行数据中的单元格数据
                    Object o = rowList.get(j);
                    int v = 0;
                    if (o instanceof URL) {
                        // 如果要导出图片的话, 链接需要传递 URL 对象
                        setCellPicture(book, row, patriarch, i, j, (URL) o);
                    } else {
                        Cell cell = row.createCell(j);
                        if (i == 0) {
                            // 第一行为表头行，采用灰色底背景
                            v = setCellValue(cell, o, headStyle);
                        } else {
                            // 其他行为数据行，默认白底色
                            v = setCellValue(cell, o, rowStyle);
                        }
                    }
                    mergeArray[i][j] = v;
                }
            }
            // 合并单元格
            mergeCells(sheet, mergeArray);
            // 设置下拉列表
            setSelect(sheet, selectMap);
            // 写数据
            if (response != null) {
                // 前端导出
                try {
                    write(response, book, fileName);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                // 本地导出
                FileOutputStream fos;
                try {
                    fos = new FileOutputStream(file);
                    ByteArrayOutputStream ops = new ByteArrayOutputStream();
                    book.write(ops);
                    fos.write(ops.toByteArray());
                    fos.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    
        /**
         * 合并当前Sheet页的单元格
         *
         * @param sheet      当前 sheet 页
         * @param mergeArray 合并单元格算法
         */
        private static void mergeCells(Sheet sheet, int[][] mergeArray) {
            // 横向合并
            for (int x = 0; x < mergeArray.length; x++) {
                int[] arr = mergeArray[x];
                boolean merge = false;
                int y1 = 0;
                int y2 = 0;
                for (int y = 0; y < arr.length; y++) {
                    int value = arr[y];
                    if (value == CELL_COLUMN_MERGE) {
                        if (!merge) {
                            y1 = y;
                        }
                        y2 = y;
                        merge = true;
                    } else {
                        merge = false;
                        if (y1 > 0) {
                            sheet.addMergedRegion(new CellRangeAddress(x, x, (y1 - 1), y2));
                        }
                        y1 = 0;
                        y2 = 0;
                    }
                }
                if (y1 > 0) {
                    sheet.addMergedRegion(new CellRangeAddress(x, x, (y1 - 1), y2));
                }
            }
            // 纵向合并
            int xLen = mergeArray.length;
            int yLen = mergeArray[0].length;
            for (int y = 0; y < yLen; y++) {
                boolean merge = false;
                int x1 = 0;
                int x2 = 0;
                for (int x = 0; x < xLen; x++) {
                    int value = mergeArray[x][y];
                    if (value == CELL_ROW_MERGE) {
                        if (!merge) {
                            x1 = x;
                        }
                        x2 = x;
                        merge = true;
                    } else {
                        merge = false;
                        if (x1 > 0) {
                            sheet.addMergedRegion(new CellRangeAddress((x1 - 1), x2, y, y));
                        }
                        x1 = 0;
                        x2 = 0;
                    }
                }
                if (x1 > 0) {
                    sheet.addMergedRegion(new CellRangeAddress((x1 - 1), x2, y, y));
                }
            }
        }
    
        private static void write(HttpServletResponse response, SXSSFWorkbook book, String fileName) throws IOException {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            String name = new String(fileName.getBytes("GBK"), "ISO8859_1") + XLSX;
            response.addHeader("Content-Disposition", "attachment;filename=" + name);
            ServletOutputStream out = response.getOutputStream();
            book.write(out);
            out.flush();
            out.close();
        }
    
        private static int setCellValue(Cell cell, Object o, CellStyle style) {
            // 设置样式
            cell.setCellStyle(style);
            // 数据为空时
            if (o == null) {
                cell.setCellType(CellType.STRING);
                cell.setCellValue("");
                return CELL_OTHER;
            }
            // 是否为字符串
            if (o instanceof String) {
                String s = o.toString();
                if (isNumeric(s)) {
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(Double.parseDouble(s));
                    return CELL_OTHER;
                } else {
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(s);
                }
                if (s.equals(ROW_MERGE)) {
                    return CELL_ROW_MERGE;
                } else if (s.equals(COLUMN_MERGE)) {
                    return CELL_COLUMN_MERGE;
                } else {
                    return CELL_OTHER;
                }
            }
            // 是否为字符串
            if (o instanceof Integer || o instanceof Long || o instanceof Double || o instanceof Float) {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(Double.parseDouble(o.toString()));
                return CELL_OTHER;
            }
            // 是否为Boolean
            if (o instanceof Boolean) {
                cell.setCellType(CellType.BOOLEAN);
                cell.setCellValue((Boolean) o);
                return CELL_OTHER;
            }
            // 如果是BigDecimal，则默认3位小数
            if (o instanceof BigDecimal) {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(((BigDecimal) o).setScale(3, RoundingMode.HALF_UP).doubleValue());
                return CELL_OTHER;
            }
            // 如果是Date数据，则显示格式化数据
            if (o instanceof Date) {
                cell.setCellType(CellType.STRING);
                cell.setCellValue(formatDate((Date) o));
                return CELL_OTHER;
            }
            // 如果是其他，则默认字符串类型
            cell.setCellType(CellType.STRING);
            cell.setCellValue(o.toString());
            return CELL_OTHER;
        }
    
        private static void setCellPicture(SXSSFWorkbook wb, Row sr, Drawing<?> patriarch, int x, int y, URL url) {
            // 设置图片宽高
            sr.setHeight((short) (IMG_WIDTH * IMG_HEIGHT));
            // （jdk1.7版本try中定义流可自动关闭）
            try (InputStream is = url.openStream(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                byte[] buff = new byte[BYTES_DEFAULT_LENGTH];
                int rc;
                while ((rc = is.read(buff, 0, BYTES_DEFAULT_LENGTH)) > 0) {
                    outputStream.write(buff, 0, rc);
                }
                // 设置图片位置
                XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, y, x, y + 1, x + 1);
                // 设置这个，图片会自动填满单元格的长宽
                anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
                patriarch.createPicture(anchor, wb.addPicture(outputStream.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    
        private static String formatDate(Date date) {
            if (date == null) {
                return "";
            }
            SimpleDateFormat format = new SimpleDateFormat(DATE_FORMAT);
            return format.format(date);
        }
    
        private static void setSelect(Sheet sheet, Map<Integer, List<String>> selectMap) {
            if (selectMap == null || selectMap.isEmpty()) {
                return;
            }
            Set<Entry<Integer, List<String>>> entrySet = selectMap.entrySet();
            for (Entry<Integer, List<String>> entry : entrySet) {
                int y = entry.getKey();
                List<String> list = entry.getValue();
                if (list == null || list.isEmpty()) {
                    continue;
                }
                String[] arr = new String[list.size()];
                for (int i = 0; i < list.size(); i++) {
                    arr[i] = list.get(i);
                }
                DataValidationHelper helper = sheet.getDataValidationHelper();
                CellRangeAddressList addressList = new CellRangeAddressList(1, 65000, y, y);
                DataValidationConstraint dvc = helper.createExplicitListConstraint(arr);
                DataValidation dv = helper.createValidation(dvc, addressList);
                if (dv instanceof HSSFDataValidation) {
                    dv.setSuppressDropDownArrow(false);
                } else {
                    dv.setSuppressDropDownArrow(true);
                    dv.setShowErrorBox(true);
                }
                sheet.addValidationData(dv);
            }
        }
    
        private static boolean isNumeric(String str) {
            if ("0.0".equals(str)) {
                return true;
            }
            for (int i = str.length(); --i >= 0; ) {
                if (!Character.isDigit(str.charAt(i))) {
                    return false;
                }
            }
            return true;
        }
    
        private static String getString(String s) {
            if (s == null) {
                return "";
            }
            if (s.isEmpty()) {
                return s;
            }
            return s.trim();
        }
    
    }


### ExcelImport

    package com.zyq.util.excel;
    
    import java.lang.annotation.ElementType;
    import java.lang.annotation.Retention;
    import java.lang.annotation.RetentionPolicy;
    import java.lang.annotation.Target;
    
    /**
     * @author sunnyzyq
     * @date 2021/12/17
     */
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelImport {
    
        /** 字段名称 */
        String value();
    
        /** 导出映射，格式如：0-未知;1-男;2-女 */
        String kv() default "";
    
        /** 是否为必填字段（默认为非必填） */
        boolean required() default false;
    
        /** 最大长度（默认255） */
        int maxLength() default 255;
    
        /** 导入唯一性验证（多个字段则取联合验证） */
        boolean unique() default false;
    
    }

### ExcelExport

    package com.zyq.util.excel;
    
    import java.lang.annotation.ElementType;
    import java.lang.annotation.Retention;
    import java.lang.annotation.RetentionPolicy;
    import java.lang.annotation.Target;
    
    /**
     * @author sunnyzyq
     * @date 2021/12/17
     */
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelExport {
    
        /** 字段名称 */
        String value();
    
        /** 导出排序先后: 数字越小越靠前（默认按Java类字段顺序导出） */
        int sort() default 0;
    
        /** 导出映射，格式如：0-未知;1-男;2-女 */
        String kv() default "";
    
        /** 导出模板示例值（有值的话，直接取该值，不做映射） */
        String example() default "";
    
    }

### ExcelClassField

    package com.zyq.util.excel;
    
    import java.util.LinkedHashMap;
    
    /**
     * @author sunnyzyq
     * @date 2021/12/17
     */
    public class ExcelClassField {
    
        /** 字段名称 */
        private String fieldName;
    
        /** 表头名称 */
        private String name;
    
        /** 映射关系 */
        private LinkedHashMap<String, String> kvMap;
    
        /** 示例值 */
        private Object example;
    
        /** 排序 */
        private int sort;
    
        /** 是否为注解字段：0-否，1-是 */
        private int hasAnnotation;
    
        public String getFieldName() {
            return fieldName;
        }
    
        public void setFieldName(String fieldName) {
            this.fieldName = fieldName;
        }
    
        public String getName() {
            return name;
        }
    
        public void setName(String name) {
            this.name = name;
        }
    
        public LinkedHashMap<String, String> getKvMap() {
            return kvMap;
        }
    
        public void setKvMap(LinkedHashMap<String, String> kvMap) {
            this.kvMap = kvMap;
        }
    
        public Object getExample() {
            return example;
        }
    
        public void setExample(Object example) {
            this.example = example;
        }
    
        public int getSort() {
            return sort;
        }
    
        public void setSort(int sort) {
            this.sort = sort;
        }
    
        public int getHasAnnotation() {
            return hasAnnotation;
        }
    
        public void setHasAnnotation(int hasAnnotation) {
            this.hasAnnotation = hasAnnotation;
        }
    
    }


[![](https://blog.csdn.net/sunnyzyq/article/details/121994504?utm_medium=distribute.pc_feed_blog_category.none-task-blog-classify_tag-6.nonecasedepth_1-utm_source=distribute.pc_feed_blog_category.none-task-blog-classify_tag-6.nonecase](https://blog.csdn.net/sunnyzyq/article/details/121994504?utm_medium=distribute.pc_feed_blog_category.none-task-blog-classify_tag-6.nonecasedepth_1-utm_source=distribute.pc_feed_blog_category.none-task-blog-classify_tag-6.nonecase)
