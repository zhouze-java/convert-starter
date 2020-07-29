### 功能
针对Word文档的一些常用操作
+ 格式转换,支持word文档转 doc,docx,pdf,html,text,jpeg,png 这些格式
+ 全局水印配置
+ 添加自定义水印
+ 克隆文档
+ txt 转 word
+ 转为只读文档
+ 模板导出/替换关键字/生成表格/插入图片
+ ...

### 自动装配
工具类, `WordConvertUtil`, `WordTemplateUtil` 已设置SpringBoot自动装配, 导包后可以直接注入使用   
  
### 依赖引入
先下载源码,之后 `mvn clean install` 安装到本地  
  
项目中引入依赖:  
```
<dependency>
    <groupId>com.easy.convert</groupId>
    <artifactId>convert-starter</artifactId>
    <version>1.0.0</version>
</dependency>
```
  
之后将 lib 目录下的 jar 包,扔到maven仓库中对应的文件夹即可    
  
### 全局水印设置
通过 application 配置文件设置,示例
```
file.config.word:
  watermarks:
    # 开启水印
    enable: true
    # 水印内容
    text: test
    ...
```
  
具体配置项可以看 `WordWatermarksProperties` 这个类  
  
### 模板导出功能
简介: 可以在 word 模板中使用占位符,之后调用方法,把占位符替换成自定义的数据.  
  
#### API 
```
/**
 * 模板渲染(包括文字，图片，表格)
 *
 * @param templatePath 模板路径
 * @param data 数据,Map<String,Object>/Map<String,Any?>
 * @param outputPath 输出路径
 * @throws Exception
 */
WordTemplateUtil.render(templatePath, data ,outputPath,watermarksEnable)
```

#### 自定义配置项
application 配置文件中,可以自定义模板相关的配置,如下:  

```
file.config.word:
  template:
    textSymbol:
    imageSymbol: 
    ...
```
  
具体可配置项可以看 `WordTemplateProperties` 这个类  
  
#### 关键字
目前一共可以处理4种关键字   
+ ${普通文本}
+ ${-带下划线的文本}
+ ${@图片}
+ ${#表格}

默认关键字规则: 语法前缀 + 类型表示 + 具体关键字 + 语法后缀   
   
这些都可以通过配置文件自定义    
  
#### 使用方法
首先需要有一个包含关键字的word文档,例如下面这样:   

![image](http://r.photo.store.qq.com/psc?/V10eEnSd0OPhSW/TmEUgtj9EK6.7V8ajmQrEHTDdHO68PWx5In.MRup0VhgwVAVLdPl2vwbkDbFXT*nBGjLnNHUrWFDPCo1.SWoiAJWHxvz2Ifca5NX5PHOFL0!/r)  

之后就是组合数据,然后调用API,示例代码如下:   
```
Map<String, Object> data = new HashMap<>();
data.put("name","zhangsan");
data.put("age", 19);

// 替换图片,指定路径和长,宽
ImageData imageData = new ImageData("C:\\Users\\Administrator\\Desktop\\testImg.png", 200.00, 200.00);
data.put("pic1", imageData);

// 表头设置
List<String> headers = new ArrayList<>();
headers.add("姓名");
headers.add("年龄");
headers.add("地址");

// 添加行
List<String> row1 = new ArrayList<>();
row1.add("张三");
row1.add("20");
row1.add("杭州");

List<String> row2 = new ArrayList<>();
row2.add("李四");
row2.add("20");
row2.add("杭州");

// 包装到 TableData 类中
List<TableData.Row> rows = new ArrayList<>();
rows.add(new TableData.Row(row1));
rows.add(new TableData.Row(row2));
TableData tableData1 = new TableData(headers, rows);

data.put("table1", tableData1);

// 调用api
wordTemplateUtil.render("/test/test.docx", data, "/test/test-out.docx");
```
  
执行后效果如下:  
  
![image](http://r.photo.store.qq.com/psc?/V10eEnSd0OPhSW/TmEUgtj9EK6.7V8ajmQrEFweXvHeLrSYWQe2XEHsRj690pQQCeAeYqRgiiExY6dnH2SxsLYAKZyI2FzOteLKL2HGUGoIM8HSBdyt4Zll.mE!/r)  
  
水印是根据全局设置添加,也可以通过重载的方法关闭   
  
