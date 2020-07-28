### 功能
针对Word文档的一些常用操作
+ 格式转换,支持word文档转 doc,docx,pdf,html,text,jpeg,png 这些格式
+ 全局水印配置
+ 添加自定义水印
+ 克隆文档
+ ...

### 自动装配
工具类, `WordConvertUtil` 已设置SpringBoot自动装配, 导包后可以直接注入使用   
  
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
  
