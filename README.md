# Excel导出与上传解析Demo

### 1.主要依赖
```xml
     <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.9</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.9</version>
        </dependency>
```
### 2.入口
入口见ExcelTestController，SpringBoot项目可以直接运行

http://localhost:8082/excel/export  
http://localhost:8081/excel/upload  

 
### 3.导出Excel 
支持导出Excel导出，支持使用注解注解对应的实体  
    @ExcelSheet注解实体类，value值作为sheet名   
    @ExcelCell注解实体属性，value值可作为表头  
    
#### 3.1 StudentExcelVo.java  
![](../excel-util-demo/src/main/resources/images/excel-1.png)  

#### 3.2 导出
http://localhost:8082/excel/export    
![](../excel-util-demo/src/main/resources/images/excel-2.png)  

    
### 4.上传解析（导入）Excel      
支持解析Excel，返回一个map  
    key作为sheet顺序编号  
    value存储第1~n行的数据（去表头），行数据保存到list中 


导入第3步导出的文件演示   
http://localhost:8082/excel/upload    
![](../excel-util-demo/src/main/resources/images/excel-3.png)  

导入结果  
![](../excel-util-demo/src/main/resources/images/excel-4.png)  


 
        



参考：  
https://blog.csdn.net/yywdys/article/details/82900815  
[Java实现Excel导入导出](https://github.com/caojx-git/learn-java-notes/blob/master/java/java%E5%AE%9E%E7%8E%B0excel%E5%AF%BC%E5%85%A5%E5%AF%BC%E5%87%BA.md)