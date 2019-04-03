# Excel导出与上传解析Demo

### 1.入口
入口见ExcelTestController
 
### 2.导出Excel 
支持导出Excel导出，支持使用注解注解对应的实体  
    @ExcelSheet注解实体类，value值作为sheet名   
    @ExcelCell注解实体属性，value值可作为表头  
### 3.上传解析Excel      
支持解析Excel，返回一个map  
    key作为sheet行编号
    value为list，list按顺序存储第0行的0list到第n行的list
        



参考：  
https://blog.csdn.net/yywdys/article/details/82900815