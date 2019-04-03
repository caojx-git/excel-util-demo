package caojx.learn.excelutildemo.pojo.vo;

import caojx.learn.excelutildemo.annotation.ExcelCell;
import caojx.learn.excelutildemo.annotation.ExcelSheet;
import lombok.Data;

import java.io.Serializable;

/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: StudentVo.java,v 1.0 2019-04-03 17:22 caojx
 * @date 2019-04-03 17:22
 */
@Data
@ExcelSheet("学生信息")
public class StudentExcelVo implements Serializable {

    @ExcelCell(value = "id")
    private Long id;

    @ExcelCell(value = "姓名")
    private String name;

    @ExcelCell(value = "年龄")
    private Integer age;
}