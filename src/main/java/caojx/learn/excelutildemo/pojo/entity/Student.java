package caojx.learn.excelutildemo.pojo.entity;

import lombok.Data;

import java.io.Serializable;

/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: Student.java,v 1.0 2019-04-03 17:34 caojx
 * @date 2019-04-03 17:34
 */
@Data
public class Student implements Serializable {
    /**
     * id
     */
    private Long id;

    /**
     * 姓名
     */
    private String name;

    /**
     * 年龄
     */
    private Integer age;
}