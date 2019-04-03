package caojx.learn.excelutildemo.pojo.search;

import lombok.Data;

import java.io.Serializable;

/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: StudentSearch.java,v 1.0 2019-04-03 17:30 caojx
 * @date 2019-04-03 17:30
 */
@Data
public class StudentSearch implements Serializable {

    /**
     * 年龄
     */
    private Integer age;
}