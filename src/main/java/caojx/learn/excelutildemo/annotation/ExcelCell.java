package caojx.learn.excelutildemo.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出实体类注解
 *
 * @author caojx
 * @version $Id: ExcelCell.java,v 1.0 2019-04-03 17:06 caojx
 * @date 2019-04-03 17:06
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelCell {

    String value() default "";

    int width() default 11;

    String timeFormat() default "yyyy-MM-dd HH:mm:ss";
}