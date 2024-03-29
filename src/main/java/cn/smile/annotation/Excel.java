package cn.smile.annotation;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    /* excel 字段名称 */
    String value();

    /* 每列的宽度 */
    int width() default 16;

    /* 行的高度 */
    int height() default 14;

    /* 排序 */
    int sort() default -1;

    String dateFormat() default "";

    CellType cellType() default CellType.STRING;

}
