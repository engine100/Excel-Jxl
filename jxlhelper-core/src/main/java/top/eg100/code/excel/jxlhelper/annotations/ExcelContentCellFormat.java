/*
 * Created by Engine100 on 2016-11-30 11:10:00.
 *
 *      https://github.com/engine100
 *
 */
package top.eg100.code.excel.jxlhelper.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * format the content.
 * usual,you can add it on method which return WritableCellFormat,
 * most times ,it doesn't fit the big picture
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({java.lang.annotation.ElementType.METHOD})
public @interface ExcelContentCellFormat {
    String titleName();
}