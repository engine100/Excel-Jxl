package com.engine100.excel.jxl.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 类成员里对应Excel里的标题
 *
 * @description </br>
 * @author ZhuChengCheng
 * @github https://github.com/engine100
 * @time 2016年11月30日 - 上午11:09:04
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ java.lang.annotation.ElementType.FIELD })
public @interface ExcelContent {
	public String titleName();

}