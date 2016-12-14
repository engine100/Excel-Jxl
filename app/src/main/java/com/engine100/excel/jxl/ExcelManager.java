package com.engine100.excel.jxl;


import com.engine100.excel.jxl.annotations.ExcelContent;
import com.engine100.excel.jxl.annotations.ExcelContentCellFormat;
import com.engine100.excel.jxl.annotations.ExcelSheet;
import com.engine100.excel.jxl.annotations.ExcelTitleCellFormat;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * @author ZhuChengCheng
 * @description 转换类</br>
 * @github https://github.com/engine100
 * @time 2016年11月29日 - 下午11:14:11
 */
public class ExcelManager {

    Map<String, Field> fieldCache = new HashMap<>();
    private Map<String, Method> contentMethodsCache;
    private Map<Integer, String> titleCache = new HashMap<>();

    /**
     * 写表格，只有一个表
     *
     * @param excelStream
     * @param datas
     * @return
     * @throws Exception
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月29日 - 下午11:26:20
     */
    public boolean toExcel(OutputStream excelStream, List<?> datas) throws Exception {
        if (datas == null || datas.size() == 0) {
            return false;
        }
        Class<?> dataType = datas.get(0).getClass();
        String sheetName = getSheetName(dataType);
        List<ExcelClassKey> keys = getKeys(dataType);

        // 创建工作簿
        WritableWorkbook workbook = Workbook.createWorkbook(excelStream);
        // 建立sheet表格
        WritableSheet sheet = workbook.createSheet(sheetName, 0);

        // 添加标题
        for (int x = 0; x < keys.size(); x++) {
            sheet.addCell((WritableCell) new Label(x, 0, keys.get(x).getTitle()));
        }
        fieldCache.clear();
        // 添加数据
        for (int y = 0; y < datas.size(); y++) {
            for (int x = 0; x < keys.size(); x++) {
                String fieldName = keys.get(x).getFieldName();

                Field field = getField(dataType, fieldName);
                Object value = field.get(datas.get(y));
                String content = value != null ? value.toString() : "";
                // 添加标题后，数据从第2行开始，所以是y+1
                sheet.addCell(new Label(x, y + 1, content));
            }
        }
        workbook.write();
        workbook.close();
        excelStream.close();
        return true;
    }

    /**
     * 写表格，只有一个表，附带格式
     *
     * @param excelStream
     * @param datas
     * @return
     * @throws Exception
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 下午4:54:13
     */
    public boolean toExcelWithFormat(OutputStream excelStream, List<?> datas) throws Exception {
        if (datas == null || datas.size() == 0) {
            return false;
        }
        Class<?> dataType = datas.get(0).getClass();
        String sheetName = getSheetName(dataType);
        List<ExcelClassKey> keys = getKeys(dataType);

        // 创建工作簿
        WritableWorkbook workbook = Workbook.createWorkbook(excelStream);
        // 建立sheet表格
        WritableSheet sheet = workbook.createSheet(sheetName, 0);

        // 添加标题
        // 获取标题格式
        Map<String, WritableCellFormat> titleFormats = getTitleFormat(dataType);
        for (int x = 0; x < keys.size(); x++) {
            String titleName = keys.get(x).getTitle();
            WritableCellFormat f = titleFormats.get(titleName);
            if (f != null) {
                sheet.addCell((WritableCell) new Label(x, 0, titleName, f));
            } else {
                sheet.addCell((WritableCell) new Label(x, 0, titleName));
            }
        }
        fieldCache.clear();
        // 添加数据
        for (int y = 0; y < datas.size(); y++) {
            for (int x = 0; x < keys.size(); x++) {
                // 当前数据
                Object data = datas.get(y);
                ExcelClassKey classKey = keys.get(x);

                // 获得内容
                String fieldName = classKey.getFieldName();
                Field field = getField(dataType, fieldName);
                Object value = field.get(data);
                String content = value != null ? value.toString() : "";

                // 获得格式
                String title = classKey.getTitle();
                WritableCellFormat contentFormat = getContentFormat(title, data);

                // 添加标题后，数据从第2行开始，所以是y+1
                if (contentFormat != null) {
                    sheet.addCell(new Label(x, y + 1, content, contentFormat));
                } else {
                    sheet.addCell(new Label(x, y + 1, content));
                }
            }
        }
        workbook.write();
        workbook.close();
        excelStream.close();
        return true;
    }

    /**
     * 获得所有的标题的格式，标题为key
     *
     * @param clazz
     * @return
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 下午4:22:09
     */
    private Map<String, WritableCellFormat> getTitleFormat(Class<?> clazz) throws Exception {
        Map<String, WritableCellFormat> titleFormat = new HashMap<>();
        Method[] methods = clazz.getDeclaredMethods();
        for (int m = 0; m < methods.length; m++) {
            // 获得有标题注解的方法
            Method method = methods[m];
            ExcelTitleCellFormat formatAnno = method.getAnnotation(ExcelTitleCellFormat.class);
            if (formatAnno == null) {
                continue;
            }

            method.setAccessible(true);
            WritableCellFormat format = null;
            // 标题注解的必须是静态的方法
            try {
                format = (WritableCellFormat) method.invoke(null);
            } catch (Exception e) {
                throw new Exception("用ExcelTitleCellFormat 注解的标题格式方法必须是static");
            }

            if (format != null) {
                String title = formatAnno.titleName();
                titleFormat.put(title, format);
            }

        }
        return titleFormat;
    }

    /**
     * 获得所有的单元格的格式的方法，标题为key
     *
     * @param clazz
     * @return
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 下午4:22:09
     */
    private Map<String, Method> getContentFormatMethods(Class<?> clazz) {
        Map<String, Method> contentMethods = new HashMap<>();
        Method[] methods = clazz.getDeclaredMethods();
        for (int m = 0; m < methods.length; m++) {
            // 获得有内容注解的方法
            Method method = methods[m];
            ExcelContentCellFormat formatAnno = method.getAnnotation(ExcelContentCellFormat.class);
            if (formatAnno == null) {
                continue;
            }
            contentMethods.put(formatAnno.titleName(), method);
        }
        return contentMethods;
    }

    /**
     * 获取数据对应的单元格格式
     *
     * @param title
     * @param data
     * @return
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 下午4:39:44
     */
    private <T> WritableCellFormat getContentFormat(String title, T data) {
        if (contentMethodsCache == null) {
            contentMethodsCache = getContentFormatMethods(data.getClass());
        }

        Method method = contentMethodsCache.get(title);
        if (method == null) {
            return null;
        }

        method.setAccessible(true);
        WritableCellFormat format = null;
        try {
            format = (WritableCellFormat) method.invoke(data);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return format;
    }

    /**
     * 获取所有的需要导入导出Excel的注解
     *
     * @param clazz
     * @return
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 上午9:59:02
     */
    private List<ExcelClassKey> getKeys(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelClassKey> keys = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            ExcelContent content = fields[i].getAnnotation(ExcelContent.class);
            if (content != null) {
                keys.add(new ExcelClassKey(content.titleName(), fields[i].getName()));
            }
        }
        return keys;

    }

    private Field getField(Class<?> type, String fieldName) throws Exception {
        Field f = null;

        if (fieldCache.containsKey(fieldName)) {
            f = fieldCache.get(fieldName);
        } else {
            f = type.getDeclaredField(fieldName);
            fieldCache.put(fieldName, f);
        }
        f.setAccessible(true);
        return f;
    }

    private String getSheetName(Class<?> clazz) {
        ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
        String sheetName = sheet.sheetName();
        return sheetName;
    }

    /**
     * @return
     * @throws Exception
     * @description 读表格，只有一个表的情况</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月29日 - 下午11:25:17
     */
    public <T> List<T> fromExcel(InputStream excelStream, Class<T> dataType) throws Exception {
        String sheetName = getSheetName(dataType);

        // 获取键值对儿
        List<Map<String, String>> title_content_values = getMapFromExcel(excelStream, sheetName);
        if (title_content_values == null || title_content_values.size() == 0) {
            return null;
        }

        Map<String, String> value0 = title_content_values.get(0);
        List<ExcelClassKey> keys = getKeys(dataType);
        // 判断在实体里的注解标题在excel表里是否存在，不存在就说明表格不能格式化为所需实体
        boolean isExist = false;
        for (int kIndex = 0; kIndex < keys.size(); kIndex++) {
            String title = keys.get(kIndex).getTitle();
            if (value0.containsKey(title)) {
                isExist = true;
                break;
            }
        }
        if (!isExist) {
            return null;
        }

        List<T> datas = new ArrayList<>();
        fieldCache.clear();
        // 键值对儿映射数据
        for (int n = 0; n < title_content_values.size(); n++) {
            Map<String, String> title_content = title_content_values.get(n);
            T data = dataType.newInstance();
            for (int k = 0; k < keys.size(); k++) {
                // 根据title和字段值映射
                String title = keys.get(k).getTitle();
                String fieldName = keys.get(k).getFieldName();
                Field field = getField(dataType, fieldName);
                field.set(data, title_content.get(title));
            }
            datas.add(data);
        }
        return datas;
    }

    /**
     * 获取表的对应字段和值,<title,content>的形式
     *
     * @param excelStream
     * @param sheetName
     * @return
     * @throws Exception
     * @description</br>
     * @author ZhuChengCheng
     * @github https://github.com/engine100
     * @time 2016年11月30日 - 上午9:38:05
     */
    public List<Map<String, String>> getMapFromExcel(InputStream excelStream, String sheetName) throws Exception {

        Workbook workBook = Workbook.getWorkbook(excelStream);
        Sheet sheet = workBook.getSheet(sheetName);

        int yNum = sheet.getRows();// 行数
        // 只有标题或者什么都没有
        if (yNum <= 1) {
            return null;
        }
        int xNum = sheet.getColumns();// 列数
        // 一个字段都没有
        if (xNum <= 0) {
            return null;
        }
        List<Map<String, String>> values = new LinkedList<>();

        titleCache.clear();
        // yNum-1是数据的大小，去掉第一行标题
        for (int y = 0; y < yNum - 1; y++) {
            Map<String, String> value = new LinkedHashMap<>();
            for (int x = 0; x < xNum; x++) {
                // 读标题
                String title = getExcelTitle(sheet, x);
                // 读数据,读数据从第2行开始读
                String content = getContent(sheet, x, y + 1);
                value.put(title, content);
            }
            values.add(value);
        }

        workBook.close();
        return values;
    }

    private String getExcelTitle(Sheet sheet, int x) {
        String title;
        if (titleCache.containsKey(x)) {
            title = titleCache.get(x);
        } else {
            title = getContent(sheet, x, 0);
            titleCache.put(x, title);
        }
        return title;
        // return getContent(sheet, x, 0);
    }

    private String getContent(Sheet sheet, int x, int y) {
        Cell contentCell = sheet.getCell(x, y);
        String content = contentCell.getContents();
        return content != null ? content : "";
    }
}
