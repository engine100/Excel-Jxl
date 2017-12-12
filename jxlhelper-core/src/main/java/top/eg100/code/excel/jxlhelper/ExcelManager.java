/*
 * Created by Engine100 on 2016-11-30 11:12:25.
 *
 *      https://github.com/engine100
 *
 */
package top.eg100.code.excel.jxlhelper;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import top.eg100.code.excel.jxlhelper.annotations.ExcelContent;
import top.eg100.code.excel.jxlhelper.annotations.ExcelContentCellFormat;
import top.eg100.code.excel.jxlhelper.annotations.ExcelSheet;
import top.eg100.code.excel.jxlhelper.annotations.ExcelTitleCellFormat;

/**
 * import from excel to class or export beans to excel
 */
public class ExcelManager {

    Map<String, Field> fieldCache = new HashMap<>();
    private Map<String, Method> contentMethodsCache;
    private Map<Integer, String> titleCache = new HashMap<>();

    /**
     * write excel to only one sheet ,no format
     */
    public boolean toExcel(OutputStream excelStream, List<?> dataList) throws Exception {
        if (dataList == null || dataList.size() == 0) {
            return false;
        }
        Class<?> dataType = dataList.get(0).getClass();
        String sheetName = getSheetName(dataType);
        List<ExcelClassKey> keys = getKeys(dataType);
        WritableWorkbook workbook = null;
        try {

            // create one book
            workbook = Workbook.createWorkbook(excelStream);
            // create sheet
            WritableSheet sheet = workbook.createSheet(sheetName, 0);

            // add titles
            for (int x = 0; x < keys.size(); x++) {
                sheet.addCell(new Label(x, 0, keys.get(x).getTitle()));
            }
            fieldCache.clear();
            // add data
            for (int y = 0; y < dataList.size(); y++) {
                for (int x = 0; x < keys.size(); x++) {
                    String fieldName = keys.get(x).getFieldName();

                    Field field = getField(dataType, fieldName);
                    Object value = field.get(dataList.get(y));
                    String content = value != null ? value.toString() : "";

                    // below the title ,the data begin from y+1
                    sheet.addCell(new Label(x, y + 1, content));
                }
            }
//            workbook.write();
//            workbook.close();
//            excelStream.close();

        } catch (Exception e) {
            throw e;
        } finally {
            if (workbook != null) {

                try {
                    workbook.write();
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }
            try {
                excelStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return true;
    }

    public boolean toExcel(String fileAbsoluteName, List<?> dataList) throws Exception {

        File file = new File(fileAbsoluteName);
        if (file.exists()) {
            if (file.isDirectory()) {
//                throw new Exception("do you want to write content into a directory named "
//                        + fileAbsoluteName + " ? , please check your filePath");
            }
        }
        File folder = file.getParentFile();
        if (!folder.exists()) {
            folder.mkdirs();
        }

        OutputStream stream = new FileOutputStream(file, false);
        return toExcel(stream, dataList);
    }
    public boolean toExcel(File file, List<?> dataList) throws Exception {

        if (file.exists()) {
            if (file.isDirectory()) {
//                throw new Exception("do you want to write content into a directory named "
//                        + fileAbsoluteName + " ? , please check your filePath");
            }
        }
        File folder = file.getParentFile();
        if (!folder.exists()) {
            folder.mkdirs();
        }

        OutputStream stream = new FileOutputStream(file, false);
        return toExcel(stream, dataList);
    }

    /**
     * write excel ,only one sheet ,with format
     */
    public boolean toExcelWithFormat(OutputStream excelStream, List<?> dataList) throws Exception {
        if (dataList == null || dataList.size() == 0) {
            return false;
        }
        Class<?> dataType = dataList.get(0).getClass();
        String sheetName = getSheetName(dataType);
        List<ExcelClassKey> keys = getKeys(dataType);

        // create one book
        WritableWorkbook workbook = Workbook.createWorkbook(excelStream);
        // create sheet
        WritableSheet sheet = workbook.createSheet(sheetName, 0);

        // add titles
        // find title format
        Map<String, WritableCellFormat> titleFormats = getTitleFormat(dataType);
        for (int x = 0; x < keys.size(); x++) {
            String titleName = keys.get(x).getTitle();
            WritableCellFormat f = titleFormats.get(titleName);
            if (f != null) {
                sheet.addCell(new Label(x, 0, titleName, f));
            } else {
                sheet.addCell(new Label(x, 0, titleName));
            }
        }
        fieldCache.clear();
        // add data
        for (int y = 0; y < dataList.size(); y++) {
            for (int x = 0; x < keys.size(); x++) {
                // current data
                Object data = dataList.get(y);
                ExcelClassKey classKey = keys.get(x);

                // add content
                String fieldName = classKey.getFieldName();
                Field field = getField(dataType, fieldName);
                Object value = field.get(data);
                String content = value != null ? value.toString() : "";

                // add format
                String title = classKey.getTitle();
                WritableCellFormat contentFormat = getContentFormat(title, data);

                // below the title ,the data begin from y+1
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
     * find all titles' WritableCellFormat
     */
    private Map<String, WritableCellFormat> getTitleFormat(Class<?> clazz) throws Exception {
        Map<String, WritableCellFormat> titleFormat = new HashMap<>();
        Method[] methods = clazz.getDeclaredMethods();
        for (int m = 0; m < methods.length; m++) {

            Method method = methods[m];
            ExcelTitleCellFormat formatAnno = method.getAnnotation(ExcelTitleCellFormat.class);
            if (formatAnno == null) {
                continue;
            }

            method.setAccessible(true);
            WritableCellFormat format = null;

            try {
                format = (WritableCellFormat) method.invoke(null);
            } catch (Exception e) {
                throw new Exception("The method added ExcelTitleCellFormat must be the static method");
            }

            if (format != null) {
                String title = formatAnno.titleName();
                titleFormat.put(title, format);
            }

        }
        return titleFormat;
    }

    /**
     * find all methods with ExcelContentCellFormat
     */
    private Map<String, Method> getContentFormatMethods(Class<?> clazz) {
        Map<String, Method> contentMethods = new HashMap<>();
        Method[] methods = clazz.getDeclaredMethods();
        for (int m = 0; m < methods.length; m++) {

            Method method = methods[m];
            ExcelContentCellFormat formatAnno = method.getAnnotation(ExcelContentCellFormat.class);
            if (formatAnno == null) {
                continue;
            }
            contentMethods.put(formatAnno.titleName(), method);
        }
        return contentMethods;
    }

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

    private List<ExcelClassKey> getKeys(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelClassKey> keys = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            ExcelContent content = fields[i].getAnnotation(ExcelContent.class);
            if(content!=null){
                keys.add(new ExcelClassKey(content.titleName(), fields[i].getName(), content.index()));
            }
        }
        //sort to control the title index in excel
        Collections.sort(keys, new Comparator<ExcelClassKey>() {
            @Override
            public int compare(ExcelClassKey t1, ExcelClassKey t2) {
                return t1.getIndex() - t2.getIndex();
            }
        });

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
        if (sheet == null) {
            throw new RuntimeException(clazz.getSimpleName() + " : lost sheet name!");
        }
        String sheetName = sheet.sheetName();
        return sheetName;
    }

    /**
     * read excel ,it is usual read by sheet name
     * the sheet name must as same as the ExcelSheet annotation's sheetName on dataType
     */
    public <T> List<T> fromExcel(InputStream excelStream, Class<T> dataType) throws Exception {
        String sheetName = getSheetName(dataType);

        // read map in excel
        List<Map<String, String>> title_content_values = getMapFromExcel(excelStream, sheetName);
        if (title_content_values == null || title_content_values.size() == 0) {
            return null;
        }

        Map<String, String> value0 = title_content_values.get(0);
        List<ExcelClassKey> keys = getKeys(dataType);

        //if there is no ExcelContent annotation in class ,return null
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

        // parse data from content
        for (int n = 0; n < title_content_values.size(); n++) {
            Map<String, String> title_content = title_content_values.get(n);
            T data = dataType.newInstance();
            for (int k = 0; k < keys.size(); k++) {

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
     * read excel by map
     */
    public List<Map<String, String>> getMapFromExcel(InputStream excelStream, String sheetName) throws Exception {

        Workbook workBook = Workbook.getWorkbook(excelStream);
        Sheet sheet = workBook.getSheet(sheetName);

        // row num
        int yNum = sheet.getRows();
        // there is only tile or nothing
        if (yNum <= 1) {
            return null;
        }
        // column num
        int xNum = sheet.getColumns();

        // none column
        if (xNum <= 0) {
            return null;
        }
        List<Map<String, String>> values = new LinkedList<>();

        titleCache.clear();

        // yNum-1 is the data size , but not title

        for (int y = 0; y < yNum - 1; y++) {
            Map<String, String> value = new LinkedHashMap<>();
            for (int x = 0; x < xNum; x++) {
                //read title name
                String title = getExcelTitle(sheet, x);

                //read data,from second row
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
