package top.eg100.code.excel.jxlhelper;

/**
 * orm in excel and java bean fields ,
 * if the field in bean has ExcelContent annotation ,it can be export to excel
 */
class ExcelClassKey {

    /**
     * title in excel
     */
    private String title;
    /**
     * field Name in java bean
     */
    private String fieldName;

    public ExcelClassKey(String title, String fieldName) {
        this.title = title;
        this.fieldName = fieldName;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

}
