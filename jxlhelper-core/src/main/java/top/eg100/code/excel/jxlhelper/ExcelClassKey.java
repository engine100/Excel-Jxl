package top.eg100.code.excel.jxlhelper;

/**
 * orm in excel and java bean fields ,
 * if the field in bean has ExcelContent annotation ,it can be export to excel
 * 2017-12-12 add field index by engine100
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

    /**
     * sort title in excel
     */
    private int index;

    public ExcelClassKey(String title, String fieldName, int index) {
        this.title = title;
        this.fieldName = fieldName;
        this.index = index;
    }

    public int getIndex(){
        return index;
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
