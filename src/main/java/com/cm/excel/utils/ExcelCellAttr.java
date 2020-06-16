package com.cm.excel.utils;

/**
 * ClassName: ExcelCellAttr
 * Function:  TODO  功能说明.
 * <p>
 * date: 2020年06月16日  09:44
 *
 * @author baize
 * @since JDK 1.8
 * <p>
 * Modified By： <修改人>
 * Modified Date: <修改日期，格式:YYYY-MM-DD>
 * Why & What is modified: <修改描述>
 */
public class ExcelCellAttr {


    /**
     * 字段标题
     */
    private String headerText;

    /**
     * 导出的字段
     */
    private String field;

    /**
     * 列宽
     */
    // private String columnWidth;

    /**
     * Date类型字段的日期格式
     */
    private String datePattern;
    /**
     * 数据格式,默认0  （具体见org.apache.poi.ss.usermodel.BuiltinFormats）
     */
    private int dataFormatIndex;

    public int getDataFormatIndex() {
        return dataFormatIndex;
    }

    public void setDataFormatIndex(int dataFormatIndex) {
        this.dataFormatIndex = dataFormatIndex;
    }

    public String getHeaderText()
    {
        return headerText;
    }

    public void setHeaderText(String headerText)
    {
        this.headerText = headerText;
    }

    public String getField()
    {
        return field;
    }

    public void setField(String field)
    {
        this.field = field;
    }

    // public String getColumnWidth()
    // {
    // return columnWidth;
    // }
    //
    // public void setColumnWidth(String columnWidth)
    // {
    // this.columnWidth = columnWidth;
    // }

    public String getDatePattern()
    {
        return datePattern;
    }

    public void setDatePattern(String datePattern)
    {
        this.datePattern = datePattern;
    }

    public ExcelCellAttr(String headerText, String field) {
        super();
        this.headerText = headerText;
        this.field = field;
    }

    public ExcelCellAttr(String headerText, String field, String datePattern) {
        super();
        this.headerText = headerText;
        this.field = field;
        this.datePattern = datePattern;
    }


    public ExcelCellAttr(String headerText, String field, int dataFormatIndex) {
        super();
        this.headerText = headerText;
        this.field = field;
        this.dataFormatIndex = dataFormatIndex;
    }
}