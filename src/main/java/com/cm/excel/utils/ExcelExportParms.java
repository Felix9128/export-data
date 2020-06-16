package com.cm.excel.utils;

import java.util.ArrayList;
import java.util.List;

/**
 * ClassName: ExcelExportParms
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
public class ExcelExportParms<T>
{

	private static int DEFAULT_COLUMN_WIDTH = 15;

	private static String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";

	private static int DEFAULT_PAGE_SIZE = 20000;

	private static String DEFAULT_SHEET_TITLE = "sheet";

	private static String DEFAULT_FILE_NAME = "excel文件";
	/**
	 * 数据集合
	 */
	private List<ExcelCellAttr> cellAttrs = new ArrayList<>();

	/**
	 * 数据集合
	 */
	private List<T> datas = new ArrayList<>();

	/**
	 * 缺省列宽
	 */
	private int defaultColumnWidth = DEFAULT_COLUMN_WIDTH;

	/**
	 * 缺省每页大小
	 */
	private int pageSize = DEFAULT_PAGE_SIZE;

	/**
	 * 缺省Date类型
	 */
	private String defaultDatePattern = DEFAULT_DATE_PATTERN;

	/**
	 * sheet标题
	 */
	private String sheetTitle = DEFAULT_SHEET_TITLE;

	/**
	 * 文件名
	 */
	private String fileName = DEFAULT_FILE_NAME;

	public int getDefaultColumnWidth()
	{
		return defaultColumnWidth;
	}

	public ExcelExportParms<T> setDefaultColumnWidth(int defaultColumnWidth)
	{
		this.defaultColumnWidth = defaultColumnWidth;
		return this;
	}

	public String getDefaultDatePattern()
	{
		return defaultDatePattern;
	}

	public ExcelExportParms<T> setDefaultDatePattern(String defaultDatePattern)
	{
		this.defaultDatePattern = defaultDatePattern;
		return this;
	}

	public int getPageSize()
	{
		return pageSize;
	}

	public ExcelExportParms<T> setPageSize(int pageSize)
	{
		this.pageSize = pageSize;
		return this;
	}

	public String getSheetTitle()
	{
		return sheetTitle;
	}

	public ExcelExportParms<T> setSheetTitle(String sheetTitle)
	{
		this.sheetTitle = sheetTitle;
		return this;
	}

	public String getFileName()
	{
		return fileName;
	}

	public ExcelExportParms<T> setFileName(String fileName)
	{
		this.fileName = fileName;
		return this;
	}

	public List<ExcelCellAttr> getCellAttrs()
	{
		return cellAttrs;
	}

	public ExcelExportParms<T> setCellAttrs(List<ExcelCellAttr> cellAttrs)
	{
		this.cellAttrs = cellAttrs;
		return this;
	}

	public List<T> getDatas()
	{
		return datas;

	}

	public ExcelExportParms<T> setDatas(List<T> datas)
	{
		this.datas = datas;
		return this;
	}

}