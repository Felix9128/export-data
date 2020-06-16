package com.cm.excel.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Slf4j
public class ExportExcel {

	// private static ThreadLocal<HttpServletResponse> responseLocal = new ThreadLocal<HttpServletResponse>();

	/**
	 * 导出excel数据
	 * @author baize
	 * @param excelAttr:
	 * @param response:
	 * @return
	 **/

	public static <T> void exportExcel(ExcelExportParms<T> excelAttr,HttpServletResponse response) throws IOException {
		// HttpServletResponse response = responseLocal.get();

		// 清空response
		response.reset();

		// 设置response的Header
		// 文件名
		String fileName = excelAttr.getFileName();
		fileName = new String(fileName.getBytes("UTF-8"), "ISO_8859_1");
		
		String sheetTitle = excelAttr.getSheetTitle();

		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

		// 每页大小
		int pageSize = excelAttr.getPageSize();
		// 日期格式
		String defaultDatePattern = excelAttr.getDefaultDatePattern();

		// 单元格属性
		List<ExcelCellAttr> attrs = excelAttr.getCellAttrs();

		// 导出数据
		List<T> dataset = excelAttr.getDatas();
        log.info("exportExcel 导出文件名:{},记录数:{}",excelAttr.getFileName(),dataset.size());
		OutputStream ouputStream = response.getOutputStream();

		// 声明一个工作薄
		XSSFWorkbook workbook = new XSSFWorkbook();

		// 生成一个样式
		XSSFCellStyle fieldHeaderStyle = workbook.createCellStyle();
		// 设置这些样式
		// fieldHeaderStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		fieldHeaderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 默认换行
		// style.setWrapText(true);

		// 生成一个字体
		XSSFFont fieldHeaderFont = workbook.createFont();
		fieldHeaderFont.setFontHeightInPoints((short) 12);
		fieldHeaderFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把字体应用到当前的样式
		fieldHeaderStyle.setFont(fieldHeaderFont);
		// 生成并设置另一个样式
		XSSFCellStyle fieldStyle = workbook.createCellStyle();

		// 默认换行
		// style.setWrapText(true);

		// fieldStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		fieldStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		fieldStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// 生成另一个字体
		XSSFFont fieldFont = workbook.createFont();
		fieldFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		// 把字体应用到当前的样式
		fieldStyle.setFont(fieldFont);
		XSSFCellStyle newStyle = workbook.createCellStyle();
		
		// 总页数
		int page = dataset.size() % pageSize == 0 ? dataset.size() / pageSize : dataset.size() / pageSize + 1;
		if (page == 0) {
			// 生成一个表格
			workbook.createSheet();
		}
		for (int j = 1; j < page + 1; j++) {
			// 生成一个表格
			XSSFSheet sheet = workbook.createSheet(sheetTitle + "_第（" + j + "）页");
			// 设置表格默认列宽度为15个字节
			sheet.setDefaultColumnWidth(15);

			// 产生表格标题行
			XSSFRow row = sheet.createRow(0);

			// 创建标题

			for (int i = 0; i < attrs.size(); i++) {
				ExcelCellAttr attr = attrs.get(i);
				XSSFCell cell = row.createCell(i);
				cell.setCellStyle(fieldHeaderStyle);
				XSSFRichTextString text = new XSSFRichTextString(attr.getHeaderText());
				cell.setCellValue(text);
			}
			List<T> list = null;
			if (page == 1) {
				list = dataset;
			} else {
				int restCount = dataset.size() - (j - 1) * pageSize;
				if (pageSize >= restCount) {
					list = dataset.subList((j - 1) * pageSize, (j - 1) * pageSize + restCount);
				} else {
					list = dataset.subList((j - 1) * pageSize, j * pageSize);
				}
			}

			int index = 0;
			for (T t : list) {
				index++;
				XSSFRow row1 = sheet.createRow(index);
				if (t instanceof Map) {
					for (int i = 0; i < attrs.size(); i++) {
						String fieldName = attrs.get(i).getField();
						Map map = (Map) t;
						XSSFCell cell = row1.createCell(i);
						
						newStyle.cloneStyleFrom(fieldStyle);
						if(attrs.get(i).getDataFormatIndex() > 0){
							newStyle.setDataFormat((short)attrs.get(i).getDataFormatIndex());
							sheet.setDefaultColumnStyle(i,newStyle);
						}else {
							newStyle.setDataFormat((short)0);
							sheet.setDefaultColumnStyle(i,newStyle);
						}
						cell.setCellStyle(newStyle);
						Object value = map.get(fieldName);

						// 判断值的类型后进行强制类型转换
						String textValue = null;
						if (value instanceof Integer) {
							int intValue = (Integer) value;
							cell.setCellValue(intValue);
						} else if (value instanceof Float) {
							float fValue = (Float) value;
							cell.setCellValue(fValue);
						} else if (value instanceof Double) {
							double dValue = (Double) value;
							cell.setCellValue(dValue);
						}

						else if (value instanceof Long) {
							long longValue = (Long) value;
							cell.setCellValue(longValue);
						} else if (value instanceof Date) {
							Date date = (Date) value;
							SimpleDateFormat sdf = new SimpleDateFormat(defaultDatePattern);
							textValue = sdf.format(date);
							cell.setCellValue(textValue);
						} else {
							// 其它数据类型都当作字符串简单处理
							textValue = value == null ? "" : value.toString();
							cell.setCellValue(textValue);
						}

					}

				} else {
					// 动态调用getXxx()
					Field[] fields = t.getClass().getDeclaredFields();
					for (int i = 0; i < attrs.size(); i++) {
						String fieldName = attrs.get(i).getField();

						List<Field> as = Arrays.asList(fields).stream().filter((a) -> {
							{
								return fieldName.equals(a.getName());
							}
						}).collect(Collectors.toList());

						if (as.isEmpty()) {
							continue;
						}
						XSSFCell cell = row1.createCell(i);

						newStyle.cloneStyleFrom(fieldStyle);
						if(attrs.get(i).getDataFormatIndex() > 0){
							newStyle.setDataFormat((short)attrs.get(i).getDataFormatIndex());
							sheet.setDefaultColumnStyle(i,newStyle);
						}else {
							newStyle.setDataFormat((short)0);
							sheet.setDefaultColumnStyle(i,newStyle);
						}
						
						cell.setCellStyle(newStyle);
						String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
						try {
							Class<?> tCls = t.getClass();
							Method getMethod = tCls.getMethod(getMethodName, new Class[]{});
							Object value = getMethod.invoke(t, new Object[]{});
							// 判断值的类型后进行强制类型转换
							String textValue = null;
							if (value instanceof Integer) {
								int intValue = (Integer) value;
								cell.setCellValue(intValue);
							} else if (value instanceof Float) {
								float fValue = (Float) value;
								cell.setCellValue(fValue);
							} else if (value instanceof Double) {
								double dValue = (Double) value;
								cell.setCellValue(dValue);
							}

							else if (value instanceof Long) {
								long longValue = (Long) value;
								cell.setCellValue(longValue);
							} else if (value instanceof Date) {
								Date date = (Date) value;
								SimpleDateFormat sdf = new SimpleDateFormat(defaultDatePattern);
								textValue = sdf.format(date);
								cell.setCellValue(textValue);
							} else {
								// 其它数据类型都当作字符串简单处理
								textValue = value == null ? "" : value.toString();
								cell.setCellValue(textValue);
							}

						} catch (SecurityException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (NoSuchMethodException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (IllegalArgumentException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (InvocationTargetException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} finally {

						}
					}
				}
			}
		}
		try {
			workbook.write(ouputStream);
			ouputStream.flush();
			ouputStream.close();
			workbook.close();
			log.info("exportExcel 导出完成");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			log.info("exportExcel 导出失败");
		}

	}

	/**
	 * 导出有固定列的数据表格且数据类型用Map<String, Object>来存储
	 * @author baize
	 * @param excelAttr:
     * @param response:
	 * @return
	 **/
	public static <T> void exportExcelMap(ExcelExportParms<T> excelAttr,HttpServletResponse response) throws IOException {
		// HttpServletResponse response = responseLocal.get();

		// 清空response
		response.reset();

		// 设置response的Header
		// 文件名
		String fileName = excelAttr.getFileName();
		fileName = new String(fileName.getBytes("UTF-8"), "ISO_8859_1");

		String sheetTitle = excelAttr.getSheetTitle();

		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

		// 每页大小
		int pageSize = excelAttr.getPageSize();
		// 日期格式
		String defaultDatePattern = excelAttr.getDefaultDatePattern();

		// 单元格属性
		List<ExcelCellAttr> attrs = excelAttr.getCellAttrs();

		// 导出数据
		List<T> dataset = excelAttr.getDatas();

		OutputStream ouputStream = response.getOutputStream();

		// 声明一个工作薄
		XSSFWorkbook workbook = new XSSFWorkbook();

		// 生成一个样式
		XSSFCellStyle fieldHeaderStyle = workbook.createCellStyle();
		// 设置这些样式
		fieldHeaderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		fieldHeaderStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// 生成一个字体
		XSSFFont fieldHeaderFont = workbook.createFont();
		fieldHeaderFont.setFontHeightInPoints((short) 12);
		fieldHeaderFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把字体应用到当前的样式
		fieldHeaderStyle.setFont(fieldHeaderFont);
		// 生成并设置另一个样式
		XSSFCellStyle fieldStyle = workbook.createCellStyle();

		fieldStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		fieldStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		fieldStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// 生成另一个字体
		XSSFFont fieldFont = workbook.createFont();
		fieldFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		// 把字体应用到当前的样式
		fieldStyle.setFont(fieldFont);
		XSSFCellStyle newStyle = workbook.createCellStyle();
		// 总页数
		int page = dataset.size() % pageSize == 0 ? dataset.size() / pageSize : dataset.size() / pageSize + 1;
		if (page == 0) {
			// 生成一个表格
			workbook.createSheet();
		}
		for (int j = 1; j < page + 1; j++) {
			// 生成一个表格
			XSSFSheet sheet = workbook.createSheet(sheetTitle + "_第（" + j + "）页");
			// 设置表格默认列宽度为15个字节
			sheet.setDefaultColumnWidth(15);

			// 产生表格标题行
			XSSFRow row = sheet.createRow(0);

			// 创建标题
			for (int i = 0; i < attrs.size(); i++) {
				ExcelCellAttr attr = attrs.get(i);
				XSSFCell cell = row.createCell(i);
				cell.setCellStyle(fieldHeaderStyle);
				XSSFRichTextString text = new XSSFRichTextString(attr.getHeaderText());
				cell.setCellValue(text);
			}
			List<T> list = null;
			if (page == 1) {
				list = dataset;
			} else {
				int restCount = dataset.size() - (j - 1) * pageSize;
				if (pageSize >= restCount) {
					list = dataset.subList((j - 1) * pageSize, (j - 1) * pageSize + restCount);
				} else {
					list = dataset.subList((j - 1) * pageSize, j * pageSize);
				}
			}

			int index = 0;
			for (T t : list) {

				index++;
				XSSFRow rowItem = sheet.createRow(index);
				if (t instanceof Map) { // 用Map存储数据
					@SuppressWarnings("unchecked")
					Map<String, Object> mapItem = (Map<String, Object>) t;
					for (int i = 0; i < attrs.size(); i++) {
						String key = attrs.get(i).getField();
						XSSFCell cell = rowItem.createCell(i);
						
						newStyle.cloneStyleFrom(fieldStyle);
						if(attrs.get(i).getDataFormatIndex() > 0){
							newStyle.setDataFormat((short)attrs.get(i).getDataFormatIndex());
							sheet.setDefaultColumnStyle(i,newStyle);
						}else {
							newStyle.setDataFormat((short)0);
							sheet.setDefaultColumnStyle(i,newStyle);
						}
						cell.setCellStyle(newStyle);

						Object value = mapItem.get(key);
						// 判断值的类型后进行强制类型转换
						String textValue = null;
						if (value instanceof Integer) {
							int intValue = (Integer) value;
							cell.setCellValue(intValue);
						} else if (value instanceof Float) {
							float fValue = (Float) value;
							cell.setCellValue(fValue);
						} else if (value instanceof Double) {
							double dValue = (Double) value;
							cell.setCellValue(dValue);
						} else if (value instanceof BigDecimal) {
							double dValue = ((BigDecimal) value).doubleValue();
							cell.setCellValue(dValue);
						} else if (value instanceof Long) {
							long longValue = (Long) value;
							cell.setCellValue(longValue);
						} else if (value instanceof Date) {
							Date date = (Date) value;
							SimpleDateFormat sdf = new SimpleDateFormat(defaultDatePattern);
							textValue = sdf.format(date);
							cell.setCellValue(textValue);
						} else {
							// 其它数据类型都当作字符串简单处理
							textValue = value == null ? "" : value.toString();
							cell.setCellValue(textValue);
						}
					}
				}
			}
		}
		try {
			workbook.write(ouputStream);
			ouputStream.flush();
			ouputStream.close();
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}