package com.dooioo.excel.export;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.dooioo.export.excel.POIExcelExport;

/** 
 * @author Corrine Cao
 *
 */
public class CustomPOIExcelExport extends POIExcelExport {
	
	private final Log log = LogFactory.getLog(this.getClass());


	public CustomPOIExcelExport() {

	}

	@Override
	public void export(OutputStream out, String columns[], String methods[],
			List links, List dataSet, String fileName) {
		SXSSFWorkbook workbook = null;
		try {
			workbook = new SXSSFWorkbook(100);
			Sheet sheet = workbook.createSheet(fileName);
			setSheet(sheet, columns, methods, dataSet, links);
			workbook.write(out);
			
		} catch (Exception e) {
			log.error(e);
			
		}finally{
			if(out != null){
				try {
					out.flush();
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			
			if(workbook != null){
				log.info(">>>>>>>>>>>>>>处理在磁盘上的临时文件备份");
				workbook.dispose();
			}
			
		}

	}


	@Override
	public void export(OutputStream out, String columns[], String methods[],
			List dataSet, String fileName) {

		SXSSFWorkbook workbook = null;
		try {
			workbook = new SXSSFWorkbook(100);
			Sheet sheet = workbook.createSheet(fileName);
			setSheet(sheet, columns, methods, dataSet);
			workbook.write(out);
		} catch (Exception e) {
			log.error("导出捕获异常",e);
		}finally{
			if(out != null){
				try {
					out.flush();
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			
			if(workbook != null){
				log.info(">>>>>>>>>>>>>>处理在磁盘上的临时文件备份");
				workbook.dispose();
			}
		}
	}

	/** 设置sheet的内容，包含links
	 * @param sheet
	 * @param columns
	 * @param methods
	 * @param dataSet
	 * @param links
	 * @throws IllegalArgumentException
	 * @throws SecurityException
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 */
	private void setSheet(Sheet sheet, String columns[], String methods[],
			List dataSet, List links) throws IllegalArgumentException,
			SecurityException, IllegalAccessException,
			InvocationTargetException, NoSuchMethodException {
		
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
		
		Row header = sheet.createRow(0);
		Class model = dataSet.get(0).getClass();
	    Cell cell = null;
		int len = columns.length;
		for (int j = 0; j < len; j++) {
			cell = header.createCell(j);
			cell.setCellValue(columns[j]);
		}

		Object obj = null;
		Row row = null;
		Object value = null;
		String textValue = null;
		Cell column = null;
		String link = null;
		String linkValue = null;
		int size = dataSet.size();
		for (int i = 0; i < size; i++) {
			row = sheet.createRow(i + 1);
			obj = dataSet.get(i);
			for (int j = 0; j < len; j++) {
				value = model.getMethod(methods[j], new Class[0]).invoke(obj,
						new Object[0]);
				if (value == null)
					textValue = "";
				else if (value instanceof Date)
					textValue = dateFormat.format(value);
				else
					textValue = value.toString();
				link = (String) links.get(j);
				column = row.createCell(j);
				column.setCellValue(textValue);
				if (!link.isEmpty()) {
					linkValue = (new StringBuilder())
							.append(model.getMethod((String) links.get(j),
									new Class[0]).invoke(obj, new Object[0]))
							.toString();
					HSSFHyperlink hyperlink = new HSSFHyperlink(1);
					hyperlink.setAddress(linkValue);
					column.setHyperlink(hyperlink);
				}
			}

		}

	}

	/** 设置sheet表的输出
	 * @param sheet
	 * @param columns
	 * @param methods
	 * @param dataSet
	 * @throws IllegalArgumentException
	 * @throws SecurityException
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 */
	private void setSheet(Sheet sheet, String columns[], String methods[],
			List dataSet) throws IllegalArgumentException, SecurityException,
			IllegalAccessException, InvocationTargetException,
			NoSuchMethodException {
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
		
		Row header = sheet.createRow(0);
		Class model = dataSet.get(0).getClass();
		Cell cell = null;
		int len = columns.length;
		for (int j = 0; j < len; j++) {
			cell = header.createCell(j);
			cell.setCellValue(columns[j]);
		}

		Object obj = null;
		Row row = null;
		Object value = null;
		String textValue = null;
		Cell column = null;
		int size = dataSet.size();
		for (int i = 0; i < size; i++) {
			row = sheet.createRow(i + 1);
			obj = dataSet.get(i);
			for (int j = 0; j < len; j++) {
				value = model.getMethod(methods[j], new Class[0]).invoke(obj,
						new Object[0]);
				if (value == null){
					textValue = "";
				}else if (value instanceof Date){
					try {
						textValue = dateFormat.format(value);
					}catch (Exception e){
						log.error("格式化出错",e);
					}
				}else{
					textValue = value.toString();
				}
				column = row.createCell(j);
				column.setCellValue(textValue);
			}

		}

	}

}
