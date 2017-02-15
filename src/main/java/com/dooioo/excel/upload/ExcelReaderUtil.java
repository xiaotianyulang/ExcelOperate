package com.dooioo.excel.upload;

import java.io.InputStream;

import org.xml.sax.SAXException;

public class ExcelReaderUtil {
	
	/**
	 * 读取Excel文件
	 * @param reader
	 * @param is
	 * @throws Exception
	 */
	public static void readExcel_2007(IRowReader reader, InputStream is) throws Exception {
		Excel2007Reader excel07 = new Excel2007Reader();
		excel07.setRowReader(reader);
		excel07.process(is);
	}
	
	/**读取Excel文件
	 * @param reader
	 * @param fileName
	 * @throws Exception
	 */
	public static void readExcel_2003(IRowReader reader, InputStream is) throws Exception {
		Excel2003Reader excel03 = new Excel2003Reader();
		excel03.setRowReader(reader);
		excel03.process(is);
		String error = excel03.getErrInfo();
		if(error.length()>0){
			throw new Exception(error);
		}
	}

}
