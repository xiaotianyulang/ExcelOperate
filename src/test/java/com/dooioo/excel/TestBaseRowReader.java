package com.dooioo.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.dooioo.excel.upload.impl.BaseRowReader;
import com.dooioo.excel.upload.ExcelReaderUtil;

public class TestBaseRowReader {

    public static void main(String[] args){
		
	    //String path = "E:\\wKgA8FZTJ_2AZbyxAAAeANSPI5I752.xlsx";
	    
	    String path = "E:\\wKgA8FdorayAFXPTAABJRTE4Fis37.xlsx";
	   // String path = "E:\\wKgA8FZTJ_2AZbyxAAAeANSPI5I752.xls";
	
		InputStream is = null;
		//获取输入流
		try {
			is = new FileInputStream(path);
		} catch (FileNotFoundException e ) {
			e.printStackTrace();
		}
		
		BaseRowReader rowReader = new BaseRowReader();
		
		//rowReader.setHasHeader(false);//表示导入的表格无标题栏，也就是从第一行数据开始就是
		//rowReader.setIs_2007(true);//默认是false,2003, .xls，否则是true,2007 .xlsx
		
		if(path.toUpperCase().endsWith(".XLSX")){
			rowReader.setIs_2007(true);
		}
		
		try {
			if(rowReader.isIs_2007()){
				ExcelReaderUtil.readExcel_2007(rowReader, is);
			}else{
				ExcelReaderUtil.readExcel_2003(rowReader, is);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}
