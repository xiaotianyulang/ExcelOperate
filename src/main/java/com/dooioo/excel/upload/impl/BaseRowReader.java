package com.dooioo.excel.upload.impl;

import java.util.List;

import com.dooioo.excel.upload.IRowReader;

/**
 * @author Corrine Cao
 *
 */
public class BaseRowReader implements IRowReader{
	
	private boolean is_2007 = false; //是否是2007true是.xlsx, false是.xls
	
	public boolean isIs_2007() {
		return is_2007;
	}

	public void setIs_2007(boolean is_2007) {
		this.is_2007 = is_2007;
	}

	@Override
	public void getRows(int sheetIndex, int curRow, List<String> rowlist)
			throws Exception {
		System.out.println(">>>>>current row:" + (curRow + 1) + ">>>"+ rowlist);
		
	}

}
