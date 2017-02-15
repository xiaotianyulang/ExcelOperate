package com.dooioo.excel.upload;

import java.util.List;

public interface IRowReader {
	
	/**业务逻辑实现方法
	 * @param sheetIndex
	 * @param curRow
	 * @param rowlist
	 */
	public void getRows(int sheetIndex,int curRow, List<String> rowlist) throws Exception;
	
}
