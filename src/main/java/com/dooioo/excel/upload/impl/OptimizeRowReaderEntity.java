package com.dooioo.excel.upload.impl;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.dooioo.excel.upload.IRowReader;
import com.dooioo.export.annotation.Column;

/**
 * @author Corrine Cao
 *
 */
public class OptimizeRowReaderEntity implements IRowReader {
	
    private boolean hasHeader = true; //是否有标题栏，有就跳过第一行，从第二行开始
	
	private boolean is_2007 = false; //是否是2007 true是.xlsx, false是.xls
	
	private Class classInfo = null; //需要返回的集合元素类型
	
	private List<Object> resultList = null; //需要返回的结果集
	
    private List<Integer> keyIndex = null; //key值，标示每一列的数据，以逗号按表中从左往右顺序隔开
	
	private InputStream is = null; //用来存放输入流
	
	private Map<Integer,Field> indexToField = null;//存放返回元素类的column对应的set方法 
	
	private boolean ignore = true;//跳过空行
	
	private final Log log = LogFactory.getLog(this.getClass());

	private List<String> errorList = new ArrayList<String>();//记录格式错误

	public OptimizeRowReaderEntity() {
		
	}

	
	public List<String> getErrorList() {
		return errorList;
	}

	public void setErrorList(List<String> errorList) {
		this.errorList = errorList;
	}

	public boolean isIgnore() {
		return ignore;
	}


	public void setIgnore(boolean ignore) {
		this.ignore = ignore;
	}


	public boolean isHasHeader() {
		return hasHeader;
	}

	public void setHasHeader(boolean hasHeader) {
		this.hasHeader = hasHeader;
	}

	public boolean isIs_2007() {
		return is_2007;
	}

	public void setIs_2007(boolean is_2007) {
		this.is_2007 = is_2007;
	}

	public List<Integer> getKeyIndex() {
		return keyIndex;
	}

	public void setKeyIndex(List<Integer> keyIndex) {
		this.keyIndex = keyIndex;
	}

	public InputStream getIs() {
		return is;
	}

	public void setIs(InputStream is) {
		this.is = is;
	}

	public Class getClassInfo() {
		return classInfo;
	}

	public void setClassInfo(Class classInfo) {
		this.classInfo = classInfo;
	}

	public List getResultList() {
		return resultList;
	}

	public void setResultList(List resultList) {
		this.resultList = resultList;
	}

	
	
	@Override
	public void getRows(int sheetIndex, int curRow, List<String> rowlist)
			throws Exception {
		boolean flag = false;
		try {
			if (hasHeader) {// 如果有标题栏，就跳过第一行
				if (curRow == 0) { // 读取第一行，则跳过，无需处理
					return;
				}
			}

			if (resultList == null) {
				resultList = new ArrayList();
			}

			if (keyIndex == null || keyIndex.size() == 0) {
				flag = true;
				throw new Exception("keyIndex没有设置");
			}

			if (classInfo == null) {
				flag = true;
				throw new Exception("classInfo没有设置");
			}

			if(rowlist.size() < keyIndex.size()){
				if(ignore){
					log.info(">>>>>>>>>>第 "+(curRow+1)+" 行有空值，跳过");
					return;	
				}
				
			}
			
			if(indexToField == null){//保证一个对象只执行一次
				//获取这个类所有的元素
				Field[] fields = classInfo.getDeclaredFields();
				//将这个类column对应的field缓存起来
				indexToField = new HashMap<Integer,Field>();
				for (int j = 0; j < fields.length; j++) {
					Field field = fields[j];
					if (field.isAnnotationPresent(Column.class)) {
						Column column = (Column) field.getAnnotation(Column.class);
						int index = column.index();
						field.getType();
						indexToField.put(index, field);

					}
				}
			}
			

			Field paramfield = null;
			String methodName = null;
			Class paramType = null;

			String cellValue = null;
			Object value = null;
			
			//根据传入的index，依次设置对象的属性
			int cellNum = rowlist.size();
			if(cellNum == 0){
				log.info(">>>>>>>>>>第 "+(curRow+1)+" 行都是空值，跳过……");
				return;
			}
			
			//新建一个对象，一行即一个对象
			Object model = null;
			
			for (int i = 0; i < keyIndex.size(); i++) {
				//获取表单中的值
				if(i > cellNum - 1){
					cellValue = "";
				}else{
					cellValue = rowlist.get(i);
				}
				
				//有空值，此行不处理
				if("".equals(cellValue.trim())){
					
					if(ignore){
						log.info(">>>>>>>>>>第 "+(curRow+1)+" 行有空值，跳过……");
						return;
					}else{
						log.info(">>>>>>>>>>第 "+(curRow+1)+" 行第"+ (i+1) +"列有空值，跳过单元格");
						continue;
					}
					
				}
				
				if(model == null){//这样做可以保证当此行出现非空值时，才新建对象
					model = classInfo.newInstance();
				}
				
				//获取column的index
				Integer perIndex = keyIndex.get(i);
				//获取对应的成员变量名
				paramfield = indexToField.get(perIndex);// 获取对应的参数名称
				// 根据成员变量名获取对应的set方法名
				methodName = new StringBuilder("set")
						.append(paramfield.getName().substring(0, 1)
								.toUpperCase())
						.append(paramfield.getName().substring(1)).toString();
				// 获取成员变量的参数类型
				paramType = paramfield.getType();

				//获取set方法对象
				Method method = classInfo.getMethod(methodName, paramType);
				if (method == null) {
					flag = true;
					throw new Exception("没有找到方法:" + methodName);
				}

				
			    try {
					value = getValue(paramType,cellValue);
					
					//调用此set方法，向model中设置值
					method.invoke(model, value);
				} catch (Exception e) {
					errorList.add((curRow + 1)+","+(i + 1));//将格式不正确的行号和列号存起来
					e.printStackTrace();
				}
			}
			
			//将model对象添加到集合中
			if(model != null){
				//add column index
				Method method = classInfo.getMethod("setColumnIndex", Integer.class);
				if (method == null) {
					log.info(">>>>没有设置行号，没有对应的set方法");
				}
				method.invoke(model, (curRow + 1));
				//add
				
				resultList.add(model);
			}
			
							
		} catch (Exception e) {
			if(flag){
				throw e;
			}else{
				log.error(e);
				throw new Exception(e);
			}
		}	
	}

	
	/** 根据类型转化数据获得value
	 * @param paramType
	 * @param cellValue
	 * @return
	 * @throws Exception
	 */
	private Object getValue(Class paramType,String cellValue) throws Exception{
		if (paramType.equals(byte.class)
				|| paramType.equals(Byte.class)) {
			return Byte.valueOf(cellValue);

		} else if (paramType.equals(short.class)
				|| paramType.equals(Short.class)) {
			return Short.valueOf(cellValue);

		} else if (paramType.equals(int.class)
				|| paramType.equals(Integer.class)) {
			return Integer.valueOf(cellValue);

		} else if (paramType.equals(long.class)
				|| paramType.equals(Long.class)) {
			return Long.valueOf(cellValue);

		} else if (paramType.equals(float.class)
				|| paramType.equals(Float.class)) {
			return new BigDecimal(cellValue).setScale(2, BigDecimal.ROUND_HALF_UP).floatValue();

		} else if (paramType.equals(double.class)
				|| paramType.equals(Double.class)) {
			return new BigDecimal(cellValue).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();

		} else if (paramType.equals(String.class)) {
			return cellValue;

		} else {
			throw new Exception("现在还不支持此种类型的数值直接读取转换：" + paramType);
		}
	}
}
