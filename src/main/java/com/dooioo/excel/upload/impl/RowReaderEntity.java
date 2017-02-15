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
public class RowReaderEntity implements IRowReader {
	
    private boolean hasHeader = true; //是否有标题栏，有就跳过第一行，从第二行开始
	
	private boolean is_2007 = false; //是否是2007 true是.xlsx, false是.xls
	
	private Class classInfo = null; //需要返回的集合元素类型
	
	private List<Object> resultList = null; //需要返回的结果集
	
    private List<Integer> keyIndex = null; //key值，标示每一列的数据，以逗号按表中从左往右顺序隔开
	
	private InputStream is = null; //用来存放输入流
	
	private boolean ignore = true;//跳过空行
	
	private final Log log = LogFactory.getLog(this.getClass());

	private List<Method> methods = null;//存放需要用到的set方法
	
	private List<Class> paramTypes = null; //存放传入参数的类型

	public RowReaderEntity() {
		
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
			
			if(methods == null){//保证一个对象只执行一次
				Column column = null;
				String methodName = null;
				
				//获取这个类所有的元素
				Field[] fields = classInfo.getDeclaredFields();
				//将这个类column对应的field缓存起来
				Map<Integer,Field> indexToField = new HashMap<Integer,Field>();
				for (int j = 0; j < fields.length; j++) {
					if (fields[j].isAnnotationPresent(Column.class)) {
						column = (Column) fields[j].getAnnotation(Column.class);
						indexToField.put(column.index(), fields[j]);
					}
				}
				
				//获取keyIndex对应的方法
				methods = new ArrayList<Method>();
				paramTypes = new ArrayList<Class>();
				Field field = null;
				
				for(Integer perIndex: keyIndex){
					//获取对应的成员变量名
					field = indexToField.get(perIndex);// 获取对应的参数名称
					// 根据成员变量名获取对应的set方法名
					methodName = new StringBuilder("set")
							.append(field.getName().substring(0, 1)
									.toUpperCase())
							.append(field.getName().substring(1)).toString();
					// 获取成员变量的参数类型

					//获取set方法对象
					Method method = classInfo.getMethod(methodName, field.getType());
					if (method == null) {
						flag = true;
						throw new Exception("没有找到方法:" + methodName);
					}
					methods.add(method);//将方法缓存起来
					paramTypes.add(field.getType()); //将参数类型缓存起来
				
				}
				
				indexToField = null; //方便回收
				fields = null;
				
				log.info(">>>当前用到的方法依次是："+ methods );
				log.info(">>>当前方法参数依次是："+ paramTypes );
				
			}

			String cellValue = null;
			Object value = null;
			
			//根据传入的index，依次设置对象的属性
			if(rowlist.size() == 0){
				log.info(">>>>>>>>>>第 "+(curRow+1)+" 行都是空值，跳过……");
				return;
			}
			
			//新建一个对象，一行即一个对象
			Object model = null;
			
			for (int i = 0; i < methods.size(); i++) {
				//获取表单中的值
				if(i > rowlist.size() - 1){
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
				
				if(model == null){//一行数据只新建一个对象
					model = classInfo.newInstance();
				}
				
			    value = getValue(paramTypes.get(i),cellValue.trim());
				
				//调用此set方法，向model中设置值
			    methods.get(i).invoke(model, value);
			}
			
			//将model对象添加到集合中
			if(model != null){
				resultList.add(model);
			}
			
							
		} catch (Exception e) {
			if(flag){
				throw e;
			}else{
				log.error(e);
				throw new Exception("第" + (curRow + 1) + "行格式有问题，请检查后重新上传");
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
