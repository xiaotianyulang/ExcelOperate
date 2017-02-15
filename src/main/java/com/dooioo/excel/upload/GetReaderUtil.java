package com.dooioo.excel.upload;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.dooioo.excel.upload.impl.OptimizeRowReaderEntity;
import com.dooioo.excel.upload.impl.RowReaderEntity;


/**
 * @author Corrine Cao
 *
 */
public class GetReaderUtil {
	
	private static final Log log = LogFactory.getLog(GetReaderUtil.class);
	
	/** 获取导入结果
	 * @param request
	 * @param classInfo
	 * @param hasHeader
	 * @param keyIndex
	 * @param ignore
	 * @return
	 * @throws Exception
	 */
	public static OptimizeRowReaderEntity getRowReader(String fileName, InputStream is,Class classInfo,boolean hasHeader,List<Integer> keyIndex,boolean ignore) throws Exception{
		OptimizeRowReaderEntity rowReaderEntity = getUploadOptimizeRowReaderEntity(fileName,is);
		
		if(rowReaderEntity == null){
			throw new Exception("导入异常，请检查EXCEL文件内容是否错误");
		}
		
		rowReaderEntity.setIgnore(ignore);
		rowReaderEntity.setClassInfo(classInfo);//表单结果封装成salary对象集合
		rowReaderEntity.setHasHeader(hasHeader);//表单中有标题栏，跳过
		rowReaderEntity.setKeyIndex(keyIndex);//设置表单中从左往右列，一次对应对象的属性
		
		try {
			if (rowReaderEntity.isIs_2007()){//如果是xlsx结尾
			    log.info(">>>>>>>>2007");
			    ExcelReaderUtil.readExcel_2007(rowReaderEntity, rowReaderEntity.getIs());
					
			}else{//如果是xls结尾
				log.info(">>>>>>>>2003");
				ExcelReaderUtil.readExcel_2003(rowReaderEntity, rowReaderEntity.getIs());
					
			}
		} catch (Exception e) {
			log.error(e);
			throw e;
		}

		return rowReaderEntity;
	}
	
	/** 将导入的excel表封装成对象集合，true有空值的行将跳过
	 * @param request
	 * @param classInfo
	 * @param hasHeader
	 * @param keyIndex
	 * @return
	 * @throws Exception
	 */
	public static List getImportList(HttpServletRequest request,Class classInfo,boolean hasHeader,List<Integer> keyIndex) throws Exception{
		return getImportList(request,classInfo,hasHeader,keyIndex,true);
	}
	
	/** 导入封装，fasle，保留有空值的行
	 * @param request
	 * @param classInfo 返回的泛型类型
	 * @param hasHeader 表单是否有头部，有为true，读取时跳过第一行
	 * @param keyIndex 表单中的单元格从左往右对应实体类的成员变量的index
	 * @param ignore 是否忽略空行
	 * @return 返回读取表单集合
	 * @throws Exception
	 */
	public static List getImportList(HttpServletRequest request,Class classInfo,boolean hasHeader,List<Integer> keyIndex,boolean ignore) throws Exception{
		RowReaderEntity rowReaderEntity = getUploadRowReaderEntity(request);
		
		if(rowReaderEntity == null){
			throw new Exception("导入异常，请检查EXCEL文件内容是否错误");
		}
		
		rowReaderEntity.setIgnore(ignore);
		rowReaderEntity.setClassInfo(classInfo);//表单结果封装成salary对象集合
		rowReaderEntity.setHasHeader(hasHeader);//表单中有标题栏，跳过
		rowReaderEntity.setKeyIndex(keyIndex);//设置表单中从左往右列，一次对应对象的属性
		
		try {
			if (rowReaderEntity.isIs_2007()){//如果是xlsx结尾
			    log.info(">>>>>>>>2007");
			    ExcelReaderUtil.readExcel_2007(rowReaderEntity, rowReaderEntity.getIs());
					
			}else{//如果是xls结尾
				log.info(">>>>>>>>2003");
				ExcelReaderUtil.readExcel_2003(rowReaderEntity, rowReaderEntity.getIs());
					
			}
		} catch (Exception e) {
			log.error(e);
			throw e;
		}

		return rowReaderEntity.getResultList();
	}
	
	/** 根据request获取上传文件，封装成reader
	 * @param request
	 * @return
	 * @throws FileUploadException
	 * @throws IOException
	 */
	public static RowReaderEntity getUploadRowReaderEntity(HttpServletRequest request) throws FileUploadException, IOException{
		ServletFileUpload sfu = new ServletFileUpload(new DiskFileItemFactory());
		@SuppressWarnings("unchecked")
		List<FileItem> itemlist = sfu.parseRequest(request);
		for(FileItem item : itemlist){
			if (item.isFormField())
	            continue;
			InputStream is = item.getInputStream();
			String fileextend = getSuffix(item.getName());
			
        	if(fileextend.equalsIgnoreCase(".xls")){//如果是03版的，返回workbook，保留之前的逻辑
        		RowReaderEntity rowReaderEntity = new RowReaderEntity();
        		rowReaderEntity.setIs_2007(false);
        		rowReaderEntity.setIs(is);
        		return rowReaderEntity;
        		
        	}else if(fileextend.equalsIgnoreCase(".xlsx")){//如果是07版的，可能是大量的数据，则采用逐行处理，返回输入流。
        		RowReaderEntity rowReaderEntity = new RowReaderEntity();
        		rowReaderEntity.setIs_2007(true);
        		rowReaderEntity.setIs(is);
        		return rowReaderEntity;
        	}
        		
		}
		return null;
	}
	
	/** 获取文件名的后缀
	 * @param filename
	 * @return
	 */
	private static String getSuffix(String filename){
		if(filename.lastIndexOf(".")==-1)
			return filename;
		return filename.substring(filename.lastIndexOf("."));
	}
	
	/** 获取rowReader
	 * @param request
	 * @return
	 * @throws FileUploadException
	 * @throws IOException
	 */
	private static OptimizeRowReaderEntity getUploadOptimizeRowReaderEntity(
			String fileName, InputStream is) throws FileUploadException,
			IOException {

		if (fileName.endsWith(".xls") || fileName.endsWith(".XLS")) {// 如果是03版的，返回workbook，保留之前的逻辑
			OptimizeRowReaderEntity rowReaderEntity = new OptimizeRowReaderEntity();
			rowReaderEntity.setIs_2007(false);
			rowReaderEntity.setIs(is);
			return rowReaderEntity;

		} else if (fileName.endsWith(".xlsx") || fileName.endsWith(".XLSX")) {// 如果是07版的，可能是大量的数据，则采用逐行处理，返回输入流。
			OptimizeRowReaderEntity rowReaderEntity = new OptimizeRowReaderEntity();
			rowReaderEntity.setIs_2007(true);
			rowReaderEntity.setIs(is);
			return rowReaderEntity;
		}

		return null;
	}

}
