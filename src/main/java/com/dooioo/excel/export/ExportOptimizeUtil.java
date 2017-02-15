package com.dooioo.excel.export;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.dooioo.export.Exportable;
import com.dooioo.export.Mime;
import com.dooioo.export.annotation.Column;
import com.dooioo.export.annotation.Export;
import com.dooioo.export.excel.AbstractExcelExport;
import com.dooioo.export.utils.StringUtils;
import com.dooioo.export.web.WebExport;

/** 海量导出数据(直接调用)
 * @author Corrine Cao 
 *
 */
public class ExportOptimizeUtil {

	//private static final Log log = LogFactory.getLog(ExportOptimizeUtil.class);

	/*public static void main(String[] args){
		HttpServletResponse response = null;
		List untreatedList = new ArrayList();
		WebExport.exportInclude(response, untreatedList, new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, "-工资-扣款明细");
	}*/

	/** 根据配置的属性导出不同的数据
	 * @param response
	 * @param dataSet
	 * @param includeIndices
	 * @param fileName
	 */
	public static void exportInclude(HttpServletResponse response,
			List dataSet, int includeIndices[], String fileName) {
		exportInclude(response, dataSet, includeIndices, fileName, false);
	}

	/** 根据配置的属性导出不同的数据
	 * @param response
	 * @param dataSet
	 * @param includeIndices
	 * @param fileName
	 * @param hasLink
	 */
	public static void exportInclude(HttpServletResponse response,
			List dataSet, int includeIndices[], String fileName, boolean hasLink) {

		if (isNullOrEmpty(dataSet)) {
			try {
				response.setContentType("text/html;charset=UTF-8");
				response.getWriter()
				.print("<script>alert('\u6CA1\u6709\u6570\u636E\u53EF\u4EE5\u5BFC\u51FA,\u70B9\u786E\u5B9A\u540E\u7CFB\u7EDF\u5C06\u81EA\u52A8\u8FD4\u56DE...');history.go(-1);</script>");
			} catch (IOException e) {
				e.printStackTrace();
			}
			return;
		}
		Class model = dataSet.get(0).getClass();
		Map params = parseForParams(model, 2, includeIndices);
		String ext = (new StringBuilder()).append(params.get("ext")).toString();
		String mime = (new StringBuilder()).append(params.get("mime"))
				.toString();
		String labels[] = (String[]) params.get("labels");
		String methods[] = (String[]) params.get("methods");
		List links = (List) params.get("links");
		Exportable processor = (Exportable) params.get("processor");
		if (fileName == null || fileName.trim().isEmpty())
			fileName = model.getSimpleName();
		response.reset();
		response.setContentType((new StringBuilder(String.valueOf(mime)))
				.append(";charset=UTF-8").toString());
		try {
			response.setHeader(
					"Content-disposition",
					(new StringBuilder("attachment; filename="))
					.append(StringUtils.toString(fileName, null,
							"ISO8859-1")).append(ext).toString());
			if (hasLink)
				processor.export(response.getOutputStream(), labels, methods,
						links, dataSet, fileName);
			else
				processor.export(response.getOutputStream(), labels, methods,
						dataSet, fileName);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/** 判断是否是空值
	 * @param data
	 * @return
	 */
	private static boolean isNullOrEmpty(List data){
		return data == null || data.isEmpty();
	}

	/** 获取参数
	 * @param model
	 * @param type
	 * @param indices
	 * @return
	 */
	private static Map parseForParams(Class model, int type, int indices[]){
		Export export = (Export)model.getAnnotation(Export.class);
		if(export == null)
			throw new IllegalArgumentException((new StringBuilder("没有为")).append(model.getCanonicalName()).append("配置@export标注。").toString());
		Field fields[] = model.getDeclaredFields();
		String label = null;
		String link = null;
		String fieldName = null;
		Column column = null;
		Map exportLablels = new TreeMap();
		Map exportMethods = new TreeMap();
		Map exportLinks = new HashMap();
		int count = 0;
		Field afield[];
		int k = (afield = fields).length;
		for(int j = 0; j < k; j++)
		{
			Field field = afield[j];
			if(field.isAnnotationPresent(Column.class))
			{
				column = (Column)field.getAnnotation(Column.class);
				fieldName = field.getName();
				int index = column.index();
				label = column.label();
				link = column.link();
				if(index == -1)
					throw new IllegalArgumentException((new StringBuilder("请填写字段")).append(fieldName).append("的index标注索引").toString());
				if(index > count)
					count = index;
				if(label == null || label.isEmpty())
					label = fieldName;
				if(link.isEmpty())
					exportLinks.put(Integer.valueOf(index), "");
				else
					exportLinks.put(Integer.valueOf(index), (new StringBuilder("get")).append(link.substring(0, 1).toUpperCase()).append(link.substring(1)).toString());
				exportLablels.put(Integer.valueOf(index), label);
				exportMethods.put(Integer.valueOf(index), (new StringBuilder("get")).append(fieldName.substring(0, 1).toUpperCase()).append(fieldName.substring(1)).toString());
			}
		}

		if(exportLablels.size() == 0)
			throw new IllegalArgumentException("没有可以导出的字段");
		count++;
		List labels = new ArrayList();
		List methods = new ArrayList();
		List links = new ArrayList();
		if(type == 1)
		{
			if(indices != null && indices.length > 0)
			{
				count -= indices.length;
				if(count <= 0)
					throw new IllegalArgumentException((new StringBuilder("excludeIndices的length不能大于等于总共要导出的列(")).append(count).append(")").toString());
				int ai[];
				int j1 = (ai = indices).length;
				for(int l = 0; l < j1; l++)
				{
					Integer i = Integer.valueOf(ai[l]);
					exportLablels.remove(i);
					exportMethods.remove(i);
					exportLinks.remove(i);
				}

			}
			Integer index;
			for(Iterator iterator = exportLablels.keySet().iterator(); iterator.hasNext(); links.add((String)exportLinks.get(index)))
			{
				index = (Integer)iterator.next();
				labels.add((String)exportLablels.get(index));
				methods.add((String)exportMethods.get(index));
			}

		} else
			if(type == 2)
			{
				if(indices != null && indices.length > count)
					throw new IllegalArgumentException((new StringBuilder("包含的列数量不能大于等于总共要导出的列(")).append(count).append(")").toString());
				int ai1[];
				int k1 = (ai1 = indices).length;
				for(int i1 = 0; i1 < k1; i1++)
				{
					Integer index = Integer.valueOf(ai1[i1]);
					labels.add((String)exportLablels.get(index));
					methods.add((String)exportMethods.get(index));
					links.add((String)exportLinks.get(index));
				}

			}
		if(labels == null || labels.size() == 0)
			throw new IllegalArgumentException("没有可以导出的列");
		String mime = "";
		String ext = "";
		Exportable processor = null;

		mime = Mime.Excel.getContent();
		ext = ".xlsx";
		processor = new CustomPOIExcelExport();

		/*switch(export.value().ordinal())
	        {
	        case 2: // '\002'
	            mime = Mime.Excel.getContent();
	            ext = ".xlsx";
	           // processor = AbstractExcelExport.getInstance();
	            processor = new CustomPOIExcelExport();
	            break;

	        case 1: // '\001'
	            mime = Mime.Word.getContent();
	            ext = Mime.Excel.getExtension();
	            // fall through

	        default:
	            mime = Mime.Txt.getContent();
	            ext = Mime.Txt.getExtension();
	            break;
	        }*/
		Map params = new HashMap();
		String ls[] = new String[0];
		String ms[] = new String[0];
		params.put("mime", mime);
		params.put("ext", ext);
		params.put("methods", ((Object) (methods.toArray(ms))));
		params.put("labels", ((Object) (labels.toArray(ls))));
		params.put("links", links);
		params.put("processor", processor);
		return params;
	}

}
