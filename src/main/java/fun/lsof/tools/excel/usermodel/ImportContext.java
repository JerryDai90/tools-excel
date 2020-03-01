package fun.lsof.tools.excel.usermodel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import fun.lsof.tools.excel.usermodel.row.Head;
import org.apache.log4j.Logger;


@SuppressWarnings({ "rawtypes", "unchecked" })
public class ImportContext
{
	static Logger logger = Logger.getLogger(ImportContext.class);
	
	public static ThreadLocal	importContext	= new ThreadLocal();
	Map							context;
	public static final String	IMPORT_INFO		= "import_info";
	public static final String	ERROR_LIST		= "error_list";		// �����б�
	public static final String	HEAD			= "head";			// Excel��ͷ

	public ImportContext(Map context) { this.context = null == context ? new HashMap() : context; }
	public ImportContext() { this.context = new HashMap(); }

	public void setErrorList(ArrayList<Map<String, String>> errorList) { put(ERROR_LIST, errorList); }
	public ArrayList<Map<String, String>> getErrorList() { return ((ArrayList<Map<String, String>>) get(ERROR_LIST)); }
	
	public void setHead(List<Head> head) { put(HEAD, head); }
	public List<Head> getHead() { return ((List<Head>) get(HEAD)); }
	
	public void setImportInfo(Object obj) { put(IMPORT_INFO, obj); }
	public Object getImportInfo() { return get(IMPORT_INFO); }

	public Object get(String key) { return context.get(key); }
	public void put(String key, Object value) { context.put(key, value); }

	public static ImportContext getContext() { return (ImportContext) importContext.get(); }
	public static void setContext(ImportContext context) { importContext.set(context); }

	public void setContextMap(Map contextMap) { getContext().context = contextMap; }
}
