package fun.lsof.common.excel.eventusermodel.xlsx.db;

import com.wordty.common.assist.utils.DBUtil;
import fun.lsof.common.excel.eventusermodel.xlsx.exception.XlsContentImproperException;
import fun.lsof.common.excel.eventusermodel.xlsx.db.vo.DBFieldType;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;


public class TestXSSF2DB
{
	public static void main(String[] args) throws Exception
	{
		System.out.println(System.currentTimeMillis());

		String tgTable = "td";
		String schema = "IPS_IMPORT";

		XSSF2DB xssf2db = new  XSSF2DB("D:\\test.xlsx", "1", tgTable);
		xssf2db.setRowColumnName("row_num_");
//		xssf2db.setInsertRowNum(true);

//		xssf2db.setConnection(DBUtil.getConnection(DBUtil.Type.jdbc, 
//				"oracle.jdbc.driver.OracleDriver", 
//				"jdbc:oracle:thin:@isclx001.ap.mars:1521:DB1378T", 
//				"wpp", 
//				"powerpos" ));

		xssf2db.setConnection(DBUtil.getConnection(DBUtil.Type.jdbc,
				"oracle.jdbc.driver.OracleDriver",
				"jdbc:oracle:thin:@jerry-svr:1521:orcl",
				"IPS_IMPORT",
				"IPS_IMPORT"));

		List<String> stencilTitle = new ArrayList<String>();
		stencilTitle.add("TD");
		stencilTitle.add("TS");

		xssf2db.setStencilTitle(stencilTitle);


		List<String> stencilDBColumn = new  ArrayList<String>();
		stencilDBColumn.add("TD");
		stencilDBColumn.add("TS");

		List<DBFieldType> dbs = new ArrayList<DBFieldType>();
		List<String> dbTypeColumnList = new ArrayList<String>();
		Map<String, DBFieldType> dbMap = xssf2db.getFieldType4DB(tgTable,
				DBUtil.getConnection(DBUtil.Type.jdbc,
						"oracle.jdbc.driver.OracleDriver",
						"jdbc:oracle:thin:@jerry-svr:1521:orcl",
						"IPS_IMPORT",
						"IPS_IMPORT"),
				schema);
		for( String str : stencilDBColumn ){
			DBFieldType dbType = dbMap.get(str);
			if( null != dbType ){
				dbs.add( dbType );
				dbTypeColumnList.add(dbType.getFieldName());
			}
		}

		//模板的字段是否和目标表的字段对应，不对应就抛错
		if( !dbTypeColumnList.containsAll(stencilDBColumn) ){
			stencilDBColumn.removeAll(dbTypeColumnList);
			throw new RuntimeException("保存的字段信息和数据库表对应不上，"+stencilDBColumn.toString()+"在目标表中没有找到！");
		}

		xssf2db.setDbFieldTypes(dbs);

		try{
			xssf2db.parse();
		}catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("error:"+xssf2db.getImproperException().errors.size());
		System.out.println("sccess:"+xssf2db.getWorkbookMap().size());

		for( XlsContentImproperException.Entry e : xssf2db.getImproperException().errors ){
			System.out.println(e.getHmData().toString());
			System.out.println(e.getOtherError().toString());
		}

//		System.out.println(xssf2db.workbookMap.toString());

		System.exit(1);
	}
}
