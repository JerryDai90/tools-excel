package fun.lsof.tools.excel.usermodel;

import fun.lsof.tools.excel.usermodel.row.Head;
import fun.lsof.tools.excel.usermodel.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileWriter;
import java.util.List;
import java.util.Map;


/*
 *
 */
public class TestAnalysisUtil {


    String xx = "";

    
    public static void main(String[] args) throws Exception
    {

//		initContext();
//
//		//读取规则
//		List<Rule> rules = new ArrayList<AnalysisUtil.Rule>();
//		rules.add(new Rule("A", "A", "number"));
//		rules.add(new Rule("B", "B", "number"));
//		rules.add(new Rule("C", "C", "number"));
//		rules.add(new Rule("D", "D", "number"));
//		rules.add(new Rule("E", "E", "number"));
//		ImportContext.getContext().setImportInfo(rules);
//
//
        String path = "/Users/jerry/Mix/共享/1.xls";
//		Workbook wb = ExcelUtil.getWorkbook(path);
//		List<Head> head = getSheetheader(wb.getSheetAt(0), 0, 0, null);
//		ImportContext.getContext().setHead(head);
//
//		List<Map<String, String>> sheetCont = getSheetContent(wb.getSheetAt(0), head, 0, 1, null, null);
//
//		logger.info(sheetCont.size());

//		getExcelContent(path, 0, 0);

        //	System.out.println(Runtime.getRuntime().freeMemory());


        //	Thread.sleep(1000 * 10L);

        System.out.println("开始");
        Workbook wb = ExcelUtil.getWorkbook(path);
        List<Head> keys = AnalysisUtil.getSheetheader(wb.getSheetAt(0), 0, 0, null);
        List<Map<String, String>> sheetCont = AnalysisUtil.getSheetContent(wb.getSheetAt(0), keys, 0, 0, null, null, "lineNumber_");
//		System.out.println( getExcelContent(path, 0, 0));

        FileWriter f = new FileWriter("c:\\数据.txt");


        for( Map<String, String> map : sheetCont ){
            //	System.out.println(JSONUtil.toString(map)+",");
//			f.write(JSONUtil.toString(map)+",\n");
        }
        f.flush();
        f.close();


//		System.out.println(Arrays.toString( keys.toArray()));
        System.out.println("读取完成");
    }

}
