package fun.lsof.common.excel.usermodel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import fun.lsof.common.excel.usermodel.interceptor.IRowInterceptor;
import fun.lsof.common.excel.usermodel.row.CellBase;
import fun.lsof.common.excel.usermodel.row.Head;
import fun.lsof.common.excel.usermodel.row.ICell;
import fun.lsof.common.excel.usermodel.util.ExcelUtil;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * 解析 excel 工具类.
 *
 * @author jerry
 * @date 2017 -06-16 18:51:29
 */
public class AnalysisUtil
{
	static Logger logger = Logger.getLogger(AnalysisUtil.class);

	/**
	 * 读取 sheet 的数据
	 *
	 * @param sheet           需要读取的工作表，必填
	 * @param head            表头数据信息，必填
	 * @param firstCellNum    开始读取的列，必填
	 * @param firstRowNum     开始读取的行（除去表头）必填
	 * @param cell            单元格拦截器处理，可为空
	 * @param iRowInterceptor 行拦截器处理，可为空
	 * @return content
	 * @throws Exception exception
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static List<Map<String, String>> getSheetContent(org.apache.poi.ss.usermodel.Sheet sheet, List<Head> head, Integer firstCellNum, Integer firstRowNum, ICell cell, IRowInterceptor iRowInterceptor) throws Exception{
		return getSheetContent(sheet, head, firstCellNum, firstRowNum, cell, iRowInterceptor, null);
	}

	/**
	 * 读取 sheet 的数据
	 *
	 * @param sheet           sheet
	 * @param head            表头
	 * @param firstCellNum    从0开始数
	 * @param firstRowNum     first row num
	 * @param cell            cell
	 * @param iRowInterceptor row interceptor
	 * @param lineNumber      line number
	 * @return content
	 * @throws Exception exception
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static List<Map<String, String>> getSheetContent(org.apache.poi.ss.usermodel.Sheet sheet, List<Head> head, Integer firstCellNum, Integer firstRowNum, ICell cell, IRowInterceptor iRowInterceptor, String lineNumber) throws Exception{
		cell = cell == null ? new CellBase(): cell;
		boolean rowInterceptor = null != iRowInterceptor;
		//保存整个sheet内容，一个Map就是行数据，每一个key和value就是一个单元格
		List<Map<String, String>> sheetContent = new ArrayList<Map<String, String>>();
		Map<String, String> loadRow = null;
		Integer lastCellNum = head.size();
		if( null == lastCellNum || null == firstCellNum ){
			throw new RuntimeException("开始的行数或者开始的列为空");
		}

		if( rowInterceptor )
			iRowInterceptor.init();

		for( int i = (firstRowNum == null ? 1 : firstRowNum), numberOfRow = sheet.getLastRowNum(); i <= numberOfRow; i++ ){
			org.apache.poi.ss.usermodel.Row row = sheet.getRow(i);
			if( rowInterceptor )
				iRowInterceptor.before();

			if( null == row )//还需要记录空的行数
				continue;
			loadRow = new HashMap<String, String>();

			for( Head _hd : head ){//根据列头的index来读取内容
				Object obj = cell.getCellVal(row.getCell(_hd.getIndex()));
				loadRow.put(_hd.getText(), null == obj ? null : obj.toString());
			}
			if( null != lineNumber ){
				//excel的行是从0开始的
				loadRow.put(lineNumber, i+1+"");
			}

			sheetContent.add(loadRow);

			if( rowInterceptor ){
				iRowInterceptor.after(loadRow, i);
				if( i % iRowInterceptor.getTruncationRow()  == 0 || i == sheet.getLastRowNum() ){
					iRowInterceptor.truncation(sheetContent, i);
					sheetContent = new ArrayList<Map<String,String>>();
				}
			}
		}
		if( rowInterceptor )
			iRowInterceptor.end();

		return sheetContent;
	}

	/**
	 * 获取表头
	 *
	 * @param sheet        sheet
	 * @param firstRowNum  first row num
	 * @param firstCellNum first cell num
	 * @param cell         获取单元格的类 ，但cell为空的时候采用内置的处理方式
	 * @return
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static List<Head> getSheetheader(org.apache.poi.ss.usermodel.Sheet sheet, Integer firstRowNum, Integer firstCellNum, ICell cell){
		cell = cell == null ? new CellBase() : cell;
		List<Head> sheetHeader = new ArrayList<Head>();
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(firstRowNum);
		int numberOfColumns = row.getLastCellNum();
		Head head = null;
		for(int i = (firstCellNum == null ? 0 : firstCellNum); i < numberOfColumns; i ++){
			Object obj = cell.getCellVal(row.getCell(i)); //还需要记录空的列数
			head = new Head(i, obj.toString());
			sheetHeader.add(head);
		}
		return sheetHeader;
	}

	/**
	 * Filter string [ ].
	 *
	 * @param rawDatas raw datas
	 * @return string [ ]
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static String[] filter(String[] rawDatas){
		List<String> depot = new ArrayList<String>();
		for(String rawData : rawDatas){
			if( null != rawData && rawData.length() != 0 )
				depot.add(rawData);
		}
		return depot.toArray(new String[0]);
	}

	/**
	 * 获取 excel 的内容.
	 *
	 * @param path         excel路径
	 * @param firstRowNum  first row num
	 * @param firstCellNum first cell num
	 * @return content
	 * @throws Exception exception
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static List<Map<String, String>> getExcelContent(String path, int firstRowNum, int firstCellNum) throws Exception{
		//step2. 解析excel
		List<Map<String, String>> content = null;
		List<Head> keys = null;
		try{
			Workbook wb = ExcelUtil.getWorkbook(path);
			org.apache.poi.ss.usermodel.Sheet firstSheet = wb.getSheetAt(0);
			keys = getSheetheader(firstSheet, firstRowNum, firstCellNum, new  ICell(){
				public Object getCellVal(Cell cell)
				{
					Object src = new CellBase().getCellVal(cell);
					return null != src ? src.toString().replaceAll(" ", "")
							.replaceAll("\\(", "")
							.replaceAll("\\)", "")
							.replaceAll("/", "")
							.replaceAll("-", "")
							.replaceAll("@", "")
							.replaceAll("'", "")
							.replaceAll("&", "")
							.replaceAll("\n", "")
							.replaceAll("\r", "")
							.replaceAll("\\.", ""): null;
				}
			});
			logger.info("列头："+(Arrays.toString(keys.toArray())));
			content = getSheetContent(firstSheet, keys, firstCellNum, firstRowNum+1, null, null);
		}catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("step2. 执行失败，失败原因："+e.getMessage());
		}

		return content;
	}


	/**
	 * Getdb field list.
	 *
	 * @param head head
	 * @return list
	 * @author jerry
	 * @date 2017 -06-16 18:51:29
	 */
	public static List<String> getdbField(List<Head> head){
		List<String> fields = new ArrayList<String>();
		for( Head cell : head ){
			if( !StringUtils.isBlank( cell.getText()) ){
				fields.add(cell.getText());
			}
		}
		return fields;
	}

}



