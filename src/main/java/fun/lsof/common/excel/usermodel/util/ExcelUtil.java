package fun.lsof.common.excel.usermodel.util;

import com.wordty.common.assist.utils.ResourceCloseUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil
{
	/**
	 * 实例化 excel workbook.
	 *
	 * @param path  excel 所在全路径
	 * @return workbook
	 * @throws Exception exception
	 * @author jerry
	 * @date 2017 -06-16 19:12:23
	 */
	public static Workbook getWorkbook(String path) throws Exception{
		FileInputStream fis = null;
		Workbook book = null;
		try{
			fis = new FileInputStream(path);
			book = WorkbookFactory.create(fis);

		}catch (Exception e) {
			throw e;
		}finally{
			ResourceCloseUtil.close(fis);
		}
		return book;
	}


	/**
	 * 使用模板来构建新的 excel
	 * @param ltable 需要写入的数据
	 * @param objKey 按循序的 key
	 * @param templateFile 模板文件路径
	 * @param writeFirstRow 从第几行开始写入
	 * @author jerry
	 * @return 字节流
	 * @throws Exception
	 */
	public static byte[] writeExcelFormTemplate(List<Map<String, Object>> ltable, List<String> objKey, String templateFile, int writeFirstRow) throws Exception {
		ByteArrayOutputStream outputStream =  null;
		try {
			Workbook wb = getWorkbook(templateFile);

			Sheet sheetAt = wb.getSheetAt(0);
			int rowCount = writeFirstRow;
			for( Map<String, Object> _row : ltable ){
				Row _newRow = sheetAt.createRow(rowCount);
				int cellCount = 0;
				for (String _key : objKey ){
					Cell _newCell = _newRow.createCell(cellCount);
					_newCell.setCellValue( _row.get(_key).toString() );
					cellCount++;
				}
				rowCount++;
			}
			outputStream = new ByteArrayOutputStream();
			wb.write(outputStream);

			return outputStream.toByteArray();
		}catch (Exception e){
			throw new RuntimeException(e);
		}finally {
			ResourceCloseUtil.close(outputStream);
		}
	}

	/**
	 * 使用模板来构建新的 excel
	 * @param ltable 需要写入的数据
	 * @param objKey 按循序的 key
	 * @param templateFile 模板文件路径
	 * @param targetFile 需要保存到的文件路径
	 * @param writeFirstRow 从第几行开始写入
	 * @author jerry
	 * @throws Exception
	 */
	public static void writeExcelFormTemplate(List<Map<String, Object>> ltable, List<String> objKey, String templateFile, String targetFile, int writeFirstRow) throws Exception {
		FileOutputStream outputStream = null;
		try {
			byte[] byts = writeExcelFormTemplate(ltable, objKey, templateFile, writeFirstRow);
			outputStream = new FileOutputStream(targetFile);
			outputStream.write(byts);
		}finally {
			ResourceCloseUtil.close(outputStream);
		}
	}

	public static void main (String[] s) throws Exception {

		List<Map<String, Object>> ltable = new ArrayList<Map<String, Object>>();

		Map<String, Object> row = new HashMap<String, Object>();
		row.put("id", "1");
		row.put("name", "2");
		row.put("loginid", "3");
		row.put("phone", "4");
		row.put("xxx", "5");

		ltable.add(row);

		Map<String, Object> row2 = new HashMap<String, Object>();
		row2.put("id", "1");
		row2.put("name", "2");
		row2.put("loginid", "3");
		row2.put("phone", "4");
		row2.put("xxx", "5");

		ltable.add(row2);


		List<String> objKey = new ArrayList<String>();
		objKey.add("id");
		objKey.add("name");
		objKey.add("loginid");
		objKey.add("phone");
		objKey.add("xxx");

		String file = "/Users/jerry/Mix/共享/networdTmpl.xlsx";

		writeExcelFormTemplate(ltable, objKey, file, "/Users/jerry/Mix/共享/networdTmpl2.xlsx",  3);

	}


}
