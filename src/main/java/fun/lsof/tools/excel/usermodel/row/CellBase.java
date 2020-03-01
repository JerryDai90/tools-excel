package fun.lsof.tools.excel.usermodel.row;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Locale;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;


/**
 * 单元格基本读取实现类.
 *
 * @author jerry
 * @date 2017 -06-16 19:11:37
 */
public class CellBase implements ICell
{
	static Logger logger = Logger.getLogger(CellBase.class);
	HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
    DecimalFormat decimalFormat = new DecimalFormat("#.##########");
    
	public Object getCellVal(Cell cell)
	{
		if( null == cell ){
			return null;
		}
		int cellType = cell.getCellType();
		
		if( cellType == Cell.CELL_TYPE_NUMERIC ){
			if(HSSFDateUtil.isCellDateFormatted(cell))
	    	{
	    		SimpleDateFormat sdf = new SimpleDateFormat("yyyy��MM��dd��", Locale.CHINA);
	    		return sdf.format(cell.getDateCellValue());
	    	}
	    	else
	    	{
	    		return decimalFormat.format(cell.getNumericCellValue());
             // return  String.valueOf(cell.getNumericCellValue());//dataFormatter.formatCellValue(cell);
	    	}
			
//			String cellDataFormatStr = cell.getCellStyle().getDataFormatString();
//			if(isDateFormatValue(cellDataFormatStr))
//			{
//				Date temp_date = cell.getDateCellValue();
//				return temp_date;
//			}
//			
//			return cell.getNumericCellValue();
		}else if(cellType == Cell.CELL_TYPE_STRING){
			return cell.getStringCellValue();
		}else if(cellType == Cell.CELL_TYPE_BLANK){
			return "";
		}else{
			return "";
		}
	}
	
	public static boolean isDateFormatValue(String cellDataFormatStr){
		if("yyyy\\-mm\\-dd;@".equals(cellDataFormatStr)
				||"d/m/yyyy;@".equals(cellDataFormatStr)
				||"m/d/yy".equals(cellDataFormatStr)){
			return true;
		}
		return false;
	}

}
