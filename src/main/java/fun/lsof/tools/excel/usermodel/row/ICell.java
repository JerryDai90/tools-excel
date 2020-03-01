package fun.lsof.tools.excel.usermodel.row;

/**
 * 单元格处理
 *
 * @author jerry
 * @date 2017 -06-16 18:49:21
 */
public interface ICell
{
	/**
	 * 获取一个单元格的值
	 *
	 * @param cell 单元格对象
	 * @return cell val
	 * @author jerry
	 * @date 2017 -06-16 18:49:21
	 */
	Object getCellVal(org.apache.poi.ss.usermodel.Cell cell);
	
}
