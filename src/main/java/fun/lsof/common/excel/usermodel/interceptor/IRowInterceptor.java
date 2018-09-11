package fun.lsof.common.excel.usermodel.interceptor;

import java.util.List;
import java.util.Map;

/**
 * excel 处理行拦截器，可对每行进行拦截处理
 *
 * @author jerry
 * @date 2017 -06-16 18:54:48
 */
public interface IRowInterceptor
{

	/**
	 * 初始化调用，只会调用一次
	 *
	 * @throws Exception exception
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	void init() throws Exception;

	/**
	 * 读取了一行数据，但是为处理的时候调用
	 *
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	void before();

	/**
	 * 处理完了一行记录之后调用
	 *
	 * @param rows rows
	 * @param row  row
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	void after(Map<String, String> rows, int row);

	/**
	 * 执行截断处理，通常用于多少批次后提交业务。和 @getTruncationRow 一起使用
	 * 注意：最后一个批次也会触发.
	 *
	 * @param sheetContent sheet content
	 * @param row          row
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	void truncation(List<Map<String, String>> sheetContent, int row);

	/**
	 * 在每逢读取到多少行的时候触发.
	 *
	 * @return truncation row
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	int getTruncationRow();

	/**
	 * 工作表读取完成后调用，值调用一次
	 *
	 * @author jerry
	 * @date 2017 -06-16 18:54:48
	 */
	void end();
}
