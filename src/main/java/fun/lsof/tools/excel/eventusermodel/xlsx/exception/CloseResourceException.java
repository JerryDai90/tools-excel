package fun.lsof.tools.excel.eventusermodel.xlsx.exception;


/**
 * 关闭出错异常类.
 *
 * @author jerry
 * @date 2017 -06-16 19:17:49
 */
public class CloseResourceException extends RuntimeException{
    public CloseResourceException() {
        super();
    }
    public CloseResourceException(String message) {
        super(message);
    }
    public CloseResourceException(Throwable cause) {
        super(cause);
    }
}
