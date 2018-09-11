package fun.lsof.common.excel.eventusermodel.xlsx.exception;


/**
 * excel 读取异常.
 *
 * @author jerry
 * @date 2018 -09-11 21:09:00
 */
public class ReadExcelException extends RuntimeException {
    public ReadExcelException() {
    }

    public ReadExcelException(String message) {
        super(message);
    }

    public ReadExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ReadExcelException(Throwable cause) {
        super(cause);
    }

    public ReadExcelException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
