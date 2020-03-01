package fun.lsof.tools.excel.utils;

import fun.lsof.tools.excel.eventusermodel.xlsx.exception.CloseResourceException;

import java.io.*;
import java.lang.reflect.Method;

/**
 * provide close all resource method. call the close method through reflection.
 *
 * @author jerry
 * @date 2017 -06-16 18:37:27
 */
public class ResourceCloseUtil {

    /**
     * Close all resource, but resource must has a close method.
     *
     * @param obj resource list, at least one.
     * @throws CloseResourceException when resource not found, throw this exception.
     * @author jerry
     * @date 2017 -06-16 18:37:27
     */
    public static void close(Object... obj) throws CloseResourceException {
        for (Object _obj : obj) {
            if (null == _obj) {
                continue;
            }
            try {
                if (!_close(_obj)) {
                    Class<? extends Object> clazz = _obj.getClass();
                    Method method = clazz.getMethod("close");
                    method.invoke(_obj);
                }
            } catch (Exception e) {
                throw new CloseResourceException("can not close resource：" + _obj.getClass());
            }
        }
    }

    /**
     * Close resource.
     *
     * @param _obj resource list
     * @return boolean
     * @throws Exception exception
     * @author jerry
     * @date 2017 -06-16 18:37:27
     */
    private static boolean _close(Object _obj) throws Exception {
        if (_obj instanceof Closeable) {
            ((Closeable) _obj).close();
        } else {
            return false;
        }
        return true;
    }

}
