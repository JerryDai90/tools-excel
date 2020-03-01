package fun.lsof.tools.excel.utils;

import junit.framework.TestFailure;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

public class ResourceCloseUtilTest {

    @Test
    public void close() throws FileNotFoundException {

        FileInputStream f = null;
        Object obj = new Object();
        ResourceCloseUtil.close(f);
    }



}
