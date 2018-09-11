package fun.lsof.common.excel.eventusermodel.xlsx.exception;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * .
 *
 * @author jerry
 * @date 2017 -06-16 19:07:16
 */
public class XlsContentImproperException implements java.io.Serializable {

    private static final long serialVersionUID = -4843374527405500423L;

    /**
     * 文件名称.
     */
    public String fileName = null;
    /**
     * sheet 名称.
     */
    public String sheetName = null;
    /**
     * 每行 excel 的数据信息.
     */
    public Collection<Entry> errors = new ArrayList<Entry>();

    public String getFileName() {
        return fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public Collection<Entry> getErrors() {
        return errors;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setErrors(Collection<Entry> errors) {
        this.errors = errors;
    }

    public static class Entry {
        /**
         * 当前的所在行.
         */
        public int rowNum;
        /**
         * 一行 excel 的数据信息.
         */
        public Map<String, String> hmData = null;

        /**
         * 验证数据是否符合规则.
         */
        public Map<String, String> otherError = new HashMap<String, String>();
        ;

        public int getRowNum() {
            return rowNum;
        }

        public Map<String, String> getHmData() {
            return hmData;
        }

        public Map<String, String> getOtherError() {
            return otherError;
        }

        public void setOtherError(Map<String, String> otherError) {
            this.otherError = otherError;
        }

        public boolean hasError() {
            return !(otherError.isEmpty());
        }

        /**
         * Put other error.
         *
         * @param key  列头名称
         * @param info 错误信息
         * @author jerry
         * @date 2017 -06-16 19:09:32
         */
        public void putOtherError(String key, String info) {
            String val = null;
            if (otherError.containsKey(key)) {
                val = otherError.get(key);
            }
            val = (null != val) ? val + "、" : "";
            otherError.put(key, val + info);
        }

    }


}
