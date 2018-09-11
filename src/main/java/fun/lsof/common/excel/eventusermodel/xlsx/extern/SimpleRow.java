package fun.lsof.common.excel.eventusermodel.xlsx.extern;

import java.util.HashMap;
import java.util.Map;

/**
 * 一行 excel 数据.
 *
 * @author jerry
 * @date 2018 -09-11 20:23:43
 */
public class SimpleRow {
    /**
     * excel 中的行数.
     */
    private int rowNum;
    /**
     * 这一行的数据.
     */
    private Map<String, SimpleCell> rowData;

    public SimpleRow(int rowNum, Map<String, SimpleCell> rowData) {
        this.rowNum = rowNum;
        this.rowData = new HashMap<String, SimpleCell>(rowData);
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public Map<String, SimpleCell> getRowData() {
        return rowData;
    }

    public void setRowData(Map<String, SimpleCell> rowData) {
        this.rowData = rowData;
    }

    public String toString() {
        return "{rowNum : " + rowNum + ", rowData:" + rowData.toString() + "}";
    }
}
