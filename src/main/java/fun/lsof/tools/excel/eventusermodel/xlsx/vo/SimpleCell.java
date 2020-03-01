package fun.lsof.tools.excel.eventusermodel.xlsx.vo;


import fun.lsof.tools.excel.eventusermodel.xlsx.handler.XSSFSheetXMLHandler;

/**
 * 单元格.
 *
 * @author jerry
 * @date 2018 -09-11 20:23:19
 */
public class SimpleCell {
    /**
     * The Cell reference.
     */
    private String cellReference;
    /**
     * excel 的真正类型.
     */
    private XSSFSheetXMLHandler.xssfDataType type;
    /**
     * 格式化后的数据.
     */
    private String formattedValue;
    /**
     * 未格式化的数据.
     */
    private String formatString;
    /**
     * 单元格值.
     */
    private String cellValue;
    /**
     * 单元格所在的下标.
     */
    private int cellIndex;

    public SimpleCell(String cellReference, XSSFSheetXMLHandler.xssfDataType type,
                      String formattedValue, String formatString, int cellIndex, String cellValue) {
        this.cellReference = cellReference;
        this.type = type;
        this.formattedValue = formattedValue;
        this.formatString = formatString;
        this.cellIndex = cellIndex;
        this.cellValue = cellValue;
    }

    public String getCellReference() {
        return cellReference;
    }

    public void setCellReference(String cellReference) {
        this.cellReference = cellReference;
    }

    public XSSFSheetXMLHandler.xssfDataType getType() {
        return type;
    }

    public void setType(XSSFSheetXMLHandler.xssfDataType type) {
        this.type = type;
    }

    public String getFormattedValue() {
        return formattedValue;
    }

    public void setFormattedValue(String formattedValue) {
        this.formattedValue = formattedValue;
    }

    public String getFormatString() {
        return formatString;
    }

    public void setFormatString(String formatString) {
        this.formatString = formatString;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getCellValue() {
        return cellValue;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }

    @Override
    public String toString() {
        return "{type : " + type + ", cellIndex:" + cellIndex + ", formattedValue:" + formattedValue + ", formatString:" + formatString + "}";
    }
}
