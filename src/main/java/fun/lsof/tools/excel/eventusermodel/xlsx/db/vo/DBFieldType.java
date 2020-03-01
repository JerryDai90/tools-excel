package fun.lsof.tools.excel.eventusermodel.xlsx.db.vo;


/**
 * .
 *
 * @author jerry
 * @date 2020 -03-01 13:22:53
 */
public class DBFieldType {
    /**
     * DB字段名称
     */
    private String fieldName;
    /**
     * 字段类型
     */
    private String type;
    /**
     * 用java.sql.Types描述的字段类型
     */
    private int _type;
    /**
     * 字段长度
     */
    private int length;
    /**
     * 是否可以为空
     */
    private boolean isNull;

    public DBFieldType(String fieldName, String type, int _type, int length, boolean isNull) {
        this.fieldName = fieldName;
        this.type = type;
        this._type = _type;
        this.length = length;
        this.isNull = isNull;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public int get_type() {
        return _type;
    }

    public void set_type(int _type) {
        this._type = _type;
    }

    public int getLength() {
        return length;
    }

    public void setLength(int length) {
        this.length = length;
    }

    public boolean isNull() {
        return isNull;
    }

    public void setNull(boolean isNull) {
        this.isNull = isNull;
    }

}
