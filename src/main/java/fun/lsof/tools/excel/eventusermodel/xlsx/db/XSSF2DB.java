//package fun.lsof.tools.excel.eventusermodel.xlsx.db;
//
//import java.sql.Connection;
//import java.sql.PreparedStatement;
//import java.sql.ResultSet;
//import java.sql.SQLException;
//import java.util.ArrayList;
//import java.util.Date;
//import java.util.HashMap;
//import java.util.List;
//import java.util.Map;
//
//import javax.naming.NamingException;
//
//import com.wordty.common.assist.utils.DBUtil;
//import fun.lsof.tools.excel.eventusermodel.xlsx.handler.ReadXSSFBaseHandler;
//import fun.lsof.tools.excel.eventusermodel.xlsx.db.vo.DBFieldType;
//import fun.lsof.tools.excel.eventusermodel.xlsx.extern.SimpleCell;
//import fun.lsof.tools.excel.eventusermodel.xlsx.extern.SimpleRow;
//import org.apache.commons.lang.StringUtils;
//import org.apache.log4j.Logger;
//import org.apache.poi.ss.usermodel.DateUtil;
//
///**
// * 读取 excel 直接写入数据
// */
//public class XSSF2DB extends ReadXSSFBaseHandler {
//
//    private static Logger LOGGER = Logger.getLogger(XSSF2DB.class);
//
//    Connection conn = null;
//    PreparedStatement statement = null;
//    String tableName = null;
//
//    /**
//     * 是否在插入的时候增加excel行号
//     */
//    boolean insertRowNum = false;
//    /**
//     * excel行号名称
//     */
//    String rowColumnName = "";
//
//    private int validInsertRowSum = 0;
//
//    int cuttingRow = 100000;
//
//    /**
//     * 和stencilTitley要一一对于的表字段名称
//     */
//    private List<DBFieldType> dbFieldTypes = new ArrayList<DBFieldType>();
//
//    /**
//     * @param path      excel的路径（包括文件名称）
//     * @param sheetName 要导入的sheet的名称
//     * @param tableName 导入的目标表
//     */
//    public XSSF2DB(String path, String sheetName, String tableName) {
//        super(path, sheetName);
//        this.tableName = tableName;
//    }
//
//    /**
//     * @param path          excel的路径（包括文件名称）
//     * @param sheetName     要导入的sheet的名称
//     * @param tableName     导入的目标表
//     * @param rowColumnName 记录excel行数的字段名称
//     * @param cuttingRow    导入行数 % cuttingRow == 0 的时候插入数据库
//     */
//    public XSSF2DB(String path, String sheetName, String tableName, String rowColumnName, int cuttingRow) {
//        this(path, sheetName, tableName);
//        if (null == rowColumnName || rowColumnName.trim().length() == 0) {
//            throw new NullPointerException("参数rowColumnName不能为空");
//        }
//        this.rowColumnName = rowColumnName;
//        this.insertRowNum = true;
//        this.cuttingRow = cuttingRow;
//    }
//
//    @Override
//    protected void initTitle() {
//        //要先初始化列头信息
//        super.initTitle();
//        initDBConfig();
//    }
//
//    @Override
//    public void endRow(int rowNum) {
//        super.endRow(rowNum);
//
//        //当数据达到10W的时候，插入一次
//        if (rowNum % cuttingRow == 0 && !globalHaveException && rowNum != 0) {
//            //如果检查出有不合格的，需要回滚
//            try {
//                LOGGER.info("----midway insert to DB, size=" + workbookMap.size());
//                processImport();
//                LOGGER.info("----end midway insert to DB, size=" + workbookMap.size());
//                workbookMap = new ArrayList<SimpleRow>();
////				System.gc();//TODO 注意，此处为手工调用GC来回收内存，因此可能会影响到性能
//            } catch (Exception e) {
//                LOGGER.error(e.getMessage(), e);
//                throw new RuntimeException(e);
//            }
//            validInsertRowSum += cuttingRow;
//        }
//    }
//
//    /**
//     * 初始化导入的数据库信息，包括预编译sql（先要初始化列头才能调用该方法）
//     */
//    private void initDBConfig() {
//        try {
//            conn.setAutoCommit(false);
//        } catch (Exception e) {
//            LOGGER.error(e.getMessage(), e);
//            throw new RuntimeException("创建数据库连接失败:" + e.getMessage());
//        }
//        String sqlInsert = "";
//        try {
//            List<String> columns = getTitle4DBColumn();
//            if (insertRowNum)
//                columns.add(rowColumnName);
//            sqlInsert = "INSERT INTO $TABLE ($COLUMNS) values ($VALUES)";
//            sqlInsert = StringUtils.replace(sqlInsert, "$TABLE", tableName);
//            sqlInsert = StringUtils.replace(sqlInsert, "$COLUMNS", StringUtils.join(columns.iterator(), ","));
//            sqlInsert = StringUtils.replace(sqlInsert, "$VALUES", StringUtils.repeat(",?", columns.size()).substring(1));
//
//            statement = conn.prepareStatement(sqlInsert);
//        } catch (Exception e) {
//            LOGGER.error(e.getMessage(), e);
//            throw new RuntimeException("创建预编译sql失败,sql=[" + sqlInsert + "], error info:" + e.getMessage());
//        }
//    }
//
//    public void processImport() throws SQLException {
//        for (SimpleRow row : workbookMap) {
//            Map<String, SimpleCell> simpMap = row.getRowData();
//            for (int i = 0; i < stencilTitle.size(); i++) {
//                String sTitle = stencilTitle.get(i);
//                SimpleCell cell = simpMap.get(sTitle);
//
//                Object value = cell.getFormattedValue();
////				String format = cell.getFormatString();
//
//                //需要格式化date然后再插入数据库
//                if (null != value && !StringUtils.isBlank(String.valueOf(value))) {
//                    //91 or 93是时间类型
//                    DBFieldType dbType = dbFieldTypes.get(i);
//                    if (dbType.get_type() == 91 || dbType.get_type() == 93) {
//                        //利用poi的dateUtil解析单元格的date类型
//                        Date date = DateUtil.getJavaDate(Double.parseDouble(cell.getCellValue().toString()));
//                        if (dbType.get_type() == 91) {
//                            value = new java.sql.Date(date.getTime());
//                        } else {
//                            value = new java.sql.Timestamp(date.getTime());
//                        }
//                    }
//                }
//                statement.setObject(i + 1, value);
//            }
//            //增加行号
//            if (insertRowNum) {
//                statement.setObject(stencilTitle.size() + 1, row.getRowNum());
//            }
//            statement.addBatch();
//        }
//        statement.executeBatch();
//    }
//
//    @Override
//    protected void cellFilter(String sTitle, int stencilTitleIndex, SimpleCell simpleCell) {
//        DBFieldType dbtype = dbFieldTypes.get(stencilTitleIndex);
//        String value = simpleCell.getFormattedValue();
//
//        //不能为空
//        if (!dbtype.isNull() && StringUtils.isEmpty(value)) {
//            error.putOtherError(sTitle, "不能为空");
//            return;
//        }
//
//        //可以为空
//        //不验证单元格为空的
//        if (StringUtils.isEmpty(value)) {
//            return;
//        }
//
//        //日期类型不验证
//        if (dbtype.get_type() == 91 || dbtype.get_type() == 93) {
//            return;
//        }
//
//
//        //只验证数字类型
//        if ("NUMBER".equals(dbtype.getType())) {
//            if (!value.matches("^[-+]?[0-9]+(\\.[0-9]+)?$")) {
//                error.putOtherError(sTitle, "只能是数字");
//            }
//        } else {//其他只验证长度
//            if (value.length() >= dbtype.getLength()) {
//                error.putOtherError(sTitle, "字符长度过长");
//            }
//        }
//    }
//
//    /**
//     * 返回和protoTitle一一对于的表字段
//     */
//    public List<String> getTitle4DBColumn() {
//        if (null == dbFieldTypes || dbFieldTypes.isEmpty()) {
//            throw new NullPointerException("表字段类型不能为空！");
//        }
//
//        List<String> dbColumn = new ArrayList<String>();
//        for (DBFieldType dbt : this.dbFieldTypes) {
//            dbColumn.add(dbt.getFieldName());
//        }
//        return dbColumn;
//    }
//
//    public static Connection getConnection() throws SQLException, NamingException, IllegalAccessException {
//        return DBUtil.getConnection(DBUtil.Type.jdbc,
//                "oracle.jdbc.driver.OracleDriver",
//                "jdbc:oracle:thin:@jerry-svr:1521:orcl",
//                "IPS_IMPORT",
//                "IPS_IMPORT");
//    }
//
//    @Override
//    public void parse() throws Exception {
//        try {
//            LOGGER.info("start parse...");
//            super.parse();
//            LOGGER.info("end parse...");
//            LOGGER.info("start insert to DB, size=" + workbookMap.size());
//            processImport();
//            LOGGER.info("end insert");
//
//            validInsertRowSum += workbookMap.size();
//
//            //此处rollback是因为分部导入了，发生错误之后要回滚
//            if (globalHaveException) {
//                try {
//                    conn.rollback();
//                } catch (Exception ex) {
//                }
//                ;
//            }
//
//            conn.commit();
//            LOGGER.info("conmmit record");
//        } catch (Exception e) {
//            try {
//                conn.rollback();
//            } catch (Exception ex) {
//            }
//            LOGGER.error(e.getMessage(), e);
//            throw e;
//        } finally {
//            DBUtil.close(conn, statement);
//            LOGGER.info("close resource");
//        }
//    }
//
//
//    /**
//     * 读取数据库表的字段信息.
//     *
//     * @param tergetTableName terget table name
//     * @param connection      connection
//     * @param schema          schema
//     * @return field type 4 db
//     * @author jerry
//     * @date 2018 -09-11 21:08:40
//     */
//    public Map<String, DBFieldType> getFieldType4DB(String tergetTableName, Connection connection, String schema) {
//        if (null == connection) {
//            throw new NullPointerException("必须调用setConnection()设置Connection!");
//        }
//
//        Map<String, DBFieldType> dbmap = new HashMap<String, DBFieldType>();
//        ResultSet rs = null;
//        try {
//            rs = connection.getMetaData().getColumns("", schema.toUpperCase(), tergetTableName.toUpperCase(), "%");
//            DBFieldType fieldType = null;
//            while (rs.next()) {
//                boolean isNull = rs.getInt("NULLABLE") == 0 ? false : true;
//                fieldType = new DBFieldType(rs.getString("COLUMN_NAME"), rs.getString("TYPE_NAME"), rs.getInt("DATA_TYPE"), rs.getInt("COLUMN_SIZE"), isNull);
//                dbmap.put(rs.getString("COLUMN_NAME"), fieldType);
//            }
//        } catch (Exception e) {
//            LOGGER.error(e.getMessage(), e);
//            throw new RuntimeException("获取表配置失败：" + e.getMessage());
//        } finally {
//            DBUtil.close(rs);
//        }
//        return dbmap;
//    }
//
//
//    public void setDbFieldTypes(List<DBFieldType> dbFieldTypes) {
//        this.dbFieldTypes = dbFieldTypes;
//    }
//
//    public void setConnection(Connection conn) {
//        this.conn = conn;
//    }
//
//    public void setInsertRowNum(boolean insertRowNum) {
//        this.insertRowNum = insertRowNum;
//    }
//
//    public void setRowColumnName(String rowColumnName) {
//        this.rowColumnName = rowColumnName;
//    }
//
//    public int getValidInsertRowSum() {
//        return validInsertRowSum;
//    }
//
//}
