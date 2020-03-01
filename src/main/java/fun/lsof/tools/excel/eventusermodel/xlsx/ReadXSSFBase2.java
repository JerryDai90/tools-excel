package fun.lsof.tools.excel.eventusermodel.xlsx;

import java.util.*;

import fun.lsof.tools.excel.eventusermodel.xlsx.exception.XlsContentImproperException;
import fun.lsof.tools.excel.eventusermodel.xlsx.vo.SimpleCell;
import fun.lsof.tools.excel.eventusermodel.xlsx.vo.SimpleRow;
import fun.lsof.tools.excel.eventusermodel.xlsx.handler.WorkbookXMLTable;
import fun.lsof.tools.excel.eventusermodel.xlsx.handler.XSSFSheetXMLHandler;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;


/**
 * <pre>
 * 1. 先获取excel的列头
 * 		1.1 提供接口给子类，让起校验列头是否合法
 *
 * 2. 开始读取数据（记录数据，要判断是否需要数据合法性）（数据记录map<列头，列值>）
 * 		2.1 提供接口给子类，让其校验数据的合法性
 * <pre>
 * @author Administrator
 */
public class ReadXSSFBase2 implements XSSFSheetXMLHandler.SheetContentsHandler {

    private XSSFReader xssfReader;
    /**
     * excel中重复的字符存放表
     */
    private ReadOnlySharedStringsTable sst;
    private StylesTable stylesTable;

    /**
     * 保存一行的数据
     */
    protected Map<String, SimpleCell> rowSimpleCellMap = null;

    /**
     * 保存成功的数据
     */
    protected List<SimpleRow> workbookMap = new ArrayList<SimpleRow>();

    /**
     * 标记已经检查出错误，不让再存储数据（减少内存的使用）
     */
    protected boolean globalHaveException = false;

    /**
     * 记录本次导入的异常
     */
    protected XlsContentImproperException improperException = new XlsContentImproperException();
    /**
     * 记录行异常
     */
    protected XlsContentImproperException.Entry error = null;

    /**
     * 当前解析的excel列头
     */
    protected List<Title> cExcelTitle = new ArrayList<Title>();
    /**
     * 带有上传excel列头下标的title
     */
    protected Map<String, String> cIndexExcelTitle = null;
    /**
     * excel中有效的行数，忽略空行
     */
    protected int validRowSum = 0;

    String sheetName = null;
    String filePath = null;


    /**
     * 上传的excel列头，用于读取列头名称
     * <p>
     * (0,0)
     * |---------------> x
     * |
     * |
     * ↓
     * y
     */
    protected int startTitleX = 0;//列
    protected int startTitleY = 0;//行


    public ReadXSSFBase2(String path, String sheetName) {
        this.filePath = path;
        this.sheetName = sheetName;
    }

    public List<SimpleRow> getWorkbookMap() {
        return workbookMap;
    }

    public XlsContentImproperException getImproperException() {
        return improperException;
    }

    public int getValidRowSum() {
        return validRowSum;
    }

    public void setStartTitleX(int startTitleX) {
        this.startTitleX = startTitleX;
    }

    public void setStartTitleY(int startTitleY) {
        this.startTitleY = startTitleY;
    }


    @Override
    public void startRow(int rowNum) {
        validRowSum++;
        rowSimpleCellMap = new HashMap<String, SimpleCell>();
    }

    @Override
    public void endRow(int rowNum) {
        error = new XlsContentImproperException.Entry();

        if (cExcelTitle.isEmpty()) {
            throw new RuntimeException("未在上传的excel中的第" + startTitleY + "找到列头");
        }
        //调用方法校验列头
        checkTitle(cExcelTitle);

        //忽略表头
        if (startTitleY == rowNum) {
            return;
        }

        //验证
        for (int i = 0; i < cExcelTitle.size(); i++) {
            String sTitle = cExcelTitle.get(i).getName();
            SimpleCell vsinp = rowSimpleCellMap.get(sTitle);
            if (null == vsinp) {
                //补充空值
                vsinp = new SimpleCell("", null, "", "", -1, "");
                rowSimpleCellMap.put(sTitle, vsinp);
            }
            cellFilter(sTitle, i, vsinp);
        }

        //处理错误行
        if (error.hasError()) {
            if (!globalHaveException) {
                globalHaveException = true;
            }
            error.rowNum = rowNum;
            error.hmData = simple2Map(rowSimpleCellMap);
            improperException.errors.add(error);
        } else {
            //一旦检查出有错误，就不在记录数据
            if (!globalHaveException) {
                workbookMap.add(new SimpleRow(rowNum, rowSimpleCellMap));
            }
        }
    }

    /**
     * 检查上传的excel列头和期望的列头是否一致
     *
     * @param title
     */
    protected void checkTitle(List<Title> title) {

    }

    /**
     * 验证每行的一个cell
     *
     * @param sTitle            上传excel列头名称
     * @param simpleCell        该单元格的值
     * @param stencilTitleIndex 模板下标
     * @例子 <pre>
     * if("时间".equals(sTitle)){
     * 	if( !simpleCell.getFormattedValue().matches("^[0-9]+$") ){
     * 		error.putOtherError(sTitle, "只能是数字");
     * 	}
     * }
     * 	<pre>
     */
    protected void cellFilter(String sTitle, int stencilTitleIndex, SimpleCell simpleCell) {
    }

    public Map<String, String> simple2Map(Map<String, SimpleCell> simpleCells) {
        Map<String, String> temp = new HashMap<String, String>();
        for (Title sTitle : cExcelTitle) {
            SimpleCell simp = simpleCells.get(sTitle.getName());
            temp.put(sTitle.getName(), simp == null ? "" : simp.getFormattedValue());
        }
        return temp;
    }

    /**
     * 是不能保证返回全部的单元格的，如果单元格为空XML文件是不会记录该值单元格的
     */
    @Override
    public void cell(String cellReference, String cellValue, String formattedValue, XSSFSheetXMLHandler.xssfDataType nextDataType, String formatString, int rowNum) {
        int cellIndex = getCellIndex(cellReference) - 1;
        /**
         * 读取单元格的时候
         * 1. 自从指定X和Y开始读取
         * 2. 读取body的时候是以列头为主（当列头缺失的时候，此列就相当于忽略）
         */
        if (startTitleY > rowNum || startTitleX > cellIndex) {
            return;
        }
        //读取列头
        if (startTitleY == rowNum) {
            cExcelTitle.add(new Title(cellIndex, formattedValue));
        } else {//body
            String _title = getTitleByIndex(cellIndex);
            rowSimpleCellMap.put(_title, new SimpleCell(cellReference, nextDataType, formattedValue, formatString, cellIndex, cellValue));
        }
    }

    /**
     * 用列头的下标来获取列头名称
     *
     * @param cellIndex
     * @return
     */
    public String getTitleByIndex(int cellIndex) {
        if (this.cIndexExcelTitle == null) {
            this.cIndexExcelTitle = new HashMap<String, String>();
            for (Title t : cExcelTitle) {
                if (cellIndex == t.getX()) {
                    this.cIndexExcelTitle.put(cellIndex + "", t.getName());
                }
            }
        }
        return cIndexExcelTitle.get(cellIndex + "");
    }


    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
    }

    /**
     * 调用此方法开始解析
     *
     * @throws Exception
     */
    public void parse() throws Exception {
        OPCPackage op = OPCPackage.open(filePath, PackageAccess.READ);
        this.xssfReader = new XSSFReader(op);
        this.sst = new ReadOnlySharedStringsTable(op);
        this.stylesTable = xssfReader.getStylesTable();

        XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(stylesTable, sst, this, new DataFormatter(), false);
        XMLReader xmlReader = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        xmlReader.setContentHandler(handler);

        //用sheet名称获取rid
        WorkbookXMLTable wbt = new WorkbookXMLTable(xssfReader);
        String sheetRId = wbt.getSheetRId(sheetName);
        if (null == sheetRId) {
            throw new NullPointerException("未能在上传的excel中找到sheet：" + sheetName);
        }

        xmlReader.parse(new InputSource(xssfReader.getSheet(sheetRId)));
    }


    //得到列索引，每一列c元素的r属性构成为字母加数字的形式，字母组合为列索引，数字组合为行索引，
    //如AB45,表示为第（A-A+1）*26+（B-A+1）*26列，45行
    public int getCellIndex(String rowStr) {
        rowStr = rowStr.replaceAll("[^A-Z]", "");
        byte[] rowAbc = rowStr.getBytes();
        int len = rowAbc.length;
        float num = 0;
        for (int i = 0; i < len; i++) {
            num += (rowAbc[i] - 'A' + 1) * Math.pow(26, len - i - 1);
        }
        return (int) num;
    }

    public static void main(String[] args) throws Exception {
        //LKA Raw Data Format
        ReadXSSFBase2 t = new ReadXSSFBase2("/Users/jerry/Mix/共享/test.xlsx", "1");

//		String title = "";
//		String[] titles = StringUtils.split(title, ",");
        List<String> tl = Arrays.asList("执行日期", "大区", "区域", "上级城市", "城市", "城市项目负责人", "联系方式", "系统名称");

//        t.stencilTitle.addAll(tl);

        t.parse();

        //	System.out.println(t.workbookMap.toString());
        System.out.println(t.improperException.errors.size());
        System.out.println(t.workbookMap.size());
        System.out.println(t.cIndexExcelTitle);


    }
}


class Title implements java.io.Serializable {
    private static final long serialVersionUID = 6207562280307461858L;

    private int x;
    private String name;

    public Title(int x, String name) {
        super();
        this.x = x;
        this.name = name;
    }

    public int getX() {
        return x;
    }

    public void setX(int x) {
        this.x = x;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "{x:" + this.x + ", name:" + this.name + "}";
    }
}
