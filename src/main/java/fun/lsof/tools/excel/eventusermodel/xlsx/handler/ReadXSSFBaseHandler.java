package fun.lsof.tools.excel.eventusermodel.xlsx.handler;


import fun.lsof.tools.excel.eventusermodel.xlsx.exception.ReadExcelException;
import fun.lsof.tools.excel.eventusermodel.xlsx.exception.XlsContentImproperException;
import fun.lsof.tools.excel.eventusermodel.xlsx.vo.SimpleCell;
import fun.lsof.tools.excel.eventusermodel.xlsx.vo.SimpleRow;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.util.*;

public class ReadXSSFBaseHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private final static Logger LOGGER = Logger.getLogger(ReadXSSFBaseHandler.class);

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
     * 是保存在数据里面的模板title
     */
    protected List<String> stencilTitle = new ArrayList<String>();
    /**
     * 当前解析的excel列头
     */
    protected List<String> title = new ArrayList<String>();
    /**
     * 带有上传excel列头下标的title
     */
    protected Map<String, String> indexTitle = new HashMap<String, String>();
    /**
     * excel中有效的行数，忽略空行
     */
    protected int validRowSum = 0;

    String sheetName = null;
    String filePath = null;

    public ReadXSSFBaseHandler(String path, String sheetName) {
        this.filePath = path;
        this.sheetName = sheetName;
    }

    public void setStencilTitle(List<String> stencilTitle) {
        this.stencilTitle = stencilTitle;
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

    /**
     * 处理表头（判断数据库存放的列头是否和读取到的列头是一致的）
     */
    protected void initTitle() {

        if (null == stencilTitle || stencilTitle.isEmpty()) {
            throw new ReadExcelException("未输入保存在数据库中的模板列头");
        }

        //在数据库中存放的列头都能在上传的excel中找到
        if (title.containsAll(stencilTitle)) {
            String[] cloneTitle = title.toArray(new String[]{});
            for (String pTitle : stencilTitle) {
                int index = ArrayUtils.indexOf(cloneTitle, pTitle);
                indexTitle.put(index + "", pTitle);
            }
        } else {
            stencilTitle.removeAll(title);
            throw new ReadExcelException("在上传的excel中未找到title：{"+stencilTitle.toString()+"}");
        }
    }


    @Override
    public void startRow(int rowNum) {
        validRowSum++;
        rowSimpleCellMap = new HashMap<String, SimpleCell>();
    }

    @Override
    public void endRow(int rowNum) {
        error = new XlsContentImproperException.Entry();

        if (rowNum != 0 && title.isEmpty()) {
            throw new ReadExcelException("未在上传的excel中的首行找到列头");
        }
        //处理表头
        if (0 == rowNum) {
            initTitle();
            return;
        }

        //验证
        for (int i = 0; i < stencilTitle.size(); i++) {
            String sTitle = stencilTitle.get(i);
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
     * 验证每行的一个cell
     *
     * @param sTitle            模板头名称
     * @param simpleCell        该单元格的值
     * @param stencilTitleIndex 模板下标
     * @例子 if("时间".equals(sTitle)){<br>
     * <PRE>&#9;if( !simpleCell.getFormattedValue().matches("^[0-9]+$") ){</PRE>
     * <PRE>&#9;&#9;error.putOtherError(sTitle, "只能是数字");</PRE>
     * <PRE>&#9;}</PRE>
     * }<br>
     */
    protected void cellFilter(String sTitle, int stencilTitleIndex, SimpleCell simpleCell) {
    }

    public Map<String, String> simple2Map(Map<String, SimpleCell> simpleCells) {
        Map<String, String> temp = new HashMap<String, String>();
        for (String sTitle : stencilTitle) {
            SimpleCell simp = simpleCells.get(sTitle);
            temp.put(sTitle, simp == null ? "" : simp.getFormattedValue());
        }
        return temp;
    }

    /**
     * 是不能保证返回全部的单元格的，如果单元格为空XML文件是不会记录该值单元格的
     */
    @Override
    public void cell(String cellReference, String cellValue, String formattedValue, XSSFSheetXMLHandler.xssfDataType nextDataType, String formatString, int rowNum) {
        if (0 == rowNum) {
            title.add(formattedValue);
        } else {
            int cellIndex = getCellIndex(cellReference) - 1;
            String title = indexTitle.get(cellIndex + "");
            if (!org.apache.commons.lang.StringUtils.isBlank(title)) {
                rowSimpleCellMap.put(title, new SimpleCell(cellReference, nextDataType, formattedValue, formatString, cellIndex, cellValue));
            }
        }
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
            throw new ReadExcelException("未能在上传的excel中找到sheet：{"+sheetName+"}");
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
        ReadXSSFBaseHandler t = new ReadXSSFBaseHandler("/Users/jerry/Mix/temp/CPAT-BW Query Interface.xlsx", "sheet");
//		String title = "";
//		String[] titles = StringUtils.split(title, ",");
        List<String> tl = Arrays.asList("L00 Responsibility Area", "L01 Responsibility Area", "L01 Posting period", "L00 Material", "L01 Material", "L01 Material Key", "L01 Material Product (Key)");

        t.stencilTitle.addAll(tl);


        t.parse();

        //	System.out.println(t.workbookMap.toString());
        System.out.println(t.improperException.errors.size());
        System.out.println(t.workbookMap.size());
    }
}
