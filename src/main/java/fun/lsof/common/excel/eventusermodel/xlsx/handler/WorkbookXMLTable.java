package fun.lsof.common.excel.eventusermodel.xlsx.handler;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class WorkbookXMLTable extends DefaultHandler {
    private Map<String, Sheet> sheets = new HashMap<String, Sheet>();

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if ("sheet".equals(qName)) {
            sheets.put(attributes.getValue("name"),
                    new Sheet(attributes.getValue("name"), attributes.getValue("sheetId"), attributes.getValue("r:id")));
        }
    }

    public WorkbookXMLTable(XSSFReader xssfReader) throws Exception {
        XMLReader xmlReader = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        xmlReader.setContentHandler(this);
        InputStream inp = xssfReader.getWorkbookData();
        xmlReader.parse(new InputSource(inp));

        inp.close();
    }


    public String getSheetRId(String sheetName) {
        Sheet sheet = sheets.get(sheetName);
        return null != sheet ? sheet.rId : null;
    }


    public static final class Sheet {
        public String name;
        public String sheetId;
        public String rId;

        public Sheet(String name, String sheetId, String rId) {
            super();
            this.name = name;
            this.sheetId = sheetId;
            this.rId = rId;
        }
    }
}
