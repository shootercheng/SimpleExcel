package com.example;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static com.example.ExcelXmlConstant.POSITION;
import static com.example.ExcelXmlConstant.ROW_TAG;

//@Slf4j
@Setter
@Getter
public class Excel07SaxReader extends AbstractExcelSaxReader<Excel07SaxReader> implements ContentHandler {

    // 填充字符串
    public static final char CELL_FILL_CHAR = '@';
    // 列的最大位数
    public static final int MAX_CELL_BIT = 3;
    // 行的最大列坐标
    private String maxColCoordinate;

    private static final String class_saxparser = "org.apache.xerces.parsers.SAXParser";

    private static final String RID_PREFIX = "rId";
    //行元素
    private static final String row = "row";
    //单元格
    private static final String C_CONTENT = "c";
    //单元格类型
    private static final String T_ATTR_VALUE = "t";
    //s属性
    private static final String S_ATTR_VALUE = "s";
    //单元格元素值或引用
    private static final String v_CONTENT="v";
    //每一列的元素内容
    List<String> rowList = new ArrayList<>();
    //当前工作薄
    private Integer sheetIndex;
    //当前行
    private Integer curRow;
    // 当前列
    private Integer curCol;
    //单元格格式
    private CellDataType nextDataType = CellDataType.SSTINDEX;
    // excel 2007 的共享字符串表,对应sharedString.xml
    private SharedStringsTable sharedStringsTable;
    // 单元格的格式表，对应style.xml
    private StylesTable stylesTable;

    private String lastContent;
    //单元格存储的索引
    private short formatIndex;
    //单元格存储的格式化字符串
    private String formatString;
    // 当前列的坐标
    private String curCoordinate;
    // 前一个列的坐标 处理空白单元格
    private String preCoordinate;
    private final DataFormatter formatter = new DataFormatter();
    private List<String> refNames = new ArrayList<>();
    //回调接口 回调行数据
    private RowDataHandler rowDatHandler;


    public Excel07SaxReader(RowDataHandler rowDataHandler){
        this.rowDatHandler= rowDataHandler;
    }

    @Override
    public Excel07SaxReader read(InputStream in, Integer relId) throws OpenXML4JException, IOException {
        return read(OPCPackage.open(in), relId);
    }

    @Override
    public void setDocumentLocator(Locator locator) {
    }

    @Override
    public void startDocument() throws SAXException {
        //  log.info("startDocument");
    }

    @Override
    public void endDocument() throws SAXException {
        //  log.info("endDocument");
    }

    @Override
    public void startPrefixMapping(String prefix, String uri) throws SAXException {
        //  log.info("startPrefixMapping");
    }

    @Override
    public void endPrefixMapping(String prefix) throws SAXException {
        // log.info("endPrefixMapping");
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes atts) throws SAXException {
//        log.info("startElement:{}", qName);

        // row
        if (ROW_TAG.equals(qName)) {
            curRow = PositionUtils.getRowByRowTagt(atts.getValue(POSITION), curRow);
        }

        if (C_CONTENT.equals(qName)) {

            String tempCurCoordinate = atts.getValue("r");//坐标
            //  log.info("tempCurCoordinate:{}", tempCurCoordinate);
            if (preCoordinate == null) {
                preCoordinate = String.valueOf(CELL_FILL_CHAR);
            } else {
                preCoordinate = curCoordinate;

            }
            curCoordinate = tempCurCoordinate;
            if (curCol > 0 && curRow > 0 && !refNames.contains(v_CONTENT)) {
                rowList.add(curCol, "");
                curCol++;
            }
            refNames = new ArrayList<>();

            setNextDataType(atts);

        }
        lastContent = "";

    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
//        log.info("endElement:{}", qName);
        final String contentStr = lastContent.trim();

        if (curRow > 0) {
            refNames.add(qName);
        }

        if (qName.equals(v_CONTENT)) {
            String value = getDataValue(contentStr, "");
            if (rowList.size() == 0) {
                int len = countNullCell(curCoordinate, "A" + curRow + 1);
                for (int i = 0; i < len + 1; i++) {
                    rowList.add(curCol, "");
                    curCol++;
                }
            }
            //补全单元格之间的空单元格
            if (!curCoordinate.equals(preCoordinate)) {
                int len = countNullCell(curCoordinate, preCoordinate);
                for (int i = 0; i < len; i++) {
                    rowList.add(curCol, "");
                    curCol++;
                }
            }
            rowList.add(value);
            curCol++;
        } else {
            //如果标签名称为 row，这说明已到行尾
            if (qName.equals("row")) {
                //默认第一行为表头，以该行单元格数目为最大数目
                if (curRow == 0) {
                    maxColCoordinate = curCoordinate;
                }
                //补全一行尾部可能缺失的单元格
                if (maxColCoordinate != null) {
                    int len = countNullCell(maxColCoordinate, curCoordinate);
                    for (int i = 0; i <= len; i++) {
                        rowList.add(curCol, "");
                        curCol++;
                    }
                }
                rowDatHandler.handle(sheetIndex,curRow,rowList);
                rowList.clear();
//                curRow++;
                curCol = 0;
                // 置空当前列坐标和前一列坐标
                preCoordinate = null;
                curCoordinate = null;
            }
        }


    }

    /**
     * 字符串的填充
     *
     * @param str
     * @param len
     * @param let
     * @param isPre
     * @return
     */
    String fillChar(String str, int len, char let, boolean isPre) {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < (len - len_1); i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < (len - len_1); i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }

    public int countNullCell(String ref, String preRef) {
        //excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = preRef.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    /**
     * 根据数据类型获取数据
     * @param value
     * @param thisStr
     * @return
     */
    public String getDataValue(String value, String thisStr)

    {
        switch (nextDataType)
        {
            //这几个的顺序不能随便交换，交换了很可能会导致数据错误
            case BOOL:
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR:
                thisStr = "\"ERROR:" + value.toString() + '"';
                break;
            case FORMULA:
                thisStr = '"' + value.toString() + '"';
                break;
            case INLINESTR:

                XSSFRichTextString rtsi = new XSSFRichTextString(value);
                thisStr = rtsi.toString();
                rtsi = null;
                break;
            case SSTINDEX:
                final Integer idx =Integer.parseInt(value);
                XSSFRichTextString xssf = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                thisStr = xssf.toString();
                xssf=null;
                break;
            case NUMBER:
                if (formatString != null){
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                }else{
                    thisStr = value;
                }
                thisStr = thisStr.replace("_", "").trim();
                break;
            case DATE:
                try{
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                }catch(NumberFormatException ex){
                    thisStr = value.toString();
                }
                thisStr = thisStr.replace(" ", "");
                break;
            default:
                thisStr = "";
                break;
        }
        return thisStr;
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        lastContent = lastContent.concat(new String(ch, start, length));
    }

    @Override
    public void ignorableWhitespace(char[] ch, int start, int length) throws SAXException {
        // log.info("ignorableWhitespace");
    }

    @Override
    public void processingInstruction(String target, String data) throws SAXException {
        //  log.info("processingInstruction");
    }

    @Override
    public void skippedEntity(String name) throws SAXException {
        // log.info("skippedEntity");
    }

    private Excel07SaxReader read(OPCPackage pkg, Integer relId) throws IOException, OpenXML4JException {
        InputStream sheetInputStream;
        final XSSFReader xssfReader = new XSSFReader(pkg);

        //共享字符串表
        sharedStringsTable = xssfReader.getSharedStringsTable();
        stylesTable = xssfReader.getStylesTable();
        if (relId > -1) {
            this.sheetIndex = relId;
            this.curRow = 0;
            this.curCol = 0;
            sheetInputStream = xssfReader.getSheet(RID_PREFIX + (relId + 1));
            parse(sheetInputStream);

        } else {
            this.sheetIndex = -1;
            Iterator<InputStream> sheetsData = xssfReader.getSheetsData();
            while (sheetsData.hasNext()) {
                this.curRow = 0;
                this.curCol = 0;
                this.sheetIndex++;
                sheetInputStream = sheetsData.next();
                parse(sheetInputStream);
            }

        }

        return this;
    }

    enum CellDataType{
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }
    /**
     * 根据element属性设置数据类型
     * @param attributes
     */
    public void setNextDataType(Attributes attributes){

        nextDataType = CellDataType.NUMBER;
        formatIndex = -1;
        formatString = null;
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");
        if ("b".equals(cellType)){
            nextDataType = CellDataType.BOOL;
        }else if ("e".equals(cellType)){
            nextDataType = CellDataType.ERROR;
        }else if ("inlineStr".equals(cellType)){
            nextDataType = CellDataType.INLINESTR;
        }else if ("s".equals(cellType)){
            nextDataType = CellDataType.SSTINDEX;
        }else if ("str".equals(cellType)){
            nextDataType = CellDataType.FORMULA;
        }
        if (cellStyleStr != null){
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();
            if ("m/d/yy".equals(formatString)){
                nextDataType = CellDataType.DATE;
                //full format is "yyyy-MM-dd hh:mm:ss.SSS";
                formatString = "yyyy-MM-dd";
            }
            if (formatString == null){
                nextDataType = CellDataType.NULL;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }
    }

    private void parse(InputStream sheetInputStream) {
        try {
            fetchSheetReader().parse(new InputSource(sheetInputStream));
        } catch (IOException | SAXException e) {
            e.printStackTrace();
        } finally {
            try {
                sheetInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private XMLReader fetchSheetReader() throws SAXException {
        XMLReader xmlReader = XMLReaderFactory.createXMLReader(class_saxparser);
        xmlReader.setContentHandler(this);
        return xmlReader;
    }

}
