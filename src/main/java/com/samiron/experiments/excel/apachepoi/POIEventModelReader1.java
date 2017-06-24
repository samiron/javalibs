package com.samiron.experiments.excel.apachepoi;
/**
 * Created by samir on 6/23/2017.
 */
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class POIEventModelReader1 {

    private CellValueListener cellValueListener = null;
    private String fileName;
    public POIEventModelReader1(String fileName){
        this.fileName = fileName;
    }

    public void setCellValueListner(CellValueListener listener){
        this.cellValueListener = listener;
    }

    public void startRead() throws OpenXML4JException, IOException, SAXException {
        OPCPackage pkg = OPCPackage.open(this.fileName);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
        while(sheets.hasNext()) {
            System.out.println("Processing new sheet:\n");
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
            System.out.println("");
        }
        this.cellValueListener.parseFinished();
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser =
                XMLReaderFactory.createXMLReader();
        ContentHandler handler = new SheetHandler(sst, this.cellValueListener);
        parser.setContentHandler(handler);
        return parser;
    }

    public interface CellValueListener {
        void cellValueFound(String reference, String value);

        void parseFinished();
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    public static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private CellValueListener cellValueListener = null;
        private String currentCellReference;
        private boolean handlingFormula;
        private String formula = "";

        private SheetHandler(SharedStringsTable sst, CellValueListener cellValueListener) {
            this.sst = sst;
            this.cellValueListener = cellValueListener;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            //System.out.println(localName);
            // c => cell
            if(name.equals("c")) {
                // Print the cell reference
                //System.out.print(attributes.getValue("r") + " - ");
                this.currentCellReference = attributes.getValue("r");
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                if(cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            } else if( name.equals("f")){
                // f => formula
                this.handlingFormula = true;
                this.formula = "";
            }
            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name)
                throws SAXException {
            //System.out.println(localName);
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v")) {
                if(this.cellValueListener != null){
                    System.out.println(String.format("Emitting %s -> %s", this.currentCellReference, this.lastContents));
                    this.cellValueListener.cellValueFound(this.currentCellReference, this.lastContents);
                }else{
                    System.out.println(String.format("Printing %s -> %s", this.currentCellReference, this.lastContents));
                }
            } else if(name.equals("c")){
                this.currentCellReference = null;
            } else if (name.equals("f")){
                this.handlingFormula = false;
                System.out.println("Formula: " + this.formula);
            }
        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
            if(!this.handlingFormula) {
                lastContents += new String(ch, start, length);
            }else{
                this.formula += new  String(ch, start, length);
            }
        }
    }

    public static void main(String[] args) throws Exception {
        String filepath = "e:\\sxssf_formula.xlsx";
        POIEventModelReader1 example = new POIEventModelReader1(filepath);
        example.startRead();
    }

}
