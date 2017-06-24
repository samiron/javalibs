package com.samiron.experiments.excel.apachepoi;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.SAXException;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by samir on 6/24/2017.
 */
public class POIHybrid {

    private POIEventModelReader1 excelReader = null;
    private SXSSFWorkbook outputExcel = null;
    private SXSSFSheet currentSheet;
    private String outputFilePath = null;

    public static void main(String[] args) throws OpenXML4JException, SAXException, IOException {
        String inputfile = "e:\\sxssf_formula.xlsx";
        String outfile = "e:\\sxssf_formula_resolved.xlsx";

        POIHybrid poiHybrid = new POIHybrid();
        poiHybrid.initReader(inputfile);
        poiHybrid.initWriter(outfile);
        poiHybrid.convert();
    }

    private void convert() throws IOException, SAXException, OpenXML4JException {
        this.excelReader.startRead();
    }

    private void initWriter(String outFile) {
        this.outputExcel = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        this.currentSheet = this.outputExcel.createSheet();
        this.outputFilePath = outFile;

    }

    private void initReader(String inputFile){
        this.excelReader = new POIEventModelReader1(inputFile);
        this.excelReader.setCellValueListner(new POIEventModelReader1.CellValueListener() {
            public void cellValueFound(String reference, String value) {
                POIHybrid.this.writeCellValue(reference, value);
            }

            public void parseFinished() {
                POIHybrid.this.parseFinished();
            }
        });
    }

    private void parseFinished()  {
        try {
            FileOutputStream out = new FileOutputStream(this.outputFilePath);
            this.outputExcel.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        POIHybrid.this.outputExcel.dispose();
    }

    private void writeCellValue(String reference, String value) {
        int col = reference.charAt(0) - 'A';
        int row = Integer.parseInt(reference.substring(1)) - 1;
        Row r = this.currentSheet.getRow(row);
        if(r == null){
            r = this.currentSheet.createRow(row);
        }
        Cell cell = r.createCell(col);
        cell.setCellValue(value);
    }
}
