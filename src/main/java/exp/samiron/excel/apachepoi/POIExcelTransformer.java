package exp.samiron.excel.apachepoi;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.SAXException;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by samiron on 6/24/2017.
 */
public class POIExcelTransformer {

    private POIEventModelReader1 excelReader = null;
    private SXSSFWorkbook outputExcel = null;
    private SXSSFSheet currentSheet;
    private String outputFilePath = null;
	private Pattern cellReferencePattern;

    public static void main(String[] args) throws OpenXML4JException, SAXException, IOException {
    	if(args.length < 1){
    		System.err.println("First parameter should contain the absolute path of input xlsx file.");
    		System.exit(1);
    	}
    	
        String inputfile = args[0];
        String outfile = inputfile.replace(".xlsx", "_resolved.xlsx");
        System.out.println(inputfile + ", " + outfile);

        POIExcelTransformer poiHybrid = new POIExcelTransformer();
        poiHybrid.initReader(inputfile);
        poiHybrid.initWriter(outfile);
        poiHybrid.convert();
        System.out.println("Done. Transformed file is saved as: " + outfile);
    }
    
    public POIExcelTransformer(){
    	this.cellReferencePattern = Pattern.compile("([A-Z]+)(\\d+)");
	}

    private void convert() throws IOException, SAXException, OpenXML4JException {
        this.excelReader.startRead();
    }

    private void initWriter(String outFile) {
        this.outputExcel = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        this.outputFilePath = outFile;
    }

    private void initReader(String inputFile){
        this.excelReader = new POIEventModelReader1(inputFile);
        this.excelReader.setCellValueListner(new POIEventModelReader1.CellValueListener() {
            public void cellValueFound(String reference, String value) {
                POIExcelTransformer.this.writeCellValue(reference, value);
            }
            public void parseFinished() {
                POIExcelTransformer.this.parseFinished();
            }

            public void sheetNameFound(int sheetIndex, String sheetName) {
                POIExcelTransformer.this.sheetNameFound(sheetIndex, sheetName);
            }
        });
    }

    private void sheetNameFound(int sheetIndex, String sheetName) {
        this.currentSheet = this.outputExcel.createSheet();
        this.outputExcel.setSheetName(sheetIndex, sheetName);
    }

    private void parseFinished()  {
        try {
            FileOutputStream out = new FileOutputStream(this.outputFilePath);
            this.outputExcel.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        POIExcelTransformer.this.outputExcel.dispose();
    }

    private void writeCellValue(String reference, String value) {
        Matcher m = this.cellReferencePattern.matcher(reference.trim().toUpperCase());
        if(m.matches()){
            int colnum = this.columnNameToNumber(m.group(1));
            int row = Integer.parseInt(reference.substring(1)) - 1;

            Row r = this.currentSheet.getRow(row);
            if(r == null){
                r = this.currentSheet.createRow(row);
            }
            Cell cell = r.createCell(colnum-1);
            cell.setCellValue(value);
        }
    }

    private int columnNameToNumber(String reference) {
        int l = reference.length();
        int colNum = 0;
        for(char c : reference.toCharArray()){
            l--;
            colNum += ((int)Math.pow(26,l)) * (c - 'A' + 1);
        }
        return colNum;
    }
}
