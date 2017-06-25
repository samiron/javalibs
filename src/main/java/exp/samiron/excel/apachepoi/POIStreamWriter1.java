package exp.samiron.excel.apachepoi;
/**
 * Created by samir on 6/23/2017.
 */
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;

public class POIStreamWriter1 {

    public static void main(String[] args) throws Throwable {
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        int cols[] = {1, 26, 27, 52, 53, 78, 677, 702};
        Sheet sh = wb.createSheet("Mysheet");
        for(int c : cols){
            Row row = sh.getRow(1);
            if(row == null){
                row = sh.createRow(1);
            }
            Cell cell = row.createCell(c-1);
            String address = new CellReference(cell).formatAsString();
            cell.setCellValue(address);
        }

//        for(int rownum = 0; rownum < 1000; rownum++){
//            Row row = sh.createRow(rownum);
//            for(int cellnum = 0; cellnum < 10; cellnum++){
//                Cell cell = row.createCell(cellnum);
//                String address = new CellReference(cell).formatAsString();
//                cell.setCellValue(address);
//            }
//
//        }

        // Rows with rownum < 900 are flushed and not accessible
        for(int rownum = 0; rownum < 900; rownum++){
        }

        // ther last 100 rows are still in memory
        for(int rownum = 900; rownum < 1000; rownum++){
        }

        FileOutputStream out = new FileOutputStream("e:\\sxssf_custom.xlsx");
        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }

}
