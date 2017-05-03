import com.sun.rowset.internal.Row;
import javafx.scene.control.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Created by maedeh.hesabgar on 03/05/2017.
 */
public class SearchWebsite {

    public static void main(String[] args){
        try {
//            File file = new File("src.");
//            for(String fileNames : file.list()) System.out.println(fileNames);
            File myFile = new File("test.xlsx");
            FileInputStream fileInputStream = new FileInputStream(myFile);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheetAt(0);
            XSSFRow row1 = worksheet.getRow(0);
            XSSFCell cellA1 = row1.getCell((short) 0);
            String a1Val = cellA1.getStringCellValue();
            XSSFCell cellB1 = row1.getCell((short) 1);
            String b1Val = cellB1.getStringCellValue();
            XSSFCell cellC1 = row1.getCell((short) 2);
            boolean c1Val = cellC1.getBooleanCellValue();
            XSSFCell cellD1 = row1.getCell((short) 3);
            Date d1Val = cellD1.getDateCellValue();

            System.out.println("A1: " + a1Val);
            System.out.println("B1: " + b1Val);
            System.out.println("C1: " + c1Val);
            System.out.println("D1: " + d1Val);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
