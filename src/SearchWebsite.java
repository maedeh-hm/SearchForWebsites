
import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import javax.net.ssl.HttpsURLConnection;


/**
 * Created by maedeh.hesabgar on 03/05/2017.
 */
public class SearchWebsite {

    public static void main(String[] args){
        try {
            File file = new File("src.");
            for(String fileNames : file.list()) System.out.println(fileNames);
            File myFile = new File("test.xlsx");
            FileInputStream fileInputStream = new FileInputStream(myFile);
            XSSFWorkbook readingWorkbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet ReadingWorksheet = readingWorkbook.getSheetAt(0);

            Iterator rows = ReadingWorksheet.rowIterator();
            String c1Val = null;

            //writeIntoFile
            FileOutputStream fileOut = new FileOutputStream("poi-test.xlsx");
            XSSFWorkbook writingWorkbook = new XSSFWorkbook();
            XSSFSheet writingWorksheet = writingWorkbook.createSheet("POI Worksheet");

            int writingRowIndex = 0;


            while(rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                XSSFCell cellC1 = row.getCell((short) 3);
                c1Val = cellC1.getStringCellValue();
                System.out.println("C1: " + c1Val);



                //************************************************************************
                String key = "AIzaSyCbi64egDqtc-Ir1qd05Jnt8DWU-_iQlHE";
                String qry = URLEncoder.encode(c1Val + " -filetype:pdf", "UTF-8");
                URL url = null;

                url = new URL(
                        "https://www.googleapis.com/customsearch/v1?key=" + key + "&cx=005119725969692011056:lmlrdefkooa&q=" + qry + "&alt=json");

                System.out.println(url);
                HttpsURLConnection connection = null;
                connection = (HttpsURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("Accept", "application/json");
                connection.setDoInput(true);
                connection.setDoOutput(true);
                BufferedReader br = null;
                br = new BufferedReader(new InputStreamReader((connection.getInputStream())));
                String output;
                System.out.println("Output from Server .... \n");
                ArrayList<String> profileNameResults = new ArrayList<>();
                while ((output = br.readLine()) != null) {


                    if (output.contains("\"link\": \"")) {
                        String link = output.substring(output.indexOf("\"link\": \"") + ("\"link\": \"").length(), output.indexOf("\","));
                        System.out.println(link);       //Will print the google search links
                        profileNameResults.add(link);
                        break;
                    }
                }

                connection.disconnect();
                //**************************************************************************


                //writing to an excel file
                Row WritingRow = writingWorksheet.createRow(writingRowIndex++);

                XSSFCell cellA1 = (XSSFCell) WritingRow.createCell(0);
                cellA1.setCellValue(c1Val);
                XSSFCellStyle cellStyle = writingWorkbook.createCellStyle();
                cellA1.setCellStyle(cellStyle);

                int writingCellIndex = 1;
                int j = 0;
                for(String s:profileNameResults){
                    if(j<5) {
                        XSSFCell cellB1 = (XSSFCell) WritingRow.createCell(writingCellIndex++);
                        cellB1.setCellValue(s);
                        cellStyle = writingWorkbook.createCellStyle();
                        cellB1.setCellStyle(cellStyle);
                        j++;
                    }else{
                        break;
                    }
                }
            }

            writingWorkbook.write(fileOut);
            fileOut.flush();
            fileOut.close();

        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


}
