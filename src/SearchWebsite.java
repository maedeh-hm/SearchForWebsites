
import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.net.www.protocol.http.HttpURLConnection;

import javax.net.ssl.HttpsURLConnection;


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


        String key="AIzaSyCbi64egDqtc-Ir1qd05Jnt8DWU-_iQlHE";
        String qry="salam";
        URL url = null;
        try {
            url = new URL(
                    "https://www.googleapis.com/customsearch/v1?key="+key+ "&cx=005119725969692011056:lmlrdefkooa&q="+ qry + "&alt=json");
            HttpsURLConnection conn = null;
            conn = (HttpsURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Accept", "application/json");
            BufferedReader br = null;
            br = new BufferedReader(new InputStreamReader((conn.getInputStream())));
            String output;
            System.out.println("Output from Server .... \n");
            while ((output = br.readLine()) != null) {

                if(output.contains("\"link\": \"")){
                    String link=output.substring(output.indexOf("\"link\": \"")+("\"link\": \"").length(), output.indexOf("\","));
                    System.out.println(link);       //Will print the google search links
                }
            }

            conn.disconnect();

        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
