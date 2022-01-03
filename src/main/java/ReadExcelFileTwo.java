import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcelFileTwo {

    public static void main(String[] args) {

        String value = null; //variable we will be storing the cell value in.

        Workbook wb = null;   //initialize Workbook as null

        try {
            //reading the data from a file
            FileInputStream file = new FileInputStream("C:\\Ages.xlsx");

            //A workbook for the file
            wb = new XSSFWorkbook(file);

        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        Sheet sheet = wb.getSheetAt(0); //getting the XSSFSheet

        Row row = sheet.getRow(3); //returns the 3rd row
        Cell cell = row.getCell(2);  //returns the 2nd cell in the 3rd row
        value = cell.getStringCellValue();  //the value in the cell

        //Printing the value
        System.out.println("The height of Akriti is: " + value);

    }
}
