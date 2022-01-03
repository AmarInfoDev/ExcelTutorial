import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcelFile {

    public static void main(String[] args) throws IOException {

        //Reading the file
        FileInputStream file = new FileInputStream(new File("C:\\Ages.xlsx"));

        //creating workbook instance that refers to .xls file
        Workbook workbook = new XSSFWorkbook(file);

        //Creating a sheet of the file to retrieve the data
        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> sheetData = new HashMap<Integer, List<String>>();

        int i = 0;

        for (Row row : sheet) {            //Iterating through each row
            sheetData.put(i, new ArrayList<String>());
            for (Cell cell : row) {                  //Iterating through each cell
                switch (cell.getCellType()) {

                    //Apache POI has different methods for reading each type of data
                    //When the cell type enum value is STRING, the content will be
                    //read using the getRichStringCellValue() method of Cell interface:
                    case STRING:
                        sheetData.get(new Integer(i)).add(cell.getRichStringCellValue().getString());
                        System.out.print(cell.getRichStringCellValue() + "\t\t");
                        break;

                    //Cells having the NUMERIC content type can contain either a date or a number
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            sheetData.get(i).add(cell.getDateCellValue() + "");
                            System.out.print(cell.getDateCellValue() + "\t\t");
                        } else {
                            sheetData.get(i).add(cell.getNumericCellValue() + "");
                            System.out.print(cell.getNumericCellValue() + "\t\t");
                        }
                        break;

                    //For BOOLEAN values
                    case BOOLEAN:
                        sheetData.get(i).add(cell.getBooleanCellValue() + "");
                        System.out.print(cell.getBooleanCellValue() + "\t\t");
                        break;

                    //For when the cell type is FORMULA
                    case FORMULA:
                        sheetData.get(i).add(cell.getCellFormula() + "");
                        System.out.println();
                        break;

                    default:
                        sheetData.get(new Integer(i)).add("");


                }
            }
            System.out.println();
        }
        i++;
    }


}
