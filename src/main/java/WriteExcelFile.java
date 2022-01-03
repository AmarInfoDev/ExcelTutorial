import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

public class WriteExcelFile {

    public static void main(String[] args) {
        try {

            //Create a workbook in .xlsx format
            //For .xsl workbooks use new HSSFWorkBook();
            Workbook workbook = new XSSFWorkbook();

            //Create a sheet
            Sheet sh = workbook.createSheet("Invoices");

            //Create a row
            //Index is 0 as the first row has an index 0
            Row titleRow = sh.createRow(0);

            Cell headerCell = titleRow.createCell(0);

            // Creating header font style
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 14);
            headerFont.setColor(IndexedColors.BROWN.index);

            //Create a CellStyle for the header
            //This denotes the style of the cell
            CellStyle headerStyle = workbook.createCellStyle();

            //Apply the font we created earlier
            headerStyle.setFont(headerFont);

            //Set the Background of the cell
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setWrapText(true);
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

            headerCell.setCellStyle(headerStyle);

            //merging header cells for a title
            sh.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

            //Assigning header value
            headerCell.setCellValue("Invoices");


            //Values for the top row with column headings
            String[] columnHeadings = {"Item id", "Item Name", "Qty", "Item Price", "Sold Date"};

            //A different font for the headings
            Font columnHeaderFont = workbook.createFont();
            columnHeaderFont.setBold(true);
            columnHeaderFont.setFontHeightInPoints((short) 12);
            columnHeaderFont.setColor(IndexedColors.BLACK.index);

            CellStyle columnHeaderStyle = workbook.createCellStyle();

            columnHeaderStyle.setFont(columnHeaderFont);

            columnHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            columnHeaderStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

            //Create the actual header row in the sheet
            Row tableHeaderRow = sh.createRow(1);

            //Iterate over the column headings to create columns
            for (int i = 0; i < columnHeadings.length; i++) {

                Cell cell = tableHeaderRow.createCell(i);  //create a cell at i in row 0
                cell.setCellValue(columnHeadings[i]);  // set the value of cell from columnHeadings
                cell.setCellStyle(columnHeaderStyle);   //apply the cell style we created earlier

            }

            //freezing header and table header rows
            sh.createFreezePane(0, 1);

            //Now we fill the data
            //for this we create a helper method which returns an arraylist of Invoices
            //this data will normally probably be received from a database query
            ArrayList<Invoices> a = createData();

            //Workbook provides a creationHelper for setting date format, hyperlinks, etc
            CreationHelper creationHelper = workbook.getCreationHelper();

            //This cellStyle will be attached to cells with dates
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("MM/dd/yyyy"));

            int rowNum = 2;    // row 0 is the header
            for (Invoices i : a) {

                //for each invoice we create a row
                Row row = sh.createRow(rowNum++);

                //create a cell for each column and pass the respective value
                row.createCell(0).setCellValue(i.getItemId());
                row.createCell(1).setCellValue(i.getItemName());
                row.createCell(2).setCellValue(i.getItemQty());
                row.createCell(3).setCellValue(i.getTotalPrice());

                Cell dateCell = row.createCell(4);
                dateCell.setCellValue(i.getItemSoldDate());
                dateCell.setCellStyle(dateStyle);

            }

            //Now to autosize the column
            for (int x = 0; x < columnHeadings.length; x++) {
                //iterate over all the columns and call sheet.autosizeColumn
                sh.autoSizeColumn(x);
            }

            Sheet sh2 = workbook.createSheet("Second"); // a new sheet if needed

            //Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(
                    "C:\\Users\\ACER.LAPTOP-CFI1FCQ8\\Desktop\\Trainee tutorial\\.Invoices.xlsx");
            workbook.write(fileOut);

            //Finally we close the workbook and the stream
            fileOut.close();
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //helper method to get hard-coded data
    private static ArrayList<Invoices> createData() throws ParseException {

        ArrayList<Invoices> a = new ArrayList();
        a.add(new Invoices(1, "Mouse", 2, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2021")));
        a.add(new Invoices(2, "Keyboard", 1, 140.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2021")));
        a.add(new Invoices(3, "Charger", 22, 120.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2021")));
        a.add(new Invoices(4, "Monitor", 40, 110.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/03/2021")));
        a.add(new Invoices(5, "Keyboard", 10, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/04/2021")));
        a.add(new Invoices(6, "Mouse", 60, 60.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/04/2021")));
        a.add(new Invoices(7, "Charger", 1, 160.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2021")));
        a.add(new Invoices(8, "Monitor", 5, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2021")));

        return a;

    }

}
