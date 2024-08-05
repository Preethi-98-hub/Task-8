package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelworkbook {

public static void main(String[] args) {
   // TODO Auto-generated method stub
   //create an object of XSSF workbook
		
   try(XSSFWorkbook workbook=new XSSFWorkbook()){
   //create sheet
   XSSFSheet sheet=workbook.createSheet("Sheet1");  
			
   //Create an Array of Objects
   Object[][] data= {
   {"Name","Age", "Email"},
   {"John Doe", 30, "john@test.com"},
   {"John Doe", 28, "john@test.com"},
   {"Bob smith", 35, "jacky@example.com"},
   {"Swapnil", 37, "swapnil@example.com"},
    };
			
   //To write the data into Excel
   int rowNum=0;
   for(Object[] rowdata:data) {
   //Create a Row in the Sheet
   XSSFRow row=sheet.createRow(rowNum++);
				
   //To insert the data into Cells
   int colNum=0;
   for(Object field:rowdata) {
   XSSFCell cell=row.createCell(colNum++);
   if(field instanceof String) {
   cell.setCellValue((String)field);
					
    }else if(field instanceof Integer) {
    cell.setCellValue((Integer)field);
    }
    }
				
    }
    //Create an Output stream object and write the data
    try(FileOutputStream os=new FileOutputStream ("Test.Xlsx")){
    workbook.write(os);
				
    }
    System.out.println("Data Added Successfully to file");
			
						
    } catch (IOException e) {
    e.printStackTrace();
		}

	}
}