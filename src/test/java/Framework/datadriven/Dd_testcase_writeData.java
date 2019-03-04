package Framework.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Dd_testcase_writeData {
	
	
	public void writeExcel(String filepath, String fileName, String sheetName, String[] dataToWrite) throws IOException{
		
		//open Excel file
		File file=new File(filepath+"\\"+fileName);
		
		//read excel file
		FileInputStream fis=new FileInputStream(file);
		
		//initialise Workbook variable
		Workbook wb=null;
		
		//splitting the file name and extension
		
		String fileExtName=fileName.substring(fileName.indexOf("."));
		
		if(fileExtName.equals(".xlsx")){
			wb=new XSSFWorkbook(fis);
		}
		else if(fileExtName.equals(".xls")){
			wb=new HSSFWorkbook(fis);
		}
		
		//read excel file by its name
		
		Sheet sheet=wb.getSheet(sheetName);
		
		
		//find number of rows in excel file. We exclude the first row.. this is the column name
		int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
		System.out.println(rowCount);
		
		Row row=sheet.getRow(0);
		
		//create a new row and append it at last of sheet
		
		Row nRow=sheet.createRow(rowCount+1);
		
		//create a loop over the cell of newly created row
		for (int j=0; j<row.getLastCellNum(); j++){
			//fill the data in row
			Cell cell=nRow.createCell(j);
			cell.setCellValue(dataToWrite[j]);
		}
		
		//close input stream
		fis.close();
		
		//create an object of FileOutputStream class to write data in excel file
		
		FileOutputStream fos=new FileOutputStream(file);
		
		//write data in the excel file
		wb.write(fos);
		
		//close output stream
		fos.close();
		
}
		
		
		
		
	 public static void main(String[] args) throws IOException{
		 			  
			 Dd_testcase_writeData writeExcelFile = new Dd_testcase_writeData();
			 
			 String[] valueToWrite= {"test11", "test21"};

			    //Prepare the path of excel file

			    String filePath = "C:\\Users\\Naresh\\oxygen-workspace\\datadriven\\src\\test\\java\\Framework\\datadriven";

			    //Call read file method of the class to read data

			    writeExcelFile.writeExcel(filePath, "testdata.xlsx", "LoginTest", valueToWrite);

		 }
		
	}


