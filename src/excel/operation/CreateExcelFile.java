package excel.operation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.util.CellRangeAddressList;


class GenerateExcel
{
//	private static final int columnIndexFromHeaderList = 0;
	static String menuFile="C:\\Users\\samirans\\Desktop\\MenuFile.txt";
	static String fileName="C:\\Users\\samirans\\Desktop\\FileName.txt";
	static String extensionName="C:\\Users\\samirans\\Desktop\\FileEetension.txt";
	static String statusname ="C:\\Users\\samirans\\Desktop\\Status.txt";

	public ArrayList<String> getMenu() throws FileNotFoundException
	{
		String currentLine = null;
		ArrayList<String>list = new ArrayList<String>();
		try(BufferedReader br = new BufferedReader(new FileReader(menuFile)))
		{
			while((currentLine = br.readLine()) != null)
			{
				list.add(currentLine);
			}
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
		
		return list;
	}
	
	public String[] getStatus() throws FileNotFoundException, IOException
	{
		String current = null;
		ArrayList<String>slist = new ArrayList<String>();
		try(BufferedReader br = new BufferedReader(new FileReader(statusname)))
		{
			while((current = br.readLine()) != null)
			{
				slist.add(current);
			}
		}
		
		catch(IOException e)
		{
			e.printStackTrace();
		}
		
		String Array[] = new String[slist.size()];
		for(int i = 0; i < slist.size(); i++)
		{
			Array[i] = slist.get(i);
		}
		
		/*for(String k : Array)
		{
			System.out.println(k);
		}*/
		return Array;
	}
	
	public void createExcel() throws IOException
	{
		ArrayList<String>arr = new ArrayList<String>();
		arr = getMenu();
		
		/*DataValidation dataValidation = null;
		DataValidationConstraint constraint = null;
		DataValidationHelper validationHelper = null;*/
		
		//@SuppressWarnings("resource")
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		
		int noOfDays = days();
		String month = null;
		if(noOfDays != 0)
		{
			int dateHold = Calendar.DAY_OF_MONTH;
			Calendar cal = Calendar.getInstance();
			if(dateHold <= 15)
			{
				month = new SimpleDateFormat("MM-DD-YYYY").format(cal.getTime());
			}
			else
			{
				month = new SimpleDateFormat("MM-DD-YYYY").format(Calendar.MONTH);
			}
		}
		
		HSSFSheet spreadsheet = workbook.createSheet(month);
	
		//XSSFCell cell = null;
		HSSFCellStyle mystyle = workbook.createCellStyle();
		HSSFRow row = spreadsheet.createRow(0);
		HSSFCell cell = row.createCell(0);
		
		int i=0;
		
		for(String menu : arr)
		{
			cell = row.createCell(i);
			cell.setCellValue(menu);
			cell.setCellStyle(mystyle);
			i++;
		}
		
		String[] Status = getStatus();
		for(int k = 0; k < Status.length; k++)
		{
			//System.out.println(Status[k]);
		}
		DVConstraint dvConstraint = null;
		CellRangeAddressList addressList = new CellRangeAddressList(1,5,0,0); //Addressing Positions(StratRow, endRow, StatCol, endCol)
		/*for(int j = 0; j < Status.length; j++)
		{
			String S = Status[j];
			System.out.println(Status[j]);
			dvConstraint = DVConstraint.createExplicitListConstraint(new String[] {Status[j]});
			
		}*/
		
		dvConstraint = DVConstraint.createExplicitListConstraint(Status);
		
		for(int j=1; j<=5;j++)
		{
			spreadsheet.createRow(j).createCell(j-1).setCellValue("Pending"); 
		}
		 //Creating default value in Dropdown
		DataValidation dataValidation = new HSSFDataValidation(addressList, dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		spreadsheet.addValidationData(dataValidation);
	
		LocalDate localDate = LocalDate.now();
		int Month = localDate.getMonthValue();
		int Year = localDate.getYear();
		String FileName = fileName();
		String Extension = FileExtension();
		FileOutputStream out = new FileOutputStream(new File("C:\\Users\\samirans\\Desktop\\"+FileName+""+Month+"-"+Year+"."+Extension));
		workbook.write(out);
		out.close();
		System.out.println(FileName+" "+Month+"-"+Year+" written successfully");
	}
	
	public String fileName() throws FileNotFoundException
	{
		String name = null;
		try(BufferedReader read = new BufferedReader(new FileReader(fileName)))
		{
			String current= null;
			while((current = read.readLine()) != null)
			{
				
				name = current;
			}
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
		return name;
	}
	
	public String FileExtension()throws FileNotFoundException
	{
		String extension = null;
		try(BufferedReader read = new BufferedReader(new FileReader(extensionName)))
		{
			String currentL=null;
			while((currentL = read.readLine()) != null)
			{
				extension = currentL;
			}
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
		return extension;
	}
	
	public int days()
	{
		Calendar c = Calendar.getInstance();
		int day = c.getActualMaximum(Calendar.DAY_OF_MONTH);
		return day;
	}
}

public class CreateExcelFile {

	public static void main(String[] args)throws Exception{
		// TODO Auto-generated method stub

		GenerateExcel obj = new GenerateExcel();
		obj.createExcel();
		//obj.getStatus();
 	}
}