import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.*;

//import jxl.*;
//import jxl.read.biff.BiffException;
//import jxl.write.*;

public class MMRTracker 
{
	private static final String EXCEL_FILE_LOCATION = "C:\\Users\\Michael\\Documents\\Dota 2\\Dota MMR Tracker.xls";
	//private static WritableWorkbook workbook = null;
	//private static Workbook original = null;
	private static Scanner sc = new Scanner(System.in);
	private static int games_played = 0;
	private static int matchID;
	private static int mmr;
	private static String another = null;
	
	private static DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM/dd/yyyy");
	private static String date = dtf.format(LocalDateTime.now());
	
	public static void main(String[] args) 
	{
		/*do
		{
	        try {
	
	            original = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
	            
	            workbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION), original);
	
	            for(WritableCell cell = workbook.getWritableCell("E" + (games_played + 2)); !(cell.getContents().equals("")); cell = workbook.getWritableCell("E" + (games_played + 2)))
	            	games_played++;
	            
	            System.out.print("Match ID: ");
	            matchID = sc.nextInt();
	            System.out.print("MMR: ");
	            mmr = sc.nextInt();
	            
	            WritableCell dateCell = workbook.getWritableCell("B" + (games_played + 2));
	            WritableCell matchIDCell = workbook.getWritableCell("C" + (games_played + 2));
	            WritableCell mmrCell = workbook.getWritableCell("E" + (games_played + 2));
	            
	            if(dateCell.getType() == CellType.LABEL)
	            {
	            	Label l = (Label) dateCell;
	            	l.setString(date);
	            }
	            if(matchIDCell.getType() == CellType.LABEL)
	            {
	            	Label l = (Label) matchIDCell;
	            	l.setString("" + matchID);
	            }
	            if(mmrCell.getType() == CellType.LABEL)
	            {
	            	Label l = (Label) mmrCell;
	            	l.setString("" + mmr);
	            }
	            
	            workbook.write();
	            workbook.close();
	            
	            System.out.print("Enter another game? Y for yes, anything else for no: ");
	            another = sc.nextLine();
	
	        } catch (IOException e) {
	            e.printStackTrace();
	        } catch (BiffException e) {
	            e.printStackTrace();
	        } catch (WriteException e) {
	        	e.printStackTrace();
	        } finally {
	        
	
	            if (original != null) {
	                original.close();
	            }
	        }
		}while(another.equals("Y"));*/
		
		do
		{
			try
			{
				FileInputStream fsIP= new FileInputStream(new File(EXCEL_FILE_LOCATION)); //Read the spreadsheet that needs to be updated
	            
	            HSSFWorkbook wb = new HSSFWorkbook(fsIP); //Access the workbook
	              
	            HSSFSheet worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
	            
	            for(Cell cell = worksheet.getRow(5).getCell(games_played + 2); cell != null; cell = worksheet.getRow(games_played + 2).getCell(5))
	            	games_played++;
	            
	            Cell dateCell = worksheet.getRow(games_played + 2).getCell(2);
	            Cell matchIDCell = worksheet.getRow(games_played + 2).getCell(3);
	            Cell mmrCell = worksheet.getRow(games_played + 2).getCell(5);
	            
	            dateCell.setCellValue(date);
	            matchIDCell.setCellValue("" + matchID);
	            mmrCell.setCellValue("" + mmr);
	            
	            fsIP.close(); //Close the InputStream
	             
	            FileOutputStream output_file =new FileOutputStream(new File(EXCEL_FILE_LOCATION));  //Open FileOutputStream to write updates
	              
	            wb.write(output_file); //write changes
	              
	            output_file.close();  //close the stream 
	            wb.close();
	            
	            System.out.print("Enter another game? Y for yes, anything else for no: ");
	            another = sc.nextLine();
			} catch (IOException e) {
	            e.printStackTrace();
	        } 
		}while(another.equals("Y"));
	}

}
