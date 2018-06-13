import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Scanner;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;

public class MMRTracker 
{
	private static final String EXCEL_FILE_LOCATION = "C:\\Users\\Michael\\Documents\\Dota 2\\Dota MMR Tracker.xls";
	private static WritableWorkbook workbook = null;
	private static Workbook original = null;
	private static Scanner sc = new Scanner(System.in);
	private static int games_played = 0;
	private static int matchID;
	private static int mmr;
	private static String another = null;
	
	private static DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM/dd/yyyy");
	private static String date = dtf.format(LocalDateTime.now());
	
	public static void main(String[] args) 
	{
		do
		{
	        try {
	
	            original = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
	            
	            workbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION), original);
	
	            for(Cell cell = workbook.getWritableCell("E" + (games_played + 2)); !(cell.getContents().equals("")); cell = workbook.getWritableCell("E" + (games_played + 2)))
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
		}while(another.equals("Y"));
	}

}
