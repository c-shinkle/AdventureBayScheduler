package excelParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParser {
	private final static int LEFT_COLUMN_ROTATIONS = 3;
	
	private final static int RIGHT_COLUMN_ROTATIONS = 8;
	
	private static ArrayList<String> listOfLifeguards = new ArrayList<>();

	private static int indexOfList = 0;

	public static void main(String[] args) {
		long time = System.currentTimeMillis();
		fillLifeguardList();
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(new FileInputStream("CloneScheduler2017.xlsx"));
			XSSFSheet s = wb.getSheet("Rotations");
			fillRotation(s, RotationColor.RED);
			fillRotation(s, RotationColor.ORANGE);
			fillRotation(s, RotationColor.BLUE);
			fillRotation(s, RotationColor.YELLOW);
			fillRotation(s, RotationColor.GREEN);
			fillRotation(s, RotationColor.VIOLET);
		} catch (FileNotFoundException e) {
			System.out.println("We couldn't edit the file because Excel is using it right now!");
			System.out.println("Close the file in Excel!");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					//write to the file, close all resources
					FileOutputStream output = new FileOutputStream("CloneScheduler2017.xlsx");
					wb.write(output);
					wb.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		System.out.println((System.currentTimeMillis() - time) + " Miliseconds");
	}

	private static void fillRotation(XSSFSheet s, RotationColor rc) {
		final int totalRotSize = findRotationSize(rc);
		final int colorOffset = findColorOffset(rc);
		final boolean isLeftColumn = findColumn(rc);
		System.out.println("Filling " + rc + " rotation...");
		for (int indexOfRotation = 0; indexOfRotation < totalRotSize
				&& indexOfList < listOfLifeguards.size(); indexOfRotation++) {
			if (!isOptionalSpot(indexOfRotation + colorOffset, isLeftColumn)) {
				final int column;
				if (isLeftColumn)
					column = LEFT_COLUMN_ROTATIONS;
				else
					column = RIGHT_COLUMN_ROTATIONS;
				s.getRow(indexOfRotation + colorOffset).getCell(column)
						.setCellValue(listOfLifeguards.get(indexOfList++));
//				System.out.println("Rotation " + rc + " gets " + listOfLifeguards.get(indexOfList - 1) + " at R"
//						+ (indexOfRotation + colorOffset - 1));
			}
		}
	}

	private static boolean isOptionalSpot(int i, boolean b) {
		if (b) {
			switch (i) {
			case 7:
			case 32:
			case 33:
				return true;
			default:
				return false;
			}
		} else {
			switch (i) {
			case 5:
			case 11:
			case 19:
			case 33:
			case 34:
			case 36:
				return true;
			default:
				return false;
			}
		}
	}

	private static int findRotationSize(RotationColor rc) {
		switch (rc) {
		case BLUE:
		case VIOLET:
		case GREEN:
			return 10;
		case RED:
		case ORANGE:
			return 11;
		case YELLOW:
			return 12;
		default:
			return -1;
		}
	}

	private static boolean findColumn(RotationColor rc) {
		switch (rc) {
		case RED:
		case ORANGE:
		case BLUE:
			return true;
		default:// Basically, any other color
			return false;
		}
	}

	private static int findColorOffset(RotationColor rc) {
		switch (rc) {
		case RED:
		case YELLOW:
			return 2;
		case ORANGE:
			return 15;
		case GREEN:
			return 16;
		case BLUE:
		case VIOLET:
			return 28;
		default:
			return -1;
		}
	}

	private static void fillLifeguardList() {
		System.out.println("The lifeguards on the schedule for today are...");
		for (int i = 0; i < 67; i++) {
			listOfLifeguards.add("Lifeguard #" + i);
			System.out.println("Lifeguard #" + i);
		}
	}
}