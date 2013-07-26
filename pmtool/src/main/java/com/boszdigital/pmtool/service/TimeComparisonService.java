package com.boszdigital.pmtool.service;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import au.com.bytecode.opencsv.CSVReader;

import com.boszdigital.pmtool.model.Project;
import com.boszdigital.pmtool.model.StaffMember;

/**
 * This service compares the time reported by the staff members on different
 * projects, from the two time tracking tools (ActiveTime and Oracle).
 * 
 * @author <a href="mailto:daniel.hoyos@boszdigital.com">Daniel Hoyos</a>
 * 
 */
public class TimeComparisonService {

	private static final char ORACLE_SEPARATOR = '\t';
	private static final String OUTPUT_NAME = "Output_Report_";
	private static final String OUTPUT_NAME_FORMAT = "yyyy-MM-dd";
	private static final String OUTPUT_NAME_EXT = ".xls";

	private String folderPath;
	private String activeTimeFileName;
	private String oracleFileName;

	private Map<String, StaffMember> staffActiveTime = new TreeMap<String, StaffMember>();
	private Map<String, StaffMember> staffOracle = new TreeMap<String, StaffMember>();

	private Map<String, String> fullStaff = new TreeMap<String, String>();
	private Map<String, String> projects = new TreeMap<String, String>();

	/**
	 * This is the method that runs the process of loading and comparing the
	 * files, and generating the report.
	 * 
	 * @param args
	 */
	public void run(String[] args) {
		long initialTime = System.currentTimeMillis();
		System.out.println("*** Initiating Time Comparison Process.");
		if (args.length == 3) {
			folderPath = args[0];
			activeTimeFileName = args[1];
			oracleFileName = args[2];
			try {
				parseActiveTimeFile(folderPath + activeTimeFileName);
				parseOracleFile(folderPath + oracleFileName, ORACLE_SEPARATOR);
				writeReport();

			} catch (IOException e) {
				System.out.println("\t" + e.getMessage());
			}
		} else {
			System.out.println("\t*** ERROR: Invalid number of arguments.");
			System.out
					.println("\t*** ERROR: java -jar pmtool.jar path/to/folder/ active_time_file_name.ext oracle_file_name.ext ");
		}

		System.out.println("*** Time Comparison Process finalized.");
		System.out.println("*** Time Spend: "
				+ (System.currentTimeMillis() - initialTime) + "ms");
	}

	/**
	 * This method parses the Active Time file into a hash map.
	 * 
	 * @param path
	 *            the path of the Active Time file.
	 * @throws IOException
	 */
	private void parseActiveTimeFile(String path) throws IOException {
		System.out.println("\tParsing the Active Time file ["
				+ activeTimeFileName + "].");
		CSVReader reader;
		try {
			reader = new CSVReader(new FileReader(path));
		} catch (FileNotFoundException e) {
			throw new IOException("*** ERROR: File not found: " + path);
		}
		String[] row;
		// Skips the first four rows of the file
		for (int i = 0; i < 4; i++) {
			row = reader.readNext();
		}
		List<StaffMember> staff = new ArrayList<StaffMember>();
		row = reader.readNext();

		int columnCount = row.length;
		for (int i = 5; i < row.length; i++) {
			StaffMember staffMember = new StaffMember(
					StaffMember.generateStaffMemberCodeName(row[i]), row[i]);
			staff.add(staffMember);
		}

		while ((row = reader.readNext()) != null) {
			if (row.length == columnCount) {
				String projectName = row[2];
				String projectCode = Project.generateProjectCode(projectName);
				for (int i = 5; i < row.length; i++) {
					float hours = Float.parseFloat(row[i]);
					if (hours != 0) {
						StaffMember person = staff.get(i - 5);
						person.addTime(projectCode, projectName, hours);
						person.addTotalTime(hours);
						staff.set(i - 5, person);
					}
				}
				if (!projects.containsKey(projectCode)) {
					projects.put(projectCode, projectName);
				}
			} else {
				break;
			}

			for (Iterator<StaffMember> i = staff.iterator(); i.hasNext();) {
				StaffMember person = i.next();
				staffActiveTime.put(person.getCodeName(), person);
				if (!fullStaff.containsKey(person.getCodeName())) {
					fullStaff.put(person.getCodeName(), person.getFullName());
				}
			}
		}
		reader.close();
		System.out.println("\tActive Time file [" + activeTimeFileName
				+ "] parsed.");
	}

	/**
	 * This method parses the oracle file into a hash map.
	 * 
	 * @param path
	 *            the path of the oracle file.
	 * @param separator
	 *            the separator used on the oracle file.
	 * @throws IOException
	 */
	private void parseOracleFile(String path, char separator)
			throws IOException {
		System.out.println("\tParsing the Oracle file [" + oracleFileName
				+ "].");

		CSVReader reader;
		try {
			reader = new CSVReader(new FileReader(path), separator);
		} catch (FileNotFoundException e) {
			throw new IOException("*** ERROR: File not found: " + path);
		}
		String[] row;
		// Skips the first row of the file
		reader.readNext();
		while ((row = reader.readNext()) != null) {
			String personName = row[2];
			String personCodeName = StaffMember
					.generateStaffMemberCodeName(personName);

			StaffMember person;
			if (staffOracle.containsKey(personCodeName)) {
				person = staffOracle.get(personCodeName);
			} else {
				person = new StaffMember(personCodeName, personName);
			}
			String projectName = row[3];
			String projectCode = Project.generateProjectCode(projectName);

			float hours = Float.parseFloat(row[13]);
			person.addTime(projectCode, projectName, hours);
			person.addTotalTime(hours);

			staffOracle.put(personCodeName, person);
			if (!projects.containsKey(projectCode)) {
				projects.put(projectCode, projectName);
			}
			if (!fullStaff.containsKey(personCodeName)) {
				fullStaff.put(personCodeName, personName);
			}
		}
		reader.close();
		System.out.println("\tOracle file [" + oracleFileName + "] parsed.");
	}

	/**
	 * This method creates the excel file, generates the report, and writes it
	 * to the disc.
	 * 
	 * @throws IOException
	 */
	private void writeReport() throws IOException {
		System.out.println("\tCreating Report.");
		ExcelReportHelper reportHelper = new ExcelReportHelper();
		// Write the output to a file
		String outputFileName = OUTPUT_NAME
				+ new SimpleDateFormat(OUTPUT_NAME_FORMAT).format(new Date(
						System.currentTimeMillis())) + OUTPUT_NAME_EXT;
		FileOutputStream fileOut = new FileOutputStream(folderPath
				+ outputFileName);
		reportHelper.writeTimeDiferenceReport().write(fileOut);
		fileOut.close();
		System.out.println("\tReport Created.");
		System.out.println("\tReport Saved to \"" + folderPath + outputFileName
				+ "\".");
	}

	/**
	 * Helper class to generate the report of the time comparison.
	 * 
	 * @author <a href="mailto:daniel.hoyos@boszdigital.com">Daniel Hoyos</a>
	 * 
	 */
	private class ExcelReportHelper {
		private final String FONT_NAME = "Calibri";
		private final short FONT_SIZE = 11;
		private final String ACTIVETIME_TITLE = "ACT";
		private final String ORACLE_TITLE = "ORA";

		private int columnPos = 0;
		private int rowPos = 0;
		private int errors = 0;

		private Workbook workbook;
		private Sheet sheet;
		private Row row;
		private Cell cell;
		private CellStyle cellStyleTitle;
		private CellStyle cellStyleTableRowLeft;
		private CellStyle cellStyleTableRowRight;
		private CellStyle cellStyleTableRowLeftOdd;
		private CellStyle cellStyleTableRowRightOdd;
		private CellStyle cellStyleTableRowError;

		/**
		 * Method that initializes the writing of the report.
		 * 
		 * @return
		 */
		public Workbook writeTimeDiferenceReport() {
			// create a new workbook
			workbook = new HSSFWorkbook();
			setStyles();
			// create a new sheet
			sheet = workbook.createSheet();
			// create rows
			createTitleRow();
			createStaffRows();

			// Auto adjust columns
			for (int i = 0; i < (projects.size() * 2) + 3; i++) {
				sheet.autoSizeColumn(i);
			}
			System.out.println("*** Number of errors found: " + errors);
			return workbook;
		}

		/**
		 * This method creates the rows of the table title.
		 */
		private void createTitleRow() {
			columnPos = 1;

			row = sheet.createRow(rowPos++);
			cell = row.createCell(columnPos++);
			cell.setCellValue(ACTIVETIME_TITLE);
			cell.setCellStyle(cellStyleTableRowLeft);
			cell = row.createCell(columnPos++);
			cell.setCellValue(ORACLE_TITLE);
			cell.setCellStyle(cellStyleTableRowRight);
			for (Iterator<String> i = projects.values().iterator(); i.hasNext(); i
					.next()) {
				cell = row.createCell(columnPos++);
				cell.setCellValue(ACTIVETIME_TITLE);
				cell.setCellStyle(cellStyleTableRowLeft);
				cell = row.createCell(columnPos++);
				cell.setCellValue(ORACLE_TITLE);
				cell.setCellStyle(cellStyleTableRowRight);
			}

			columnPos = 0;
			row = sheet.createRow(rowPos++);
			cell = row.createCell(columnPos++);
			cell.setCellValue("Staff Name");
			cell.setCellStyle(cellStyleTitle);
			cell = row.createCell(columnPos++);
			cell.setCellValue("Total");
			cell.setCellStyle(cellStyleTitle);
			cell = row.createCell(columnPos++);
			cell.setCellValue("");
			cell.setCellStyle(cellStyleTitle);
			sheet.addMergedRegion(new CellRangeAddress(rowPos - 1, rowPos - 1,
					columnPos - 2, columnPos - 1));

			for (Iterator<String> i = projects.values().iterator(); i.hasNext();) {
				cell = row.createCell(columnPos++);
				cell.setCellValue(i.next());
				cell.setCellStyle(cellStyleTitle);
				cell = row.createCell(columnPos++);
				cell.setCellValue("");
				cell.setCellStyle(cellStyleTitle);

				sheet.addMergedRegion(new CellRangeAddress(rowPos - 1,
						rowPos - 1, columnPos - 2, columnPos - 1));
			}
		}

		/**
		 * This method creates a cell containing the name of the staff. Even
		 * rows are shown in white. Odd row are shown in light turquoise. Cells
		 * with errors are shown in maroon.
		 */
		public void createStaffRows() {
			for (Iterator<String> i = fullStaff.keySet().iterator(); i
					.hasNext();) {
				columnPos = 0;
				String staffKey = i.next();
				row = sheet.createRow(rowPos++);
				cell = row.createCell(columnPos++);
				cell.setCellValue(fullStaff.get(staffKey));
				if (rowPos % 2 == 0) {
					cell.setCellStyle(cellStyleTableRowLeft);
				} else {
					cell.setCellStyle(cellStyleTableRowLeftOdd);
				}

				StaffMember staffMemberActiveTime = staffActiveTime
						.get(staffKey);
				StaffMember staffMemberOracle = staffOracle.get(staffKey);

				createTotalCells(staffMemberActiveTime, staffMemberOracle);
				for (Iterator<String> j = projects.keySet().iterator(); j
						.hasNext();) {
					String projectKey = j.next();
					createTimeCells(projectKey, staffMemberActiveTime,
							staffMemberOracle);
				}
			}
		}

		/**
		 * This method creates a cell containing the total time. Even rows are
		 * shown in white. Odd row are shown in light turquoise. Cells with
		 * errors are shown in maroon.
		 * 
		 * @param staffMemberActiveTime
		 * @param staffMemberOracle
		 */
		private void createTotalCells(StaffMember staffMemberActiveTime,
				StaffMember staffMemberOracle) {
			float timeActiveTime = 0;
			float timeOracle = 0;
			CellStyle cellStyleLeft;
			CellStyle cellStyleRight;
			if (staffMemberActiveTime != null) {
				timeActiveTime = staffMemberActiveTime.getTotalTime();
			}
			if (staffMemberOracle != null) {
				timeOracle = staffMemberOracle.getTotalTime();
			}
			if (timeActiveTime == timeOracle) {
				if (rowPos % 2 == 0) {
					cellStyleLeft = cellStyleTableRowLeft;
					cellStyleRight = cellStyleTableRowRight;
				} else {
					cellStyleLeft = cellStyleTableRowLeftOdd;
					cellStyleRight = cellStyleTableRowRightOdd;
				}
			} else {
				cellStyleLeft = cellStyleTableRowError;
				cellStyleRight = cellStyleTableRowError;
			}
			cell = row.createCell(columnPos++);
			cell.setCellValue(timeActiveTime);
			cell.setCellStyle(cellStyleLeft);
			cell = row.createCell(columnPos++);
			cell.setCellValue(timeOracle);
			cell.setCellStyle(cellStyleRight);
		}

		/**
		 * This method creates a cell containing the time. Even rows are shown
		 * in white. Odd row are shown in light turquoise. Cells with errors are
		 * shown in maroon.
		 * 
		 * @param projectKey
		 *            The key of the project.
		 * @param staffMemberActiveTime
		 *            The staff member from the active time report.
		 * @param staffMemberOracle
		 *            The staff member from the oracle report.
		 */
		private void createTimeCells(String projectKey,
				StaffMember staffMemberActiveTime, StaffMember staffMemberOracle) {
			float timeActiveTime = 0;
			float timeOracle = 0;
			CellStyle cellStyleLeft;
			CellStyle cellStyleRight;
			if (staffMemberActiveTime != null) {
				Project project = staffMemberActiveTime.getProjects().get(
						projectKey);
				if (project != null) {
					timeActiveTime = project.getTime();
				}
			}
			if (staffMemberOracle != null) {
				Project project = staffMemberOracle.getProjects().get(
						projectKey);
				if (project != null) {
					timeOracle = project.getTime();
				}
			}
			if (timeActiveTime == timeOracle) {
				if (rowPos % 2 == 0) {
					cellStyleLeft = cellStyleTableRowLeft;
					cellStyleRight = cellStyleTableRowRight;
				} else {
					cellStyleLeft = cellStyleTableRowLeftOdd;
					cellStyleRight = cellStyleTableRowRightOdd;
				}
			} else {
				errors++;
				cellStyleLeft = cellStyleTableRowError;
				cellStyleRight = cellStyleTableRowError;
			}
			cell = row.createCell(columnPos++);
			cell.setCellValue(timeActiveTime);
			cell.setCellStyle(cellStyleLeft);
			cell = row.createCell(columnPos++);
			cell.setCellValue(timeOracle);
			cell.setCellStyle(cellStyleRight);
		}

		/**
		 * Method to create cell styles according to the needs of the different
		 * cells
		 * 
		 * @param boldWeight
		 *            the short value of the bold weight
		 * @param backgroundColor
		 *            the short value of the background color
		 * @param border
		 *            the short values of the borders border[0]: Bottom Border.
		 *            border[1]: Left Border. border[2]: Right Border.
		 *            border[3]: Top Border.
		 * 
		 * @return the cell style.
		 */
		private CellStyle createCellStyle(short boldWeight,
				short backgroundColor, short... border) {
			// Create a new font and alter it.
			Font font = workbook.createFont();
			font.setFontHeightInPoints(FONT_SIZE);
			font.setFontName(FONT_NAME);
			font.setBoldweight(boldWeight);
			// Create a new style
			CellStyle style = workbook.createCellStyle();
			style.setFont(font);
			style.setFillForegroundColor(backgroundColor);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			// Borders
			style.setBorderBottom(border[0]);
			style.setBorderLeft(border[1]);
			style.setBorderRight(border[2]);
			style.setBorderTop(border[3]);

			return style;
		}

		/**
		 * Initializes all the styles used to set the cells.
		 */
		private void setStyles() {
			cellStyleTitle = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.WHITE.getIndex(), CellStyle.BORDER_MEDIUM,
					CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM,
					CellStyle.BORDER_MEDIUM);
			cellStyleTableRowLeft = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.WHITE.getIndex(), CellStyle.BORDER_THIN,
					CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN,
					CellStyle.BORDER_THIN);
			cellStyleTableRowRight = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.WHITE.getIndex(), CellStyle.BORDER_THIN,
					CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM,
					CellStyle.BORDER_THIN);
			cellStyleTableRowLeftOdd = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.LIGHT_TURQUOISE.getIndex(),
					CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM,
					CellStyle.BORDER_THIN, CellStyle.BORDER_THIN);
			cellStyleTableRowRightOdd = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.LIGHT_TURQUOISE.getIndex(),
					CellStyle.BORDER_THIN, CellStyle.BORDER_THIN,
					CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN);
			cellStyleTableRowError = createCellStyle(Font.BOLDWEIGHT_BOLD,
					IndexedColors.MAROON.getIndex(), CellStyle.BORDER_THIN,
					CellStyle.BORDER_THIN, CellStyle.BORDER_THIN,
					CellStyle.BORDER_THIN);
		}
	}
}
