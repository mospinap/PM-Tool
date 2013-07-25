package com.boszdigital.pmtool.service;

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
	
	private String folderPath;
	private String activeTimeFileName;
	private String oracleFileName;

	private Map<String, StaffMember> staffActiveTime = new TreeMap<String, StaffMember>();
	private Map<String, StaffMember> staffOracle = new TreeMap<String, StaffMember>();

	private Map<String, String> fullStaff = new TreeMap<String, String>();
	private Map<String, String> projects = new TreeMap<String, String>();

	public void run(String[] args) {
		folderPath = args[0];
		activeTimeFileName = args[1];
		oracleFileName = args[2];
		try {
			parseActiveTimeFile(folderPath + activeTimeFileName);//"staffoutput_0815-0821.csv");
			System.out.println("**************************"); // TODO Delete
			System.out.println("**************************"); // TODO Delete
			//parseOracleFile("fnd_gfm_0815-0821.tsv", ORACLE_SEPARATOR);  // TODO Delete
			parseOracleFile(folderPath + oracleFileName, ORACLE_SEPARATOR);

			System.out.println("**************************"); // TODO Delete
			System.out.println("**************************"); // TODO Delete
			System.out.println(fullStaff); // TODO Delete
			System.out.println(projects); // TODO Delete
			writeReport();
			System.out.println(new SimpleDateFormat("yyyy-MM-dd").format(new Date(System.currentTimeMillis())));
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void parseActiveTimeFile(String path) throws IOException {
		CSVReader reader = new CSVReader(new FileReader(path));
		String[] row;
		// Skips the first four rows of the file
		for (int i = 0; i < 4; i++) {
			row = reader.readNext();
		}
		List<StaffMember> staff = new ArrayList<StaffMember>();
		row = reader.readNext();
		// System.out.println(row[0]); // TODO Delete
		int columnCount = row.length;
		for (int i = 5; i < row.length; i++) {
			StaffMember staffMember = new StaffMember(
					StaffMember.generateStaffMemberCodeName(row[i]), row[i]);
			staff.add(staffMember);
			// System.out.println("Person: " + person.getCodeName()); //TODO
			// Delete
		}

		while ((row = reader.readNext()) != null) {
			if (row.length == columnCount) {
				String projectName = row[2];
				String projectCode = Project.generateProjectCode(projectName);
				// System.out.print(projectCode + "|"); //TODO Delete
				for (int i = 5; i < row.length; i++) {
					float hours = Float.parseFloat(row[i]);
					if (hours != 0) {
						StaffMember person = staff.get(i - 5);
						person.addTime(projectCode, projectName, hours);
						person.addTotalTime(hours);
						staff.set(i - 5, person);
						// System.out.print(hours + "|"); //TODO Delete
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

		// TODO delete this section
		for (Iterator<StaffMember> i = staff.iterator(); i.hasNext();) {
			StaffMember person = i.next();
			String cad = person.getFullName();
			for (Iterator<Project> j = person.getProjects().values().iterator(); j
					.hasNext();) {
				Project project = j.next();
				cad += "\n\t" + project.getCode() + ": " + project.getTime();
			}
			System.out.println(cad);
		}
		// TODO End of section
		reader.close();
	}

	private void parseOracleFile(String path, char separator)
			throws IOException {
		CSVReader reader = new CSVReader(new FileReader(path), separator);
		String[] row;
		// Skips the first row of the file
		reader.readNext();
		while ((row = reader.readNext()) != null) {
			// System.out.println(row[2]); //TODO Delete
			String personName = row[2];
			String personCodeName = StaffMember.generateStaffMemberCodeName(personName);

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

		// TODO delete this section
		for (Iterator<StaffMember> i = staffOracle.values().iterator(); i
				.hasNext();) {
			StaffMember person = i.next();
			String cad = person.getFullName();
			for (Iterator<Project> j = person.getProjects().values().iterator(); j
					.hasNext();) {
				Project project = j.next();
				cad += "\n\t" + project.getCode() + ": " + project.getTime();
			}
			System.out.println(cad);
		}
		// TODO End of section

		reader.close();
	}

	private void writeReport() throws IOException {
		ExcelReportHelper reportHelper = new ExcelReportHelper();

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(folderPath + "workbook.xls"); // FIXME The name of the output constant and date. 
		reportHelper.writeTimeDiferenceReport().write(fileOut);
		fileOut.close();
	}

	private class ExcelReportHelper {
		private final String FONT_NAME = "Calibri";
		private final short FONT_SIZE = 11;
		private final String ACTIVETIME_TITLE = "ACT";
		private final String ORACLE_TITLE = "ORA";
		
		private int columnPos = 0;
		private int rowPos = 0;
		
		private Workbook workbook;
		private Sheet sheet;
		private Row row;
		private Cell cell;
		private CellStyle cellStyleNormal;
		private CellStyle cellStyleTitle;
		private CellStyle cellStyleTableRowLeft;
		private CellStyle cellStyleTableRowRight;
		private CellStyle cellStyleTableRowLeftOdd;
		private CellStyle cellStyleTableRowRightOdd;
		private CellStyle cellStyleTableRowError;

		public Workbook writeTimeDiferenceReport() {
			// create a new workbook
			workbook = new HSSFWorkbook();
			setStyles();
			// create a new sheet
			sheet = workbook.createSheet();
			// create the title row
			createTitleRow();

			createStaffRows();
			
			// Auto adjust columns
			for(int i = 0; i < (projects.size() * 2) + 3;i++){
				sheet.autoSizeColumn(i);
			}

			return workbook;
		}

		private void createTitleRow() {
			columnPos = 1;

			row = sheet.createRow(rowPos++);
			cell = row.createCell(columnPos++);
			cell.setCellValue(ACTIVETIME_TITLE);
			cell.setCellStyle(cellStyleTableRowLeft);
			cell = row.createCell(columnPos++);
			cell.setCellValue(ORACLE_TITLE);
			cell.setCellStyle(cellStyleTableRowRight);
			//row.createCell(columnPos++).setCellValue(ACTIVETIME_TITLE); //TODO Delete
			//row.createCell(columnPos++).setCellValue(ORACLE_TITLE); //TODO Delete
			for (Iterator<String> i = projects.values().iterator(); i.hasNext(); i.next()) {
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
			//row.createCell(columnPos++).setCellValue("Staff Name"); //TODO Delete
			//row.createCell(columnPos++).setCellValue("Total");
			//row.createCell(columnPos++).setCellValue("");
			sheet.addMergedRegion(new CellRangeAddress(rowPos - 1, rowPos - 1,
					columnPos - 2, columnPos - 1));

			for (Iterator<String> i = projects.values().iterator(); i.hasNext();) {
				cell = row.createCell(columnPos++);
				cell.setCellValue(i.next());
				cell.setCellStyle(cellStyleTitle);
				cell = row.createCell(columnPos++);
				cell.setCellValue("");
				cell.setCellStyle(cellStyleTitle);
				//row.createCell(columnPos++).setCellValue(i.next());
				//row.createCell(columnPos++).setCellValue("");
				sheet.addMergedRegion(new CellRangeAddress(rowPos - 1,
						rowPos - 1, columnPos - 2, columnPos - 1));
			}
		}

		public void createStaffRows() {
			for (Iterator<String> i = fullStaff.keySet().iterator(); i
					.hasNext();) {
				columnPos = 0;
				String staffKey = i.next();
				row = sheet.createRow(rowPos++);
				cell = row.createCell(columnPos++);
				cell.setCellValue(fullStaff.get(staffKey));
				if(rowPos % 2 == 0 ){
					cell.setCellStyle(cellStyleTableRowLeft);
				} else {
					cell.setCellStyle(cellStyleTableRowLeftOdd);
				}
				//row.createCell(columnPos++).setCellValue(fullStaff.get(staffKey));
				System.out.println(staffKey); //TODO Delete
				StaffMember staffMemberActiveTime = staffActiveTime.get(staffKey);
				StaffMember staffMemberOracle = staffOracle.get(staffKey);
				
				createTotalCells(staffMemberActiveTime, staffMemberOracle);
				for(Iterator<String> j = projects.keySet().iterator(); j.hasNext(); ){
					String projectKey = j.next();
					
					createTimeCells(projectKey, staffMemberActiveTime, staffMemberOracle);
				}
				
				
				
			}

		}
		
		private void createTotalCells(StaffMember staffMemberActiveTime, StaffMember staffMemberOracle){
			float timeActiveTime = 0;
			float timeOracle = 0;
			CellStyle cellStyleLeft;
			CellStyle cellStyleRight;
			if(staffMemberActiveTime != null){
				timeActiveTime = staffMemberActiveTime.getTotalTime();
			}
			if(staffMemberOracle != null){
				timeOracle = staffMemberOracle.getTotalTime();
			}
			if(timeActiveTime == timeOracle){
				if(rowPos % 2 == 0 ){
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
			//row.createCell(columnPos++).setCellValue(timeActiveTime);
			//row.createCell(columnPos++).setCellValue(timeOracle);
		}
		
		private void createTimeCells(String projectKey, StaffMember staffMemberActiveTime, StaffMember staffMemberOracle){
			float timeActiveTime = 0;
			float timeOracle = 0;
			CellStyle cellStyleLeft;
			CellStyle cellStyleRight;
			if(staffMemberActiveTime != null){
				Project project = staffMemberActiveTime.getProjects().get(projectKey);
				if(project != null){
					timeActiveTime = project.getTime();
				}
			}
			if(staffMemberOracle != null){
				Project project = staffMemberOracle.getProjects().get(projectKey);
				if(project != null){
					timeOracle = project.getTime();
				}
			}
			if(timeActiveTime == timeOracle){
				if(rowPos % 2 == 0 ){
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
			//row.createCell(columnPos++).setCellValue(timeActiveTime);
			//row.createCell(columnPos++).setCellValue(timeOracle);
		}
		
		private CellStyle createCellStyle(short boldWeight, short backgroundColor, short... border){
			// Create a new font and alter it.
		    Font font = workbook.createFont();
		    font.setFontHeightInPoints(FONT_SIZE);
		    font.setFontName("FONT_NAME");
		    font.setBoldweight(boldWeight);
		    // Create a new style
			CellStyle style = workbook.createCellStyle();
			style.setFont(font);
			// style.setFillBackgroundColor(backgroundColor);
			//style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
			style.setFillForegroundColor(backgroundColor);
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			// Borders
			style.setBorderBottom(border[0]);
		    style.setBorderLeft(border[1]);
		    style.setBorderRight(border[2]);
		    style.setBorderTop(border[3]);
			
			
			return style;
		}
		
		private void setStyles(){
			cellStyleNormal = createCellStyle(Font.BOLDWEIGHT_NORMAL, IndexedColors.WHITE.getIndex(), CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE);
			cellStyleTitle = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.WHITE.getIndex(), CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM);
			cellStyleTableRowLeft = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.WHITE.getIndex(), CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN);
			cellStyleTableRowRight = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.WHITE.getIndex(), CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN);
			cellStyleTableRowLeftOdd = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.LIGHT_TURQUOISE.getIndex(), CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN);
			cellStyleTableRowRightOdd = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.LIGHT_TURQUOISE.getIndex(), CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_THIN);
			cellStyleTableRowError = createCellStyle(Font.BOLDWEIGHT_BOLD, IndexedColors.MAROON.getIndex(), CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN);
		}
	}
}
