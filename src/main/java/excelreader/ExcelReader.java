package excelreader;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;
import excelreader.sheets.MonthlySheetTransposed;
import excelreader.sheets.ScorecardHutsWForecastSheet;
import excelreader.sheets.ScorecardOverallwForecastITNNameSheet;
import excelreader.sheets.ScorecardOverallwForecastOverallSheet;
import excelreader.sheets.SubAccountKeySheet;
import excelreader.sheets.AssignedPmSheet;
import excelreader.sheets.DumpReportSheet;
import excelreader.sheets.MileageByPdByMonthSheet;
import excelreader.sheets.MileageByPdByMonthSheetTransposed;
import excelreader.sheets.MonthlySheet;
import excelreader.sheets.SubAccountSheet;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelReader {

	private static Logger logger = Logger.getLogger(ExcelReader.class);

	public static void main(String[] args) throws IOException, InvalidFormatException {

		long start = 0;
		long end = 0;
		long completeStart = 0;
		long completeEnd = 0;

		completeStart = System.currentTimeMillis();
		int failureCount = 0;
		SimpleDateFormat formatter = new SimpleDateFormat("dd_MMM_yyyy HH_mm_ss");
		Date date = new Date();
		String currentDateAndTime = formatter.format(date);

		ApplicationProperties properties = new ApplicationProperties();

		// dump file started
		String SAMPLE_DUMP_XLSX_FILE_PATH = properties.getProperty("app.dumpFileName");
		File dumpXlsxFile = null;
		try {
			dumpXlsxFile = new File(SAMPLE_DUMP_XLSX_FILE_PATH + ".xlsx");
		} catch (Exception e) {
			logger.debug("Unable to read file -->" + SAMPLE_DUMP_XLSX_FILE_PATH + " error->" + e.toString());
			return;
		}
		Workbook workbook2 = WorkbookFactory.create(dumpXlsxFile);

		start = System.currentTimeMillis();
		Sheet dumpSheet = workbook2.getSheet("Sheet1");
		failureCount = new DumpReportSheet(properties).processDumpAccountSheet(dumpSheet); // dumpaccount
		end = System.currentTimeMillis();
		logger.debug("Failures in dump sheet  ->" + failureCount + "-< Time taken->" + (end - start) + " mili seconds");

		workbook2.close();

		// financials file started
		String SAMPLE_XLSX_FILE_PATH = properties.getProperty("app.fileName");

		logger.info("Inside Main Method for file->" + SAMPLE_XLSX_FILE_PATH);

		File xlsxFile = null;
		Workbook workbook = null;
		try {
			xlsxFile = new File(SAMPLE_XLSX_FILE_PATH + ".xlsx");
			workbook = WorkbookFactory.create(xlsxFile);
		} catch (Exception e) {

			logger.debug("Unable to read file -->" + SAMPLE_XLSX_FILE_PATH + " exception->" + e.toString());
			Connection connection = SqlServerConnection.getConnection();
			if (connection != null) {
				SqlServerConnection.closeConnection(connection);
				logger.info("Connection closed inside exception block");
			}
			return;
		}

		start = System.currentTimeMillis();
		Sheet subAccountSheet = workbook.getSheet("Subaccount");
		failureCount = new SubAccountSheet(properties).processSubaccountSheet(subAccountSheet);
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet Subaccount  ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		Sheet monthlySheet = workbook.getSheet("Monthly");
		start = System.currentTimeMillis();
		failureCount = new MonthlySheetTransposed(properties).processMonthlySheetTransposed(monthlySheet); // Monthly
																											// pivot
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet monthlySheet pivoted table ->" + failureCount + "-< Time taken->"
				+ (end - start) + " mili seconds");

		start = System.currentTimeMillis();
		failureCount = new MonthlySheet(properties).processMonthlySheet(monthlySheet); // monthly sheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet monthlySheet normal table ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		Sheet milegeByPdByMonthSheet = workbook.getSheet("Mileage by PD by Month");
		start = System.currentTimeMillis();
		failureCount = new MileageByPdByMonthSheet(properties).processMileageByPdByMonthSheet(milegeByPdByMonthSheet); // monthly
																														// sheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet milegeByPdByMonthSheet  ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		start = System.currentTimeMillis();
		failureCount = new MileageByPdByMonthSheetTransposed(properties)
				.processMileageByPdByMonthSheetTransposed(milegeByPdByMonthSheet); // monthly sheet transposed
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet milegeByPdByMonthSheet transposed  ->" + failureCount + "-< Time taken->"
				+ (end - start) + " mili seconds");

		Sheet subAccountKeySheet = workbook.getSheet("Subaccount Key");
		start = System.currentTimeMillis();
		failureCount = new SubAccountKeySheet(properties).processSubaccountKeySheet(subAccountKeySheet); // subAccountKeySheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet subAccountKeySheet  ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		Sheet scorecardHutsSheet = workbook.getSheet("Scorecard Huts w Forecast");
		start = System.currentTimeMillis();
		failureCount = new ScorecardHutsWForecastSheet(properties)
				.processScorecardHutsWForecastSheet(scorecardHutsSheet); // scorecardHutsSheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet scorecardHutsSheet  ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		Sheet scorecardOverallwForecastITNNameSheet = workbook.getSheet("Scorecard Overall w Forecast");
		start = System.currentTimeMillis();
		failureCount = new ScorecardOverallwForecastITNNameSheet(properties)
				.processScorecardOverallwForecastITNNameSheet(scorecardOverallwForecastITNNameSheet); // scorecardOverallwForecastITNNameSheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet scorecardOverallwForecastITNNameSheet  ->" + failureCount + "-< Time taken->"
				+ (end - start) + " mili seconds");

		Sheet scorecardOverallwForecastOverallSheet = workbook.getSheet("Scorecard Overall w Forecast");
		start = System.currentTimeMillis();
		failureCount = new ScorecardOverallwForecastOverallSheet(properties)
				.processScorecardOverallwForecastOverallSheet(scorecardOverallwForecastOverallSheet); // scorecardOverallwForecastOverallSheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet scorecardOverallwForecastOverallSheet  ->" + failureCount + "-< Time taken->"
				+ (end - start) + " mili seconds");

		Sheet assignedPmSheet = workbook.getSheet("Assigned PM");
		start = System.currentTimeMillis();
		failureCount = new AssignedPmSheet(properties).processAssignedPmSheet(assignedPmSheet); // assignedPmSheet
		end = System.currentTimeMillis();
		logger.debug("Failures in sheet assignedPmSheet  ->" + failureCount + "-< Time taken->" + (end - start)
				+ " mili seconds");

		workbook.close();

		String processedFile = SAMPLE_XLSX_FILE_PATH + "_processedat_" + currentDateAndTime + ".xlsx";
		logger.debug("processedFile-->" + processedFile);
		File xlsxFileRenamed = new File(processedFile);
		boolean rename = xlsxFile.renameTo(xlsxFileRenamed);
		if (rename) {
			logger.info("File Renamed ----");
		} else {
			logger.info("File Not Renamed ----");
		}

		Connection connection = SqlServerConnection.getConnection();
		boolean close = SqlServerConnection.closeConnection(connection);
		if (close) {
			logger.info("Connection closed successfully");
		} else {
			logger.info("Connection not closed");
		}

		completeEnd = System.currentTimeMillis();
		logger.debug("Time taken to execute->" + (completeEnd - completeStart) + " mili seconds");

		System.out.println("Hi application terminated");
	}

}
