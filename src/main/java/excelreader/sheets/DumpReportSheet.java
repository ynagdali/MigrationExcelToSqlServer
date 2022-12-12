package excelreader.sheets;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.Date;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;

public class DumpReportSheet extends Sheets {

	private static Logger logger = Logger.getLogger(SubAccountSheet.class);
	ApplicationProperties properties = null;

	public DumpReportSheet(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processDumpAccountSheet(Sheet sheet) {
		DataFormatter dataFormatter = new DataFormatter();
		String ITNName = "";
		HashMap<Integer, Object> paramteres = new HashMap<Integer, Object>();
		truncateTable(this.getClass().getSimpleName(), properties);

		for (Row row : sheet) {
			int rowNumber = row.getRowNum() + 1;
			logger.debug("rownumber-->" + rowNumber);
			logger.debug("ITNName-------->" + ITNName);
			
			for (Cell cell : row) {

				int columnNumber = cell.getColumnIndex() + 1;
				if (columnNumber == 1) {
					ITNName = dataFormatter.formatCellValue(cell);
					if (ITNName.equalsIgnoreCase("Portfolio")) {
						break;
					}
					paramteres.put(columnNumber, "" + ITNName);
				}

				else if (columnNumber > 1 && columnNumber <= 33) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					paramteres.put(columnNumber, cellValue);
				}

				else if (columnNumber > 33 && columnNumber <= 46) {

					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
					} catch (ClassCastException e) {
						cellValue = generateCellValue(cell, dataFormatter);
					}
					paramteres.put(columnNumber, cellValue);
					// RichTextString cellValueRich = cell.getRichStringCellValue();
					// String cellValue = cellValueRich.getString();
					// Object cellValue = generateCellValue(cell, dataFormatter);
					// paramteres.put(columnNumber, cellValue);
				}

				else if (columnNumber > 46 && columnNumber <= 49) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					paramteres.put(columnNumber, cellValue);
				}
				else if (columnNumber == 50) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					Date obj = (Date) cellValue;
					java.sql.Timestamp timestamp =  new java.sql.Timestamp(obj.getTime());
					String cellValueAfterCOnversion = timestamp.toString();
					paramteres.put(columnNumber, cellValueAfterCOnversion);
				}
				else if (columnNumber > 50 && columnNumber <= 63) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					paramteres.put(columnNumber, cellValue);
				}

				else {
					logger.debug("code inside else column number" + columnNumber);
					break;
				}

			}

			validateRowsAndInsertInDatabase(paramteres, ITNName);
			paramteres.clear();
		}
		return failureCount;
	}

	public String insertQueryForSubAccountSheet() {
		String query = "INSERT INTO [dbo]. \r\n"
				+ properties.getProperty("app.table." + this.getClass().getSimpleName()) + "([Portfolio]\r\n"
				+ "           ,[Department ID]\r\n" + "           ,[Department]\r\n" + "           ,[Level 1]\r\n"
				+ "           ,[Level 2]\r\n" + "           ,[Level 3]\r\n" + "           ,[Level 4]\r\n"
				+ "           ,[Level 5]\r\n" + "           ,[Level 6]\r\n" + "           ,[Level 7]\r\n"
				+ "           ,[Level 8]\r\n" + "           ,[Level 9]\r\n" + "           ,[Level 10]\r\n"
				+ "           ,[Department Active Flag]\r\n" + "           ,[Account Type]\r\n"
				+ "           ,[Account Category]\r\n" + "           ,[Cost Type]\r\n" + "           ,[ITN Product]\r\n"
				+ "           ,[System]\r\n" + "           ,[Investment Category]\r\n"
				+ "           ,[Executive Category]\r\n" + "           ,[Centrally Managed]\r\n"
				+ "           ,[ITN Plan Type]\r\n" + "           ,[ITN]\r\n" + "           ,[ITNName]\r\n"
				+ "           ,[ITN Owner]\r\n" + "           ,[ITN Status]\r\n"
				+ "           ,[Level 2 Subaccount]\r\n" + "           ,[Level 3 Subaccount]\r\n"
				+ "           ,[Subaccount]\r\n" + "           ,[Labor Resource Type]\r\n"
				+ "           ,[Financial Plan Type]\r\n" + "           ,[Year]\r\n" + "           ,[Jan]\r\n"
				+ "           ,[Feb]\r\n" + "           ,[Mar]\r\n" + "           ,[Apr]\r\n" + "           ,[May]\r\n"
				+ "           ,[Jun]\r\n" + "           ,[Jul]\r\n" + "           ,[Aug]\r\n" + "           ,[Sept]\r\n"
				+ "           ,[Oct]\r\n" + "           ,[Nov]\r\n" + "           ,[Dec]\r\n"
				+ "           ,[Total]\r\n" + "           ,[Max Risk]\r\n" + "           ,[Discretionary]\r\n"
				+ "           ,[Commitments]\r\n" + "           ,[In-Service Date]\r\n" + "           ,[Tax Repair]\r\n"
				+ "           ,[Base/Major Project and IT Tower]\r\n" + "           ,[Program and IT Portfolio]\r\n"
				+ "           ,[PPM Number]\r\n" + "           ,[Need Date]\r\n"
				+ "           ,[PHISCO ITN Association]\r\n" + "           ,[Budget Category]\r\n"
				+ "           ,[PHI Sponsor Group]\r\n" + "           ,[PHI Category Lead]\r\n"
				+ "           ,[PJM Project Type]\r\n" + "           ,[PJM Project Number]\r\n"
				+ "           ,[TEAC Date]\r\n" + "           ,[Servicenow])\r\n"
				+ "     VALUES(?,?,?,?,?,?,?,?,?,?,\r\n" + "	 ?,?,?,?,?,?,?,?,?,?,\r\n"
				+ "	 ?,?,?,?,?,?,?,?,?,?,\r\n" + "	 ?,?,?,?,?,?,?,?,?,?,\r\n" + "	 ?,?,?,?,?,?,?,?,?,?,\r\n"
				+ "	 ?,?,?,?,?,?,?,?,?,?,\r\n" + "	 ?,?,?) ";

		return query;

	}

	public void insertDataForSubAccountSheet(HashMap<Integer, Object> hm) {

		Connection connection = SqlServerConnection.getConnection();
		String query = insertQueryForSubAccountSheet();

		logger.info("\n\n\n\nvalues after hm");
		hm.forEach((key, value) -> logger.info(key + " = " + value));

		try {
			PreparedStatement preparedStatement;
			preparedStatement = SqlServerConnection.prepareAStatement(connection, query, hm);
			int execution = preparedStatement.executeUpdate();
			if (execution == 1) {
				logger.info("Successfully inserted");
			} else {
				failureCount++;
				logger.debug("failureCOunt---->>>>" + failureCount);

			}
		} catch (Exception exp) {
			failureCount++;
			logger.debug(
					"failureCOunt inside exception---->>>>" + exp.toString() + " \n failuer count-->" + failureCount);
		} finally {
			// SqlServerConnection.closeConnection(connection);
		}

	}

	public void validateRowsAndInsertInDatabase(HashMap<Integer, Object> parameters, String columnOneValue) {

		String itnNameFromMap = (String) parameters.get(1);
		// String subAccountNameFromMap = (String) parameters.get(2);
		logger.debug("itnNameFromMap-->" + itnNameFromMap);
		if (itnNameFromMap == null) {
			logger.info("Code Inside here intNameFromMap is null");
			parameters.put(1, columnOneValue);
			itnNameFromMap = columnOneValue;
		}

		if (itnNameFromMap.equalsIgnoreCase("Portfolio")) {
			logger.info("code inside this if");
		}

		else if (itnNameFromMap != null) {
			insertDataForSubAccountSheet(parameters);
		}

	}

}
