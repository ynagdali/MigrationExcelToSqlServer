package excelreader.sheets;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;

public class SubAccountSheet extends Sheets {

	private static Logger logger = Logger.getLogger(SubAccountSheet.class);
	ApplicationProperties properties = null;

	public SubAccountSheet(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processSubaccountSheet(Sheet sheet) {
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
					paramteres.put(columnNumber, "" + ITNName);
				}

				else if (columnNumber > 1 && columnNumber <= 3) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					paramteres.put(columnNumber, cellValue);
				} else if (columnNumber > 3 && columnNumber <= 12) {
					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
					} catch (ClassCastException e) {
						cellValue = generateCellValue(cell, dataFormatter);
					}
					paramteres.put(columnNumber, cellValue);
				} else {
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
		String query = "INSERT INTO \r\n" + properties.getProperty("app.table." + this.getClass().getSimpleName())
				+ "           ([ITNName]\r\n" + "           ,[Subaccount Group]\r\n" + "           ,[Concat]\r\n"
				+ "           ,[Values Sum of Monthly Budget]\r\n" + "           ,[Sum of Month Actual]\r\n"
				+ "           ,[Sum of Month Variance]\r\n" + "           ,[Sum of YTD Budget]\r\n"
				+ "           ,[Sum of YTD Actual]\r\n" + "           ,[Sum of YTD Variance]\r\n"
				+ "           ,[Sum of YE Budget]\r\n" + "           ,[Sum of YE Forecast]\r\n"
				+ "           ,[Sum of YE Variance])\r\n" + "     \r\n" + "           values(?,?,?,?,?,?,?,?,?,?,?,?)";

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
		String subAccountNameFromMap = (String) parameters.get(2);
		logger.debug("itnNameFromMap-->" + itnNameFromMap);
		logger.debug("subAccountNameFromMap-->" + subAccountNameFromMap);
		if (itnNameFromMap == null) {
			logger.info("Code Inside here intNameFromMap is null");
			parameters.put(1, columnOneValue);
			itnNameFromMap = columnOneValue;
		}

		if (itnNameFromMap.contains("Account Category") || itnNameFromMap.contains("YTD Variance")
				|| itnNameFromMap.contains("Total") || itnNameFromMap.contains("ITNName")) {
			logger.info("code inside this if");
		}

		else if (("#N/A").equalsIgnoreCase(subAccountNameFromMap)) {

		} else if (itnNameFromMap != null) {
			if (itnNameFromMap.contains("Total")) {

			} else
				insertDataForSubAccountSheet(parameters);
		}

	}

}
