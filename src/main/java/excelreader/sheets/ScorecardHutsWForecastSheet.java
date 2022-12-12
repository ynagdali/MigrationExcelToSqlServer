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

public class ScorecardHutsWForecastSheet extends Sheets {

	private static Logger logger = Logger.getLogger(SubAccountSheet.class);
	ApplicationProperties properties = null;

	public ScorecardHutsWForecastSheet(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processScorecardHutsWForecastSheet(Sheet sheet) {
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
				logger.debug("column number" + columnNumber);
				if (columnNumber == 1) {
					ITNName = dataFormatter.formatCellValue(cell);
					paramteres.put(columnNumber, "" + ITNName);

				}

				else if (("").equalsIgnoreCase(ITNName)) {
					break;
				}

				/*
				 * else if (columnNumber > 1 && columnNumber <= 3) { Object cellValue =
				 * generateCellValue(cell, dataFormatter); paramteres.put(columnNumber,
				 * cellValue); }
				 */else if (columnNumber > 1 && columnNumber <= 13) {
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
		String query = "INSERT INTO [dbo].\r\n" + properties.getProperty("app.table." + this.getClass().getSimpleName())
				+ "           ([ITN Name/Type]\r\n" + "           ,[Jan]\r\n" + "           ,[Feb]\r\n"
				+ "           ,[Mar]\r\n" + "           ,[MTD Budget/Apr]\r\n" + "           ,[MTD Actual/May]\r\n"
				+ "           ,[MTD Variance/Jun]\r\n" + "           ,[YTD Budget/Jul]\r\n"
				+ "           ,[YTD Actual/Aug]\r\n" + "           ,[YTD Variance/Sep]\r\n"
				+ "           ,[YE Budget/Oct]\r\n" + "           ,[YE Actual/Nov]\r\n"
				+ "           ,[YE Variance/Dec])\r\n" + "     VALUES\r\n" + "           (?,?,?,?,?,?,?,?,?,?,?,?,?) ";

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

		logger.debug("itnNameFromMap-->" + itnNameFromMap);
		if (itnNameFromMap == null) {
			logger.info("Code Inside here intNameFromMap is null");
			itnNameFromMap = "";
		}

		if (itnNameFromMap.equalsIgnoreCase("Report Month:") || itnNameFromMap.equalsIgnoreCase("ITN Name")
			|| itnNameFromMap.equalsIgnoreCase("REACTS Portfolio Summary")) {
			logger.info("code inside this if");
		}

		else if (("").equalsIgnoreCase(itnNameFromMap)) {

		} else {
			insertDataForSubAccountSheet(parameters);
		}
	}

}
