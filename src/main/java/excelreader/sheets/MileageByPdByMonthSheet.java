package excelreader.sheets;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;

public class MileageByPdByMonthSheet extends Sheets {

	private static Logger logger = Logger.getLogger(MileageByPdByMonthSheet.class);
	ApplicationProperties properties = null;

	String columnTwoValue = "";

	public MileageByPdByMonthSheet(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processMileageByPdByMonthSheet(Sheet sheet) {
		DataFormatter dataFormatter = new DataFormatter();
		int columnIndex = 0;
		int columnNumber = 0;
		HashMap<Integer, Object> paramteres = new HashMap<Integer, Object>();
		truncateTable(this.getClass().getSimpleName(), properties);

		for (Row row : sheet) {
			int rowNumber = row.getRowNum() + 1;
			logger.debug("rownumber-->" + rowNumber);

			for (int j = 0; j < 18; j++) {
				columnNumber = j + 1;
				columnIndex = columnNumber + 2;
				Cell cell = row.getCell(j);
				if (cell == null) {
					paramteres.put(columnIndex, 0);
					continue;
				}
				if (columnNumber == 1) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					if (cellValue.equals("") || cellValue.equals(null)) {
						break;
					} else {
						paramteres.put(columnIndex, "" + cellValue);
					}
				}

				else if (columnNumber == 2) {
					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
						paramteres.put(columnIndex, cellValue);
					} catch (Exception e) {
						logger.debug("Inside this exception block--> columnNumber == 2" + e.toString() + " cell "
								+ columnIndex);
						cellValue = generateCellValue(cell, dataFormatter);
						logger.debug("cellValue ---->" + cellValue);
						paramteres.put(columnIndex, cellValue);
					}

				}

				else if (columnNumber > 2 && columnNumber <= 5) {
					RichTextString cellValueRich = cell.getRichStringCellValue();
					String cellValue = cellValueRich.getString();
					logger.debug("cellValue ---->" + cellValue);
					paramteres.put(columnIndex, cellValue);
				} else if (columnNumber > 5 && columnNumber <= 18) {
					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
						paramteres.put(columnIndex, cellValue);
					} catch (Exception e) {
						logger.debug("Inside this exception block  columnNumber > 3 && columnNumber <= 18-->"
								+ e.toString() + " cell " + columnIndex + " column number--> " + columnNumber);
						cellValue = generateCellValue(cell, dataFormatter);
						logger.debug("cellValue ---->" + cellValue);
						paramteres.put(columnIndex, cellValue);
					}

				} else {
					logger.debug("code inside else column number" + columnNumber);
					break;
				}

			}

			validateRowsAndInsertInDatabase(paramteres);
			paramteres.clear();
		}
		return failureCount;
	}

	public String insertQueryForSubAccountSheet() {
		String query = "INSERT INTO [dbo].\r\n" + properties.getProperty("app.table." + this.getClass().getSimpleName())
				+ "           ([Budget]\r\n" + "           ,[CATG]\r\n" + "           ,[SNo]\r\n"
				+ "           ,[QTR]\r\n" + "           ,[ITN]\r\n" + "           ,[PD]\r\n" + "           ,[Name]\r\n"
				+ "           ,[Jan]\r\n" + "           ,[Feb]\r\n" + "           ,[Mar]\r\n" + "           ,[Apr]\r\n"
				+ "           ,[May]\r\n" + "           ,[Jun]\r\n" + "           ,[Jul]\r\n" + "           ,[Aug]\r\n"
				+ "           ,[Sep]\r\n" + "           ,[Oct]\r\n" + "           ,[Nov]\r\n" + "           ,[Dec]\r\n"
				+ "           ,[Total])\r\n" + "     VALUES\r\n"
				+ "           (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) ";

		return query;

	}

	public void insertDataForSubAccountSheet(HashMap<Integer, Object> hm) {

		Connection connection = SqlServerConnection.getConnection();
		String query = insertQueryForSubAccountSheet();

		logger.info("\n\n\n\nvalues in hm while inserting");
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

	public void validateRowsAndInsertInDatabase(HashMap<Integer, Object> parameters) {

		logger.info("\n\n\n\nvalues in hm while validating");
		parameters.forEach((key, value) -> logger.info(key + " = " + value));

		String itnNameFromMap = "";
		try {
			itnNameFromMap = (String) parameters.get(5);
		}

		catch (Exception e) {
			logger.debug("Inside this expection block under validateRowsAndInsertInDatabase" + "<<Class>>"
					+ this.getClass() + e.toString());
			return;
		}
		logger.debug("itnNameFromMap-->" + itnNameFromMap);
		if (itnNameFromMap == null) {
			logger.info("Code Inside here intNameFromMap is null");
			return;
		} else if (itnNameFromMap.equalsIgnoreCase("ITN")) {
			logger.info("Code Inside here intNameFromMap is ITN columnTwoValue->" + columnTwoValue);
			columnTwoValue = generateCellValueForPDByMonth(columnTwoValue);

		}

		else if (itnNameFromMap.contains("Account Category") || itnNameFromMap.contains("YTD Variance")
				|| itnNameFromMap.contains("Total") || itnNameFromMap.contains("ITNName")) {
			logger.info("code inside this if");
		} else {
			parameters.put(1, "Budget");
			parameters.put(2, columnTwoValue);
			insertDataForSubAccountSheet(parameters);
		}

	}

}
