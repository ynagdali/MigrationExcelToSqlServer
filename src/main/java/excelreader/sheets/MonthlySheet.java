package excelreader.sheets;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.HashMap;
import java.util.Optional;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;

public class MonthlySheet extends Sheets {

	private static Logger logger = Logger.getLogger(MonthlySheet.class);

	ApplicationProperties properties = null;

	public MonthlySheet(ApplicationProperties properties) {
		this.properties = properties;
	}
	
	public int processMonthlySheet(Sheet sheet) {
		DataFormatter dataFormatter = new DataFormatter();
		truncateTable(this.getClass().getSimpleName(), properties);
		String ITNName = "";
		HashMap<Integer, Object> paramteres = new HashMap<Integer, Object>();

		for (Row row : sheet) {
			int rowNumber = row.getRowNum() + 1;
			logger.debug("rownumber-->" + rowNumber);
			logger.debug("ITNName-------->" + ITNName);

			/*
			 * if (rowNumber <= 5) continue; if (rowNumber > 7) // for 150 values break;
			 */

			for (Cell cell : row) {
				int columnNumber = cell.getColumnIndex() + 1;
				if (columnNumber == 1) {
					logger.info("code inside first if edited");
					logger.info(dataFormatter.formatCellValue(cell));
					ITNName = dataFormatter.formatCellValue(cell);
					paramteres.put(columnNumber, "" + ITNName);
				}

				else if (columnNumber > 1 && columnNumber <= 25) {

					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
					} catch (ClassCastException e) {
						cellValue = generateCellValue(cell, dataFormatter);
					}
					paramteres.put(columnNumber, cellValue);
				} else {
					logger.debug("code inside here column number->" + columnNumber);
					break;
				}

			}
			
			Optional<String> optionalForValueAtColumn1 = Optional.ofNullable((String) paramteres.get(1));
			
			if (optionalForValueAtColumn1.isPresent() && optionalForValueAtColumn1.get().trim().equalsIgnoreCase("Grand Total"))
			{
				break;
			}
			
			validateRowsAndInsertInDatabase(paramteres, ITNName);
			paramteres.clear();
		}
		return failureCount;

	}

	public String insertQueryForMonthlySheet() {
		String query = "INSERT INTO [dbo].\r\n"  + properties.getProperty("app.table."+this.getClass().getSimpleName())
	+ "           ([ITNName]\r\n" + "           ,[Jan]\r\n"
				+ "           ,[Feb]\r\n" + "           ,[Mar]\r\n" + "           ,[Apr]\r\n" + "           ,[May]\r\n"
				+ "           ,[Jun]\r\n" + "           ,[Jul]\r\n" + "           ,[Aug]\r\n" + "           ,[Sep]\r\n"
				+ "           ,[Oct]\r\n" + "           ,[Nov]\r\n" + "           ,[Dec]\r\n"
				+ "           ,[Forecast Jan]\r\n" + "           ,[Forecast Feb]\r\n" + "           ,[Forecast Mar]\r\n"
				+ "           ,[Forecast Apr]\r\n" + "           ,[Forecast May]\r\n" + "           ,[Forecast Jun]\r\n"
				+ "           ,[Forecast Jul]\r\n" + "           ,[Forecast Aug]\r\n" + "           ,[Forecast Sep]\r\n"
				+ "           ,[Forecast Oct]\r\n" + "           ,[Forecast Nov]\r\n"
				+ "           ,[Forecast Dec])\r\n" + "     VALUES\r\n"
				+ "           (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) ";

		return query;

	}

	public void insertDataForMonthlySheet(HashMap<Integer, Object> hm) {

		Connection connection = SqlServerConnection.getConnection();
		String query = insertQueryForMonthlySheet();

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
			logger.debug("failureCOunt inside exception---->>>>" +exp.toString()+" \n failuer count-->"+ failureCount);
		} finally {
			//SqlServerConnection.closeConnection(connection);
		}

	}

	public void validateRowsAndInsertInDatabase(HashMap<Integer, Object> parameters, String columnOneValue) {
		String itnNameFromMap = (String) parameters.get(1);
		logger.debug("itnNameFromMap-->" + itnNameFromMap);
		if (itnNameFromMap == null) {
			logger.info("here--->");
			parameters.put(1, columnOneValue);
			itnNameFromMap = columnOneValue;
		}

		if (itnNameFromMap.contains("Account Category") || itnNameFromMap.contains("YTD Variance")
				|| itnNameFromMap.contains("Total") || itnNameFromMap.contains("ITNName")
				|| itnNameFromMap.equalsIgnoreCase("") || itnNameFromMap.equalsIgnoreCase("Subaccount Group")
				|| itnNameFromMap.equalsIgnoreCase("#N/A")) {
			logger.info("code inside this if");
		}

		else if (itnNameFromMap != null) {
			if (itnNameFromMap.contains("Total")) {

			} else
				insertDataForMonthlySheet(parameters);
		}

	}

}
