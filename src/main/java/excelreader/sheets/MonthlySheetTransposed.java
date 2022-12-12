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

public class MonthlySheetTransposed extends Sheets {
	
	private static Logger logger = Logger.getLogger(MonthlySheetTransposed.class);
	HashMap<Integer, Object> parameters = new HashMap<Integer, Object>();
	String ITNName = "";
	
	ApplicationProperties properties = null;

	public MonthlySheetTransposed(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processMonthlySheetTransposed(Sheet sheet) {
		DataFormatter dataFormatter = new DataFormatter();
		truncateTable(this.getClass().getSimpleName(), properties);
		for (Row row : sheet) {
			int rowNumber = row.getRowNum() + 1;
			logger.debug("rownumber-->" + rowNumber);
			logger.debug("ITNName-------->" + ITNName);

			boolean totalRecordsReached = populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jan", 2, 14);
			
			if(totalRecordsReached)
			{
				break;
			}
			
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Feb", 3, 15);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Mar", 4, 16);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Apr", 5, 17);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "May", 6, 18);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jun", 7, 19);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jul", 8, 20);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Aug", 9, 21);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Sep", 10, 22);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Oct", 11, 23);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Nov", 12, 24);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Dec", 13, 25);

		}
		return failureCount;
	}

	private boolean populateInsertValuesForMonthlyBudget(Row row, DataFormatter dataFormatter, String month,
			int columnNumberBudget, int columnNumberForecastBudget) {

		boolean flag = false;
		parameters = populateValuesForMonthlyBudget(row, dataFormatter, month, columnNumberBudget,
				columnNumberForecastBudget);

		logger.info("here inside popolate data verify populated data");
		parameters.forEach((key, value) -> logger.info(key + "======= " + value));
		validateRowsAndInsertInDatabase(parameters, ITNName);
		
		//String itnNameFromMap = (String) parameters.get(1);  
		
		Optional<String> optionalForValueAtColumn1 = Optional.ofNullable((String) parameters.get(1));
		
		parameters.clear();
		if (optionalForValueAtColumn1.isPresent() && optionalForValueAtColumn1.get().trim().equalsIgnoreCase("Grand Total"))
		{
			return true;
		}
		return flag;
	}

	private HashMap<Integer, Object> populateValuesForMonthlyBudget(Row row, DataFormatter dataFormatter, String month,
			int columnNumberBudget, int columnNumberForecastBudget) {
		HashMap<Integer, Object> paramteres = new HashMap<Integer, Object>();
		for (Cell cell : row) {
			int columnNumber = cell.getColumnIndex() + 1;
			if (columnNumber == 1) {
				logger.info("code inside first if edited for moneth transpose");
				logger.info(dataFormatter.formatCellValue(cell));
				ITNName = dataFormatter.formatCellValue(cell);
				paramteres.put(columnNumber, "" + ITNName);
				paramteres.put(2, month);
				if (ITNName.contains("Account Category") || ITNName.contains("YTD Variance")
						|| ITNName.contains("Total") || ITNName.contains("ITNName"))
					break;

			}

			else if (columnNumber == columnNumberBudget || columnNumber == columnNumberForecastBudget) {
				Object cellValue = "";
				try {
					cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
				} catch (ClassCastException e) {
					cellValue = generateCellValue(cell, dataFormatter);
				}
				if (columnNumber == columnNumberBudget) {
					paramteres.put(3, cellValue);
				} else {
					paramteres.put(4, cellValue);
				}

			}

		}
		return paramteres;
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
		logger.debug("itnNameFromMap-->" + itnNameFromMap + "  <columnOneValue0>" + columnOneValue);
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

	public String insertQueryForMonthlySheet() {
		String query = "INSERT INTO [dbo].\r\n" + properties.getProperty("app.table."+this.getClass().getSimpleName()) 
	+ "           ([ITNName]\r\n"
				+ "           ,[month]\r\n" + "           ,[budget]\r\n" + "           ,[forecast_budget])\r\n"
				+ "     VALUES\r\n" + "           (?,?,?,?); ";
		return query;

	}

}
