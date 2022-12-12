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

public class MileageByPdByMonthSheetTransposed extends Sheets {

	String columnOneValue = "";

	private static Logger logger = Logger.getLogger(MonthlySheetTransposed.class);
	HashMap<Integer, Object> parameters = new HashMap<Integer, Object>();
	String ITNName = "";

	ApplicationProperties properties = null;

	public MileageByPdByMonthSheetTransposed(ApplicationProperties properties) {
		this.properties = properties;
	}

	public int processMileageByPdByMonthSheetTransposed(Sheet sheet) {
		DataFormatter dataFormatter = new DataFormatter();
		truncateTable(this.getClass().getSimpleName(), properties);
		for (Row row : sheet) {
			int rowNumber = row.getRowNum() + 1;

			
			Cell cell = row.getCell(2);

			if (cell != null) {
				
				ITNName = dataFormatter.formatCellValue(cell);
				logger.info("ITNName-->"+ITNName);
				if (ITNName.equalsIgnoreCase("ITN") || ITNName.equalsIgnoreCase("C3") || ITNName.equalsIgnoreCase("C21")  ) {
					logger.info("code here for putting cell 2 value");
					logger.info("Code inside row iterator columnTwoValue->" + columnOneValue);
					columnOneValue = generateCellValueForPDByMonth(columnOneValue);
					logger.info("columnOneValue -->" + columnOneValue);
				}
			}
			logger.debug("rownumber-->" + rowNumber);
			logger.debug("ITNName-------->" + ITNName);

			boolean totalRecordsReached = populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jan", 6);

			if (totalRecordsReached) {
				break;
			}

			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Feb", 7);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Mar", 8);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Apr", 9);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "May", 10);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jun", 11);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Jul", 12);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Aug", 13);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Sep", 14);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Oct", 15);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Nov", 16);
			populateInsertValuesForMonthlyBudget(row, dataFormatter, "Dec", 17);

		}
		return failureCount;
	}

	private boolean populateInsertValuesForMonthlyBudget(Row row, DataFormatter dataFormatter, String month,
			int columnNumberBudget) {

		boolean flag = false;
		parameters = populateValuesForMonthlyBudget(row, dataFormatter, month, columnNumberBudget);

		logger.info("here inside popolate data verify populated data");
		parameters.forEach((key, value) -> logger.info(key + "======= " + value));
		validateRowsAndInsertInDatabase(parameters);

		Optional<String> optionalForValueAtColumn1 = Optional.ofNullable((String) parameters.get(1));

		parameters.clear();
		if (optionalForValueAtColumn1.isPresent()
				&& optionalForValueAtColumn1.get().trim().equalsIgnoreCase("Grand Total")) {
			return true;
		}
		return flag;
	}

	private HashMap<Integer, Object> populateValuesForMonthlyBudget(Row row, DataFormatter dataFormatter, String month,
			int columnNumberBudget) {
		HashMap<Integer, Object> paramteres = new HashMap<Integer, Object>();

		for (int j = 0; j < 18; j++) {
			Cell cell = row.getCell(j);

			if (cell != null) {
				int columnNumber = cell.getColumnIndex() + 1;
				if (columnNumber == 1) {
					Object cellValue = generateCellValue(cell, dataFormatter);
					if (cellValue.equals("") || cellValue.equals(null)) {
						break;
					}
				}

				if (columnNumber == 3) {
					logger.info("code inside first if edited for moneth transpose");
					logger.info(dataFormatter.formatCellValue(cell));
					ITNName = dataFormatter.formatCellValue(cell);
					paramteres.put(columnNumber - 1, "" + ITNName);
					paramteres.put(3, month);
					if (ITNName.contains("Account Category") || ITNName.contains("YTD Variance")
							|| ITNName.contains("Total") || ITNName.contains("ITNName")) {
						break;

					}

				} else if (columnNumber == columnNumberBudget) {
					Object cellValue = "";
					try {
						cellValue = new BigDecimal((Double) generateCellValue(cell, dataFormatter));
					} catch (ClassCastException e) {
						cellValue = generateCellValue(cell, dataFormatter);
					}
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
			itnNameFromMap = (String) parameters.get(2);
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
			logger.info("code inside this validation for ITN");
		}

		else if (itnNameFromMap.contains("Account Category") || itnNameFromMap.contains("YTD Variance")
				|| itnNameFromMap.contains("Total") || itnNameFromMap.contains("ITNName")
				|| itnNameFromMap.equalsIgnoreCase("C3") || itnNameFromMap.equalsIgnoreCase("C21")) {
			logger.info("code inside this validation for different sceanarios");
		} else {
			parameters.put(1, columnOneValue);
			try {
				BigDecimal number = (BigDecimal) parameters.get(4);
				if (number == null) {
					parameters.put(4, "0.0");
				}
			} catch (Exception e) {
				logger.debug("inside exception block for coverting value for parameter 4-->" + e.toString());
			}

			insertDataForMonthlySheet(parameters);
		}
	}

	public String insertQueryForMonthlySheet() {
		String query = "INSERT INTO [dbo].\r\n" + properties.getProperty("app.table." + this.getClass().getSimpleName())
				+ "           ([catg]\r\n" + "           ,[pd]\r\n" + "           ,[mont]\r\n"
				+ "           ,[miles])\r\n" + "     VALUES (?,?,?,?)  ";
		return query;

	}

}
