package excelreader.sheets;



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

	public class AssignedPmSheet extends Sheets {

		private static Logger logger = Logger.getLogger(SubAccountSheet.class);
		ApplicationProperties properties = null;

		public AssignedPmSheet(ApplicationProperties properties) {
			this.properties = properties;
		}

		public int processAssignedPmSheet(Sheet sheet) {
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

					else if (columnNumber >= 2 || columnNumber <= 6) {
						
						RichTextString cellValueRich = cell.getRichStringCellValue();
						String cellValue = cellValueRich.getString();
						//Object cellValue = generateCellValue(cell, dataFormatter);
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
			String query = "INSERT INTO [dbo].\r\n" + properties.getProperty("app.table."+this.getClass().getSimpleName()) 
					+ " ([WPT Code]\r\n"
					+ "           ,[PD Number]\r\n"
					+ "           ,[PD Name]\r\n"
					+ "           ,[Program Manager]\r\n"
					+ "           ,[Project Manager]\r\n"
					+ "           ,[LRE])\r\n"
					+ "     VALUES\r\n"
					+ "           (?,?,?,?,?,? ) \r\n" ;
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
				//SqlServerConnection.closeConnection(connection);
			}

		}

		public void validateRowsAndInsertInDatabase(HashMap<Integer, Object> parameters, String columnOneValue) {

			String itnNameFromMap = (String) parameters.get(1);
			//String subAccountNameFromMap = (String) parameters.get(2);
			logger.debug("itnNameFromMap-->" + itnNameFromMap);
			if (itnNameFromMap == null) {
				logger.info("Code Inside here intNameFromMap is null");
				parameters.put(1, columnOneValue);
				itnNameFromMap = columnOneValue;
			}

			if (itnNameFromMap.equalsIgnoreCase("WPT Code")) {
				logger.info("code inside this if");
			}

			else if (itnNameFromMap != null) {
							insertDataForSubAccountSheet(parameters);
			}

		}

	}

