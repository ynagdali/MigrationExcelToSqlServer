package excelreader.sheets;

import java.sql.Connection;
import java.sql.PreparedStatement;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

import config.ApplicationProperties;
import dbconnection.SqlServerConnection;


public class Sheets  {

	int failureCount = 0;
	private static Logger logger = Logger.getLogger(Sheets.class);
	
	
	public  Object generateCellValue(Cell cell, DataFormatter formatter) {
		switch (cell.getCellType()) {
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case STRING:
			return cell.getStringCellValue();
			//return cell.getRichStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
				//return cell.getLocalDateTimeCellValue();
			} else {
				return cell.getNumericCellValue();
			}
			
		//case 	
		case FORMULA: {
			switch (cell.getCachedFormulaResultType()) {
			case BOOLEAN:
				return cell.getBooleanCellValue();				
			case NUMERIC:
				return cell.getNumericCellValue();
			case STRING:
				return cell.getRichStringCellValue();
			default:
				return "";
			}
		}
		case BLANK:
			return "";	
		default:
			return "";

		}

	}
	
	
	public String generateCellValueForPDByMonth(String cellValue)
	{
		switch (cellValue) {
		case "": {
			
			logger.info("Code Inside here intNameFromMap is ITN columnTwoValue->"+cellValue);
			return "FIBER";
			
		}
		case "FIBER": {
			return "PERFORM";
		}
		case "PERFORM": {
			return "ENGINEER";
		}

		}
		return "";
	}
	
	
	public void truncateTable(String tableName, ApplicationProperties properties)
	{
		Connection connection = SqlServerConnection.getConnection();
		String query = "truncate table "+properties.getProperty("app.table."+tableName)+" ;";
		
		try {
			PreparedStatement preparedStatement;
			preparedStatement = SqlServerConnection.prepareAStatement(connection, query, null);
			boolean execution = preparedStatement.execute();
			if (execution ) {
				logger.info("Successfully truncated table"+tableName);
			} else {
				failureCount++;
				logger.debug("table not truncated failureCOunt---->>>>" + failureCount);

			}
		} catch (Exception exp) {
			failureCount++;
			logger.debug("failureCOunt inside exception---->>>>" +exp.toString()+" \n failuer count-->"+ failureCount);
		} finally {
			//SqlServerConnection.closeConnection(connection);
		}

	}

}
