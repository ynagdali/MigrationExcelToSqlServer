package dbconnection;

import config.ApplicationProperties;

import java.math.BigDecimal;
import java.sql.*;
import java.util.HashMap;

import org.apache.log4j.Logger;

public class SqlServerConnection {

	private static Connection connection = null;
	private static Logger logger = Logger.getLogger(SqlServerConnection.class);

	static {

		try {
			Class.forName("net.sourceforge.jtds.jdbc.Driver");
			ApplicationProperties properties = new ApplicationProperties();
			String connectionString = properties.getProperty("app.db.connectionString");
			logger.debug("connectionString--->" + connectionString);
			connection = DriverManager.getConnection(connectionString);
		} catch (Exception e) {
			logger.debug("SQL Exception while creating connection -->" + e.toString());
			e.printStackTrace();
			System.exit(0);
		}
	}

	public static Connection getConnection() {
		return connection;
	}

	public static boolean closeConnection(Connection connection) {
		boolean closed = false;
		try {
			connection.close();
			closed = true;
		} catch (SQLException e) {
			logger.debug("SQL Exception while closing connection -->" + e.toString());
		}
		return closed;
	}

	public static PreparedStatement prepareAStatement(Connection connection, String query,
			HashMap<Integer, Object> params) {

		/*
		 * TODO: if the same method executes multiple queries, we don't need to
		 * initialize all this stuff again, but i don't think we need these here since
		 * all of these operations could be done in just one query. some other time.
		 */

		PreparedStatement preparedStatement;

		try {
			preparedStatement = connection.prepareStatement(query, Statement.RETURN_GENERATED_KEYS);

			// insert parameters in the query;
			if (params != null) {
				params.forEach((key, value) -> {
					try {
						if (value instanceof Integer) {
							// TODO: support for more datatype
							preparedStatement.setInt(key, (Integer) value);
						} else if (value instanceof String) {
							preparedStatement.setString(key, (String) value);
						} else if (value instanceof BigDecimal) {
							preparedStatement.setBigDecimal(key, (BigDecimal) value);
						}
					} catch (SQLException e) {
						e.printStackTrace();
					}
				});
			}

			return preparedStatement;
		} catch (SQLException e) {
			e.printStackTrace();
		}

		return null;
	}

}
