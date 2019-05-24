package sichuan.ytf.db.util;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.UUID;
import java.util.regex.Pattern;

public class DataBaseUtil {
	public static void close(Connection conn) {
		if (conn != null) {
			try {
				conn.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public static void close(ResultSet rs) throws SQLException {
		if (rs != null) {
			rs.close();
		}
	}

	public static void close(Statement stmt) throws SQLException {
		if (stmt != null) {
			stmt.close();
		}
	}

	public static void closeQuietly(Connection conn, Statement stmt, ResultSet rs) {
		closeQuietly(rs);
		closeQuietly(stmt);
		close(conn);
	}

	public static void closeQuietly(ResultSet rs) {
		try {
			close(rs);
		} catch (SQLException localSQLException) {
		}
	}

	public static void closeQuietly(Statement stmt) {
		try {
			close(stmt);
		} catch (SQLException localSQLException) {
		}
	}

	public static void commit(Connection conn) throws SQLException {
		if (conn != null) {
			conn.commit();
		}
	}

	public static void rollback(Connection conn) {
		if (conn != null) {
			try {
				conn.rollback();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public static String getUUID(String uuidType) {
		String uuid = UUID.randomUUID().toString().replaceAll("-", "");
		if ("menu".equals(uuidType)) {
			uuid = getNumericUUID(uuid);
		}
		return uuid;
	}

	private static String getNumericUUID(String uuid) {
		boolean isNum = false;
		while (!(isNum)) {
			String firstStr = uuid.substring(0, 1);
			if (isNumeric(firstStr)) {
				isNum = true;
			} else {
				uuid = getUUID(null);
			}
		}

		return uuid;
	}

	private static boolean isNumeric(String str) {
		Pattern pattern = Pattern.compile("[0-9]*");
		return pattern.matcher(str).matches();
	}

	public static void main(String[] args) {
		for (int i = 0; i < 10; ++i) {
			System.out.println(getUUID(null));
		}
	}
}