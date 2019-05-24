package sichuan.ytf.db.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionPool {

	private static String USERNAMR = "ceaadmin";
	private static String PASSWORD = "ceaadmin";
	private static String DRVIER = "oracle.jdbc.OracleDriver";
	private static String URL = "jdbc:oracle:thin:@172.31.67.14:1521:default";

	static {
		try {

			Class.forName(DRVIER);
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		}
	}

	// static {
	// try {
	// Class.forName("com.ibm.db2.jcc.DB2Driver");
	// } catch (ClassNotFoundException e) {
	// e.printStackTrace();
	// }
	// }

	public static Connection getConnection() {
		try {

			return DriverManager.getConnection(URL, USERNAMR, PASSWORD);
			
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}

	// public static Connection getConnection2() {
	// try {
	// return
	// DriverManager.getConnection("jdbc:mysql://9.186.56.147:3306/demo_yili?"
	// + "user=root&password=passw0rd" +
	// "&useUnicode=true&characterEncoding=utf-8"
	// + "&rewriteBatchedStatements=true" + "&useServerPrepStmts=false");
	// } catch (Exception e) {
	// e.printStackTrace();
	// }
	// return null;
	// }

}