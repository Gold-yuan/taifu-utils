package sichuan.ytf.db.util;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.util.Date;

public class TestDbUtil {
	public static void main(String[] args) {
		TestDbUtil tdu = new TestDbUtil();
//		tdu.delete();
//		tdu.insert();
		tdu.select();
	}
	
	public void delete() {
		Connection conn = null;
		PreparedStatement st = null;
		ResultSet rs = null;
		try {
			conn = ConnectionPool.getConnection();
			conn.setAutoCommit(false);
			String sql = "delete from T_M_IMPACT_FACTOR_temp where Fid=0";
			st = conn.prepareStatement(sql);
			st.executeUpdate();
			conn.commit();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	}
	
	public void insert() {
		Connection conn = null;
		PreparedStatement st = null;
		ResultSet rs = null;
		try {
			conn = ConnectionPool.getConnection();
			conn.setAutoCommit(false);
			String sql = "insert into T_M_IMPACT_FACTOR_temp (FID, FTYPE, FNAME, CREATER,CREATETIME,UPDATETIME,ISDELETE,IFUSE) values (?,?,?,?,?,?,?,?)";
			st = conn.prepareStatement(sql);
			int number = 1;
			
			for (int j = 0; j < 2; j++) {
				int i = 1;
				st.setInt(i++, j);
				st.setString(i++, "value");
				st.setString(i++, "value");
				st.setString(i++, "value");
				st.setString(i++, "value");
				Timestamp t = new Timestamp(new Date().getTime());
				st.setTimestamp(i++, t);
				st.setInt(i++, 1);
				st.setInt(i++, 1);
				st.addBatch();
				number++;
				if (number % 1000 == 0) {
					st.executeBatch();
				}
			}
			st.executeBatch();
			conn.commit();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	}
	public void select() {
		Connection conn = null;
		ResultSet rs = null;
		Statement st = null;
		try {
			conn = ConnectionPool.getConnection();
			conn.setAutoCommit(false);
			String sql = "select sysdate from dual";
			st = conn.createStatement();
			rs = st.executeQuery(sql);
			while (rs.next()) {
				System.out.println("query date:" + rs.getDate(1));
			}
			
			conn.commit();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	}
	public void test() {
		Connection conn = null;
		PreparedStatement st = null;
		ResultSet rs = null;
		Statement s = null;
		try {
			conn = ConnectionPool.getConnection();
			conn.setAutoCommit(false);
			String sql = "delete from T_M_IMPACT_FACTOR_temp where Fid=0";
			st = conn.prepareStatement(sql);
			st.executeUpdate();
			st.close();

			sql = "insert into T_M_IMPACT_FACTOR_temp (FID, FTYPE, FNAME, CREATER,CREATETIME,UPDATETIME,ISDELETE,IFUSE) values (?,?,?,?,?,?,?,?)";
			st = conn.prepareStatement(sql);
			int number = 1;

			for (int j = 0; j < 2; j++) {
				int i = 1;
				st.setInt(i++, j);
				st.setString(i++, "value");
				st.setString(i++, "value");
				st.setString(i++, "value");
				st.setString(i++, "value");
				Timestamp t = new Timestamp(new Date().getTime());
				st.setTimestamp(i++, t);
				st.setInt(i++, 1);
				st.setInt(i++, 1);
				st.addBatch();
				number++;
				if (number % 1000 == 0) {
					st.executeBatch();
				}
			}
			st.executeBatch();
			st.close();

			
			sql = "select sysdate from dual";
			s = conn.createStatement();
			rs = s.executeQuery(sql);
			while (rs.next()) {
				System.out.println("query date:" + rs.getDate(1));
			}

			conn.commit();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			DataBaseUtil.closeQuietly(s);
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	}
}
