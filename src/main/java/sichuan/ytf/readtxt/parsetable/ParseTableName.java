package sichuan.ytf.readtxt.parsetable;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import sichuan.ytf.db.util.ConnectionPool;
import sichuan.ytf.db.util.DataBaseUtil;

public class ParseTableName {
	public static void main(String[] args) throws Exception {
		ParseTableName p = new ParseTableName();
		Set<String> o = p.getTableName("表.txt");
		p.checkTableNameExist(o);
	}
	
	public Set<String> checkTableNameExist(Set<String> set) {
		Set<String> setnew = new HashSet<String>();
		String sqlbase = "select 1 from table where rownum<1";

		Connection conn = null;
		ResultSet rs = null;
		Statement st = null;
		try {
			conn = ConnectionPool.getConnection();
			st = conn.createStatement();

			for(String tableName : set) {
				String sql = sqlbase.replace("table", tableName);
				try {
					rs = st.executeQuery(sql);
					setnew.add(tableName);
				}catch (Exception e) {
					if(e.getMessage().contains("表或视图不存在")) {
						
					}else {
						e.printStackTrace();
						System.out.println("tableName:"+tableName);
					}
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	
	
		System.out.println(setnew.size() + " " + setnew);
		System.out.println(set.size() + " " + set);
		return setnew;
	}
	
	public Set<String> getTableName(String fileName) throws Exception {
		File file = new File(fileName);
		
		
		
		BufferedReader br = new BufferedReader(new FileReader(file));
		String line = "";
		Set<String> set = new HashSet<String>();
		while((line = br.readLine())!=null) {
			int start = line.indexOf("t_");
			if(start==-1) {
				start = line.indexOf("T_");
			}
			if(start!=-1) {
				String tNameLine = line.substring(start);
				Pattern p = Pattern.compile(".*?[\\s\"\',\\)\\(^=^表^}^{]");
				Matcher m = p.matcher(tNameLine);
				if(m.find()) {
					String l = m.group();
					String tName = l.substring(0, l.length()-1);
					set.add(tName.toUpperCase());
				}
			}
		}
		br.close();
		System.out.println(set+" "+set.size());
		return set;
	}

	public Set<String> getTableNames(String fileName) throws Exception {
		return null;
	}
}
