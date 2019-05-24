package sichuan.ytf.excel.util;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import sichuan.ytf.db.util.ConnectionPool;
import sichuan.ytf.db.util.DataBaseUtil;

public class ParseTMMapZrdwSairport {
	
	public static void main(String[] args) throws Exception {
		ParseTMMapZrdwSairport p = new ParseTMMapZrdwSairport();
		p.add1();
	}
	/**
	 * 第二次导入的数据
	 * @throws Exception
	 */
	public void add1() throws Exception{
		// 重复的不导入
		String filePath = "F:/project/东航/文档/销售责任单位对应表.xlsx";
		ParseExcelUtil37 pe = new ParseExcelUtil37();
		ParseTMMapZrdwSairport ps = new ParseTMMapZrdwSairport();
		
		pe.setStartCell(0);
		pe.setStartRow(1);
		String[] names = {"xsdw", "isJs", "sairport", "xsqy", "gnzb"};
		List<Map<String, String>> listOld = pe.parseExcel(filePath, names);
		
		System.out.println(listOld.size()+"  "+listOld);
		ps.preparData(listOld);
		System.out.println(listOld.size()+"  "+listOld);
		
		

		filePath = "国际.xlsx";
		pe.setStartCell(0);
		pe.setStartRow(1);
		List<Map<String, String>> listNew = pe.parseExcel(filePath, names);
		
		System.out.println(listNew.size()+"  "+listNew);
		ps.preparData(listNew);
		System.out.println(listNew.size()+"  "+listNew);
		

		Iterator<Map<String, String>> iter = listNew.iterator();
		while(iter.hasNext()) {
			Map<String, String> map = iter.next();
			String xsdw = map.get("xsdw").trim();
			String sairport = map.get("sairport").trim();
			for(Map<String, String> mapOld : listOld) {
				String xsdwOld = mapOld.get("xsdw").trim();
				String sairportOld = mapOld.get("sairport").trim();
				if(xsdw.equals(xsdwOld) && sairport.equals(sairportOld)) {
					iter.remove();
					break;
				}
			}
		}
		System.out.println(listNew.size()+"  "+listNew);
//		ps.insertData(listNew);
	}
	
	/**
	 * 第一次导入的数据
	 * @throws Exception
	 */
	public void initOne() throws Exception{
		String filePath = "F:/project/东航/文档/销售责任单位对应表.xlsx";
		ParseExcelUtil37 pe = new ParseExcelUtil37();
		pe.setStartCell(0);
		pe.setStartRow(1);
		String[] names = {"xsdw", "isJs", "sairport", "xsqy", "gnzb"};
		List<Map<String, String>> list = pe.parseExcel(filePath, names);
		
		System.out.println(list.size()+"  "+list);
		ParseTMMapZrdwSairport ps = new ParseTMMapZrdwSairport();
		ps.preparData(list);
		System.out.println(list.size()+"  "+list);
		ps.insertData(list);
	
	}
	public void preparData(List<Map<String, String>> list) {
		List<Map<String, String>> listNew = new ArrayList<Map<String, String>>();
		Iterator<Map<String, String>> iter = list.iterator();
		while(iter.hasNext()) {
			Map<String, String> map = iter.next();
			String sairport = map.get("sairport");
			if(sairport.indexOf(",") !=-1) {
				String[] sairports = sairport.split(",");
				for(String sp:sairports) {
					Map<String, String> map2 = new HashMap<String, String>();
					map2.put("xsdw", map.get("xsdw"));
					map2.put("isJs", map.get("isJs"));
					map2.put("sairport", sp);
					map2.put("xsqy", map.get("xsqy"));
					map2.put("gnzb", map.get("gnzb"));
					listNew.add(map2);
				}
				iter.remove();
			}
		}
		list.addAll(listNew);
	}
	public void insertData(List<Map<String, String>> list) {
		Connection conn = null;
		PreparedStatement st = null;
		ResultSet rs = null;

		try {
			conn = ConnectionPool.getConnection();
			conn.setAutoCommit(false);

			String sql = "insert into T_M_MAP_ZRDW_SAIRPORT (XSDW,IS_JS,SAIRPORT,XSQY,GNZB) values (?,?,?,?,?)";
			st = conn.prepareStatement(sql);

			for (Map<String, String> map : list) {
				int i = 1;
				st.setString(i++, map.get("xsdw"));
				st.setString(i++, map.get("isJs"));
				st.setString(i++, map.get("sairport"));
				st.setString(i++, map.get("xsqy"));
				st.setString(i++, map.get("gnzb"));
				st.addBatch();
			}
			st.executeBatch();
			conn.commit();
			// st.close();
		} catch (SQLException e) {

			e.printStackTrace();

		} finally {
			DataBaseUtil.closeQuietly(conn, st, rs);
		}
	
	}
}
