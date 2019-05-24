package sichuan.ytf.web.test;

import java.util.List;

import com.google.gson.JsonElement;

/**
 * 测试数据封装类
 */
public class TestRequest {
	/** 名称 */
	private String name;
	/** 请求地址 */
	private String path;
	/** 请求类型 get/post/put/delete */
	private String type;
	/** 请求参数 */
	private List<Paramter> parameters;
	/** 结果记录 */
	private List<Record> records;
	/** 期望返回代码 */
	private int expCode;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public List<Paramter> getParameters() {
		return parameters;
	}

	public void setParameters(List<Paramter> parameters) {
		this.parameters = parameters;
	}

	public List<Record> getRecords() {
		return records;
	}

	public void setRecord(List<Record> records) {
		this.records = records;
	}

	public int getExpCode() {
		return expCode;
	}

	public void setExpCode(int expCode) {
		this.expCode = expCode;
	}

	@Override
	public String toString() {
		return "RaRequestTestVo [name=" + name + ", path=" + path + ", type=" + type + ", parameters=" + parameters
				+ ", records=" + records + ", expCode=" + expCode + "]";
	}

	/**
	 * 参数
	 */
	public static class Paramter {
		private String paramType;
		private String name;
		private JsonElement data;

		public String getParamType() {
			return paramType;
		}

		public void setParamType(String paramType) {
			this.paramType = paramType;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public JsonElement getData() {
			return data;
		}

		public void setData(JsonElement data) {
			this.data = data;
		}

		@Override
		public String toString() {
			return "Paramter [paramType=" + paramType + ", name=" + name + ", data=" + data + "]";
		}

	}

	/**
	 * 返回数据记录
	 *
	 */
	public static class Record {
		private String name;
		private String prop;

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getProp() {
			return prop;
		}

		public void setProp(String prop) {
			this.prop = prop;
		}

		@Override
		public String toString() {
			return "Record [name=" + name + ", prop=" + prop + "]";
		}
	}
}
