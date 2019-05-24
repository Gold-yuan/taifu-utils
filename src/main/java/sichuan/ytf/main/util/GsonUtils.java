package sichuan.ytf.main.util;

import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

/**
 * Gson工具类
 */
public class GsonUtils {

	private static final Gson gson;

	static {
		gson = new GsonBuilder().enableComplexMapKeySerialization().create();
	}

	/**
	 * Json转Bean
	 * 
	 * @param jsonStr
	 *            字符串格式的Json数据
	 * @param clazz
	 *            数据类型
	 * @param <T>
	 *            泛型
	 * @return Bean
	 */
	public static <T> T fromJson(String jsonStr, Class<T> clazz) {
		if (StringUtils.isBlank(jsonStr)) {
			return null;
		}
		return gson.fromJson(jsonStr, clazz);
	}

	/**
	 * Bean转Json
	 * 
	 * @param object
	 *            Bean
	 * @return
	 */
	public static String toJson(Object object) {
		return gson.toJson(object);
	}

	/**
	 * Json转Bean
	 * 
	 * @param jsonStr
	 *            字符串格式的Json数据
	 * @param clazz
	 *            数据类型
	 * @param <T>
	 *            泛型
	 * @return Bean集合
	 */
	public static <T> List<T> fromJsonDataToList(String jsonStr, Class<T> clazz) {
		if (StringUtils.isBlank(jsonStr)) {
			return new ArrayList<T>();
		}
		JsonArray array = new JsonParser().parse(jsonStr).getAsJsonObject().getAsJsonArray("data");
		if (array.size() == 0) {
			return new ArrayList<T>();
		}
		List<T> list = new ArrayList<>();
		for (final JsonElement elem : array) {
			list.add(gson.fromJson(((JsonObject) elem), clazz));
		}
		return list;
	}

	public static <T> T fromJson(String jsonStr, Type type) {
		if (StringUtils.isBlank(jsonStr)) {
			return null;
		}
		return gson.fromJson(jsonStr, type);
	}

	public static <T> List<T> fromJsonToList(String jsonStr, Class<T> clazz) {
		if (StringUtils.isBlank(jsonStr)) {
			return new ArrayList<T>();
		}
		JsonArray array = new JsonParser().parse(jsonStr).getAsJsonArray();
		if (array.size() == 0) {
			return new ArrayList<T>();
		}
		List<T> list = new ArrayList<>();
		for (final JsonElement elem : array) {
			list.add(gson.fromJson(((JsonObject) elem), clazz));
		}
		return list;
	}
}
