package sichuan.ytf.web.test;

import static java.nio.charset.StandardCharsets.UTF_8;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.http.client.methods.HttpDelete;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.methods.HttpPut;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ObjectUtils;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

import sichuan.ytf.main.util.HttpClientUtils;
import sichuan.ytf.web.test.TestRequest.Paramter;
import sichuan.ytf.web.test.TestRequest.Record;

/**
 * 测试所需工具类
 */
public class HttpTestUtils {
	private static final Logger log = LoggerFactory.getLogger(HttpTestUtils.class);
	private static HttpClientUtils client = HttpClientUtils.getInstance();;

	/**
	 * 发送测试请求，返回需要保存的测试结果数据
	 * 
	 * @param testRequests
	 *            测试请求
	 * @return 测试结果数据map
	 * @throws IOException
	 * @throws InterruptedException
	 */
	public static Map<String, String> sendTestRequests(List<TestRequest> testRequests)
			throws IOException, InterruptedException {
		Map<String, String> respDataMap = new HashMap<>();

		for (TestRequest req : testRequests) {
			log.info("测试内容： {}", req.getName());

			boolean isJson = false;
			String paramJson = "";
			String getMethodParam = "";
			Map<String, String> paramMap = new HashMap<String, String>(10);

			for (Paramter reqParameter : req.getParameters()) {
				String paramName = reqParameter.getName();
				String data = replaceParameters(reqParameter.getData().toString(), respDataMap);
				if ("json".equals(reqParameter.getParamType())) {
					isJson = true;
					paramJson = data;
				} else {
					paramMap.put(paramName, data);
					getMethodParam += "&" + paramName + "=" + data;
				}
			}

			String reqPath = replaceParameters(req.getPath(), respDataMap);
			String resp = "";
			String url = "";
			switch (req.getType().toUpperCase()) {
			case HttpPost.METHOD_NAME:
				url = reqPath;
				log.info("地址：{}", url);
				if (isJson) {
					log.info("参数: {}", paramJson);
					resp = client.sendHttpJsonPost(url, paramJson);
				} else {
					log.info("参数: {}", paramMap);
					resp = client.sendHttpPost(url, paramMap);
				}
				break;
			case HttpGet.METHOD_NAME:
				getMethodParam = getMethodParam.replaceFirst("&", "?");
				getMethodParam = URLEncoder.encode(getMethodParam, UTF_8.toString());
				url = reqPath + getMethodParam;
				log.info("地址：{}", url);
				resp = client.sendHttpGet(url);
				break;
			case HttpPut.METHOD_NAME:
				url = reqPath;
				log.info("地址：{}", url);
				if (isJson) {
					log.info("参数: {}", paramJson);
					resp = client.sendHttpJsonPut(url, paramJson);
				} else {
					log.info("参数: {}", paramMap);
					resp = client.sendHttpPut(url, paramMap);
				}
				break;
			case HttpDelete.METHOD_NAME:
				url = reqPath;
				log.info("地址：{}", url);
				resp = client.sendHttpDelete(url);
				break;
			default:
				log.error("无此类型 {}", req.getType());
				return null;
			}

			log.info("返回结果：\n{}", resp);

			// 结果比对
			JsonObject jsonResp = new JsonParser().parse(resp).getAsJsonObject();
			int respCode = jsonResp.get("code").getAsInt();
			if (req.getExpCode() == respCode) {
				log.info("\n与期望结果一致");
			} else {
				log.warn("\n!!!与期望结果不一致!!!");
				continue;
			}

			// 保存数据到Map
			if (!ObjectUtils.isEmpty(req.getRecords())) {
				saveDataToMap(req.getRecords(), resp, respDataMap);
			}

			if (reqPath.contains("/channels/") && reqPath.contains("/updates/")) {
				Thread.sleep(3000);
			}
		}

		return respDataMap;
	}

	private static String replaceParameters(String data, Map<String, String> respDataMap) {
		if (StringUtils.isBlank(data)) {
			return "";
		}
		Pattern pattern = Pattern.compile("(?<=(\\$\\{))[^/]+?(?=\\})");
		Matcher dataMatcher = pattern.matcher(data);
		while (dataMatcher.find()) {
			String name = dataMatcher.group();
			String val = respDataMap.get(name);
			if (val.startsWith("[") || val.startsWith("{")) {
				data = data.replace("\"${" + name + "}\"", val);
			} else {
				data = data.replace("${" + name + "}", val);
			}
		}
		return data;
	}

	private static void saveDataToMap(List<Record> dataToSave, String resp, Map<String, String> respDataMap) {
		for (Record record : dataToSave) {
			String key = record.getName();
			String prop = record.getProp();
			String[] ps = prop.split("\\.");
			JsonObject temp = new JsonParser().parse(resp).getAsJsonObject();
			for (int x = 0; x < ps.length; x++) {
				JsonElement subElement;
				if (ps[x].contains("[")) {
					String subkey = ps[x].substring(0, ps[x].indexOf("["));
					int index = Integer.parseInt(ps[x].substring(ps[x].indexOf("[") + 1, ps[x].indexOf("]")));
					subElement = temp.get(subkey).getAsJsonArray().get(index);
				} else {
					subElement = temp.get(ps[x]);
				}

				if ((x + 1) == ps.length) {
					String val = "";
					if (subElement.isJsonObject() || subElement.isJsonArray()) {
						val = subElement.toString();
					} else {
						val = subElement.getAsString();
					}
					val = val.replace("\n", "\\n");
					val = val.replace("\r", "\\r");
					respDataMap.put(key, val);
				} else {
					temp = subElement.getAsJsonObject();
				}
			}
		}
	}
}
