package sichuan.ytf.main.util;

import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;

import org.apache.commons.lang3.StringUtils;

/**
 * 实体处理工具类
 */
public class BeanUtils {
	/**
	 * 去掉实体类中参数为{@link String}类型的属性的前后空格
	 *
	 * @param model
	 *            实体bean
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void stripStringProperty(Object model) {
		try {
			Class clazz = model.getClass();

			Method[] methods = clazz.getMethods();// 获取所有方法

			String getterMethodName;// getter方法
			String setterMethodName;// setter方法
			Object propertyValue;// 属性值

			for (Method m : methods) {

				if (m.getName().startsWith("get") && m.getModifiers() == Modifier.PUBLIC) {// 获取public类型的getter方法

					getterMethodName = m.getName();

					propertyValue = m.invoke(model, new Object[] {});// 获取属性值

					if (propertyValue == null) {
						continue;
					}
					Class propertyValueClass = propertyValue.getClass();
					// String 类型
					if (propertyValueClass.getName().equals("java.lang.String")) {

						setterMethodName = "set" + getterMethodName.substring(3);// 拼setter方法
						if (setterMethodName.equals("setKey")) {
							continue;
						}
						Method setterMethod = clazz.getMethod(setterMethodName, new Class[] { String.class });// 反射得到setter方法

						if (isSetterMethodExist(clazz, setterMethodName)) {
							setterMethod.invoke(model, StringUtils.strip((String) propertyValue));
						}
					} else if (propertyValueClass.isArray()) {// 如果属性值的类型为数组
						// String[]数组类型
						Class elementType = propertyValueClass.getComponentType();
						if (elementType.getName().equals("java.lang.String")) {// 数组元素类型为String
							int length = Array.getLength(propertyValue);// 数组长度
							for (int i = 0; i < length; i++) {
								String arrayElementValue = (String) Array.get(propertyValue, i);// 获取数组元素值
								Array.set(propertyValue, i, StringUtils.strip(arrayElementValue));// 去前后空格后，再设回数组
							}
						}
					} else if (propertyValueClass.getName().equals("java.util.ArrayList")) {
						// 是ArrayList
						ArrayList<?> list = (ArrayList<?>) propertyValue;
						if (list.size() == 0) {
							continue;
						}
						Object val = list.get(0);
						if (val instanceof String) {
							// 属性是string
							ArrayList<String> arr = (ArrayList<String>) propertyValue;
							for (int i = 0; i < arr.size(); i++) {
								arr.set(i, arr.get(i).trim());
							}
						} else {
							// 属性是其他
							for (int i = 0; i < list.size(); i++) {
								stripStringProperty(list.get(i));
							}
						}
					} else {
						// 是对象
						stripStringProperty(propertyValue);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 判断CLASS中是否包含public的setter方法，且返回值为string类型
	 *
	 * @param clazz
	 * @param setterName
	 * @return boolean
	 */
	@SuppressWarnings("rawtypes")
	private static boolean isSetterMethodExist(Class clazz, String setterName) {

		Method[] methods = clazz.getMethods();
		for (Method m : methods) {

			if (m.getName().equals(setterName) && m.getModifiers() == Modifier.PUBLIC) {// 匹配setter方法且是public

				Class[] paramTypes = m.getParameterTypes();
				for (Class c : paramTypes) {
					if (c.getName().equals("java.lang.String")) {// setter方法参数类型为String
						return true;
					}
				}
			}
		}

		return false;
	}
}
