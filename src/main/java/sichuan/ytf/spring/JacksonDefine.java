package sichuan.ytf.spring;

import java.math.BigInteger;
import java.util.List;

import org.springframework.http.converter.HttpMessageConverter;
import org.springframework.http.converter.json.MappingJackson2HttpMessageConverter;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurerAdapter;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.ToStringSerializer;

/**
 * SpringBoot,返回JSON,Long前端精准度丢失,Jackson自定义序列化规则
 * 
 * @author tfyuan
 *
 */
@SuppressWarnings("deprecation")
@Component
public class JacksonDefine extends WebMvcConfigurerAdapter {
	@Override
	public void configureMessageConverters(List<HttpMessageConverter<?>> converters) {
		MappingJackson2HttpMessageConverter messageConverter = new MappingJackson2HttpMessageConverter();
		ObjectMapper objectMapper = new ObjectMapper();
		SimpleModule simpleModule = new SimpleModule();

		/**
		 * 将Long,BigInteger序列化的时候,转化为String
		 */
		simpleModule.addSerializer(Long.class, ToStringSerializer.instance);
		simpleModule.addSerializer(Long.TYPE, ToStringSerializer.instance);
		simpleModule.addSerializer(BigInteger.class, ToStringSerializer.instance);

		objectMapper.registerModule(simpleModule);

		messageConverter.setObjectMapper(objectMapper);

		converters.add(messageConverter);
	}
}
