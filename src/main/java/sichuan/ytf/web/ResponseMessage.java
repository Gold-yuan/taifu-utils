package sichuan.ytf.web;

import java.io.Serializable;

/**
 * 请求响应类
 * 
 * @param <T> T:响应体数据类型
 */
public class ResponseMessage<T> implements Serializable {
    private static final long serialVersionUID = 1L;

    /** 状态码，0：成功，非0为异常代码 */
    private String status;
    /** 异常信息 */
    private String message;
    /** 响应信息 */
    private T data;

    public ResponseMessage() {
    }

    /**
     * 私有构造器
     * 
     * @param status  状态码
     * @param message 状态描述
     * @param data    数据
     */
    private ResponseMessage(String status, String message, T data) {
        this.status = status;
        this.message = message;
        this.data = data;
    }

    /**
     * 构造成功消息
     * 
     * @param data    返回数据
     * @param message 提示信息
     * @param         <T> 返回数据类型
     * @return 消息
     */
    public static <T> ResponseMessage<T> success(String message, T data) {
        return new ResponseMessage<T>("0", message, data);
    }

    /**
     * 构造成功消息
     * 
     * @param data 返回数据
     * @param      <T> 返回数据类型
     * @return 消息
     */
    public static <T> ResponseMessage<T> success(T data) {
        return success("ok", data);
    }

    /**
     * 构造错误消息
     * 
     * @param errorCode    异常状态码
     * @param errorMessage 异常描述信息
     * @param              <T> 返回数据类型
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(String errorCode, String message, T data) {
        return new ResponseMessage<T>(errorCode, message, data);
    }

    /**
     * 构造错误消息
     * 
     * @param errorCode    异常状态码
     * @param errorMessage 异常描述信息
     * @param              <T> 返回数据类型
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(String errorCode, String message) {
        return fail(errorCode, message, null);
    }

    /**
     * 构造错误消息
     * 
     * @param e 自定义异常
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(ExceptionEnum e, T data) {
        return fail(e.getErrorCode(), e.getErrorMsg(), data);
    }

    /**
     * 构造错误消息
     * 
     * @param e 自定义异常
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(ExceptionEnum e) {
        return fail(e, null);
    }

    /**
     * 构造错误消息
     * 
     * @param e 自定义异常
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(BaseException e, T data) {
        return fail(e.getErrorCode(), e.getErrMsg(), data);
    }

    /**
     * 构造错误消息
     * 
     * @param e 自定义异常
     * @return 消息
     */
    public static <T> ResponseMessage<T> fail(BaseException e) {
        return fail(e, null);
    }

    public String getStatus() {
        return status;
    }

    public String getMessage() {
        return message;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    @Override
    public String toString() {
        return "ResponseMessage{" + "status=" + status + ", message='" + message + '\'' + ", data=" + data + '}';
    }
}