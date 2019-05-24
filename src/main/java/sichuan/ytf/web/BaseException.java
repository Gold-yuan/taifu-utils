package sichuan.ytf.web;

/**
 * 基础异常类，自定义异常均继承此
 */
public class BaseException extends Exception {

    private static final long serialVersionUID = 1L;

    /** 错误代码 */
    protected String errorCode;
    /** 错误描述 */
    protected String errMsg;

    public BaseException(String errorCode, String msg) {
        super(msg);
        this.errMsg = msg;
        this.errorCode = errorCode;
    }

    public BaseException(String errorCode, String msg, Throwable throwable) {
        super(msg, throwable);
        this.errMsg = msg;
        this.errorCode = errorCode;
    }

    public String getErrorCode() {
        return errorCode;
    }

    public void setErrorCode(String errorCode) {
        this.errorCode = errorCode;
    }

    public String getErrMsg() {
        return errMsg;
    }

    public void setErrMsg(String errMsg) {
        this.errMsg = errMsg;
    }

    @Override
    public String toString() {
        return super.toString() + " BaseException [errorCode=" + errorCode + ", errMsg=" + errMsg + "]";
    }
}