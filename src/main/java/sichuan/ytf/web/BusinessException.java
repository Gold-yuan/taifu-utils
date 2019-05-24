package sichuan.ytf.web;

public class BusinessException extends BaseException {
    private static final long serialVersionUID = 1L;

    /** 错误信息 */
    private ExceptionEnum exceptionEnum;

    public BusinessException(ExceptionEnum exceptionEnum) {
        super(exceptionEnum.getErrorCode(), exceptionEnum.getErrorMsg());
        this.exceptionEnum = exceptionEnum;
    }

    public BusinessException(String errorCode, String msg) {
        super(errorCode, msg);
    }

    public BusinessException(String errorCode, String msg, Throwable throwable) {
        super(errorCode, msg, throwable);
    }

    public ExceptionEnum getExceptionEnum() {
        return exceptionEnum;
    }

    public void setExceptionEnum(ExceptionEnum exceptionEnum) {
        this.exceptionEnum = exceptionEnum;
    }

    @Override
    public String toString() {
        return super.toString() + " BusinessException [exceptionEnum=" + exceptionEnum + "]";
    }
}
