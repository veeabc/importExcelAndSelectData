package com.tools.doExcel.Controller;

/**
 * 组合返回结果
 */
public class ResultHolder {

    public ResultHolder() {
        this.success = true;
    }

    public ResultHolder(Object data) {
        this.data = data;
        this.success = true;
    }

    public ResultHolder(boolean success, String msg) {
        this.success = success;
        this.message = msg;
    }

    public ResultHolder(boolean success, String msg, Object data) {
        this.success = success;
        this.message = msg;
        this.data = data;
    }

    // 请求是否成功
    private boolean success = false;
    // 描述信息
    private String message;
    // 返回数据
    private Object data = "";

    public boolean isSuccess() {
        return this.success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public static ResultHolder success(Object obj) {
        return new ResultHolder(obj);
    }

    public static ResultHolder error(Object obj) {
        return new ResultHolder(false, obj.toString(), obj);
    }
}
