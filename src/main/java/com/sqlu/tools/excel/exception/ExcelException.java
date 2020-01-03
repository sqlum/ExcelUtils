package com.sqlu.tools.excel.exception;

/**
 * @author: stonelu
 * @create: 2019-12-21 11:03
 **/
public class ExcelException extends RuntimeException {
    public ExcelException() {
        super();
    }

    public ExcelException(String msg) {
        super(msg);
    }

    public ExcelException(Throwable ex) {
        super(ex);
    }

    public ExcelException(String msg, Throwable ex) {
        super(msg, ex);
    }
}
