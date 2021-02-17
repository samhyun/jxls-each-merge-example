package com.excel;

import java.io.IOException;

public class ExcelGenerateException extends RuntimeException{
    public ExcelGenerateException(String message, Throwable cause) {
        super(message, cause);
    }
}
