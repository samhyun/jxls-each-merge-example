package com.example;

import com.excel.EachMergeCommand;
import com.excel.ExcelGenerateException;
import lombok.extern.slf4j.Slf4j;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.util.List;

@Slf4j
public class ExcelGenerator {

    private final String templatePath;

    private final Context context;

    public ExcelGenerator(String templatePath) {
        this.templatePath = templatePath;
        this.context = new Context();
//        this.context.putVar(varName, var);
        XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);
    }

    public void addMappingValue(String varName, Object mappingValue) {
        this.context.putVar(varName, mappingValue);
    }

    public void generate(String outputPath) {
        log.info("start employee excel generate");
        try {
            InputStream is = this.getClass().getClassLoader().getResourceAsStream(this.templatePath);
            OutputStream os = new FileOutputStream(outputPath);
            try {
                JxlsHelper.getInstance().processTemplate(is, os, this.context);
            } catch (IOException e) {
                throw new ExcelGenerateException("occurred error in jxls process template", e);
            }
        } catch (FileNotFoundException e) {
            throw new ExcelGenerateException("not found template file", e);
        }
    }
}
