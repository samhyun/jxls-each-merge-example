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
public class EmployeeExcelGenerator {

    public EmployeeExcelGenerator() {
        XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);
    }

    public void generate(List<Department> departments, String outputPath) {
        log.info("start employee excel generate");
        try {
            InputStream is = this.getClass().getClassLoader().getResourceAsStream("each-merge-template.xlsx");
            OutputStream os = new FileOutputStream(outputPath);

            Context context = new Context();
            context.putVar("departments", departments);
            try {
                JxlsHelper.getInstance().processTemplate(is, os, context);
            } catch (IOException e) {
                throw new ExcelGenerateException("occurred error in jxls process template", e);
            }
        } catch (FileNotFoundException e) {
            throw new ExcelGenerateException("not found template file", e);
        }
    }
}
