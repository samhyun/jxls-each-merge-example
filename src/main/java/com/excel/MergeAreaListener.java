package com.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;

import java.lang.reflect.Field;
import java.util.Collection;

@Slf4j
public class MergeAreaListener implements AreaListener {

    private final PoiTransformer transformer;

    private final String varName;

    private final String childItemsName;

    public MergeAreaListener(Transformer transformer, String var, String childItems) {
        this.transformer = (PoiTransformer) transformer;
        this.varName = var;
        this.childItemsName = childItems;
    }

    @Override
    public void beforeApplyAtCell(CellRef cellRef, Context context) {

    }

    @Override
    public void afterApplyAtCell(CellRef cellRef, Context context) {
        Sheet sheet = this.transformer.getXSSFWorkbook().getSheet(cellRef.getSheetName());
        int from = cellRef.getRow();
        Object var = context.getVar(this.varName);
        try {
            Field childField = var.getClass().getDeclaredField(childItemsName);
            childField.setAccessible(true);
            Collection<?> children = (Collection<?>) childField.get(var);
            CellRangeAddress region = new CellRangeAddress(from, from + children.size() - 1, cellRef.getCol(), cellRef.getCol());
            sheet.addMergedRegion(region);

        } catch (IllegalAccessException e) {
            log.error("occurred error access child items : {} ", this.childItemsName);
        } catch (NoSuchFieldException e) {
            log.error("not found child list : {} ", this.childItemsName);
        }
    }

    @Override
    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {

    }

    @Override
    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {

    }


}
