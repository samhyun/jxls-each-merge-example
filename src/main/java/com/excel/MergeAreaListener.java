package com.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;

@Slf4j
public class MergeAreaListener implements AreaListener {

    private final CellRef commandCell;

    private final Sheet sheet;

    private CellRef lastRowCellRef;

    public MergeAreaListener(Transformer transformer, CellRef cellRef) {
        this.commandCell = cellRef;
        this.sheet = ((PoiTransformer) transformer).getXSSFWorkbook().getSheet(cellRef.getSheetName());
    }


    @Override
    public void afterApplyAtCell(CellRef cellRef, Context context) {
        // child cell
        if (commandCell.getCol() != cellRef.getCol()) {
            this.setLastRowCellRef(cellRef);
        } else {
            if (existMerged(cellRef)) {
                return;
            }
            merge(cellRef);
        }
    }

    private void merge(CellRef cellRef) {
        int from = cellRef.getRow();

        int lastRow = sheet.getMergedRegions().stream()
                .filter(address -> address.isInRange(this.lastRowCellRef.getRow(), this.lastRowCellRef.getCol()))
                .mapToInt(CellRangeAddressBase::getLastRow).findFirst().orElse(this.lastRowCellRef.getRow());

        log.debug("this :{}, merged start row : {} | end row : {} | col :{} ", this.toString(), from, lastRow, cellRef.getCol());

        CellRangeAddress region = new CellRangeAddress(from, lastRow, cellRef.getCol(), cellRef.getCol());
        sheet.addMergedRegion(region);
        applyStyle(sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol()));
    }

    private void setLastRowCellRef(CellRef cellRef) {
        if (this.lastRowCellRef == null || this.lastRowCellRef.getRow() < cellRef.getRow()) {
            this.lastRowCellRef = cellRef;
        }
    }

    private void applyStyle(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    @Override
    public void beforeApplyAtCell(CellRef cellRef, Context context) {
    }

    @Override
    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    }

    @Override
    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    }

    private boolean existMerged(CellRef cell) {
        return sheet.getMergedRegions().stream()
                .anyMatch(address -> address.isInRange(cell.getRow(), cell.getCol()));
    }

}
