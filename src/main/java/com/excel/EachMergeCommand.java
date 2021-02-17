package com.excel;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.jxls.area.Area;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.util.UtilWrapper;

@Slf4j
public class EachMergeCommand extends EachCommand {

    public static final String COMMAND_NAME = "each-merge";

    private UtilWrapper util = new UtilWrapper();

    @Getter
    @Setter
    private String childItems;

    public EachMergeCommand() {
    }

    public EachMergeCommand(String var, String items, String childItems, Area area) {
        super(var, items, area);
        this.childItems = childItems;
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        MergeAreaListener listener = new MergeAreaListener(this.getTransformer(), this.getVar(), this.getChildItems());
//        Iterable<Object> objects = util.transformToIterableObject(getTransformationConfig().getExpressionEvaluator(), this.getItems(), context);
        this.getAreaList().get(0).addAreaListener(listener);
        return super.applyAt(cellRef, context);
    }
}
