package com.excel;

import lombok.extern.slf4j.Slf4j;
import org.jxls.area.Area;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

import java.util.List;
import java.util.stream.Collectors;

@Slf4j
public class EachMergeCommand extends EachCommand {

    public static final String COMMAND_NAME = "each-merge";

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        List<Area> childAreas = this.getAreaList().stream()
                .flatMap(area1 -> area1.getCommandDataList().stream())
                .flatMap(commandData -> commandData.getCommand().getAreaList().stream())
                .collect(Collectors.toList());

        MergeAreaListener listener = new MergeAreaListener(this.getTransformer(), cellRef);

        this.getAreaList().get(0).addAreaListener(listener);

        childAreas.forEach(childArea -> {
            childArea.addAreaListener(listener);
        });
        return super.applyAt(cellRef, context);
    }
}
