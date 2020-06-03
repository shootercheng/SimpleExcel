package com.example.test.handlers;

import com.example.RowDataHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * @author James
 */
public class DefineRowHandler implements RowDataHandler {

    private List<List<String>> dataList = new ArrayList<>();

    @Override
    public void handle(int sheetIndex, int curRow, List<String> rowData) {
        List<String> newData = new ArrayList<>(rowData);
        dataList.add(newData);
    }

    public List<List<String>> getDataList() {
        return dataList;
    }
}
