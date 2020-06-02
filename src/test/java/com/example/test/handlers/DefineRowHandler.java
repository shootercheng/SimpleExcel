package com.example.test.handlers;

import com.example.RowDataHandler;

import java.util.List;

/**
 * @author James
 */
public class DefineRowHandler implements RowDataHandler {

    private List<List<String>> dataList;

    @Override
    public void handle(int sheetIndex, int curRow, List<String> rowData) {
        dataList.add(rowData);
    }

    public List<List<String>> getDataList() {
        return dataList;
    }
}
