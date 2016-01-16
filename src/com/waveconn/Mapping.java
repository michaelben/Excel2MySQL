package com.waveconn;

import java.util.List;
import java.util.stream.Collectors;

/**
 * DataType Mapping from Excel to DB
 *
 * This should be populated from the init file or from DB meta data
 *
 * Created by Michael Z. on 2015/6/26.
 *
 */
public class Mapping implements Comparable<Mapping> {
    String excel_sym;       //Excel column name
    int excel_col;          //Excel column index
    String db_col;          //DB column name
    Type type;              //column type, default is STRING
    int len;                //column len if type is STRING, default is 256

    Mapping(String excel_sym, String db_col) {
        this.excel_sym = excel_sym;
        this.db_col = db_col;
        toInt();
        type = Type.STRING;
        len = 256;
    }

    Type getType() {
        return this.type;
    }

    void setType(Type type) {
        this.type = type;
    }

    int getLen() {
        return this.len;
    }

    void setLen(int len) {
        this.len = len;
    }

    String getExcel_sym() {
        return this.excel_sym;
    }

    int getExcel_col() {
        return this.excel_col;
    }

    String getDb_col() {
        return this.db_col;
    }

    //does this Excel column has mapping recorded?
    private boolean hasSym(String excel_sym) {
        return this.excel_sym.equalsIgnoreCase(excel_sym);
    }

    //convert Excel column name to excel column index
    private void toInt() {
        for (int i = 0; i < excel_sym.length(); i++)
            excel_col = excel_col * 26 + excel_sym.charAt(i) - 'A';
    }

    public int compareTo(Mapping m) {
        //ascending order
        return this.getExcel_col() - m.getExcel_col();
    }

    //get db mapping for this Excel column name
    static Mapping getMapping(List<Mapping> dbMap, String excel_sym) {
        return dbMap.stream().filter(m -> m.hasSym(excel_sym)).findFirst().orElse(null);
    }

    //get SQL string prepared for SQL insert statement
    static String getInsertString(List<Mapping> dbMap, String db_table) {
        int numCols = dbMap.size();

        String values = new String(new char[numCols]).replace("\0", "?,").substring(0, numCols * 2 - 1);

        StringBuilder columns = new StringBuilder();
        for (Mapping m : dbMap)
            if (columns.length() == 0)
                columns.append(m.getDb_col());
            else
                columns.append("," + m.getDb_col());

        String insertString =
                "INSERT INTO " + db_table + " (" + columns.toString() + ")" +
                        " VALUES (" + values + ")";

        return insertString;
    }

    //get db mapping for this Excel column index
    static Mapping getMapping(List<Mapping> dbMap, int excel_col_index) {
        return dbMap.stream()
                .filter(m -> m.getExcel_col() == excel_col_index)
                .findFirst().orElse(null);
    }

    //does this Excel column index get mapped into DB?
    static boolean isDb_Col(List<Mapping> dbMap, int excel_col_index) {
        if (dbMap == null)
            return false;

        if (import_excel_cols == null)
            import_excel_cols = dbMap.stream().map(m -> m.getExcel_col()).collect(Collectors.toList());

        return import_excel_cols.contains(excel_col_index);
    }

    //performance hack
    static private List<Integer> import_excel_cols;
}
