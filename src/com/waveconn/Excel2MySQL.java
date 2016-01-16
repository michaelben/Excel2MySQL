/* Copyright 2015 Michael Zhang

        Licensed under the Apache License, Version 2.0 (the "License");
        you may not use this file except in compliance with the License.
        You may obtain a copy of the License at

        http://www.apache.org/licenses/LICENSE-2.0

        Unless required by applicable law or agreed to in writing, software
        distributed under the License is distributed on an "AS IS" BASIS,
        WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
        See the License for the specific language governing permissions and
        limitations under the License.
*/

package com.waveconn;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.*;

/**
 * Excel2MySQL validator
 *
 * Valid rows and selected columns are imported into DB, invalid rows are saved into error Excel.
 *

1. column is named according excel convention, i.e., A_Z, AA_AZ, BA_BZ, ...

2. for each column, there can be COL_A, COL_A_TYPE, COL_A_LEN
 if the column is missing, do not add it to db;
 if there is no value, do not add this column to db;
 if there is no type, default is string type;
 if there is no length; default is 256.

3. we assume the table already exits in the db; the table schema must comply with those specified in the init file.
  only insert not update.

4. supported Type definition refer to Type ENUM file.

5. if there is error, the error file is created each time with uniqe name based on that specified in the init file.
 *
 * Created by Michael Z. on 2015/6/26.
 *
 */

public class Excel2MySQL {
    public static final String DB_URL = "DB_URL";
    public static final String DB_USER_NAME = "DB_USER_NAME";
    public static final String DB_PASSWORD = "DB_PASSWORD";
    public static final String DB_NAME = "DB_NAME";
    public static final String DB_TABLE = "DB_TABLE";
    public static final String EXCEL_FILE_PATH = "EXCEL_FILE_PATH";
    public static final String EXCEL_ERROR_FILE_PATH = "EXCEL_ERROR_FILE_PATH";
    public static final String IS_READ_FIRST_LINE = "IS_READ_FIRST_LINE";
    public static final String BULK_SIZE = "BULK_SIZE";

    public static final int DB_STRING_LEN_DEFAULT = 256;
    public static final int DB_DATE_LEN = 64;
    public static final int DB_BOOL_LEN = 5;

    Properties properties = new Properties();

    String db_url;
    String db_user_name;
    String db_password;
    String db_name;
    String db_table;
    String excel_file_path = null;
    String excel_error_file_path = null;
    boolean is_read_first_line;
    int bulk_size;

    ArrayList<Mapping> dbMap;

    Workbook workbook = null;
    FormulaEvaluator evaluator = null;
    DataFormatter formatter = null;

    ArrayList<ArrayList<String>> correctRows = new ArrayList();
    ArrayList<ArrayList<String>> errorRows = new ArrayList();

    public static void main(String[] args) {

        if (args.length != 1) {
            System.out.println("Usage: java Excel2MySQL <init_file>");
            System.exit(-1);
        }

        Excel2MySQL app = new Excel2MySQL();

        //read in init file from command line
        app.init(args[0]);

        app.dbImport();
    }

    /**
     * read in init file
     *
     * Column can be omitted and not import into DB;
     * If there is empty value for a column, it is ignored;
     * The order is NOT important;
     * If there is no type, default is string type;
     * If there is no length; default is 256;
     * Supported types are INT, NUM, STR, DATE, BOOL
     * All are case-insensitive and start-with matched;
     * Separator can be either . or _
     */
    void init(String initFile) {
        try {
            properties.load(new FileInputStream(initFile));
        } catch (IOException e) {
            System.err.println("Something wrong with .init file");
            System.exit(-2);
        }

        //Mapping for columns
        Map<String, String> colMap = new HashMap();

        //properties can be in arbitrary order in the property init file
        for (String key : properties.stringPropertyNames()) {
            String value = properties.getProperty(key);
            switch (key) {
                case DB_URL:
                    db_url = value;
                    break;
                case DB_USER_NAME:
                    db_user_name = value;
                    break;
                case DB_PASSWORD:
                    db_password = value;
                    break;
                case DB_NAME:
                    db_name = value;
                    break;
                case DB_TABLE:
                    db_table = value;
                    break;
                case EXCEL_FILE_PATH:
                    excel_file_path = value;
                    break;
                case EXCEL_ERROR_FILE_PATH:
                    excel_error_file_path = value;
                    break;
                case IS_READ_FIRST_LINE:
                    is_read_first_line = Boolean.parseBoolean(value);
                    break;
                case BULK_SIZE:
                    bulk_size = Integer.parseInt(value);
                    break;
                default:
                    String[] tokens = key.split("[_.]");
                    if (tokens.length > 1 && tokens[0].equalsIgnoreCase("COL"))
                        colMap.put(key, value);
                    break;
            }
        }

        //Populates all valid Excel to DB mapping.
        //We only care about and validate on those columns whose header is not empty in the init file
        dbMap = colMap.entrySet().stream()
                .filter(e ->
                                e.getKey().split("[_.]").length == 2 && !e.getValue().trim().isEmpty()
                )
                .map(e -> new Mapping(e.getKey().split("[_.]")[1].trim().toUpperCase(), e.getValue()))
                .collect(ArrayList::new, ArrayList::add, ArrayList::addAll);

        Collections.sort(dbMap);

        //Then populates all types and lengths information into dbMap mapping
        colMap.entrySet().stream()
                .filter(e -> e.getKey().split("[_.]").length == 3)
                .forEach(e -> {
                    String[] tokens = e.getKey().split("[_.]");
                    String excel_sym = tokens[1].trim().toUpperCase();
                    Mapping m = Mapping.getMapping(dbMap, excel_sym);
                    if (m == null) return;

                    String token = tokens[2].trim().toLowerCase();
                    switch (token) {
                        case "type":
                            String typeStr = e.getValue().trim().toUpperCase();
                            if (typeStr.length() >= 3) {
                                typeStr = typeStr.substring(0, 3);
                                m.setType(Type.getType(typeStr));
                            } else {
                                m.setType(Type.STRING); //default to STR if invalid
                            }

                            break;
                        case "len":
                            int len = DB_STRING_LEN_DEFAULT;
                            try {
                                len = Integer.parseInt(e.getValue());
                            } catch (NumberFormatException ex) {
                                //do nothing
                            }

                            m.setLen(len);

                            break;
                    }
                });

        showinfo();
    }

    private void showinfo() {
        System.out.println("DB_URL=" + db_url);
        System.out.println("DB_USER_NAME=" + db_user_name);
        System.out.println("DB_PASSWORD=" + db_password);
        System.out.println("DB_NAME=" + db_name);
        System.out.println("DB_TABLE=" + db_table);
        System.out.println("EXCEL_FILE_PATH=" + excel_file_path);
        System.out.println("EXCEL_ERROR_FILE_PATH=" + excel_error_file_path);
        System.out.println("IS_READ_FIRST_LINE=" + is_read_first_line);
        System.out.println("BULK_SIZE=" + bulk_size);

        System.out.println("Excel   " + "DB   " + "Type   " + "Length");
        for (Mapping m : dbMap) {
            System.out.print("COL_" + m.getExcel_sym() + "   ");
            System.out.print(m.getDb_col() + "   ");
            System.out.print(m.getType() + "   ");
            System.out.print(m.getLen() + "   ");
            System.out.println();
        }

        System.out.println();
    }

    //read and validate Excel, and import into DB
    void dbImport() {
        FileInputStream excel_file = null;
        try {
            excel_file = new FileInputStream(new File(excel_file_path));
        } catch (FileNotFoundException e) {
            System.out.println("File not found: " + excel_file_path);
            System.exit(-3);
        }

        try {
            workbook = WorkbookFactory.create(excel_file);
            evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            formatter = new DataFormatter(true);

            Sheet sheet = null;
            Row row = null;
            int lastRowNum = 0;

            System.out.println("Reading excel file content from " + excel_file_path);

            // Discover how many sheets there are in the workbook....
            int numSheets = workbook.getNumberOfSheets();

            // and then iterate through them.
            for (int i = 0; i < numSheets; i++) {

                // Get a reference to a sheet and check to see if it contains any rows.
                sheet = workbook.getSheetAt(i);
                if (sheet.getPhysicalNumberOfRows() > 0) {

                    // Note down the index number of the bottom-most row and
                    // then iterate through all of the rows on the sheet starting
                    // from the very first row - number 1 - even if it is missing.
                    // Recover a reference to the row and then call another method
                    // which will strip the data from the cells and build lines
                    lastRowNum = sheet.getLastRowNum();

                    int start = 0;
                    if (!is_read_first_line)
                        start = 1;

                    for (int j = start; j <= lastRowNum; j++) {
                        row = sheet.getRow(j);
                        this.rowToData(row);
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("IOException: " + excel_file_path);
            System.exit(-4);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            System.out.println("Invalid Format: " + excel_file_path);
            System.exit(-5);
        } finally {
            if (excel_file != null) {
                try {
                    excel_file.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    System.out.println("IOException: " + excel_file_path);
                    System.exit(-6);
                }
            }
        }

        //put valid rows into DB
        System.out.println("Inserting valid rows into DB table " + db_url + "/" + db_table);
        insertDB();

        System.out.println();

        //save invalid rows if any
        int errs = errorRows.size();
        if (errs > 0) {
            saveError();
        } else {
            System.out.println("There is no invalid row");
        }
    }

    private void insertDB() {
        try (
                Connection con = DriverManager.getConnection(
                        db_url,
                        db_user_name,
                        db_password)) {

            String insertString = Mapping.getInsertString(dbMap, db_table);

            con.setAutoCommit(false);

            try (PreparedStatement insertRows = con.prepareStatement(insertString);) {
                int j = 0;
                while (j < correctRows.size()) {
                    int batch_index = 0;
                    while (batch_index < bulk_size && j < correctRows.size()) {
                        int param_index = 0;
                        for (Mapping m : dbMap) {
                            param_index++;
                            switch (m.getType()) {
                                case INTEGER:
                                    insertRows.setLong(param_index, Long.parseLong(correctRows.get(j).get(m.getExcel_col())));
                                    break;
                                case NUMBER:
                                    insertRows.setDouble(param_index, Double.parseDouble(correctRows.get(j).get(m.getExcel_col())));
                                    break;
                                case STRING:
                                case DATE:
                                case BOOLEAN:
                                    insertRows.setString(param_index, correctRows.get(j).get(m.getExcel_col()));
                                    break;
                            }
                        }

                        insertRows.addBatch();
                        batch_index++;
                        j++;
                    }

                    System.out.println("batch insert.");

                    int[] numUpdates = insertRows.executeBatch();

                    int total = 0;
                    for (int n : numUpdates)
                        if (n > 0) total += n;

                    System.out.println("batch insert " + total + " rows");

                    con.commit();
                }

                System.out.println("total insert " + j + " rows");
            } catch (BatchUpdateException b) {
                System.out.println("BatchUpdateException");
            } catch (SQLException b) {
                System.out.println("SQLException");
            }

            con.setAutoCommit(true);
        } catch (SQLException e) {
            e.printStackTrace();
            System.exit(-10);
        }
    }

    //save error into error file which is unique by TIMESTAMP
    private void saveError() {
        String error_file;
        long now = System.currentTimeMillis();
        if (excel_error_file_path.endsWith(".xlsx"))
            error_file = excel_error_file_path.split("[.]")[0] + "_" + now + ".xlsx";
        else
            error_file = excel_error_file_path + "_" + now + ".xlsx";

        try (Workbook wb = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(error_file)) {

            Sheet sheet = wb.createSheet("Errors");

            for (int i = 0; i < errorRows.size(); i++) {
                Row row = sheet.createRow(i);
                ArrayList<String> rowData = errorRows.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(rowData.get(j));
                }
            }

            // Write the output to a file
            wb.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(errorRows.size() + " invalid rows found. Saved to " + error_file);
    }

    //validate Excel data based on data type Mapping from the init file
    private void rowToData(Row row) {
        Cell cell = null;
        int lastCellNum = 0;
        ArrayList<String> line = new ArrayList();
        ArrayList<Object> correctLine = new ArrayList();

        boolean error = false;

        // Check to ensure that a row was recovered from the sheet as it is
        // possible that one or more rows between other populated rows could be
        // missing - blank. If the row does contain cells then...
        if (row != null) {

            // Get the index for the right most cell on the row and then
            // step along the row from left to right recovering the contents
            // of each cell, converting that into a formatted String and
            // then storing the String into the line ArrayList.
            lastCellNum = row.getLastCellNum();
            for (int i = 0; i < lastCellNum; i++) {
                cell = row.getCell(i);
                if (cell == null) {
                    line.add("");
                } else {
                    if (cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
                        line.add(this.formatter.formatCellValue(cell));
                    } else {
                        line.add(this.formatter.formatCellValue(cell, this.evaluator));
                    }
                }
            }

            //check if there is an error cell in this line and set a flag.
            error = false;
            for (int i = 0; i <= lastCellNum; i++) {
                //ignore the column if it is not in db table
                if (!Mapping.isDb_Col(dbMap, i))
                    continue;

                Mapping m = Mapping.getMapping(dbMap, i);

                switch (m.getType()) {
                    case INTEGER: //INT (int or long)
                        try {
                            int tmp = Integer.parseInt(line.get(i));
                            correctLine.add(tmp);
                        } catch (NumberFormatException e) {
                            try {
                                long tmp = Long.parseLong(line.get(i));
                                correctLine.add(tmp);
                                break;
                            } catch (NumberFormatException e1) {
                                error = true;
                                break;
                            }
                        }
                        break;
                    case NUMBER: //NUM (int or long or float or double)
                        try {
                            int tmp = Integer.parseInt(line.get(i));
                            correctLine.add(tmp);
                        } catch (NumberFormatException e) {
                            try {
                                long tmp = Long.parseLong(line.get(i));
                                correctLine.add(tmp);
                                break;
                            } catch (NumberFormatException e1) {
                                try {
                                    Float tmp = Float.parseFloat(line.get(i));
                                    correctLine.add(tmp);
                                    break;
                                } catch (NumberFormatException e2) {
                                    try {
                                        Double tmp = Double.parseDouble(line.get(i));
                                        correctLine.add(tmp);
                                        break;
                                    } catch (NumberFormatException e3) {
                                        error = true;
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    case STRING: //STR
                        int len = m.getLen();
                        if (len == -1) len = DB_STRING_LEN_DEFAULT;

                        String v = line.get(i);
                        if (v.length() > len)
                            v = v.substring(0, len);
                        correctLine.add(v);
                        break;
                    case DATE: //DATE not validated currently
                        v = line.get(i);
                        if (v.length() > DB_DATE_LEN)
                            v = v.substring(0, DB_DATE_LEN);
                        correctLine.add(v);
                        break;
                    case BOOLEAN: //BOOL
                        v = line.get(i);
                        if (v.length() > DB_BOOL_LEN)
                            v = v.substring(0, DB_BOOL_LEN);
                        if ("true".equalsIgnoreCase(v)
                                || "false".equalsIgnoreCase(v)
                                || "t".equalsIgnoreCase(v)
                                || "f".equalsIgnoreCase(v)
                                || "yes".equalsIgnoreCase(v)
                                || "no".equalsIgnoreCase(v)
                                || "y".equalsIgnoreCase(v)
                                || "n".equalsIgnoreCase(v)
                                ) {
                            correctLine.add(v);
                            break;
                        } else {
                            error = true;
                            break;
                        }
                }
            }
        }

        if (error)
            this.errorRows.add(line);
        else
            this.correctRows.add(line);
    }
}
