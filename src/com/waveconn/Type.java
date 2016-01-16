package com.waveconn;

/**
 * Database TYPE to be supported and validated
 *
 * COL_TYPE : java type : MySQL type

 INT : int or long : BIGINT
 NUM : int or long or float or double : DOUBLE
 STR : string with COL_LEN : VARCHAR(COL_LEN) default 256
 DATE : date :VARCHAR(64)
 BOOL : true/false or t/f or yes/no or y/n (all strings case-insensitive) : VARCHAR(5)

 This means that:
 if the program see an INT it will insert BIGINT in MySQL table;
 if the program see a NUM it will insert DOUBLE in MySQL table;
 if the program see a STR it will insert VARCHAR(COL_LEN) in MySQL table;
 if the program see a DATE it will insert VARCHAR(64) in MySQL table;
 if the program see a BOOL it will insert VARCHAR(5) in MySQL table;

 DATE is currently supported but not invalidated.
 *
 * Created by Michael Z. on 2015/6/26.
 *
 */

public enum Type {
    INTEGER, NUMBER, STRING, DATE, BOOLEAN;

    public static Type getType(String type) {
        type = type.trim().toUpperCase();
        switch (type) {
            case "INT":
                return Type.INTEGER;
            case "NUM":
                return Type.NUMBER;
            case "STR":
                return Type.STRING;
            case "DAT":
                return Type.DATE;
            case "BOO":
                return Type.BOOLEAN;
            default:
                return Type.STRING;
        }
    }
}