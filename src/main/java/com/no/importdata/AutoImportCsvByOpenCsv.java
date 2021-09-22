package com.no.importdata;


import com.opencsv.CSVReader;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;


public class AutoImportCsvByOpenCsv {

    Statement statement = null;
    Connection conn = null;
    String tableName = "";
    int lastCellNum = 0;


    public static void main(String[] args) {


        AutoImportCsvByOpenCsv autoImportExcel = new AutoImportCsvByOpenCsv();

        String path = "./csv_test.csv";
        autoImportExcel.read(path);
    }

    public void ready(){
        conn = createConnect();
        try {
            if (!conn.isClosed()) {
                statement = conn.createStatement();
            }
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }

    /**
     * 执行获取文件列和行
     */
    public void read(String path) {

        ready();
        File file = new File(path);

        CSVReader reader = null;
            try {
                reader = new CSVReader(new FileReader(path));
                String[] lineData;
                int i = 0;
                while ((lineData = reader.readNext()) != null) {
                    if (i == 0) { //header
                        String tableSql = getTableSql(lineData, file.getName());
                        if (tableSql == "") {
                            continue;
                        }
                        statement.execute(tableSql); //创建表
                    } else {
                        //导入数据
                        importData(lineData, file.getName());
                        System.out.println(lineData);
                        //System.out.println("Country [id= " + line[0] + ", code= " + line[1] + " , name=" + line[2] + "]");
                    }
                    i++;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }







    }

    public String importData(String[] columns, String filename) {

        String values = "";
        for (int i = 0; i < lastCellNum; i++) {

            values += String.format("'%s',", columns[i].replace("\"",""));
        }
        values = StringUtils.substring(values, 0, -1);

        String sqlInsertRow = "INSERT INTO " + tableName + " VALUES( " + values + " )";
        System.out.println(sqlInsertRow);
        try {
            statement.execute(sqlInsertRow);
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }


        return "";
    }


    public Row getHeadRow(Sheet sheetAt) {
        //获取第一行作为表字段
        int headRowId = 0;
        Row rowHead = sheetAt.getRow(headRowId);
        return rowHead;
    }

    public String getTableName(Sheet sheetAt, String filename) {

        String sheetName = sheetAt.getSheetName();
        filename = filename.substring(0, filename.lastIndexOf(".") - 1);
        return filename + "_" + sheetName;
    }


    public String getTableSql(String[] columns, String filename) {
        lastCellNum = columns.length;
        if (lastCellNum == 0) return "";
        System.out.println(lastCellNum);

        //拼接sql
        tableName = filename.substring(0, filename.lastIndexOf("."));
        if(tableName.length()>=64)   tableName = tableName.substring(0, 63);
        String tableSql = " CREATE TABLE `" + tableName + "` (";
        for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
            tableSql += "`" + columns[cellIndex].replace("\"","") + "` varchar(255) DEFAULT NULL,";
        }
        tableSql = StringUtils.substring(tableSql, 0, -1);
        tableSql += ")  ;";
        System.out.println(tableSql);
        return tableSql;
    }



    /**
     * 获取数据库连接
     */
    public Connection createConnect() {
        String driver = "com.mysql.cj.jdbc.Driver";
        String url = "jdbc:mysql://192.168.27.118:3306/test2?useUnicode=yes&characterEncoding=UTF-8&allowMultiQueries=true&serverTimezone=UTC&useSSL=false";
        String user = "root";
        String password = "123456";
        Connection conn = null;
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, user, password);
            return conn;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return conn;
    }


}