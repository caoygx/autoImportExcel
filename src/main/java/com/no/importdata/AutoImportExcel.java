package com.no.importdata;


import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.sql.*;


public class AutoImportExcel {

    Statement statement = null;
    Connection conn = null;

    /**
     * 获取文件夹路径
     */
    public static void main(String[] args) {

        AutoImportExcel autoImportExcel = new AutoImportExcel();

        String path = "xsl_test.xsl";
        autoImportExcel.readExcel(path);
    }



    /**
     * 执行获取文件列和行
     */
    public void readExcel(String path) {
        File file = new File(path);
        if (!file.exists()) {
            System.out.println("文件不存在!");
        }
        FileInputStream fis = null;
        Workbook workBook = null;
        PreparedStatement pStemt = null;
        try {

            conn = createConnect();
            if (!conn.isClosed()) {
                statement = conn.createStatement();
            }

            workBook = WorkbookFactory.create(new FileInputStream(file));
            int numberOfSheets = workBook.getNumberOfSheets();
            // sheet工作表
            for (int s = 0; s < numberOfSheets; s++) {
                Sheet sheetAt = workBook.getSheetAt(s);
                String tableSql = getTableSql(sheetAt,file.getName());
                if (tableSql == "") {
                    continue;
                }
                //创建表
                try {
                    statement.execute(tableSql);
                    //导入数据
                    importData(sheetAt,file.getName());
                }catch (Exception e){

                }




            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String importData(Sheet sheetAt,String filename) {

        //获取当前Sheet的总行数
        int rowsOfSheet = sheetAt.getPhysicalNumberOfRows();
        System.out.println("当前表格的总行数:" + rowsOfSheet);
        int lastCellNum = getTheNumberOfColumns(sheetAt);
        String tableName = getTableName(sheetAt,filename);
        for (int r = 1; r < rowsOfSheet; r++) {
            Row row = sheetAt.getRow(r);
            if (row == null) {
                continue;
            }

            //INSERT INTO `test2`.`Sheet1` VALUES	( 'a', 'b', NULL );

            String values="";
            for (int i = 0; i < lastCellNum; i++) {
                Cell cell = row.getCell(i);
                values+=String.format("'%s',",cell);
            }
            values = StringUtils.substring(values,0,-1);

            String sqlInsertRow = "INSERT INTO "+tableName+" VALUES( "+values+" )";
            System.out.println(sqlInsertRow);
            try {
                statement.execute(sqlInsertRow);
            } catch (SQLException throwables) {
                throwables.printStackTrace();
            }

        }
        return "";
    }


    public Row getHeadRow(Sheet sheetAt){
        //获取第一行作为表字段
        int headRowId = 0;
        Row rowHead = sheetAt.getRow(headRowId);
        return rowHead;
    }

    public String getTableName(Sheet sheetAt,String filename){

        String sheetName = sheetAt.getSheetName();
        filename = filename.substring(0,filename.lastIndexOf(".")-1);
        return filename +"_"+sheetName;
    }

    public int getTheNumberOfColumns(Sheet sheetAt){
        Row headRow = getHeadRow(sheetAt);
        if (headRow == null) return 0;
        short lastCellNum = headRow.getLastCellNum();
        return lastCellNum;
    }

    public String getTableSql(Sheet sheetAt,String filename) {

        String sheetName = sheetAt.getSheetName();         //获取工作表名称
        int lastCellNum = getTheNumberOfColumns(sheetAt);
        if(lastCellNum==0) return "";
        System.out.println(lastCellNum);

        Row headRow = getHeadRow(sheetAt);
        //拼接sql
        String tableName = getTableName(sheetAt,filename);
        String tableSql = " CREATE TABLE `" + tableName + "` (";
        for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
            Cell cell = headRow.getCell(cellIndex);
            tableSql += "`" + cell.toString() + "` varchar(255) DEFAULT NULL,";
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