package utils.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
* @author: lijinlong
* @Date: 2019/9/5 16:38
* @Description 从数据库空将指定表及表结构和数据读出来写入到以表名为名的Excel中
* @version 1.0
*/
@Component
public final class PIOOutputExcel {

    private static final String SELECT = "select";
    private static final String FROM = "FROM";
    private static String sql = "";
    static Logger logger = LoggerFactory.getLogger(PIOOutputExcel.class);
    static String table = "";
    //把值写死或者另加一个方法可以配合周期调度实现每天定时将指定表的数据备份到Excel中
    @Scheduled(cron = "0 * * * * ? ")
    public void runTest(){
        PIOOutputExcel.OutputExcel("userDemo","xls","D:");
    }

    protected static void OutputExcel(String tablename,String excelVersion,String filePath){
        //判断输入表名是否为空
        if("".equals(tablename) || null == tablename){
            if(tablename.length() <= 0){
                logger.info("导出表名为空,导出失败");
                return;
            }
            logger.info("导出表名为空,导出失败");
            return;
        }
        //获取数据库中的表名
        if(!getDBTableName(tablename)){
            logger.info("表不存在");
            return;
        }
        //取到了字段可以写入Excel的第一行
        List<String> dbTableColumn = getDBTableColumn(table);
        if(dbTableColumn.isEmpty()){
            logger.info("表中字段为空");
            return;
        }
        boolean xlsx = outPutTitle(dbTableColumn, excelVersion, filePath);
        //标题未写入成功
        if(!xlsx){
            logger.info("标题写入失败，请确认后续数据是否成功写入，若数据未正确写入请确认是否参数正确，且排查错误");
            return;
        }
        //将表中数据放入map中，key未数据，value为字段名
        Map<List<String>, Integer> stringStringMap = OutputData(dbTableColumn);
        //将数据追加写入到Excel中
        if(!OutputDBData(stringStringMap,excelVersion,filePath)){
            logger.info("追加写入数据时出现异常");
            return;
        }
        logger.info("success");
    }

    /**
     * @Description 根据输入表名来去数据库中找到对应表名
     * @Author lijinlong
     * @Date   2019/9/4 9:46
     * @Param  [tablename]
     * @Return boolean  如果表名存在将它赋值给table并返回true，如果表不存在就返回false
     * @Exception
     */
    protected static boolean getDBTableName(String tablename){
        Connection connection = JdbcConnectionUtil.getConnection();
        ResultSet resultSet = null;
        try {
            DatabaseMetaData metaData = connection.getMetaData();
            resultSet = metaData.getTables("","","",new String[]{"TABLE"});
            while(resultSet.next()){
                //获取表名
                String tname = resultSet.getString(3);
                //如果输入表名和数据库表名相同，那就输要找的表
                if(tname.toLowerCase().equals(tablename.toLowerCase())){
                    //用变量保存表名，然后下一步就是根据表明查出表结构作为excel的标题行，然后再根据表结构查数据放入指定列中
                    table = tname;
                    return true;
                }
            }
//            sql = SELECT + "" + FROM + tablename;
        } catch (SQLException e) {
            logger.info("查询数据库所有表名时出现异常");
            return false;
        }finally {
            JdbcConnectionUtil.close(connection,null,resultSet);
        }
        return false;
    }

    /**
     * @Description 查询指定表中的字段名
     * @Author lijinlong
     * @Date   2019/9/4 13:51
     * @Param  [tablename]
     * @Return java.util.List<java.lang.String>
     * @Exception
     */
    protected static List<String> getDBTableColumn(String tablename){
        Connection connection = JdbcConnectionUtil.getConnection();
        ResultSet resultSet = null;
        PreparedStatement preparedStatement = null;
        List<String> tablecolumn = new ArrayList<>();
        String sql = SELECT +" "+"*"+" " + FROM+ " " + tablename;
        try {
            preparedStatement = connection.prepareStatement(sql);
            ResultSetMetaData metaData = preparedStatement.getMetaData();
            int columnCount = metaData.getColumnCount();
            //参数为1就从第一列取
            for(int i = 1; i <= columnCount; i++){
                tablecolumn.add(metaData.getColumnName(i));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }finally {
            JdbcConnectionUtil.close(connection,preparedStatement,resultSet);
        }
        return tablecolumn;
    }

    /**
     * @Description 根据输入的字段名、Excel的版本、文件路径来将字段类型给写入指定Excel版本的标题行，然后将以表名为文件名输出到指定路径下
     * @Author lijinlong
     * @Date   2019/9/4 15:19
     * @Param  [columnName字段名, excelVersion版本, filePath路径]
     * @Return boolean 为true写出成功，false写出失败
     * @Exception
     */
    protected static boolean outPutTitle(List<String> columnName,String excelVersion,String filePath){
        Workbook workbook = null;
        OutputStream os = null;
        //如果没有指定版本默认为2007版本，或者指定为2007版本
        if("".equals(excelVersion) || null == excelVersion || "xlsx".equals(excelVersion)){
            workbook = new XSSFWorkbook();
        //指定版本为2003
        }else if("xls".equals(excelVersion)){
            workbook = new HSSFWorkbook();
        }else{
            logger.info("所输入Excel版本不存在");
            return false;
        }
        //将标题写入到Excel
        if(!tTAS(columnName,workbook)){
            logger.info("字段名为空");
            return false;
        }
        try {
            //创建输出流将指定Excel输出到指定目录
            os = new FileOutputStream(filePath.endsWith("/")?filePath+table+"."+excelVersion:filePath+"/"+table+"."+excelVersion);
            workbook.write(os);
            return true;
        } catch (Exception e) {
            logger.info("写出Excel时出现异常");
            return false;
        }finally {
            try {
                workbook.close();
                os.close();
            } catch (IOException e) {
                logger.info("关闭输出流时出现异常");
                return false;
            }
        }
    }

    /**
     * @Description 处理Excel
     * @Author lijinlong
     * @Date   2019/9/4 16:07
     * @Param  [columnName, workbook]
     * @Return org.apache.poi.ss.usermodel.Workbook
     * @Exception
     */
    protected static boolean tTAS(List<String> columnName,Workbook workbook){
        if(columnName.isEmpty()){
            return false;
        }
        //创建一个单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        //设置单元格居中
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        //创建一个字体样式
        Font font = workbook.createFont();
        //设置字体为加粗
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        //将样式附着在单元格样式上
        cellStyle.setFont(font);
        //创建一个Sheet页
        Sheet sheet = workbook.createSheet();
        //设置列宽为30
        sheet.setDefaultColumnWidth(30);
        //创建第一行,标题行
        Row row = sheet.createRow(0);
        //循环将字段名和样式放入第一行的所用单元格中
        for(int i = 0; i<columnName.size(); i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(columnName.get(i));
            cell.setCellStyle(cellStyle);
        }
        return true;
    }

    /**
     * @Description 将数据从表中读到Map集合中 key为数据，value为第几行
     * @Author lijinlong
     * @Date   2019/9/4 17:26
     * @Param  [columnName]
     * @Return java.util.Map<java.lang.String,java.lang.String>
     * @Exception
     */
    protected static Map<List<String>,Integer> OutputData(List<String> columnName){

        if(columnName.isEmpty()){
            return null;
        }
        Connection connection = JdbcConnectionUtil.getConnection();
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;
        Map<List<String>,Integer> data = new HashMap<>();
        sql = SELECT + " ";
        //拼sql
        for(int i=0; i<columnName.size(); i++){
            //最后的字段时空一格不加逗号
            if(i == columnName.size()-1){
                sql = sql + columnName.get(i) + " ";
            //不为最后一行就用都行将字段隔开
            }else{
                sql = sql + columnName.get(i) + ",";
            }
        }
        sql = sql + FROM + " " + table;
        try {
            preparedStatement = connection.prepareStatement(sql);
            resultSet = preparedStatement.executeQuery();
            //用来初始第几行
            int columnNum = 0;
            while(resultSet.next()){
                //用来接收数据
                List<String> hdata = new ArrayList<>();
                for(int i=0; i<columnName.size(); i++){
                    //获取每列参数对应的值，用的ArrayList是有序的
                    String value = resultSet.getString(columnName.get(i));
                    //如果字段值为空或者空字符串就设置值为空字符串，用来避免cell中无值串列
                    if(value==null||"".equals(value)){
                        value = "";
                    }
                    //将每一列的值放入List集合中
                    hdata.add(value);
                    //data.put(value,columnName.get(i));
                }
                //将每一行放入key，第几行放入value
                data.put(hdata,columnNum);
                //让行加1
                columnNum++;
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }finally {
            JdbcConnectionUtil.close(connection,preparedStatement,resultSet);
        }
        return data;
    }

    /**
     * @Description 将数据追加写入到Excel中
     * @Author lijinlong
     * @Date   2019/9/4 18:30
     * @Param  [data, excelVersion, filePath]
     * @Return boolean
     * @Exception
     */
    protected static boolean OutputDBData(Map<List<String>,Integer> data,String excelVersion,String filePath){
        InputStream inputStream;
        Workbook workbook = null;
        OutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(filePath.endsWith("/")?filePath+table+"."+excelVersion:filePath+"/"+table+"."+excelVersion);
            if("xlsx".equals(excelVersion)){
                workbook = new XSSFWorkbook(inputStream);
            }else if("xls".equals(excelVersion)){
                workbook = new HSSFWorkbook(inputStream);
            }
            Sheet sheet = workbook.getSheetAt(0);
            inputStream.close();
            data.forEach((datas,size)->{
                ++size;
                Row row = sheet.createRow(size);
                for(int i=0; i<datas.size(); i++){
                    row.createCell(i).setCellValue(datas.get(i));
                }
            });
            //问题:FileOutputStream第二个参数为true追加模式时，数据不会写入到Excel中
            outputStream = new FileOutputStream(filePath.endsWith("/")?filePath+table+"."+excelVersion:filePath+"/"+table+"."+excelVersion);
            workbook.write(outputStream);
            return true;
        } catch (Exception e) {
            logger.info("写入数据时出现异常");
            return false;
        }finally {
            try {
                if(workbook!=null){
                    workbook.close();
                }
                if(outputStream!=null){
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


//    public static void main(String[] args) {
//        PIOOutputExcel.OutputExcel("userDemo","xls","D:");
//    }

}
