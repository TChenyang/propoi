import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

/**
 * @author Administrator
 * @Title: PoiTest
 * @ProjectName propoi
 * @Description: TODO
 * @date 2019/4/11 001119:41
 */
public class PoiTest{
    private  static final String url="jdbc:mysql://localhost:3306/cyang";
    private  static final String root = "root";
    private  static final String password = "123321";
    /*private  static String outPutFile = "."+ File.separator+"user.xls";*/
    private  static final String outPutFile = "D:\\DEV\\Repository\\IOlibrary";
    private  static final int[] colWidth = {
            2000,2000,5000,5000,
            5000,5000,2000,2000};
    @Test
    public void poiTest() throws IOException {
        //穿件工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        //工作表
        XSSFSheet sheet = wb.createSheet("学生信息表");
        //标头行 第一行
        XSSFRow header = sheet.createRow(0);
        //创建单元格 0代表第一行第一列
        XSSFCell cell  = header.createCell(0);
        cell.setCellValue("学号");
        header.createCell(1).setCellValue("姓名");
        header.createCell(2).setCellValue("专业");
        header.createCell(3).setCellValue("班级");
        header.createCell(4).setCellValue("身份证");
        header.createCell(5).setCellValue("宿舍号");
        header.createCell(6).setCellValue("报道日期");
        //设置列的宽度
        //ofCells()代表这行有多少包含数据的列
        for(int i = 0 ;i<header.getPhysicalNumberOfCells();i++){
            // POI设置列宽度时比较特殊，它的基本单位是1/255个字符大小，
            // 因此我们要想让列能够盛的下20个字符的话，就需要用255*20
            sheet.setColumnWidth(i,255*20);
        }
        //设置行高,30像素
        header.setHeightInPoints(30);
        //输出文件要么是\\要么/否则报错
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\DEV\\Repository\\IOlibrary\\poi_demo.xls");
        //向指定文件写入内容
        wb.write(fileOutputStream);
        //close stream
        fileOutputStream.close();
    }
    /**
     * export list/data to excel
     */
    @Test
    public void exportData(){
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;
        try {
            Class.forName("com.mysql.jdbc.Driver");
            connection = DriverManager.getConnection(url,root,password);
            String sql = "select * from t_house_relation";
            preparedStatement = connection.prepareStatement(sql);
            resultSet = preparedStatement.executeQuery();
            ResultSetMetaData resultSetMetaData = resultSet.getMetaData();
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
            HSSFSheet sheet = hssfWorkbook.createSheet("t_house");
            for(int i =0;i<colWidth.length;i++){
                sheet.setColumnWidth(i,colWidth[i]);
            }
            //单元格样式对象
            HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
            hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = null;
            for(int i=0;i<colWidth.length;i++){
                cell = row.createCell(i);
                cell.setCellValue(resultSetMetaData.getColumnLabel(i+1));
                cell.setCellStyle(hssfCellStyle);
            }
            int rowIndex = 1;
            while(resultSet.next()){
                row = sheet.createRow(rowIndex);
                for(int i = 0; i<colWidth.length;i++){
                    cell =row.createCell(i);
                    cell.setCellValue(resultSet.getString(i+1));
                    cell.setCellStyle(hssfCellStyle);
                }
                rowIndex++;
            }
            FileOutputStream fileOutputStream = new FileOutputStream("D:\\DEV\\Repository\\IOlibrary\\t_house_relation.xls");
            hssfWorkbook.write(fileOutputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}








