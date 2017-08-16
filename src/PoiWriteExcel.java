import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileOutputStream;

public class PoiWriteExcel {

    public static void main(String[] args) {

        String[] title = {"id", "name", "sex"};

        //创建一个Excel
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建一个工作表
        HSSFSheet sheet = workbook.createSheet();
        //创建第一行
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = null;

        //插入表头
        for(int i=0; i<title.length; i++){
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }

        //插入数据
        for(int i=1; i<=10; i++){
            HSSFRow nextrow = sheet.createRow(i);
            HSSFCell cell2 = nextrow.createCell(0);

            cell2.setCellValue("a" + i);
            cell2 = nextrow.createCell(1);
            cell2.setCellValue("user" + i);
            cell2 = nextrow.createCell(2);
            cell2.setCellValue("男");
        }

        //创建文件
        try{
            File file = new File("D:\\work\\excel\\poi_test.xls");
            //将内容写入文件
            file.createNewFile();
            FileOutputStream stream = FileUtils.openOutputStream(file);
            workbook.write(stream);

            stream.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
