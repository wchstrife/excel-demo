import com.sun.org.apache.xpath.internal.SourceTree;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;

/**
 * Created by wangchenghao on 2017/8/16.
 */
public class PoiReadExcel {

    public static void main(String[] args){
        File file = new File("D:\\work\\excel\\poi_test.xls");
        try {
            //创建Excel,读取文件
            HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
            //获取工作页
            HSSFSheet sheet = workbook.getSheetAt(0);
            int firstRowNum = 0;
            //获取当前页的左后一行
            int lastRowNum = sheet.getLastRowNum();
            for (int i=0; i<=lastRowNum; i++){
                HSSFRow row = sheet.getRow(i);
                //获取一行有多少单元格
                int lastCellNum = row.getLastCellNum();
                for(int j=0; j<lastCellNum; j++){
                    HSSFCell cell = row.getCell(j);
                    String value = cell.getStringCellValue();
                    System.out.print(value + " ");
                }
                System.out.println();
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
