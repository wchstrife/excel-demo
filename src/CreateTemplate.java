import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;

import java.io.File;

/**
 * Created by wangchenghao on 2017/8/16.
 */
public class CreateTemplate {

    /**
     * 创建模板文件
     * @param args
     */
    public static void main(String[] args){
        //获取解析xml文件的路径
        String path = System.getProperty("user.dir")+"/test/student.xml";
        File file = new File(path);
        SAXBuilder builder = new SAXBuilder();
        try{
            //解析xml
            Document parse = builder.build(file);
            //创建Excel
            HSSFWorkbook wb = new HSSFWorkbook();
            //创建sheet
            HSSFSheet sheet = wb.createSheet("Sheet0");

            //获取xml文件根节点
            Element root = parse.getRootElement();
            //获取模板的名称
            String templateName = root.getAttribute("name").getValue();

            int rownum = 0;
            int column = 0;
            //设置列宽
            Element colgroup = root.getChild("colgroup");


        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 设置列宽
     * @param sheet
     * @param colgroup
     */
    private static void setColumnWidth(HSSFSheet sheet, Element colgroup){

    }
}
