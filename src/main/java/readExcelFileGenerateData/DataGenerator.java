package readExcelFileGenerateData;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.logging.Logger;

/**
 * 读取excel内容
 */
public class DataGenerator {

    private static Logger logger = Logger.getLogger(ExcelReader.class.getName()); // 日志打印类

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    /**
     * 传递一个文件来，根据其具有的字段生成随机数据
     * @param fileName 源文件名
     * @return  生成的数据
     */
    public static ExcelData generatorRandomlyData(String fileName){

        Random random = new Random();

        ExcelData data = new ExcelData();
        String excelFileName = "F:\\project\\tools\\src\\main\\resources\\a.xlsx";
        List<ExcelData> readResult = ExcelReader.readExcel(excelFileName);

        int size = readResult.size();

        data.setBackground(readResult.get(random.nextInt(size)).getBackground());
        data.setScenes(readResult.get(random.nextInt(size)).getScenes());
        data.setTarget(readResult.get(random.nextInt(size)).getTarget());
        data.setPoint(readResult.get(random.nextInt(size)).getPoint());
        data.setHabit(random.nextBoolean());
        data.setBehavior(readResult.get(random.nextInt(size)).getBehavior());
        data.setValue(readResult.get(random.nextInt(size)).getValue());

        return data;
    }
}
