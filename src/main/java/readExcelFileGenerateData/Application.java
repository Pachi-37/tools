package readExcelFileGenerateData;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class Application {
    private static Logger logger = Logger.getLogger(Application.class.getName());


    /**
     * 将源文件数据量扩展 scale 数据规模的数据
     *
     * @param sourceFile 源文件路径
     * @param scale 数据规模
     * @param targetFile 目标文件路径
     */
    private static void generator(String sourceFile, int scale, String targetFile) {
        // 创建需要写入的数据列表
        List<ExcelData> dataList = new ArrayList<>(2);
        for (int i = 0; i < scale; i++) {
            ExcelData data = DataGenerator.generatorRandomlyData(sourceFile);
            dataList.add(data);
        }


        // 写入数据到工作簿对象内
        Workbook workbook = ExcelWriter.exportData(dataList);

        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            String exportFilePath = targetFile;
            File exportFile = new File(exportFilePath);
            if (!exportFile.exists()) {
                exportFile.createNewFile();
            }

            fileOut = new FileOutputStream(exportFilePath);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            logger.warning("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.warning("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }

    }

    public static void main(String[] args) {
        generator("F:\\project\\tools\\src\\main\\resources\\a.xlsx",1000,"F:\\project\\tools\\src\\main\\resources\\b.xlsx");
    }
}
