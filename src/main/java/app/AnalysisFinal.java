package app;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class AnalysisFinal {

    private static String institutionidPropPath = "src/main/resources/institutionid.properties";
    private static Map<String, String> institutionidMap = new HashMap<>();

    static {
        printAllProperty(institutionidPropPath);
    }

    public static void main(String[] args) {
//        File file = new File("src/main/resources/excel/000001.xls");
//        List<List<List<String>>> sheetsData = readXls(file);
//        if (sheetsData != null) {
//            for (List<List<String>> sheetData : sheetsData) {
//                for (List<String> sheet : sheetData) {
//                    for (String cell : sheet) {
//                        System.out.println(cell);
//                    }
//                }
//            }
//        }
        System.out.println(institutionidMap);
    }

    /**
     * 读取xls格式文件，结果嵌套为：每个sheet数据（行数据（每个表格数据））
     */
    private static List<List<List<String>>> readXls(File file) {
        try {
            ArrayList<List<List<String>>> result = new ArrayList<>();
            InputStream is = new FileInputStream(file);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            // 获取每一个工作薄
            for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                ArrayList<List<String>> sheetData = new ArrayList<>();
                // 获取当前工作薄的每一行
                for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (hssfRow != null) {
                        ArrayList<String> data = new ArrayList<String>();
                        //行号
//                        HSSFCell index = hssfRow.getCell(0);
                        //读取第0列数据
                        HSSFCell time = hssfRow.getCell(0);
                        data.add(time.getDateCellValue().toString());
                        //读取第1列数据
                        HSSFCell company = hssfRow.getCell(1);
                        data.add(company.getStringCellValue());
                        sheetData.add(data);
                    }
                }
                result.add(sheetData);
            }
            return result;
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("文件：" + file.getName() + "读取失败");
        }
        return null;
    }

    private static void printAllProperty(String filePath) {
        try {

            Properties prop = new Properties();
            InputStream in = new BufferedInputStream(new FileInputStream(filePath));
            prop.load(new InputStreamReader(in, StandardCharsets.UTF_8));
            Set<Map.Entry<Object, Object>> entries = prop.entrySet();
            for (Map.Entry<Object, Object> entry : entries) {
                String key = (String) entry.getKey();
                String value = (String) entry.getValue();
                institutionidMap.put(key.trim(), value.trim());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
