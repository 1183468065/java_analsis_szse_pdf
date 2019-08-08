package app;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import utils.FileUtil;
import utils.StringUtil;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * 用于将所有的fianl xls表整合到一张表里面
 */
public class AnalysisFinalToTotal {

    //final文件们所在目录
    private static String basePath = "src/main/resources/final/";
    //解析完成后，解析结果存放目录
    private static String baseAnalysisTotalPath = "src/main/resources/total/";

    //excel 配置
    private static List<String> sheetNames = new ArrayList<>();
    private static List<String> title = new ArrayList<>();
    //excel 表头
    // 访问时间
    private static final String time = "访问时间";
    //对应该列的索引位置
    private static final int timeIndex = 0;
    // 接待对象
    private static final String reception = "接待对象";
    private static final int receptionIndex = 1;
    // Brokern
    private static final String Brokern = "Brokern";
    // Institutionid
    private static final String Institutionid = "Institutionid";

    static {
        sheetNames.add("total");

        title.add(time);
        title.add(reception);
        title.add(Brokern);
        title.add(Institutionid);
    }

    public static void main(String[] args) {
        ArrayList<File> files = new ArrayList<>();
        FileUtil.listAllFiles(basePath, files);
        ArrayList<List<String>> sheetData = new ArrayList<>();

        if (CollectionUtils.isNotEmpty(files)) {
            for (File file : files) {
                List<List<Map<Integer, String>>> lists = readXls(file);
                for (List<Map<Integer, String>> sheet : lists) {
                    for (Map<Integer, String> rows : sheet) {
                        ArrayList<String> rowData = new ArrayList<>();
                        for (Map.Entry<Integer, String> entry : rows.entrySet()) {
                            rowData.add(entry.getValue());
                        }
                        sheetData.add(rowData);
                    }
                }
            }
            String filename = baseAnalysisTotalPath + "total.xls";
            boolean succ = createExcelXls(filename, sheetNames, title);
            if (succ) {
                AnalysisResultToExcel.writeToExcelInTurn(filename, "total", sheetData);
                System.out.println("写入成功");
            }
        }
    }

    /**
     * 创建新excel(xls).
     *
     * @param fileDir    excel的路径
     * @param sheetNames sheet names
     * @param titles     excel的第一行即表格头
     */
    public static boolean createExcelXls(String fileDir, List<String> sheetNames, List<String> titles) {
        File file = new File(fileDir);
        if (!file.exists()) {
            //先得到文件的上级目录，并创建上级目录，在创建文件
            file.getParentFile().mkdir();
            try {
                //创建文件
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //创建workbook
        HSSFWorkbook workbook = new HSSFWorkbook();
        //新建文件
        FileOutputStream fileOutputStream = null;
        HSSFRow row = null;
        try {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
            for (int i = 0; i < sheetNames.size(); i++) {
                workbook.createSheet(sheetNames.get(i));
                workbook.getSheet(sheetNames.get(i)).createRow(0);
                //添加表头, 创建第一行
                row = workbook.getSheet(sheetNames.get(i)).createRow(0);
                row.setHeight((short) (20 * 20));
                for (short j = 0; j < titles.size(); j++) {
                    HSSFCell cell = row.createCell(j, CellType.BLANK);
                    cell.setCellValue(titles.get(j));
                    cell.setCellStyle(cellStyle);
                }
                fileOutputStream = new FileOutputStream(fileDir);
                workbook.write(fileOutputStream);
            }
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fileOutputStream != null) {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return false;
    }

    /**
     * 读取xls格式文件，结果嵌套为：每个sheet数据（行数据（每个表格列索引，数据））
     */
    private static List<List<Map<Integer, String>>> readXls(File file) {
        try {
            ArrayList<List<Map<Integer, String>>> result = new ArrayList<>();
            InputStream is = new FileInputStream(file);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            // 获取每一个工作薄
            for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                ArrayList<Map<Integer, String>> sheetData = new ArrayList<>();
                // 获取当前工作薄的每一行，第一行是表头不要
                for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (hssfRow != null) {
                        short lastCellNum = hssfRow.getLastCellNum();
                        HashMap<Integer, String> dataEntry = new HashMap<>();
                        for (int i = 0; i < lastCellNum; i++) {
                            HSSFCell cell = hssfRow.getCell(i);
                            dataEntry.put(i, cell.getStringCellValue());
                        }
                        sheetData.add(dataEntry);
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

}
