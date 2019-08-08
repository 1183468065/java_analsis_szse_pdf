package app;

import org.apache.commons.io.FileUtils;
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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 用于将所有的txt内容，通过分析，写入到每张xls内，用于后续处理
 */
public class AnalysisResultToExcel {

    //result文件们所在目录
    private static String basePath = "src/main/resources/result/";
    //解析完成后，解析结果存放目录
    private static String baseAnalysisResultPath = "src/main/resources/excel/";
    //stock 长度
    private static int stockLen = 6;

    //excel 配置
    private static List<String> sheetNames = new ArrayList<>();
    private static List<String> title = new ArrayList<>();
    //excel 表头
    // 访问时间
    private static final String time = "访问时间";
    // 接待对象
    private static final String reception = "接待对象";

    // 比较严格的匹配年月日
    private static final Pattern patternYMDLong = Pattern
            .compile("^((\\d{2}(([02468][048])|([13579][26]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?(" +
                    "(0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))" +
                    "|(0?2[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|([13579][01345789]))" +
                    "[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])" +
                    "|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|(1[0-9])|" +
                    "(2[0-8]))))))(\\s(((0?[0-9])|([1-2][0-3]))\\:([0-5]?[0-9])((\\s)|(\\:([0-5]?[0-9])))))?$");
    // 匹配年月日
    private static final Pattern patternYMD = Pattern
            .compile("^((?!0000)[0-9]{4}(-|\\/|)((0[1-9]|1[0-2]|[1-9])(-|\\/|)" +
                    "([1-9]|1[0-9]|2[0-8]|0[1-9]|1[0-9]|2[0-8])|(0[13-9]|1[0-2])-" +
                    "(29|30)|(0[13578]|1[02])-31)|([0-9]{2}(0[48]|[2468][048]|[13579][26])|" +
                    "(0[48]|[2468][048]|[13579][26])00)-02-29)$");
    // 匹配年月
    private static final Pattern patternYM = Pattern
            .compile("^((?!0000)[0-9]{4}(-|\\/|)(([1-9]|0[1-9]|1[0-2])))$");

    static {
        sheetNames.add("result");

        title.add(time);
        title.add(reception);
    }

    public static void main(String[] args) throws IOException {
        ArrayList<File> files = new ArrayList<>();
        FileUtil.listAllFiles(basePath, files);
        for (File file : files) {
            System.out.println("开始分析文件：" + file);
            String xlsName = baseAnalysisResultPath + file.getName().substring(0, AnalysisPDFToResult.stockLen) +
                    ".xls";
            contentToExcel(file, xlsName);
        }
    }

    /**
     * 文件分析并写入xls
     */
    private static boolean contentToExcel(File resultFile, String xlsPath) {
        String txt = null;
        try {
            txt = FileUtils.readFileToString(resultFile).trim();
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("文件：" + resultFile + "读取失败");
            return false;
        }
        txt = replaceAllEnter(txt);
        String[] split = txt.split(" ");
        ArrayList<String> list = new ArrayList<>();
        for (int i = 0; i < split.length; i++) {
            String s = split[i].trim();
            if (StringUtil.isNotEmptyOrBlank(s)) {
                list.add(s);
            }
        }
        int rowIndex = 0;//行数
        HashMap<Integer, List<String>> dataMap = new HashMap<>();
        //如果不是日期格式的话，就从map中取出该行数据，然后add
        //如果是日期格式的话，就新建list加入map中
        for (String cellData : list) {
            cellData = cellData.replace(".", "/");//将.换成/这样excel可以格式化日期格式
            if (patternYMD.matcher(cellData).matches() || patternYM.matcher(cellData).matches()) {
                cellData = cellData.replace("-", "/");
                rowIndex = rowIndex + 1;
                List<String> rowData = dataMap.get(rowIndex);
                if (rowData == null) {
                    rowData = new ArrayList<>();
                }
                rowData.add(cellData);
                dataMap.put(rowIndex, rowData);

            } else {
                List<String> rowData = dataMap.get(rowIndex);
                if (rowData != null) {
                    rowData.add(cellData);
                } else {
                    System.out.println("忽略掉单元格数据：" + cellData);
                }
            }
        }
        ArrayList<List<String>> dataList = new ArrayList<>();
        for (Integer row : dataMap.keySet()) {
            List<String> rowList = dataMap.get(row);
            if (rowList != null) {
                dataList.add(rowList);
            }
        }
        boolean succ = createExcelXls(xlsPath, sheetNames, title);
        HashMap<String, List<List<String>>> map = new HashMap<>();
        map.put("result", dataList);//result sheet中的数据
        if (succ) {
            for (String key : map.keySet()) {
                writeToExcelInTurn(xlsPath, key, map.get(key));
            }
        }
        return true;
    }

    private static String replaceAllEnter(String content) {
        Pattern pattern = Pattern.compile("(\r\n|\r|\n|\n\r)");
        Matcher m = pattern.matcher(content);
        if (m.find()) {
            return m.replaceAll(" ");
        }
        return content;
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
        if (file.exists()) {
            System.out.println("文件：" + fileDir + "已存在");
            return false;
        }
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
     * 往excel(xls)中写入(和表头对应写入).
     */
    public static void writeToExcel(String fileDir, String sheetName, List<Map<String, String>> mapList) {

        //创建workbook
        File file = new File(fileDir);
        FileOutputStream fileOutputStream = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
            //文件流
            HSSFSheet sheet = workbook.getSheet(sheetName);
            //获取表头的列数
            int columnCount = sheet.getRow(0).getLastCellNum();
            // 获得表头行对象
            HSSFRow titleRow = sheet.getRow(0);
            //创建单元格显示样式
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            if (titleRow != null) {
                for (int rowId = 0; rowId < mapList.size(); rowId++) {
                    Map<String, String> map = mapList.get(rowId);
                    HSSFRow newRow = sheet.createRow(rowId + 1);
                    newRow.setHeight((short) (20 * 20));//设置行高  基数为20

                    for (short columnIndex = 0; columnIndex < columnCount; columnIndex++) {  //遍历表头
                        String mapKey = titleRow.getCell(columnIndex).toString().trim();
                        HSSFCell cell = newRow.createCell(columnIndex);
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(map.get(mapKey));
                    }
                }
            }

            fileOutputStream = new FileOutputStream(fileDir);
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 往excel(xls)中写入(不和表头对应，顺序写入).
     */
    public static void writeToExcelInTurn(String fileDir, String sheetName, List<List<String>> dataList) {

        //创建workbook
        File file = new File(fileDir);
        FileOutputStream fileOutputStream = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
            //文件流
            HSSFSheet sheet = workbook.getSheet(sheetName);
            //获取表头的列数
            int columnCount = sheet.getRow(0).getLastCellNum();
            // 获得表头行对象
            HSSFRow titleRow = sheet.getRow(0);
            //创建单元格显示样式
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            if (titleRow != null) {
                for (int rowNo = 0; rowNo < dataList.size(); rowNo++) {
                    HSSFRow newRow = sheet.createRow(rowNo + 1);
                    newRow.setHeight((short) (20 * 20));//设置行高  基数为20
                    List<String> rowData = dataList.get(rowNo);
                    for (int index = 0; index < rowData.size(); index++) {
                        HSSFCell cell = newRow.createCell(index);
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(rowData.get(index));
                    }
                }
            }

            fileOutputStream = new FileOutputStream(fileDir);
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
