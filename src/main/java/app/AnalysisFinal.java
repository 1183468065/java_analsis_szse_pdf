package app;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import utils.StringUtil;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class AnalysisFinal {

    private static String institutionidPropPath = "src/main/resources/institutionid.properties";
    private static Map<String, String> institutionidMap = new HashMap<>();

    //result文件们所在目录
    private static String basePath = "src/main/resources/result/";
    //解析完成后，解析结果存放目录
    private static String baseAnalysisFinalPath = "src/main/resources/final/";
    //stock 长度
    private static int stockLen = 6;

    //excel 配置
    private static HSSFWorkbook workbook = null;
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
    private static final int BrokernIndex = 2;
    // Institutionid
    private static final String Institutionid = "Institutionid";
    private static final int InstitutionIDIndex = 3;

    //分隔表示符，例如：齐鲁证券，浙商证券、南京；123这种脑残做法
    private static List<String> regexList = new ArrayList<>();

    static {
        loadAllProperty(institutionidPropPath);

        sheetNames.add("result");

        title.add(time);
        title.add(reception);
        title.add(Brokern);
        title.add(Institutionid);

        regexList.add("，");
        regexList.add("、");
    }

    public static void main(String[] args) {
        File file = new File("src/main/resources/excel/000001.xls");
        List<List<Map<Integer, String>>> sheetsData = readXls(file);
        //结果
        ArrayList<List<String>> rowResults = new ArrayList<>();
        int resultRowNo = 1;
        //结果中，每行对应的cellResults
        HashMap<Integer, List<String>> rowAndCellResults = new HashMap<>();
        // 结果顺序，日期，待查询公司，查询到的公司名称，查询到的公司id，查询到的公司名称，查询到的公司id...
        //读取xls格式文件，结果嵌套为：每个sheet数据（行数据（每个表格列索引，数据））
        if (sheetsData != null) {
            for (List<Map<Integer, String>> sheetData : sheetsData) {
                for (Map<Integer, String> rowMap : sheetData) {
                    for (Map.Entry<Integer, String> cellEntry : rowMap.entrySet()) {
                        if (cellEntry.getKey() == timeIndex) {
                            ArrayList<String> rowCellsResults = new ArrayList<>();
                            rowCellsResults.add(cellEntry.getValue());
                            rowAndCellResults.put(resultRowNo, rowCellsResults);
                            continue;
                        }
                        if (cellEntry.getKey() == receptionIndex) {
                            List<String> cellsResults = rowAndCellResults.get(resultRowNo);
                            String time = cellsResults.get(0);//第0个必然是时间
                            //对文本内容切割，分隔符有、，等，这里递归切割，防止脑残用多种符号分割
                            List<String> splits = recursionSplit(0, cellEntry.getValue());
                            for (int i = 0; i < splits.size(); i++) {
                                String reception = splits.get(i);
                                List<String> analysisList = getAnalysisList(reception);
                                if (i == 0) {
                                    // 把当前行补充完整然后rowNo+1，然后新建
                                    cellsResults.add(reception);
                                    cellsResults.addAll(analysisList);
                                    resultRowNo = resultRowNo + 1;
                                } else {
                                    // 新建
                                    ArrayList<String> rowCellsResults = new ArrayList<>();
                                    rowCellsResults.add(time);
                                    rowCellsResults.add(reception);
                                    rowCellsResults.addAll(analysisList);
                                    rowAndCellResults.put(resultRowNo, rowCellsResults);
                                    resultRowNo = resultRowNo + 1;
                                }
                            }
                        }
                    }
                }
            }
        }

        if (rowAndCellResults.size() != 0) {
            for (Map.Entry<Integer, List<String>> entry : rowAndCellResults.entrySet()) {
                List<String> rowData = entry.getValue();
                rowResults.add(rowData);
            }
        }

        String finalXlsName = baseAnalysisFinalPath + file.getName().substring(0, AnalysisPDFMain.stockLen) + ".xls";
        boolean succ = AnalysisResultToExcel.createExcelXls(finalXlsName, sheetNames, title);
        if (succ) {
            AnalysisResultToExcel.writeToExcelInTurn(finalXlsName, "result", rowResults);
        }
    }

    private static List<String> recursionSplit(int regexIndex, String str) {
        String regex = regexList.get(regexIndex);
        String[] split = str.split(regex);
        if (regexIndex == regexList.size() - 1) {
            //最后一个分隔符，把结果收集起来
            return new ArrayList<>(Arrays.asList(split));
        }
        ArrayList<String> result = new ArrayList<>();
        for (int i = 0; i < split.length; i++) {
            String s = split[i];
            List<String> nextRes = recursionSplit(regexIndex + 1, s);
            if (nextRes != null) {
                result.addAll(nextRes);
            }
        }
        return result;
    }

    //顺序为Brokern，Institutionid，Brokern，Institutionid...
    private static List<String> getAnalysisList(String content) {
        List<String> institutionids = getInstitutionids(content);
        if (CollectionUtils.isNotEmpty(institutionids)) {
            return institutionids;
        }
        return new ArrayList<>();
    }

    //顺序为Brokern，Institutionid，Brokern，Institutionid...
    private static List<String> getInstitutionids(String Brokern) {
        ArrayList<String> result = new ArrayList<>();
        String institutionid = institutionidMap.get(Brokern);
        if (StringUtil.isEmptyOrBlank(institutionid)) {
            //遍历找相似的，因为相似的可能有很多，这里都记录下
            for (String key : institutionidMap.keySet()) {
                if (key.contains(Brokern)) {
                    result.add(key);
                    result.add(institutionidMap.get(key));
                }
            }
            if (CollectionUtils.isEmpty(result)) {
                System.out.println(Brokern + "，未找到相同或相似公司");
            }
            return result;
        } else {
            result.add(Brokern);
            result.add(institutionid);
            return result;
        }

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
                        HashMap<Integer, String> dataEntry = new HashMap<>();
                        //读取第0列数据
                        HSSFCell time = hssfRow.getCell(timeIndex);
                        dataEntry.put(timeIndex, time.getStringCellValue());
                        //读取第1列数据
                        HSSFCell reception = hssfRow.getCell(receptionIndex);
                        dataEntry.put(receptionIndex, reception.getStringCellValue());
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

    private static void loadAllProperty(String filePath) {
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
