package app;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import utils.StringUtil;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AnalysisResultToExcel {

    //result文件们所在目录
    private static String basePath = "src/main/resources/result/";
    //解析完成后，解析结果存放目录
    private static String baseAnalysisResultPath = "src/main/resources/excels/";
    //stock 长度
    private static int stockLen = 6;

    //excel 配置
    private static HSSFWorkbook workbook = null;
    private static List<String> sheetNames = new ArrayList<>();
    private static List<String> title = new ArrayList<>();
    //excel 表头
    // 访问时间
    private static final String time = "访问时间";
    // 接待对象
    private static final String reception = "接待对象";
    // 访问时间
    private static final String InstitutionID = "InstitutionID";

    static {
        sheetNames.add("result");
        sheetNames.add("result2");

        title.add(time);
        title.add(reception);
        title.add(InstitutionID);
    }

    public static void main(String[] args) throws IOException {
//        List<String> allFilenameStocks = AnalysisPDFMain.getAllFilenameStocks(basePath);
//        for (String stock : allFilenameStocks) {
//
//        }
        String fileDir = "1.xls";

        createExcelXls(fileDir, sheetNames, title);
        List<List<String>> userList1 = new ArrayList<>();
        ArrayList<String> list1 = new ArrayList<>();
        list1.add("111");
        list1.add("111");
        list1.add("张三");
        list1.add("张三");
        list1.add("111！@#");
        ArrayList<String> list2 = new ArrayList<>();
        list2.add("222");
        list2.add("222");
        list2.add("张三");
        list2.add("111！@#");
        ArrayList<String> list3 = new ArrayList<>();
        list3.add("33");
        list3.add("张三");
        list3.add("张三");
        list3.add("111！@#");
        userList1.add(list1);
        userList1.add(list2);
        userList1.add(list3);

        Map<String, List<List<String>>> users = new HashMap<>();

        users.put("result", userList1);
        users.put("result2", userList1);

        for (String sheetName : sheetNames) {
            List<List<String>> datas = users.get(sheetName);
            writeToExcelInTurn(fileDir, sheetName, datas);
            System.out.println("成功写入文件：" + fileDir + "，sheet：" + sheetName);
        }

//        for (int j = 0; j < sheetNames.size(); j++) {
//            writeToExcelInTurn(fileDir, sheetNames.get(j), users.get(sheetNames.get(j)));
//            System.out.println("成功写入文件：" + fileDir + "，sheet：" + sheetNames.get(j));
//        }

//        String txt = FileUtils.readFileToString(new File("src/main/resources/result/000001.txt")).trim();
//        txt = replaceAllEnter(txt);
//        String[] split = txt.split(" ");
//        ArrayList<String> list = new ArrayList<>();
//        for (int i = 0; i < split.length; i++) {
//            String s = split[i].trim();
//            if (StringUtil.isNotEmptyOrBlank(s)) {
//                list.add(s);
//            }
//        }
//        Pattern pattern = Pattern
//                .compile("^((\\d{2}(([02468][048])|([13579][26]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?(" +
//                        "(0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)
//                        ))" +
//                        "|(0?2[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|
//                        ([13579][01345789]))" +
//                        "[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|((
//                        (0?[469])" +
//                        "|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|(1[0-9])|" +
//                        "(2[0-8]))))))(\\s(((0?[0-9])|([1-2][0-3]))\\:([0-5]?[0-9])((\\s)|(\\:([0-5]?[0-9])))))?$");
//        for (String s : list) {
//            System.out.println(s);
//            Matcher m2 = pattern.matcher(s);
//            if (m2.matches()) {
////                System.out.println("是日期"+s);
//            } else {
////                System.out.println("不是日期" + s);
//            }
//        }

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
    public static void createExcelXls(String fileDir, List<String> sheetNames, List<String> titles) {
        //创建workbook
        workbook = new HSSFWorkbook();
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
                for (short j = 0; j < title.size(); j++) {
                    HSSFCell cell = row.createCell(j, CellType.BLANK);
                    cell.setCellValue(titles.get(j));
                    cell.setCellStyle(cellStyle);
                }
                fileOutputStream = new FileOutputStream(fileDir);
                workbook.write(fileOutputStream);
            }
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
    }

    /**
     * 往excel(xls)中写入(和表头对应写入).
     */
    public static void writeToExcel(String fileDir, String sheetName, List<Map<String, String>> mapList) {

        //创建workbook
        File file = new File(fileDir);
        FileOutputStream fileOutputStream = null;
        try {
            workbook = new HSSFWorkbook(new FileInputStream(file));
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
            workbook = new HSSFWorkbook(new FileInputStream(file));
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
