package app;

import factory.PDFReaderFactory;
import org.apache.commons.io.FileUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import utils.FileUtil;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 用于将所有的PDF内容截取到result txt里面
 */
public class AnalysisPDFToResult {

    //pdf文件们所在目录
    private static String basePath = "src/main/resources/PDF/";
    //解析完成后，解析结果存放目录
    private static String baseAnalysisResultPath = "src/main/resources/result/";
    //pdf文件内容切割，从……
    private static List<String> subStrFromList = new ArrayList<String>();
    //pdf文件内容切割，到……
    private static List<String> subStrToList = new ArrayList<String>();
    //stock 长度
    protected static int stockLen = 6;

    static {
        subStrFromList.add("内容及提供的资料");
        subStrFromList.add("接待对象");

        subStrToList.add("一、");
        subStrToList.add("三、");
    }

    public static void main(String[] args) {
        List<String> allFilenameStocks = getAllFilenameStocks(basePath);
        ArrayList<File> files = new ArrayList<File>();
        //加载所有file
        FileUtil.listAllFiles(basePath, files);
        for (String stock : allFilenameStocks) {
            StringBuilder sb = new StringBuilder();
            //同一个stock的pdf结果写入同一个txt结果中
            for (File file : files) {
                if (file.getName().startsWith(stock)) {
                    String subStringByFile = subStringByFile(file);
                    sb.append(subStringByFile);
                }
            }
            String result2Write = sb.toString();
            String resultName = stock + ".txt";
            boolean writeSuc = writeResultToFile(baseAnalysisResultPath + resultName, result2Write);
            if (writeSuc) {
                System.out.println("公司：" + stock + "分析完成，写入文件" + resultName);
            }
        }
    }

    /**
     * 读取pdf中文字信息(全部)
     *
     * @param inputFile
     */
    private static String READPDF(File inputFile) {
        //创建文档对象
        PDDocument doc;
        String content;
        try {
            //加载一个pdf对象
            doc = PDDocument.load(inputFile);
            //获取一个PDFTextStripper文本剥离对象
            PDFTextStripper textStripper = PDFReaderFactory.getPDFTextStripper();
            content = textStripper.getText(doc);
            doc.close();
            return content;
        } catch (Exception e) {
            System.out.println("" + inputFile.getPath() + "读取失败");
            return null;
        }
    }

    protected static List<String> getAllFilenameStocks(String path) {
        ArrayList<String> filenames = new ArrayList<String>();
        FileUtil.listAllFilename(path, filenames);
        ArrayList<String> result = new ArrayList<String>();
        if (filenames.size() != 0) {
            for (String filename : filenames) {
                String stock = filename.substring(0, stockLen);
                if (!result.contains(stock)) {
                    result.add(stock);
                }
            }
        }
        return result;
    }


    /**
     * files subString
     */
    private static String subStringByFiles(List<File> files) {
        StringBuilder sb = new StringBuilder();
        for (File file : files) {
            sb.append(subStringByFile(file));
        }
        return sb.toString();
    }

    private static String subStringByFile(File file) {
        String pdfContent = READPDF(file);
        if (pdfContent == null) {
            System.out.println("文件：" + file.getName() + "无法读取内容");
            return "";
        }
        int from = -1;
        for (String subStrFrom : subStrFromList) {
            from = pdfContent.indexOf(subStrFrom);
            if (from != -1) {
                break;
            }
        }
        if (from == -1) {
            System.out.println("文件：" + file.getName() + "无法找到内容切割起始位置");
            return "";
        }
        String fromStr = pdfContent.substring(from);

        int to = -1;
        for (String subStrTo : subStrToList) {
            to = fromStr.indexOf(subStrTo);
            if (to != -1) {
                break;
            }
        }
        if (to == -1) {
            System.out.println("文件：" + file.getName() + "无法找到内容切割终止位置，直接截取到文章末尾，请手动处理");
            return fromStr;
        } else {
            return fromStr.substring(0, to);
        }
    }

    private static List<File> listCompanyByEqualStock(List<File> files, String stock) {
        ArrayList<File> result = new ArrayList<File>();
        for (File file : files) {
            String fileName = file.getName();
            if (fileName.startsWith(stock)) {
                result.add(file);
            }
        }
        return result;
    }

    private static boolean writeResultToFile(String path, String data) {
        try {
            File file = new File(path);
            if (file.exists()) {
                System.out.println("文件：" + path + "已存在");
                return false;
            }
            FileUtils.writeStringToFile(file, data);
            return true;
        } catch (IOException e) {
            System.out.println("文件：" + path + "写入失败");
        }
        return false;
    }
}
