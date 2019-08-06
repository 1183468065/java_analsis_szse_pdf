package utils;

import java.io.File;
import java.util.List;

public class FileUtil {
    /**
     * 列出文件夹中所有文件
     */
    public static void listAllFiles(String path, List<File> files) {
        File file = new File(path);
        File[] tempList = file.listFiles();
        if (tempList == null) {
            return;
        }
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile()) {
                files.add(tempList[i]);
            }
            if (tempList[i].isDirectory()) {
                listAllFiles(tempList[i].getAbsolutePath(), files);
            }
        }
    }

    /**
     * 列出文件夹中所有文件名
     */
    public static void listAllFilename(String path, List<String> filenames) {
        File file = new File(path);
        File[] tempList = file.listFiles();
        if (tempList == null) {
            return;
        }
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile()) {
                filenames.add(tempList[i].getName());
            }
            if (tempList[i].isDirectory()) {
                listAllFilename(tempList[i].getAbsolutePath(), filenames);
            }
        }
    }
}
