package com.hwolf.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * Recursive get all files in the path
 * @author hwolf
 * @email h.wolf@qq.com
 * @date 2017/11/25.
 */
public class ReadFile {
    /**
     *
     * @param absolutePath
     * @return
     * @throws FileNotFoundException
     */
    public List<File> ReadAllFile(String absolutePath) throws FileNotFoundException {
        // InputStream inp = new FileInputStream("workbook.xls");
        File f = new File(absolutePath);
        // read all files
        File[] files = f.listFiles();
        // file list
        List<File> list = new ArrayList<>();
        for (File file : files) {
            // recursive
            if (file.isDirectory()) {
                ReadAllFile(file.getPath());
            } else {
                // add
                list.add(file);
            }
        }
        return list;
    }

    /**
     * read excel
     * @param f
     * @param list
     * @return
     */
    public List<File> fileToList(File f, List<File> list) {
        File[] files = f.listFiles();
        if (files == null){
            return list;
        }
        for (File file : files) {
            // recursive
            if (file.isDirectory()){
                fileToList(file, list);
            }
            else {
                // your wish type
                if (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) {
                    list.add(file.getAbsoluteFile());
                }
            }
        }
        return list;
    }
}
