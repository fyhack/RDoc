package com.fyhack.rdoc;


/**
 * Main
 * <p/>
 *
 * @author elc_simayi
 * @since 2015/11/5
 */
public class Main {
    public static void main(String args[]){
        System.out.print("本机系统编码:" + System.getProperty("file.encoding") + "\n");

        SearchFile searchFiles = new SearchFile("C:\\Users\\elc_simayi\\Desktop\\5","姓名","doc");
        searchFiles.startSearchContent();

    }
}
