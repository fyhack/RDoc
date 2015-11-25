package com.fyhack.rdoc;


import com.fyhack.rdoc.vo.PersonnelInfo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;

/**
 * Main
 * <p/>
 *
 * @author elc_simayi
 * @since 2015/11/5
 */
public class Main {
    public static String file_path = "C:\\Users\\elc_simayi\\Desktop\\审核2";
    public static String file_type = "doc";
    public static String output_xls_file = "C:\\Users\\elc_simayi\\Desktop\\审核2\\test1.xlsx";
    public static String src_xls_file = "C:\\Users\\elc_simayi\\Desktop\\审核2\\test.xlsx";


    private static Workbook workbook;
    private static XSSFSheet sheet;
    private static XSSFRow row;
    private static XSSFCell code;

    public static void main(String args[]){
        System.out.println("检索程序开始: \t" + "目标文件夹位置 " + file_path + ",目标文件类型 " + file_type +
                ", ps'本机系统编码 " + System.getProperty("file.encoding"));

        SearchFile searchFiles = new SearchFile(file_path,file_type);
        ArrayList<PersonnelInfo> list = (ArrayList<PersonnelInfo>) searchFiles.startSearchContent();

        writeXSL(list);
    }

    private static void writeXSL(ArrayList<PersonnelInfo> list){
        try {
            workbook = WorkbookFactory.create(new FileInputStream(src_xls_file));
            FileOutputStream fos = new FileOutputStream(output_xls_file);

            Sheet sheet = workbook.getSheetAt(0);

            for(int r=0;r<list.size();r++){
                PersonnelInfo personnelInfo = list.get(r);
                setValue(sheet,r+1,personnelInfo);
                System.out.println((r+1)+"_"+personnelInfo.getName() + "," + personnelInfo.getWork_units_and_positions()
                        + "," + personnelInfo.getWork_level() + "," + personnelInfo.getAudit_opinion());

            }
            workbook.write(fos);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setValue(Sheet sheet, int r, PersonnelInfo personnelInfo){
        for(int c=1;c<=4;c++){
            Row row = sheet.getRow(r);
            if(row==null)
                row = sheet.createRow(r);
            Cell cell = row.getCell(c);
            if(cell==null)
                cell = row.createCell(c);

            switch (c){
                case 1:
                    cell.setCellValue(personnelInfo.getName());
                    break;
                case 2:
                    cell.setCellValue(personnelInfo.getWork_units_and_positions());
                    break;
                case 3:
                    cell.setCellValue(personnelInfo.getWork_level());
                    break;
                case 4:
                    cell.setCellValue(personnelInfo.getAudit_opinion());
                    break;
            }
        }
    }
}
