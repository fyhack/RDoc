package com.fyhack.rdoc;

import com.fyhack.rdoc.vo.CadreAppointmentAndRemovalApprovalInfo;
import com.fyhack.rdoc.vo.PersonnelArchivesSpecialAuditInfo;
import com.hankcs.textrank.TextRankKeyword;
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
    public static String file_path = "C:\\Users\\elc_simayi\\Desktop\\新建文件夹";
    public static String file_type = "doc";
    public static String output_xls_file = "C:\\Users\\elc_simayi\\Desktop\\output\\product.xlsx";
    public static String muban_xls_file = "C:\\Users\\elc_simayi\\Desktop\\output\\CadreAppointmentAndRemovalApprovalInfo.xlsx"; //模板文件

    private static Workbook workbook;
    private static XSSFSheet sheet;
    private static XSSFRow row;
    private static XSSFCell code;

    private static StringBuffer stringBuffer;

    public static void main(String args[]){
        System.out.println("检索程序开始: \t" + "目标文件夹位置 " + file_path + ",目标文件类型 " + file_type +
                ", ps'本机系统编码 " + System.getProperty("file.encoding"));

        switch (3){
            case 1:
                SearchFileByPersonnelArchivesSpecialAuditInfo searchFiles1 = new SearchFileByPersonnelArchivesSpecialAuditInfo(file_path,file_type);
                ArrayList<PersonnelArchivesSpecialAuditInfo> list1 = (ArrayList<PersonnelArchivesSpecialAuditInfo>) searchFiles1.startSearchContent();
//                for(int r=0;r<list1.size();r++){
//                    PersonnelArchivesSpecialAuditInfo PersonnelArchivesSpecialAuditInfo = list.get(r);
//                    System.out.println((r+1)+"|"+PersonnelArchivesSpecialAuditInfo.getName() + "|" + PersonnelArchivesSpecialAuditInfo.getWork_units_and_positions()
//                            + "|" + PersonnelArchivesSpecialAuditInfo.getWork_level() + "|" + PersonnelArchivesSpecialAuditInfo.getOther_opinion() + "|" + PersonnelArchivesSpecialAuditInfo.getAudit_opinion());
//                }
                writeXSLToPersonnelArchivesSpecialAuditInfo(list1);
                break;
            case 2:
                SearchFileByCadreAppointmentAndRemovalApprovalInfo searchFiles2 = new SearchFileByCadreAppointmentAndRemovalApprovalInfo(file_path,file_type);
                ArrayList<CadreAppointmentAndRemovalApprovalInfo> list2 = (ArrayList<CadreAppointmentAndRemovalApprovalInfo>) searchFiles2.startSearchContent();
                for(int r=0;r<list2.size();r++){
                    CadreAppointmentAndRemovalApprovalInfo cadreAppointmentAndRemovalApprovalInfo = list2.get(r);
                    System.out.println((r+1)+"|"+cadreAppointmentAndRemovalApprovalInfo.name + "|" + cadreAppointmentAndRemovalApprovalInfo.sex
                            + "|" + cadreAppointmentAndRemovalApprovalInfo.birthday + "|" + cadreAppointmentAndRemovalApprovalInfo.nation + "|" + cadreAppointmentAndRemovalApprovalInfo.birthplace
                            + "|" + cadreAppointmentAndRemovalApprovalInfo.partytime + "|" + cadreAppointmentAndRemovalApprovalInfo.worktime
                            + "|" + cadreAppointmentAndRemovalApprovalInfo.positions + "|" + cadreAppointmentAndRemovalApprovalInfo.education
                            + "|" + cadreAppointmentAndRemovalApprovalInfo.school + "|" + cadreAppointmentAndRemovalApprovalInfo.workEducation
                            + "|" + cadreAppointmentAndRemovalApprovalInfo.workSchool + "|" + cadreAppointmentAndRemovalApprovalInfo.job);
                }
                writeXSLToCadreAppointmentAndRemovalApprovalInfo(list2);
                break;
            case 3:
                SearchFileByCadreAppointmentAndRemovalApprovalInfo searchFiles3 = new SearchFileByCadreAppointmentAndRemovalApprovalInfo(file_path,file_type);
                ArrayList<CadreAppointmentAndRemovalApprovalInfo> list3 = (ArrayList<CadreAppointmentAndRemovalApprovalInfo>) searchFiles3.startSearchContent();
//                writeXSLToCadreAppointmentAndRemovalApprovalInfo(list2);
                break;
        }

        //输出excel
//        writeXSL(list);

        // 检测词频
        /*stringBuffer = new StringBuffer();
        for (PersonnelArchivesSpecialAuditInfo p:list){
            stringBuffer.append(p.getAudit_opinion());
        }
        rankKeyword(stringBuffer.toString());

        List<String> phraseList = HanLP.extractPhrase(stringBuffer.toString(), 20);
        System.out.println("前20短语: "+phraseList);*/
    }

    private static void writeXSLToCadreAppointmentAndRemovalApprovalInfo(ArrayList<CadreAppointmentAndRemovalApprovalInfo> list){
        try {
            workbook = WorkbookFactory.create(new FileInputStream(muban_xls_file));
            FileOutputStream fos = new FileOutputStream(output_xls_file);

            Sheet sheet = workbook.getSheetAt(0);

            for(int r=0;r<list.size();r++){
                CadreAppointmentAndRemovalApprovalInfo cadreAppointmentAndRemovalApprovalInfo = list.get(r);
                for(int c=1;c<=13;c++){
                    Row row = sheet.getRow(r+1);
                    if(row==null)
                        row = sheet.createRow(r+1);
                    Cell cell = row.getCell(c);
                    if(cell==null)
                        cell = row.createCell(c);

                    switch (c){
                        case 1:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.name);
                            break;
                        case 2:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.sex);
                            break;
                        case 3:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.birthday);
                            break;
                        case 4:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.nation);
                            break;
                        case 5:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.birthplace);
                            break;
                        case 6:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.partytime);
                            break;
                        case 7:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.worktime);
                            break;
                        case 8:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.positions);
                            break;
                        case 9:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.education);
                            break;
                        case 10:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.school);
                            break;
                        case 11:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.workEducation);
                            break;
                        case 12:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.workSchool);
                            break;
                        case 13:
                            cell.setCellValue(cadreAppointmentAndRemovalApprovalInfo.job);
                            break;
                    }
                }
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

    private static void writeXSLToPersonnelArchivesSpecialAuditInfo(ArrayList<PersonnelArchivesSpecialAuditInfo> list){
        try {
            workbook = WorkbookFactory.create(new FileInputStream(muban_xls_file));
            FileOutputStream fos = new FileOutputStream(output_xls_file);

            Sheet sheet = workbook.getSheetAt(0);

            for(int r=0;r<list.size();r++){
                PersonnelArchivesSpecialAuditInfo PersonnelArchivesSpecialAuditInfo = list.get(r);
                setValue(sheet,r+1,PersonnelArchivesSpecialAuditInfo);
//                System.out.println((r+1)+"|"+PersonnelArchivesSpecialAuditInfo.getName() + "|" + PersonnelArchivesSpecialAuditInfo.getWork_units_and_positions()
//                        + "|" + PersonnelArchivesSpecialAuditInfo.getWork_level() + "|" + PersonnelArchivesSpecialAuditInfo.getAudit_opinion());

                System.out.println((r + 1) + "|" + PersonnelArchivesSpecialAuditInfo.getName() + "|" + PersonnelArchivesSpecialAuditInfo.getAudit_opinion());
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

    private static void setValue(Sheet sheet, int r, PersonnelArchivesSpecialAuditInfo PersonnelArchivesSpecialAuditInfo){
        for(int c=1;c<=5;c++){
            Row row = sheet.getRow(r);
            if(row==null)
                row = sheet.createRow(r);
            Cell cell = row.getCell(c);
            if(cell==null)
                cell = row.createCell(c);

            switch (c){
                case 1:
                    cell.setCellValue(PersonnelArchivesSpecialAuditInfo.getName());
                    break;
                case 2:
                    cell.setCellValue(PersonnelArchivesSpecialAuditInfo.getWork_units_and_positions());
                    break;
                case 3:
                    cell.setCellValue(PersonnelArchivesSpecialAuditInfo.getWork_level());
                    break;
                case 4:
                    cell.setCellValue(PersonnelArchivesSpecialAuditInfo.getOther_opinion());
                    break;
                case 5:
                    cell.setCellValue(PersonnelArchivesSpecialAuditInfo.getAudit_opinion());
                    break;
            }
        }
    }


    /**
     * 需要在hanlp配置文件中配置data源
     * @param text
     */
    private static void rankKeyword(String text){
        System.out.println("开始检测词频: \t");
        System.out.println("前20关键词: "+new TextRankKeyword().getKeyword("", text));
    }

}
