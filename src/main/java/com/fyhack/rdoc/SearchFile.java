package com.fyhack.rdoc;


import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * SearchFiles
 * <p/>
 *
 * @author elc_simayi
 * @since 2015/11/5
 */
public class SearchFile {
    private File f=null;     //要查找的目录对象
    private String filename=null;   //要查找的目录路径
    private BufferedWriter bw=null;
    private String findtxt=null;    //要查找的文本内容
    private String fileType=null;   //要查找的文件类型
    private int totalFileCount=0;   //共搜索的文件数
    private int findedFileCount=0;  //搜索到有用的文件数
    private int findContentCount=0; //搜索到的有用信息数目
    private int columns = 0;    //用于写入xsl的列号

    XSSFWorkbook workBook;
    XSSFSheet sheet;
    XSSFRow row;
    XSSFCell code;

    /**构造函数，
     @param filename 要查找目录的对象
     @param findtxt    要查找的关键字
     @param fileType 要查找的文件类型
     */
    public SearchFile( String  filename ,String findtxt, String fileType )
    {
        this.filename=filename;
        this.findtxt=findtxt;
        this.fileType=fileType;
    }

    private void writeXSL(int l , String text){
        // 在指定的索引处创建一列（单元格）
        code = row.createCell(l);
        // 定义单元格为字符串类型
        code.setCellType(XSSFCell.CELL_TYPE_STRING);

        // 在单元格输入内容
        XSSFRichTextString codeContent = new XSSFRichTextString(getFormatText(text));
        code.setCellValue(codeContent);

    }

    //暴露的公共接口，开始在指定的目录中搜索关键字
    public void startSearchContent()
    {
        // 创建工作薄
        workBook = new XSSFWorkbook();
        // 在工作薄中创建一工作表
        sheet = workBook.createSheet();

        try
        {
            f=new File( filename );
            listFile( f );
            System.out.println("搜索完毕");

            FileOutputStream fos = new FileOutputStream("C:/Users/elc_simayi/Desktop/hos.xlsx");
            workBook.write(fos);
            fos.flush();
            //操作结束，关闭流
            fos.close();
        }
        catch( Exception e)
        {
            e.printStackTrace();
            System.out.println("搜索出错！！！");
        }


    }

    /*
    通过递归搜索目录，搜索过程分两种情况：
    1.如果是目录，则通过递归继续查找目录下的文件
    2.如果是文件，则先判断是否是fileType类型文件，如果是的话就搜索文件内容
    */
    private void listFile( File f )
    {
        File[] files = f.listFiles();
        for(int x=0; x<files.length; x++)
        {
            if(files[x].isDirectory())
                listFile( files[x] );
            else
            {
                //判断文件名是否以fileType结尾
                if( files[x].getName().endsWith( fileType ))
                {
                    FindTxt( files[x]);
                }
            }
        }
    }

    private String printFindtxt(String text , String start_c , String end_c){
        int name_start = text.indexOf(start_c)+start_c.length();
        if(name_start==-1)
            return text;
        int name_end = text.indexOf(end_c,name_start);
        if(name_end==-1)
            return text;

        String name = text.substring(name_start, name_end);
        name = htmlRemoveTag(name);
        System.out.println(name);
        writeXSL(columns,name);

            //TODO 尾部判断
        return text = text.substring(name_end);
    }

    /*
    从文件中搜索制定的内容，分下面几步
    1.使用自定义的山寨版LineNumberReader类，读取文件的每一行
    2.
    */
    private void FindTxt(File f )
    {
        System.out.println(f.getName() + ":");

//        String text = getTextContent(f);
//
//        findInfo(text);

//        System.out.println(text);

        columns = 0;
        // 在指定的索引处创建一行
        row = sheet.createRow(totalFileCount);
        test(f);
        totalFileCount++; //搜索到的文件数加1
    }

    private void findInfo(String text){

        System.out.print("应聘职位: ");
        String office_start_c = "应聘职位：";
        String office_end_c = "\n";
        text = printFindtxt(text,office_start_c,office_end_c);

        System.out.print("工作地点: ");
        String site_start_c = "工作地点：";
        String site_end_c = "\n";
        text = printFindtxt(text,site_start_c,site_end_c);

        System.out.print("应聘部门： ");
        String depart_start_c = "应聘部门：" ;
        String depart_end_c = "\n" +
                "                            \n" +
                "            \n" +
                "                \n" +
                "                        \n" +
                "            \n" +
                "            \n" +
                "               ";
        text = printFindtxt(text,depart_start_c,depart_end_c);

        text = text.substring(text.indexOf("更新日期"));

        System.out.print("姓名：");
        String name_start_c = "\n" +
                "            \n" +
                "        \n" +
                "    \n" +
                "\n" +
                "\n" +
                "\n" +
                "    \n" +
                "        \n" +
                "            \n" +
                "                ";
        String name_end_c = "                \n" +
                "                \n" +
                "            \n" +
                "                \n";
        text = printFindtxt(text,name_start_c,name_end_c);

        System.out.print("个人信息: ");
        String introduction_start_c = "                \n" +
                "                \n" +
                "            \n" +
                "                \n";
        String introduction_end_c = "验";
        text = printFindtxt(text,introduction_start_c,introduction_end_c);

        System.out.print("手机: ");
        String tel_start_c = "手机：";
        String tel_end_c = "电子邮件";
        text = printFindtxt(text, tel_start_c, tel_end_c);

        System.out.print("工作经历: ");
        String workhistory1_start_c = "工作经历";
        String workhistory1_end_c = "教育背景";
        text = printFindtxt(text,workhistory1_start_c,workhistory1_end_c);

//        String workhistory2_start_c = "<span style='font-family:SimSun;mso-ascii-font-family:Calibri;mso-hansi-font-family: Calibri'>";
//        String workhistory2_end_c = "</span>";
//        text = printFindtxt(text, workhistory2_start_c, workhistory2_end_c);

        System.out.print("教育背景: ");
        String educational1_start_c = "教育背景";
        String educational1_end_c = "专业描述";
        text = printFindtxt(text,educational1_start_c,educational1_end_c);

//        String educationa2_start_c = "<span style='font-family:SimSun;  mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri'>";
//        String educationa2_end_c = "</span>";
//        text = printFindtxt(text,educationa2_start_c,educationa2_end_c);

        System.out.println("-----------------------------------");
        System.out.print("\n");
    }

    public void test(File f){
        try {

            FileReader fileReader = new FileReader(f);
            char[] ch = new char[1024 * 200];
            int len = fileReader.read(ch);
            fileReader.close();

            String text = new String(ch,0,len);

            System.out.print("应聘职位: ");
            columns = 2;
            String office_start_c = "应聘职位：<b style='mso-bidi-font-weight:normal'>";
            String office_end_c = "</b>";
            text = printFindtxt(text,office_start_c,office_end_c);

            System.out.print("工作地点: ");
            columns = 3;
            String site_start_c = "工作地点：<b style='mso-bidi-font-weight:normal'>";
            String site_end_c = "</b>";
            text = printFindtxt(text,site_start_c,site_end_c);

//            String depart_start_c = "应聘部门：<b\n" +
//                    "                            style='mso-bidi-font-weight:normal'>";
//            String depart_end_c = "</b>";
//            text = printFindtxt(text,depart_start_c,depart_end_c);

            System.out.print("姓名: ");
            columns = 1;
            String name_start_c = "<span style='font-size:36.0pt;mso-bidi-font-size:22.0pt; line-height:115%;font-family:SimSun;mso-ascii-font-family:Calibri;mso-hansi-font-family: Calibri'>";
            String name_end_c = "</span>";
            text = printFindtxt(text,name_start_c,name_end_c);

            System.out.print("个人信息: ");
            columns = 4;
            String introduction_start_c = "<span style='font-family:SimSun;mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri'>";
            String introduction_end_c = "</span>";
            text = printFindtxt(text,introduction_start_c,introduction_end_c);

            System.out.print("手机: ");
            columns = 5;
            String tel_start_c = "手机：</span></p></td><td width=572 valign=top style='width:428.9pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><span lang=EN-US>";
            String tel_end_c = "</span>";
            text = printFindtxt(text, tel_start_c, tel_end_c);

            System.out.print("工作经历: ");
            columns = 6;
            String workhistory1_start_c = "<span style='font-family:SimSun;mso-ascii-font-family:Calibri;mso-hansi-font-family: Calibri'>";
            String workhistory1_end_c = "]";
            text = printFindtxt(text,workhistory1_start_c,workhistory1_end_c);

            columns = 7;
            String workhistory2_start_c = "<span style='font-family:SimSun;mso-ascii-font-family:Calibri;mso-hansi-font-family: Calibri'>";
            String workhistory2_end_c = "]";
            text = printFindtxt(text, workhistory2_start_c, workhistory2_end_c);

            columns = 8;
            System.out.print("教育背景: ");
            String educational1_start_c = "<span style='font-family:SimSun;  mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri'>";
            String educational1_end_c = "                    ";
            text = printFindtxt(text,educational1_start_c,educational1_end_c);

            columns = 9;
            String educationa2_start_c = "<span style='font-family:SimSun;  mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri'>";
            String educationa2_end_c = "                    ";
            text = printFindtxt(text,educationa2_start_c,educationa2_end_c);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("-----------------------------------");
        System.out.print("\n");
    }

    public String getTextContent(File f){
        try {
            FileReader fileReader = new FileReader(f);
            char[] ch = new char[1024 * 200];
            int len = 0;
            len = fileReader.read(ch);
            fileReader.close();

            String text = new String(ch,0,len);

            return htmlRemoveTag(text);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    private String htmlRemoveTag(String inputString) {
        if (inputString == null)
            return null;
        String htmlStr = inputString; // 含html标签的字符串
        String textStr = "";
        java.util.regex.Pattern p_script;
        java.util.regex.Matcher m_script;
        java.util.regex.Pattern p_style;
        java.util.regex.Matcher m_style;
        java.util.regex.Pattern p_html;
        java.util.regex.Matcher m_html;
        try {
            //定义script的正则表达式{或<script[^>]*?>[\\s\\S]*?<\\/script>
            String regEx_script = "<[\\s]*?script[^>]*?>[\\s\\S]*?<[\\s]*?\\/[\\s]*?script[\\s]*?>";
            //定义style的正则表达式{或<style[^>]*?>[\\s\\S]*?<\\/style>
            String regEx_style = "<[\\s]*?style[^>]*?>[\\s\\S]*?<[\\s]*?\\/[\\s]*?style[\\s]*?>";
            String regEx_html = "<[^>]+>"; // 定义HTML标签的正则表达式
            p_script = Pattern.compile(regEx_script, Pattern.CASE_INSENSITIVE);
            m_script = p_script.matcher(htmlStr);
            htmlStr = m_script.replaceAll(""); // 过滤script标签
            p_style = Pattern.compile(regEx_style, Pattern.CASE_INSENSITIVE);
            m_style = p_style.matcher(htmlStr);
            htmlStr = m_style.replaceAll(""); // 过滤style标签
            p_html = Pattern.compile(regEx_html, Pattern.CASE_INSENSITIVE);
            m_html = p_html.matcher(htmlStr);
            htmlStr = m_html.replaceAll(""); // 过滤html标签
            textStr = htmlStr;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return textStr;// 返回文本字符串
    }

    public String getFormatText(String inputString){
        String dest = null;
        if (inputString!=null) {
            Pattern p = Pattern.compile("\\s*|\t|\r|\n");
            Matcher m = p.matcher(inputString);
            dest = m.replaceAll("");
        }
        return dest;
    }

    //输出搜索的统计信息
    private void showInfo() throws IOException
    {
        bw.write( "        搜索关键字："+findtxt);
        bw.newLine();
        bw.write("        共搜索的" + fileType + "文件数：" + totalFileCount);
        bw.newLine();
        bw.write( "        关键文件数："+findedFileCount);
        bw.newLine();
        bw.write( "        搜索到的关键字数目："+findContentCount );
        bw.newLine();
    }
}
