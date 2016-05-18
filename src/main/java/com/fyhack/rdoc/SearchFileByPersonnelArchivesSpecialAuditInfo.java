package com.fyhack.rdoc;


import com.fyhack.rdoc.vo.PersonnelArchivesSpecialAuditInfo;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * SearchFiles
 * <p/>
 *
 * @author elc_simayi
 * @since 2015/11/5
 */
public class SearchFileByPersonnelArchivesSpecialAuditInfo {
    private boolean DEBUG = false;
    private String filename=null;   //要查找的目录路径
    private BufferedWriter bw=null;
    private String[] fileType=null;   //要查找的文件类型
    private int count =0;

    private ArrayList<PersonnelArchivesSpecialAuditInfo> list;  //


    /**构造函数，
     @param filename 要查找目录的对象
     @param fileType 要查找的文件类型
     */
    public SearchFileByPersonnelArchivesSpecialAuditInfo( String  filename , String[] fileType )
    {
        this.filename=filename;
        this.fileType=fileType;
    }

    //暴露的公共接口，开始在指定的目录中搜索关键字
    public List<PersonnelArchivesSpecialAuditInfo> startSearchContent()
    {
        list = new ArrayList<PersonnelArchivesSpecialAuditInfo>();
        count = 0;
        try
        {
            File f=new File( filename );
            listFile(f);
            System.out.println("检索程序完毕: \t查找到."+fileType+"文件数目"+count+",检索出有效名单数目"+list.size());
        }
        catch( Exception e)
        {
            e.printStackTrace();
            System.out.println("检索程序出错！！！");
        }

        return list;
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
                if( files[x].getName().endsWith( fileType[0] ))
                {
                    FindTxt( files[x]);
                    count++;
                }
            }
        }
    }

    private TmpString printFindtxt(TmpString tmpString , String start_c , String end_c ,boolean checkColon){
        tmpString.value = null;
        int value_start = tmpString.text.indexOf(start_c)+start_c.length();
        if(value_start==-1)
            return tmpString;
        int value_end = tmpString.text.indexOf(end_c, value_start);
        if(value_end==-1)
            return tmpString;

        TmpString newTmpString = new TmpString();
        String value = tmpString.text.substring(value_start, value_end);
        value = getFormatText(htmlRemoveTag(value));

        if(checkColon){
            //过滤冒号
            int colon = value.indexOf("：");
            if (colon == -1)
                colon = value.indexOf(":");
            if (colon!=-1)
                value = value.substring(colon+1);
        }

        newTmpString.value = value;
        if(DEBUG) System.out.println(newTmpString.value);

            //TODO 尾部判断
        newTmpString.text = tmpString.text.substring(value_end);

        return newTmpString;
    }

    private void FindTxt(File f )
    {
        if(DEBUG) System.out.println(f.getName() + ":");
        String text = getTextContentByExtractors(f);
        searchInfo(text);
        if(DEBUG) System.out.println("end.");
    }

    private void searchInfo(String text){
        TmpString tmpString = new TmpString(getFormatText(text));
        PersonnelArchivesSpecialAuditInfo PersonnelArchivesSpecialAuditInfo = null;

        if(DEBUG) System.out.print("姓名：");
        String office_start_c = "姓名";
        String office_end_c = "工作单位及职务";
        tmpString = printFindtxt(tmpString,office_start_c,office_end_c,true);
        if (tmpString.value!=null)
            PersonnelArchivesSpecialAuditInfo = new PersonnelArchivesSpecialAuditInfo();
        if (PersonnelArchivesSpecialAuditInfo!=null)
            PersonnelArchivesSpecialAuditInfo.name = tmpString.value;

        if(DEBUG) System.out.print("工作单位及职务：");
        String site_start_c = "工作单位及职务";
        String site_end_c = "级别";
        tmpString = printFindtxt(tmpString,site_start_c,site_end_c,true);
        if (PersonnelArchivesSpecialAuditInfo!=null)
            PersonnelArchivesSpecialAuditInfo.work_units_and_positions = tmpString.value;

        if(DEBUG) System.out.print("级别：");
        String name_start_c = "级别";
        String name_end_c = "项目";
        tmpString = printFindtxt(tmpString,name_start_c,name_end_c,true);
        if (PersonnelArchivesSpecialAuditInfo!=null)
            PersonnelArchivesSpecialAuditInfo.work_level = tmpString.value;

        if(DEBUG) System.out.print("其他问题：");
        String other_start_c ="其他问题";
        String other_end_c = "审核意见";
        tmpString = printFindtxt(tmpString,other_start_c,other_end_c,false);
        if (PersonnelArchivesSpecialAuditInfo!=null)
            PersonnelArchivesSpecialAuditInfo.other_opinion = tmpString.value;

        if(DEBUG) System.out.print("审核意见: ");
        String introduction_start_c = "审核意见";
        String introduction_end_c = "初审人";
        tmpString = printFindtxt(tmpString,introduction_start_c,introduction_end_c,false);
        if (PersonnelArchivesSpecialAuditInfo!=null)
            PersonnelArchivesSpecialAuditInfo.audit_opinion = tmpString.value;

        if (PersonnelArchivesSpecialAuditInfo!=null)
           list.add(PersonnelArchivesSpecialAuditInfo);
    }

    public String getTextContentByExtractors(File f){
        FileInputStream in = null;
        String text = null;
        try {
            in = new FileInputStream(f);
            // 创建WordExtractor
            WordExtractor extractor = new WordExtractor(in);
            // 对doc文件进行提取
            text = extractor.getText();

//            XWPFDocument doc2007;
//            XWPFWordExtractor word2007;
//            doc2007 = new XWPFDocument(in);
//            word2007 = new XWPFWordExtractor(doc2007);

//            HWPFDocument doc2003;
//            WordExtractor word2003;
//            doc2003 = new HWPFDocument(in);
//            word2003 = new WordExtractor(doc2003);
//
//            text = word2003.getText();

//            System.out.println(text);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return text;
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

    private class TmpString{
        public String text;
        public String value;

        public TmpString(){}

        public TmpString(String text){
            this.text = text;
        }
    }

}
