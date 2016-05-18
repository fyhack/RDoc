package com.fyhack.rdoc;

import com.fyhack.rdoc.vo.CandidateInfo;
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
public class SearchFileByCandidateInfo {
    private boolean DEBUG = false;
    private String filename=null;   //要查找的目录路径
    private BufferedWriter bw=null;
    private String[] fileType=null;   //要查找的文件类型
    private int count =0;

    private ArrayList<CandidateInfo> list;  //


    /**构造函数，
     @param filename 要查找目录的对象
     @param fileType 要查找的文件类型
     */
    public SearchFileByCandidateInfo(String filename, String[] fileType)
    {
        this.filename=filename;
        this.fileType=fileType;
    }

    //暴露的公共接口，开始在指定的目录中搜索关键字
    public List<CandidateInfo> startSearchContent()
    {
        list = new ArrayList<CandidateInfo>();
        count = 0;
        try
        {
            File f=new File( filename );
            listFile(f);  //检索文件
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
                for (String ftype : fileType){
                    if( files[x].getName().endsWith( ftype ))
                    {
                        FindTxt( files[x]); //检索文件内容
                        count++;
                    }
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
        CandidateInfo info = null;

        if(DEBUG) System.out.print("姓名：");
        String name_start_c = "姓名";
        String name_end_c = "性别";
        tmpString = printFindtxt(tmpString,name_start_c,name_end_c,true);
        if (tmpString.value!=null)
            info = new CandidateInfo();
        if (info!=null)
            info.name = tmpString.value;

        if(DEBUG) System.out.print("性别：");
        String sex_start_c = "性别";
        String sex_end_c = "民族";
        tmpString = printFindtxt(tmpString,sex_start_c,sex_end_c,true);
        if (info!=null)
            info.sex = tmpString.value;

        if(DEBUG) System.out.print("出生年月(岁)：");
        String birthday_start_c = "出生日期";
        String birthday_end_c = "工作时间";
        tmpString = printFindtxt(tmpString,birthday_start_c,birthday_end_c,true);
        if (info!=null)
            info.birthday = tmpString.value;

        if(DEBUG) System.out.print("技术职称: ");
        String jishuzhicheng_start_c = "技术职称";
        String jishuzhicheng_end_c = "户籍";
        tmpString = printFindtxt(tmpString,jishuzhicheng_start_c,jishuzhicheng_end_c,false);
        if (info!=null)
            info.jishuzhicheng = tmpString.value;


        if(DEBUG) System.out.print("现单位: ");
        String worktime_start_c = "现单位及进入时间";
        String worktime_end_c = "现岗位及聘任时间";
        tmpString = printFindtxt(tmpString,worktime_start_c,worktime_end_c,false);
        if (info!=null)
            info.exp = tmpString.value;

        if(DEBUG) System.out.print("手机/宅电: ");
        String positions_start_c = "手机/宅电";
        String positions_end_c = "电子邮箱";
        tmpString = printFindtxt(tmpString,positions_start_c,positions_end_c,false);
        if (info!=null)
            info.tel = tmpString.value;

        if(DEBUG) System.out.print("学历/学位: ");
        String school_start_c = "学历/学位";
        String school_end_c = "工作履历";
        tmpString = printFindtxt(tmpString,school_start_c,school_end_c,false);
        if (info!=null)
            info.school = tmpString.value;

        if(DEBUG) System.out.print("第一志愿： ");
        String education_start_c = "第一志愿：";
        String education_end_c = "请明确岗位所属单位";
        tmpString = printFindtxt(tmpString,education_start_c,education_end_c,false);
        if (info!=null)
            info.gangwei = tmpString.value;

//        if(DEBUG) System.out.print("全日制教育毕业院校系及专业: ");
//        String school_start_c = "毕业院校系及专业";
//        String school_end_c = "在职教育";
//        tmpString = printFindtxt(tmpString,school_start_c,school_end_c,false);
//        if (cadreAppointmentAndRemovalApprovalInfo!=null)
//            cadreAppointmentAndRemovalApprovalInfo.school = tmpString.value;
//
//        if(DEBUG) System.out.print("在职教育: ");
//        String workEducation_start_c = "在职教育";
//        String workEducation_end_c = "毕业院校";
//        tmpString = printFindtxt(tmpString,workEducation_start_c,workEducation_end_c,false);
//        if (cadreAppointmentAndRemovalApprovalInfo!=null)
//            cadreAppointmentAndRemovalApprovalInfo.workEducation = tmpString.value;
//
//        if(DEBUG) System.out.print("在职教育毕业院校系及专业: ");
//        String workSchool_start_c = "毕业院校系及专业";
//        String workSchool_end_c = "现任职务";
//        tmpString = printFindtxt(tmpString,workSchool_start_c,workSchool_end_c,false);
//        if (cadreAppointmentAndRemovalApprovalInfo!=null)
//            cadreAppointmentAndRemovalApprovalInfo.workSchool = tmpString.value;
//
//        if(DEBUG) System.out.print("现任职务: ");
//        String job_start_c = "现任职务";
//        String job_end_c = "拟任职务";
//        tmpString = printFindtxt(tmpString,job_start_c,job_end_c,false);
//        if (cadreAppointmentAndRemovalApprovalInfo!=null)
//            cadreAppointmentAndRemovalApprovalInfo.job = tmpString.value;

        if (info!=null)
           list.add(info);
    }

    public String getTextContentByExtractors(File f){
        FileInputStream in = null;
        String text = null;
        try {
            in = new FileInputStream(f);

            if (f.getName().endsWith("docx")){
                XWPFDocument doc2007;
                XWPFWordExtractor word2007;
                doc2007 = new XWPFDocument(in);
                word2007 = new XWPFWordExtractor(doc2007);
                text = word2007.getText();
            } else {
                // 创建WordExtractor
                WordExtractor extractor = new WordExtractor(in);
                // 对doc文件进行提取
                text = extractor.getText();
            }

//            HWPFDocument doc2003;
//            WordExtractor word2003;
//            doc2003 = new HWPFDocument(in);
//            word2003 = new WordExtractor(doc2003);
//
//            text = word2003.getText();

//            System.out.println(text);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            System.out.println(f.getName());
            e.printStackTrace();
        }

        return text;
    }

    private String htmlRemoveTag(String inputString) {
        if (inputString == null)
            return null;
        String htmlStr = inputString; // 含html标签的字符串
        String textStr = "";
        Pattern p_script;
        Matcher m_script;
        Pattern p_style;
        Matcher m_style;
        Pattern p_html;
        Matcher m_html;
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
