package setswordsandsexcelstyle;

import java.io.File;
import java.io.IOException;

/**
 * @author 辰轩
 * Date 2020/4/13 9:35
 * 鲜衣怒马少年时
 */


public class Text
{
    public static void main(String[] args) throws IOException,InterruptedException
    {
        //创建WorkStyle对象，用来规范格式（调用其中的方法）
        DocxStyle docxstyle =new DocxStyle();
        //用来打开doc文件
        OpenDocFile docfile = new OpenDocFile();
        //用来打开docx文件
        OpenDocxFile docxFile = new OpenDocxFile();
        //返回系统中所有的字体
        ShowAllFonts saf = new ShowAllFonts();
        //主菜单
        MainMeun mm = new MainMeun();

        //文档所在的文件夹
        String filepath = "./File/";
        //将系统中的字体存储起来
        String[] fontnames = saf.returnAllFonts();
        //存储所有doc文件
        File[] docfiles = docfile.openDocFiles(filepath);
        /*这里加一个doc文件导入提示，输出文件列表*/
        /*这里加一个doc转docx的方法，然后再扫描docx文件*/
        for (File x:docfiles)
        {
            System.out.println("导入文件:"+x.getName());
            DocToDocx.docToDocx(x);
        }

        //存储所有docx文件
        File[] docxfiles = docxFile.openDocxFiles(filepath);
        /*这里加一个docx文件导入提示,输出文件列表*/
        for (File x:docxfiles)
        {
            System.out.println("导入文件:"+x.getName());
        }

        //序号
        int serialNumber = 0;
        System.out.println("\n文件目录：------------------------------");
        for (File x:docxfiles)
        {
            System.out.println(serialNumber+++"."+x.getName());
        }
        System.out.println("----------------------------------------");

        //暂停
        GetValue.pause();

        //主菜单开始
        mm.startMainMuen(docxfiles,docxstyle,fontnames);

    }
}
