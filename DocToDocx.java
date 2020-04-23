package setswordsandsexcelstyle;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;

/**
 * Date 2020/4/14 14:32
 * 鲜衣怒马少年时
 * @author 辰轩
 *
 * doc文件转换为docx
 */

public class DocToDocx extends DocxStyle
{
    static void docToDocx(File docfile) throws IOException
    {
        //创建对象hwpf，用来操作doc文件
        HWPFDocument hwpf = new HWPFDocument(new FileInputStream(docfile));
        //获取doc文件的range,并存储在range变量中 range：范围
        Range range = hwpf.getRange();
        //方便输出为docx格式,存储传入的doc文件数据
        XWPFDocument docx = new XWPFDocument();

        //将doc文件段落里的字符串输入到docx段落的run中
        for (int i=0;i<range.numParagraphs();i++)
        {
            docx.createParagraph().createRun().setText(range.getParagraph(i).text());
        }

        //创建文件输出流
        FileOutputStream fileOutputStream = new FileOutputStream("./File/Change_"+
                GetValue.removeSuffix(docfile)+".docx");
        //将docx对象中存储的数据写入文件
        docx.write(fileOutputStream);
        //输出提示信息
        System.out.println("已将文件"+docfile.getName()+"转换为Change_"+GetValue.removeSuffix(docfile)+".docx");
        //清空输出缓冲区
        fileOutputStream.flush();
        //关闭文件输出流
        fileOutputStream.close();
    }
}
