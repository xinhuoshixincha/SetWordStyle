package setswordsandsexcelstyle;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.io.*;
import java.math.BigInteger;
import java.util.List;

/**
 * @author 辰轩
 * Date 2020/4/13 11:53
 * 鲜衣怒马少年时
 *
 * 设置样式
 */


public class DocxStyle
{
    /**设置标题字体*/
    private String titlefont = "宋体";
    /**设置正文字体*/
    private String bodyfont = "仿宋_GB2312";
    /**设置标题字号*/
    private int titlefontsize = 22;
    /**设置正文字号*/
    private int bodyfontsize = 16;
    /**设置页边距*/
    private long margintop = 37L,marginbottom = 35L,marginleft = 28L,marginright = 23L;
    /**设置行间距*/
    private int linespacing = 28;
    /**设置标题是否加粗*/
    private boolean isboldTitle = true;
    /**设置正文是否加粗*/
    private boolean isIsboldBody = false;
    /**设置行前间距*/
    private int lineBefore = 0;
    /**设置行后间距*/
    private int lineAfter = 0;


    /**设置docx样式*/
    void setWordStyle(File path, int titlenum) throws IOException
    {
        /*导入word----------------------------------------------------*/
        System.out.println("导入文档:"+path.getName());
        XWPFDocument doc = new XWPFDocument(new FileInputStream(path));
        //存储全部段落
        List<XWPFParagraph> paras = doc.getParagraphs();
        /*-----------------------------------------------------------*/

        /*设置页边距--------------------------------------------------*/
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        CTPageMar ctPageMar = sectPr.getPgMar();

        //底部35mm+-1mm
        ctPageMar.setBottom(BigInteger.valueOf(marginbottom*57L));
        //顶部37mm+-1mm
        ctPageMar.setTop(BigInteger.valueOf(margintop*57L));
        //左侧28mm+-1mm
        ctPageMar.setLeft(BigInteger.valueOf(marginleft*57L));
        //右侧26mm+-1mm
        ctPageMar.setRight(BigInteger.valueOf(marginright*57L));
        /*-----------------------------------------------------------*/

        /*设置行间距---------------------------------------------*/
        for (XWPFParagraph para : paras)
        {
            CTPPr ppr = para.getCTP().getPPr();
            if (ppr == null)
            {
                System.out.println(false);
                ppr = para.getCTP().addNewPPr();
            }
            CTSpacing spacing = ppr.isSetSpacing()?ppr.getSpacing():ppr.addNewSpacing();
            spacing.setAfter(BigInteger.valueOf(lineAfter*20));
            spacing.setBefore(BigInteger.valueOf(lineBefore*20));
            //设置行间距格式为静态
            spacing.setLineRule(STLineSpacingRule.EXACT);
            //linespacing磅
            spacing.setLine(BigInteger.valueOf(linespacing*20));
        }
        /*--------------------------------------------------------------*/

        /*设置标题样式---------------------------------------------------*/
        for (int i=0;i<titlenum;i++)
        {
            List<XWPFRun> runs = paras.get(i).getRuns();
            //段落居中
            paras.get(i).setAlignment(ParagraphAlignment.CENTER);
            for (XWPFRun x : runs)
            {
                CTFonts fonts = x.getCTR().addNewRPr().addNewRFonts();
                //设置中文字体
                fonts.setEastAsia(titlefont);
                //设置英文以及数字字体
                x.setFontFamily(titlefont);

                //设置字号
                x.setFontSize(titlefontsize);
                //加粗
                x.setBold(isboldTitle);
            }
        }
        /*-----------------------------------------------------------*/

        /*设置正文字体，跳过标题行------------------------------------------------*/
        for (int i=titlenum;i<paras.size();i++)
        {
            List<XWPFRun> runs = paras.get(i).getRuns();
            //向左对齐
            paras.get(i).setAlignment(ParagraphAlignment.LEFT);
            //段首缩进 字号为16 缩进2格
            paras.get(i).setIndentationFirstLine(16*20*2);
            //设置除标题以外全部段落的样式
            for (XWPFRun x:runs)
            {
                CTFonts fonts = x.getCTR().addNewRPr().addNewRFonts();
                //设置正文中文字体
                fonts.setEastAsia(bodyfont);
                //设置正文英文及数字字体
                x.setFontFamily(bodyfont);
                //设置正文字号
                x.setFontSize(bodyfontsize);
                //是否加粗
                x.setBold(isIsboldBody);
            }
        }
        /*------------------------------------------------------------*/

        /*将规范样式后的docx文档输出到文件中-----------------------------------*/
        try
        {
            //相对路径 jar包下面的out文件夹中的对应名字的docx文档
            String newfilename = "./out/"+path.getName();
            //创建输出流
            FileOutputStream outputStream = new FileOutputStream(newfilename);
            //将导入的文档输出到文件中
            doc.write(outputStream);
            //输出提示信息
            System.out.println("已输出文件"+path.getName()+"\n");
            outputStream.flush();
        }
        catch (IOException ie)
        {
            ie.printStackTrace();
        }
        /*------------------------------------------------------------*/
    }

    /**设置标题字体*/
    void setTitleFont(String titlefont)
    {
        this.titlefont = titlefont;
    }

    /**设置正文字体*/
    void setBodyFont(String otherFont)
    {
        this.bodyfont = otherFont;
    }

    /**设置标题字号*/
    void setTitleFontSize(int size)
    {
        this.titlefontsize = size;
    }

    /**设置正文字号*/
    void setBodyFontSize(int size)
    {
        this.bodyfontsize = size;
    }

    /**设置页边距*/
    void setMargin(long top, long bottom, long left, long right)
    {
        this.margintop = top;
        this.marginbottom = bottom;
        this.marginleft = left;
        this.marginright = right;
    }

    /**设置行间距*/
    void setLineSpace(int line)
    {
        this.linespacing = line;
    }

    /**设置标题是否加粗*/
    void setIsBoldTitle(boolean isbold)
    {
        this.isboldTitle = isbold;
    }

    /**设置正文是否加粗*/
    void setIsBoldBody(boolean isbold)
    {
        this.isIsboldBody = isbold;
    }

}
