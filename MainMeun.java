package setswordsandsexcelstyle;

import java.io.File;
import java.io.IOException;
import java.util.InputMismatchException;

/**
 * @author 辰轩
 * Date 2020/4/16 11:08
 * 鲜衣怒马少年时
 */


public class MainMeun
{
    private GetValue gv = new GetValue();
    private ShowAllFonts saf = new ShowAllFonts();

    public void startMainMuen(File[] docxfiles, DocxStyle docxStyle, String[] fontnames) throws InterruptedException, IOException
    {
        try
        {
            for (int i = 0; i < docxfiles.length; i++)
            {
                ClearScreen.cls();
                System.out.println("欢迎来到格式生成器，请选择格式\n1、公文格式");
                //存储格式编号
                int num = gv.getStyle(docxfiles[i]);
                switch (num)
                {
                    case 1:
                        //公文格式
                        officeDocumentFormatMainMeum(docxfiles[i], docxStyle,fontnames);
                        GetValue.pause();
                        break;
                    default:
                        i--;
                        break;
                }
            }
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            startMainMuen(docxfiles,docxStyle,fontnames);
            return;
        }
    }


    private void officeDocumentFormatMainMeum(File docxfile, DocxStyle docxStyle, String[] fontnames) throws IOException, InterruptedException//公文格式主菜单
    {
        ClearScreen.cls();
        System.out.println("公文格式菜单：\n1、标准公文格式\n2、自定义公文格式");
        try
        {
            int num=gv.getStyle(docxfile);
            switch (num)
            {
                case 1:
                    //调用标准公文格式
                    officeDocumentFormatStaticMeum(docxfile,docxStyle);
                    GetValue.pause();
                    break;
                case 2:
                    //调用自定义公文格式
                    officeDocumentFormatCustomizeMeum(docxfile,docxStyle,fontnames);
                    GetValue.pause();
                    break;
                default:
                    officeDocumentFormatMainMeum(docxfile,docxStyle,fontnames);
                    return;
            }
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            officeDocumentFormatMainMeum(docxfile,docxStyle,fontnames);
            return;
        }

    }

    /**静态公文格式*/
    public void officeDocumentFormatStaticMeum(File docxfile, DocxStyle docxStyle) throws IOException, InterruptedException
    {
        try
        {
            ClearScreen.cls();
            System.out.println("标准公文格式——"+docxfile.getName());
            //存储标题行数
            int titlenum = gv.getTitleNumber(docxfile);
            docxStyle.setWordStyle(docxfile,titlenum);
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            officeDocumentFormatStaticMeum(docxfile,docxStyle);
            return;
        }
    }

    /**自定义公文格式*/
    public void officeDocumentFormatCustomizeMeum(File docxfile, DocxStyle docxStyle, String[] fontnames)
            throws IOException, InterruptedException//自定义公文格式
    {
        //清屏
        ClearScreen.cls();

        //一个字体名称占位为35
        int placeholder = 35;
        //每行的字体数量
        int number_of_fonts = 3;

        System.out.println("自定义公文格式"+docxfile.getName());

        //输出全部的字体名称
        saf.showAllFonts(fontnames);
        System.out.println("");

        //输出分割线
        for (int i=0;i<placeholder*number_of_fonts;i++)
        {
            System.out.print("-");
        }
        System.out.println("\n自定义公文格式"+docxfile.getName());
        //换行
        System.out.println("");
        //输出分割线
        for (int i=0;i<placeholder*number_of_fonts;i++)
        {
            System.out.print("-");
        }

        /*设置样式---------------------------------------------------------------------*/

        System.out.println("\n请选择输入字体方式：\n1、输入字体编号\n2、输入字体名称");

        //设置标题以及正文字体
        customizeMeumSetFont(docxfile,docxStyle,fontnames);
        //设置是否加粗
        customizeMeumSetBold(docxfile,docxStyle);
        //设置页边距
        customizeMeumSetMargin(docxfile,docxStyle);
        //设置行间距
        customizeMeumSetLineSpace(docxfile, docxStyle);

        /*------------------------------------------------------------------------------------*/

        ClearScreen.cls();
        System.out.println("自定义公文格式"+docxfile.getName());

        //存储标题行数
        int titlenum = gv.getTitleNumber(docxfile);
        //调用setWordStyle方法将自定义的样式应用并输出到文件中
        docxStyle.setWordStyle(docxfile,titlenum);
    }

    /**自定义公文格式菜单设置字体*/
    public void customizeMeumSetFont(File docxfile, DocxStyle docxStyle, String[] fontnames) throws IOException, InterruptedException
    {
        try {
            switch (gv.setHowToSetFont()) {
                case 1:
                    //根据编号设置字体
                    docxStyle.setTitleFont(fontnames[gv.getFontsNumber(true)]);
                    docxStyle.setBodyFont(fontnames[gv.getFontsNumber(false)]);
                    break;
                case 2:
                    //根据名称设置字体
                    String titlename = gv.getFontName(true);
                    if (!determineValueInTheFontArray(fontnames,titlename))
                    {
                        System.out.println("\n输入错误!请输入有效名称!\n");
                        //清空输入缓存
                        gv = new GetValue();
                        GetValue.pause();
                        officeDocumentFormatCustomizeMeum(docxfile,docxStyle,fontnames);
                        return;
                    }
                    String bodyname = gv.getFontName( false);
                    if (!determineValueInTheFontArray(fontnames,titlename))
                    {
                        System.out.println("\n输入错误!请输入有效名称!\n");
                        //清空输入缓存
                        gv = new GetValue();
                        GetValue.pause();
                        officeDocumentFormatCustomizeMeum(docxfile,docxStyle,fontnames);
                        return;
                    }
                    docxStyle.setTitleFont(titlename);
                    docxStyle.setBodyFont(bodyname);
                    break;
                default:
                    customizeMeumSetFont(docxfile, docxStyle, fontnames);
                    return;
            }

            ClearScreen.cls();
            System.out.println("自定义公文格式" + docxfile.getName());
            docxStyle.setTitleFontSize(gv.getFontSize(true));
            //设置标题字体大小
            docxStyle.setBodyFontSize(gv.getFontSize(false));
            //设置正文字体大小
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            officeDocumentFormatCustomizeMeum(docxfile,docxStyle,fontnames);
            return;
        }
        catch (ArrayIndexOutOfBoundsException indexout)
        {
            System.out.println("\n输入错误!请输入有效字号!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            officeDocumentFormatCustomizeMeum(docxfile,docxStyle,fontnames);
            return;
        }
    }

    /**自定义公文格式菜单设置是否加粗*/
    public void customizeMeumSetBold(File docxfile, DocxStyle docxStyle) throws IOException, InterruptedException
    {
        try
        {
            ClearScreen.cls();
            System.out.println("自定义公文格式"+docxfile.getName());
            System.out.println("是否加粗：\n1、加粗\n2、不加粗");
            //设置标题是否加粗
            switch (gv.getIsBold(true))
            {
                case 1:
                    docxStyle.setIsBoldTitle(true);
                    break;
                case 2:
                    docxStyle.setIsBoldTitle(false);
                    break;
                default:
                    customizeMeumSetBold(docxfile,docxStyle);
                    return;
            }
            ClearScreen.cls();
            System.out.println("自定义公文格式"+docxfile.getName());
            System.out.println("是否加粗：\n1、加粗\n2、不加粗");
            //设置正文是否加粗
            switch (gv.getIsBold(false))
            {
                case 1:
                    docxStyle.setIsBoldBody(true);
                    break;
                case 2:
                    docxStyle.setIsBoldBody(false);
                    break;
                default:
                    customizeMeumSetBold(docxfile,docxStyle);
                    return;
            }

        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            customizeMeumSetBold(docxfile,docxStyle);
            return;
        }
    }

    /**自定义公文格式设置页边距*/
    public void customizeMeumSetMargin(File docxfile, DocxStyle docxStyle) throws IOException, InterruptedException
    {
        ClearScreen.cls();
        System.out.println("自定义公文格式"+docxfile.getName());

        try
        {
            //设置页边距
            docxStyle.setMargin(gv.getMargin(0),gv.getMargin(1),gv.getMargin(2),gv.getMargin(3));
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            customizeMeumSetMargin(docxfile,docxStyle);
            return;
        }
    }

    /**自定义公文格式菜单设置行间距*/
    public void customizeMeumSetLineSpace(File docxfile, DocxStyle docxStyle) throws IOException, InterruptedException
    {
        ClearScreen.cls();
        System.out.println("自定义公文格式"+docxfile.getName());
        try
        {
            //设置行间距
            docxStyle.setLineSpace(gv.getLineSpace());
        }
        catch (InputMismatchException ime)
        {
            System.out.println("\n输入错误!请输入有效数字!\n");
            //清空输入缓存
            gv = new GetValue();
            GetValue.pause();
            customizeMeumSetLineSpace(docxfile,docxStyle);
            return;
        }
    }

    /**在字符数组中判断是否有这个名字*/
    private boolean determineValueInTheFontArray(String[] fontnames,String fontname)
    {
        for (String x:fontnames)
        {
            //如果查找到这个字体就返回true
            if (x.equals(fontname))
            {
                return true;
            }
        }
        return false;
    }
}
