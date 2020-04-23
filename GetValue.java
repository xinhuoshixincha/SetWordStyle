package setswordsandsexcelstyle;

import java.io.File;
import java.util.InputMismatchException;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.util.regex.Pattern.*;

/**
 * @author 辰轩
 * Date 2020/4/13 11:53
 * 鲜衣怒马少年时
 */


public class GetValue
{
    private Scanner sc = new Scanner(System.in);

    /**从键盘获取文档标题所占行数*/
    int getTitleNumber(File file)
    {

        System.out.print("请输入"+file.getName()+"文档的标题所占行数：");

        return sc.nextInt();
    }

    /**获取字体编号*/
    int getFontsNumber(boolean istitle)
    {
        String name = selectTitle(istitle);

        System.out.print("请输入"+name+"字体编号:");

        return sc.nextInt();
    }

    /**获取想要生成的格式编号*/
    int getStyle(File file) throws InputMismatchException
    {
        System.out.print("请输入文件"+file.getName()+"的格式编号: ");

        return sc.nextInt();
    }

    /**获得标题以及正文字体名字*/
    String getFontName(boolean istitle)
    {
        //清空输入流缓冲
        clearTheInputStreamBuffer();

        String name = selectTitle(istitle);

        System.out.print("请输入"+name+"字体名称:");

        return sc.nextLine();
    }

    /**获得字号*/
    int getFontSize(boolean istitle)
    {
        String name = selectTitle(istitle);

        System.out.print("请输入"+name+"字号:");

        return sc.nextInt();
    }

    /**获得页边距*/
    int getMargin(int n)
    {
        String[] name = {"上","下","左","右"};

        System.out.print("请输入"+name[n]+"侧页边距(单位为mm):");

        return sc.nextInt();
    }

    /**获得行间距*/
    int getLineSpace()
    {
        System.out.print("请输入行间距(单位为磅):");

        return sc.nextInt();
    }

    /**获得是否加粗*/
    int getIsBold(boolean istitle)
    {
        String name = selectTitle(istitle);

        System.out.print("请输入"+name+"序号:");

        return sc.nextInt();
    }

    /**筛选是否为标题*/
    private String selectTitle(boolean istitle)
    {
        if (!istitle)
        {
            return "正文";
        }
        else {
            return "标题";
        }
    }

    /**输入字体的方式，按序号或者按照名字*/
    int setHowToSetFont()
    {
        System.out.print("输入字体输入方式：");

        return sc.nextInt();
    }

    /**暂停*/
    static void pause()
    {
        //键入一个字符
        System.out.println("请输入回车继续");
        new Scanner(System.in).nextLine();
    }

    /**移除文件后缀*/
    static String removeSuffix(File filename)
    {
        //设置正则表达式去除文件后缀
        final Pattern REMOVESUF = compile(".*(?=\\.doc)");
        Matcher m = REMOVESUF.matcher(filename.getName());

        //判断是不是文件夹
        if (filename.isDirectory())
        {
            return null;
        }

        //如果可以匹配到就返回匹配值，否则返回null
        if (m.lookingAt())
        {
            return m.group();
        }
        else
            {
                return null;
            }
    }

    /**清空输入流缓冲*/
    private void clearTheInputStreamBuffer()
    {
        sc = new Scanner(System.in);
    }
}
