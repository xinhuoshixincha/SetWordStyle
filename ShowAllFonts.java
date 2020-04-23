package setswordsandsexcelstyle;

import java.awt.*;

/**
 * @author 辰轩
 * Date 2020/4/16 11:03
 * 鲜衣怒马少年时
 */


class ShowAllFonts
{

    String[] returnAllFonts()
    {
        //获取本地图形环境
        GraphicsEnvironment e = GraphicsEnvironment.getLocalGraphicsEnvironment();
        //获得所有可用的字体名称并返回
        return e.getAvailableFontFamilyNames();
    }

    void showAllFonts(String[] fonts)
    {
        int num = 0;
        int n=0;
        for (String x:fonts)
        {
            System.out.printf("%-1d%-35s\t",n++,"."+x);
            num++;
            if (num>=3)
            {
                System.out.println();
                num=0;
            }
        }
    }
}
