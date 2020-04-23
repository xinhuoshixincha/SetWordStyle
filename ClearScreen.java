package setswordsandsexcelstyle;

import java.io.IOException;

/**
 * @author 辰轩
 * Date 2020/4/17 11:23
 * 鲜衣怒马少年时
 *
 * 清屏
 */


public class ClearScreen
{
    /**清屏方法*/
    public static void cls() throws IOException,InterruptedException
    {
        new ProcessBuilder("cmd","/c","cls").inheritIO().start().waitFor();
    }
}
