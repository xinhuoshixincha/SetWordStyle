package setswordsandsexcelstyle;

import java.io.File;
import java.io.FileFilter;

/**
 * @author 辰轩
 * Date 2020/4/14 15:31
 * 鲜衣怒马少年时
 *
 * 判断是不是doc文件
 */


public class IsDoc implements FileFilter
{

    @Override
    public boolean accept(File pathname)
    {
        return pathname.isFile() && pathname.getName().endsWith(".doc");
    }
}
