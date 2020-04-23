package setswordsandsexcelstyle;

import java.io.File;
import java.io.FileFilter;

/**
 * @author 辰轩
 * Date 2020/4/14 15:18
 * 鲜衣怒马少年时
 */


public class IsDOCX implements FileFilter
{

    @Override
    public boolean accept(File pathname)
    {
        return pathname.isFile() && pathname.getName().endsWith(".docx");
    }
}
