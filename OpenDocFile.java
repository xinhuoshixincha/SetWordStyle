package setswordsandsexcelstyle;

import java.io.File;

/**
 * @author 辰轩
 * Date 2020/4/14 15:07
 * 鲜衣怒马少年时
 */


public class OpenDocFile
{
    /**打开所有的doc文件*/
    File[] openDocFiles(String path)
    {
        //目录
        File folder = new File(path);
        //声明IsDoc对象
        IsDoc filter = new IsDoc();
        //返回doc文件
        return folder.listFiles(filter);
    }
}
