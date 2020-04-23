package setswordsandsexcelstyle;

import java.io.File;

/**
 * @author 辰轩
 * Date 2020/4/14 14:11
 * 鲜衣怒马少年时
 */


public class OpenDocxFile
{
    /**打开所有的docx文件*/
    File[] openDocxFiles(String path)
    {
        //目录
        File folder = new File(path);
        //声明IsDocx对象
        IsDOCX filter = new IsDOCX();
        //返回docx文件
        return folder.listFiles(filter);
    }

}
