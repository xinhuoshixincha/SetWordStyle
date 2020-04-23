package setswordsandsexcelstyle;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * @author 辰轩
 * Date 2020/4/20 10:41
 * 鲜衣怒马少年时
 */


public class OutputDocumentContent
{
     /**输出全部段落的内容，测试用*/
     void outputDocumentContent(XWPFParagraph[] paras)
     {
          /*测试用，输出文档内容-----------------------------------------*/
          for (XWPFParagraph para : paras)
          {
               System.out.println(para.getText());

               System.out.println();
          }
          /*----------------------------------------------------------*/

     }
}
