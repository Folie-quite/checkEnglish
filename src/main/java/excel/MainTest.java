package excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Author: Dreamer-1
 * Date: 2019-03-01
 * Time: 10:13
 * Description: 示例程序入口类
 */
public class MainTest {

    private static Logger logger = Logger.getLogger(MainTest.class.getName());

    public static void main(String[] args) {
    	long startTime=System.currentTimeMillis();
    	
    	check(1);
    	
    	long endTime=System.currentTimeMillis();
    	System.out.println("程序运行时间： "+(endTime - startTime)+"ms");
    }

    private static void check(Integer type) {
    	/************** 读取Excel流程 ******************/
        // 设定Excel文件所在路径
        String excelFileName = "c:/test/English1.xlsx";
        // 读取Excel文件内容
        List<ExcelDataVO> readResult = ExcelReader.readExcel(excelFileName);
        List<ExcelDataVO> writeList = new ArrayList<ExcelDataVO>();
        List<ExcelDataVO> repetition = new ArrayList<ExcelDataVO>();
        // todo 进行业务操作
        if(1 == type) {
	        for (ExcelDataVO excelDataVO : readResult) {
	        	String english = excelDataVO.getEnglish();
	        	String answer = excelDataVO.getAnswer();
	        	if(StringUtils.isNotBlank(english)) {
	        		if(StringUtils.isBlank(answer)) {
	        			answer = "";
	        		}
	        		if(english.equals(answer)) {
	        			repetition.add(excelDataVO);
	        		} else {
	        			writeList.add(excelDataVO);
	        		}
	        	}
	    	} 
	        ExcelDataVO vo = new ExcelDataVO();
        	vo.setChinese("--end--");
        	vo.setEnglish("错误单词:" + writeList.size() + "个");
        	vo.setAnswer("--end--");
        	writeList.add(vo);
        	//将正确的单词写入excel尾部
        	for(ExcelDataVO excelDataVO : repetition) {
        		writeList.add(excelDataVO);
        	}
        	vo = new ExcelDataVO();
        	vo.setChinese("--end--");
        	vo.setEnglish("正确:" + repetition.size() + "个！");
        	vo.setAnswer("--end--");
        	writeList.add(vo);
		} 
        
        if(2 == type) {
        	Map<String, String> map = new LinkedHashMap<String, String>();
        	//利用map的key唯一特性去重
        	for (ExcelDataVO excelDataVO : readResult) {
        		String english = excelDataVO.getEnglish();
        		String string = map.get(english);
        		if(StringUtils.isNotBlank(string)) {
        			repetition.add(excelDataVO);
        		} else {
        			map.put(excelDataVO.getEnglish(), excelDataVO.getChinese());
        		}
 	    	} 
        	//将map写入list集合中，准备写入excel文件
        	for(String key : map.keySet()) {
        		ExcelDataVO result = new ExcelDataVO();
        		result.setChinese(map.get(key));
        		result.setEnglish(key);
        		writeList.add(result);
        	}
        	//去重结束线
        	ExcelDataVO vo = new ExcelDataVO();
        	vo.setChinese("--end--");
        	vo.setEnglish("剩余单词:" + writeList.size() + "个");
        	vo.setAnswer("--end--");
        	writeList.add(vo);
        	//将重复的单词写入excel尾部
        	for(ExcelDataVO excelDataVO : repetition) {
        		writeList.add(excelDataVO);
        	}
        	vo = new ExcelDataVO();
        	vo.setChinese("--end--");
        	vo.setEnglish("共去重:" + repetition.size() + "个！");
        	vo.setAnswer("--end--");
        	writeList.add(vo);
		}
        
        /************** 写入Excel流程 ******************/

        // 写入数据到工作簿对象内
        Workbook workbook = ExcelWriter.exportData(writeList);

        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            String exportFilePath = "c:/test/writeExample.xlsx";
            File exportFile = new File(exportFilePath);
            if (!exportFile.exists()) {
                exportFile.createNewFile();
            }

            fileOut = new FileOutputStream(exportFilePath);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            logger.warning("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.warning("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }

    }
    
    
}
