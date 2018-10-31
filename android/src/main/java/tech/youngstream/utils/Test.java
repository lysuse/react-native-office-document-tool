package tech.youngstream.utils;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test {
      
    public static void main(String[] args) throws Exception {
        testExcel2();
    }

    public static void testWord () throws Exception {
        Map<String, Object> param = new HashMap<String, Object>();
        param.put("${project.name}", "康辉110KV变电站建设");
        param.put("${project.experiment}", "环流测试");

        Map<String,Object> signA = new HashMap<String, Object>();
        signA.put("width", 120);
        signA.put("height", 40);
        signA.put("type", "png");
        signA.put("content", "C:\\Users\\YoungStream\\Desktop\\signA.png");
        param.put("${signA}",signA);

        Map<String,Object> signB = new HashMap<String, Object>();
        signB.put("width", 120);
        signB.put("height", 40);
        signB.put("type", "png");
        signB.put("content", "C:\\Users\\YoungStream\\Desktop\\signB.png");
        param.put("${signB}",signB);
        XWPFDocument doc = WordUtil.generateWord(param, "C:\\Users\\YoungStream\\Desktop\\现场勘察记录.docx");
        FileOutputStream fopts = new FileOutputStream("C:\\Users\\YoungStream\\Desktop\\result.docx");
        doc.write(fopts);
        fopts.close();
    }
    
    public static void testExcel () throws Exception {
        XSSFWorkbook book = new XSSFWorkbook();
        ExcelUtil.createSheetByTitles(book, "项目部员工违章档案",  new String[] {"序号", "违章时间", "违章内容（简要说明）", "违章性质", "违章责任", "本次记分", "累计记分", "备注"});
        FileOutputStream fopts = new FileOutputStream("C:\\Users\\YoungStream\\Desktop\\result.xls");
        book.write(fopts);
        fopts.close();
    }
    
    public static void testExcel2 () throws Exception {
        List<Map> dataList = new ArrayList<>();
        for (int i = 0; i < 100; i++)  {
            Map map = new HashMap();
            map.put("序号", i + 1);
            map.put("违章时间", new Date().toString());
            map.put("违章内容（简要说明）",  "描述内容 ===> " + (i + 1));
            map.put("违章性质",  "操作违章");
            map.put("违章责任",  "个人责任");
            map.put("本次记分",  i+ 2);
            map.put("累计记分",  i+ 4);
            map.put("备注",  "下不为例" + i);
            dataList.add(map);
        }
        ExcelUtil.writeToExcel("C:\\Users\\YoungStream\\Desktop\\result.xls", "项目部员工违章档案", new String[] {"序号", "违章时间", "违章内容（简要说明）", "违章性质", "违章责任", "本次记分", "累计记分", "备注"}, dataList);
    }
}  