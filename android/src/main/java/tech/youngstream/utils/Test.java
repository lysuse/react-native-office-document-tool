package tech.youngstream.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class Test {
      
    public static void main(String[] args) throws Exception {  
          
        Map<String, Object> param = new HashMap<String, Object>();
        param.put("${username}", "张三");
        param.put("${date}", new Date().toString());
          
        Map<String,Object> header = new HashMap<String, Object>();  
        header.put("width", 100);  
        header.put("height", 150);  
        header.put("type", "png");
        header.put("content", "C:\\Users\\KMYoungStream\\Desktop\\sign.png");
        param.put("${sign}",header);
          
        XWPFDocument doc = WordUtil.generateWord(param, "C:\\Users\\KMYoungStream\\Desktop\\test.docx");
        FileOutputStream fopts = new FileOutputStream("C:\\Users\\KMYoungStream\\Desktop\\result.docx");
        doc.write(fopts);  
        fopts.close();  
    }  
}  