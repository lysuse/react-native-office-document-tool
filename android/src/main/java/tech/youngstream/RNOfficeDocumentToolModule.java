
package tech.youngstream;

import com.facebook.react.bridge.Promise;
import com.facebook.react.bridge.ReactApplicationContext;
import com.facebook.react.bridge.ReactContextBaseJavaModule;
import com.facebook.react.bridge.ReactMethod;
import com.facebook.react.bridge.ReadableArray;
import com.facebook.react.bridge.ReadableMap;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;

import tech.youngstream.utils.ExcelUtil;
import tech.youngstream.utils.ReadableUtil;
import tech.youngstream.utils.WordUtil;

public class RNOfficeDocumentToolModule extends ReactContextBaseJavaModule {

  private final ReactApplicationContext reactContext;

  public RNOfficeDocumentToolModule(ReactApplicationContext reactContext) {
    super(reactContext);
    this.reactContext = reactContext;
  }

  @Override
  public String getName() {
    return "RNOfficeDocumentTool";
  }
  @ReactMethod
  public void createDocx(String templatePath, String destPath, ReadableMap params, Promise promise) {
    try {
      XWPFDocument document = WordUtil.generateWord(ReadableUtil.toHashMap(params), templatePath);
      FileOutputStream fopts = new FileOutputStream(destPath);
      document.write(fopts);
      fopts.close();
      promise.resolve(1);
    } catch (IOException e) {
      promise.reject(e);
    }
  }

  @ReactMethod
  public void writeToExcel(String destPath, String title, String columns, ReadableArray  datas, Promise promise) {
    try {
      ExcelUtil.writeToExcel(destPath, title, columns.split(","), ReadableUtil.toListMap(datas));
    } catch (Exception e) {
      promise.reject(e);
    }
  }
}