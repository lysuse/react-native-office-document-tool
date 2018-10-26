
package tech.youngstream;

import com.facebook.react.bridge.ReactApplicationContext;
import com.facebook.react.bridge.ReactContextBaseJavaModule;

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
}