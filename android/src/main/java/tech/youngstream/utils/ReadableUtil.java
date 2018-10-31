package tech.youngstream.utils;

import com.facebook.react.bridge.ReadableArray;
import com.facebook.react.bridge.ReadableMap;
import com.facebook.react.bridge.ReadableMapKeySetIterator;
import com.facebook.react.bridge.ReadableType;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadableUtil {

    public static Map toHashMap(ReadableMap readableMap) {
        Map map = new HashMap();
        if (readableMap == null) return map;
        ReadableMapKeySetIterator iterator = readableMap.keySetIterator();
        while (iterator.hasNextKey()) {
            String key = iterator.nextKey();
            if (readableMap.getType(key) ==   ReadableType.Map) {
                map.put(key,  toHashMap(readableMap.getMap(key)));
            }
            if (readableMap.getType(key) ==   ReadableType.String || readableMap.getType(key) ==   ReadableType.Number || readableMap.getType(key) ==   ReadableType.Null) {
                map.put(key,  readableMap.getString(key));
            }
            if (readableMap.getType(key) ==   ReadableType.Boolean) {
                map.put(key,  readableMap.getBoolean(key));
            }
            if (readableMap.getType(key) ==   ReadableType.Array) {
                ReadableArray array  = readableMap.getArray(key);
                List list = new ArrayList();
                if (array.size() > 0) {
                    for (int i = 0; i < array.size(); i++) {
                        if (array.getType(i) == ReadableType.Map) {
                            list.add(toHashMap(array.getMap(i)));
                        }
                        if (array.getType(i) ==   ReadableType.String || array.getType(i) ==   ReadableType.Number || array.getType(i) ==   ReadableType.Null) {
                            list.add(array.getString(i));
                        }
                        if (array.getType(i) ==   ReadableType.Boolean) {
                            list.add(array.getBoolean(i));
                        }
                    }
                    map.put(key, list);
                }
            }
        }
        return map;
    }

    public static List<Map> toListMap (ReadableArray readableArray) {
        List<Map> list = new ArrayList<>();
        if (readableArray == null) return list;
        for (int i =0; i < readableArray.size(); i++) {
            list.add(toHashMap(readableArray.getMap(i)));
        }
        return  list;
    }
}
