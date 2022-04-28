package helper;

import filtersheets.SheetFilter;
import org.apache.poi.ss.usermodel.Row;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Map;
import java.util.Set;

public class Driver {
    public static void main(String[] args) {
        XLSReader removeFromReader = new XLSReader(System.getProperty("user.home"), "removefrom.xlsx","");
        XLSReader removableReader = new XLSReader(System.getProperty("user.home"), "removable.xlsx","");
/*        JSONObject json = xlsReader.getSheetJSON(0, true);
        System.out.println(json);
        JSONArray result = (JSONArray) json.get("result");
        //System.out.println(result);
        System.out.println("No of results : "+result.length());
        xlsReader.cleanUp();*/
        Set<String> sheetMap = removableReader.convertSheetToSet(0,removableReader.getWorkSheet(0));
        System.out.println(sheetMap);
        SheetFilter.removeDuplicate(removeFromReader.getWorkSheet(0),sheetMap,0);
        removeFromReader.createWorkBook(System.getProperty("user.home"),"newfile.xlsx");
        removableReader.cleanUp();
        removeFromReader.cleanUp();
    }
}
