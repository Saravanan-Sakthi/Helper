package helper;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.*;

public class XLSReader {
    private Workbook workBook;
    private String directoryPath = System.getProperty("user.home");
    private String path = directoryPath;
    private PrintWriter printWriter = null;

    public XLSReader(String directoryPath, String fileName, String password){
        Biff8EncryptionKey.setCurrentUserPassword(password);
        if(path!=null && !path.isEmpty()) {
            this.path = directoryPath+"/"+fileName;
        }
        loadWorkBook();
    }

    private void loadWorkBook() {
        FileInputStream inputStream = null;
        try{
            inputStream = new FileInputStream(this.path);
            if(this.path.endsWith(".xls")) {
                this.workBook = new HSSFWorkbook(inputStream);
            } else if(this.path.endsWith(".xlsx")) {
                this.workBook = new XSSFWorkbook(inputStream);
            } else{
                System.out.println("Invalid path to the file");
            }
        } catch (FileNotFoundException e) {
            System.out.println("File is not present in the given path");
        }catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
                System.out.println("Closed Input Stream");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public Sheet getWorkSheet(int sheetIndex) {
        return workBook.getSheetAt(sheetIndex);
    }

    public Set<String> convertSheetToSet(int columnIndex , Sheet workSheet) {
        Set<String> set = new HashSet<>();
        for (Row row : workSheet) {
            if(row.getRowNum() == 0) {
                continue;
            }
            set.add(row.getCell(columnIndex).toString());
        }
        return set;
    }

    public Map<String,Row> convertSheetToMap(int keyColumnIndex , Sheet workSheet) {
        Map<String, Row> sheetMap = new HashMap<>();
        for(Row row : workSheet) {
            if(row.getRowNum() == 0) {
                continue;
            }
            sheetMap.put(row.getCell(keyColumnIndex).toString(),row);
        }
        return sheetMap;
    }

    private JSONObject convertSheetToJSON(Sheet workSheet, boolean createFile){
        JSONArray jsonArray = new JSONArray();
        for(Row row : workSheet) {
            if(row.getRowNum() == 0) {
                continue;
            }
            JSONObject json = new JSONObject();
            Row firstRow = workSheet.getRow(0);
            Iterator<Cell> headings = firstRow.cellIterator();
            for(Cell cell : row) {
                json.put(headings.next().toString(),cell.toString());
            }
            jsonArray.put(json);
            if(createFile) {
                writeToFile(json);
            }
        }
        try {
            this.workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new JSONObject().put("result", jsonArray);
    }

    private void writeToFile(JSONObject json) {
        if(this.printWriter == null) {
            try {
                File outputFile = new File(directoryPath+"/output.txt");
                if(outputFile.exists()){
                    outputFile.delete();
                    System.out.println("Deleted existing file with same name");
                }
                if(outputFile.createNewFile()) {
                    System.out.println("Created new file");
                }
                this.printWriter = new PrintWriter(outputFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        this.printWriter.println(json);
    }

    public JSONObject getSheetJSON(String sheetName, boolean createFile){
        return convertSheetToJSON(workBook.getSheet(sheetName), createFile);
    }

    public JSONObject getSheetJSON(int sheetIndex, boolean createFile){
        return convertSheetToJSON(workBook.getSheetAt(sheetIndex), createFile);
    }

    public void createWorkBook (String path, String fileName) {
        System.out.println("creating workbook");
        String fullPath = path+"/"+fileName;
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(fullPath);
            workBook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
            }
        }
    }

    public void cleanUp(){
        try {
            if(this.printWriter != null) {
                this.printWriter.close();
                System.out.println("Closed print writer");
            }
            if(this.workBook != null){
                this.workBook.close();
                System.out.println("Closed workbook");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
