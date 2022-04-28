package filtersheets;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Map;
import java.util.Set;

public class SheetFilter {
    public static Sheet removeDuplicate(Sheet removeFromSheet, Set<String> removableDataSheet, int removableDataIndex) {
        System.out.println("removing duplicates");
        ArrayList<Integer> indices = new ArrayList<>();
        for(Row row : removeFromSheet) {
            if(row.getRowNum() == 0) {
                continue;
            }
            String removable = row.getCell(removableDataIndex).toString();
            if(removableDataSheet.contains(removable)) {
                System.out.println(row.getRowNum());
                indices.add(row.getRowNum());
            }
        }
        for (Integer index : indices) {
            removeFromSheet.removeRow(removeFromSheet.getRow(index));
        }
        return removeFromSheet;
    }
}
