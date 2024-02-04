import com.fasterxml.jackson.databind.node.ArrayNode;
import lombok.Data;
import java.util.ArrayList;
//Json data from Excel file
@Data
public class ParsedSheet {
    private String sheetName;
    private ArrayList<String> headerList;
    private ArrayNode sheetData;

    /**
     * Constructor of ParsedSheet
     *
     * @param sheetName The name of the sheet.
     * @param headerList List of all the headers in the given sheet.
     * @param sheetData A JSON list of the rows parsed as ObjectNode.
     */
    public ParsedSheet(String sheetName, ArrayList<String> headerList, ArrayNode sheetData) {
        this.sheetName = sheetName;
        this.headerList = headerList;
        this.sheetData = sheetData;
    }
    /**
     * Fetches the count of the number of data rows of the given sheet(excluding header).
     *
     * @return count of the data rows.
     */
    public int getPhysicalNumberOfRows(){
        if(sheetData.isEmpty()) return 0;
        return sheetData.size();
    }
    /**
     * Fetches the index of the last data row.
     *
     * @return if the sheet has no data then -1 is returned otherwise return sheetSize-1
     */
    public int getLastRowNum(){
        if(sheetData.isEmpty()) return -1;
        return sheetData.size()-1;
    }
}
