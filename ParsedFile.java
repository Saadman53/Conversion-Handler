import lombok.Data;
import java.util.ArrayList;

@Data
public class ParsedFile {
    private String fileName;
    private String fileExtension;
    private ArrayList<ParsedSheet> parsedSheets;
    /**
     * Constructor of ParsedFile
     *
     * @param fileName The name of the file.
     * @param fileExtension The extension of the file without the dot(.).
     * @param parsedSheets List of all the sheets parsed from the xls/xlsx file.
     */
    public ParsedFile(String fileName, String fileExtension, ArrayList<ParsedSheet> parsedSheets) {
        this.fileExtension = fileExtension;
        this.fileName = fileName;
        this.parsedSheets = parsedSheets;
    }
    /**
     * Fetches the sheet from the parsedSheets that is wanted by the user.
     *
     * @param sheetName The name of the desired sheet to be fetched.
     * @return The desired sheet to be fetched.
     */
    public ParsedSheet getSelectedSheet(String sheetName){
        ParsedSheet selectedParsedSheet = null;
        for(ParsedSheet parsedSheet: parsedSheets){
            if(parsedSheet.getSheetName().equals(sheetName)){
                selectedParsedSheet =  parsedSheet;
                break;
            }
        }
        return selectedParsedSheet;
    }
}
