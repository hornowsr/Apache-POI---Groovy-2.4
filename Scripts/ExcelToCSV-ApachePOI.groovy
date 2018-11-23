import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


//Variables iniciation
DataFormatter formatter = null;

Workbook wb = null;

List sheetList = null;

StringBuilder sb = null;

Sheet sheet = null;

//End
// Read input into variables
inp = new FileInputStream("wb.xls");
reader = new BufferedReader(inp);
//End

//Upload a work book to work with
formatter = new DataFormatter(Locale.default);
wb = WorkbookFactory.create(inp);
sheetList = wb.sheets;
sb = new StringBuilder();
//END

//Read data into StringBuilder variable depending on cell type
for (int j = 0; j < sheetList.size(); j++) {
  sheet = wb.getSheetAt(j);
  for (Row row: sheet) {
   for (Cell cell: row) {
    switch (cell.getCellType()) {
     case Cell.CELL_TYPE_NUMERIC:
      if (DateUtil.isCellDateFormatted(cell)) {
       sb.append("" + formatter.formatCellValue(cell) + "");
      } else {
       sb.append("" + cell.getNumericCellValue() + "");
      }
      break;
     case Cell.CELL_TYPE_STRING:
      sb.append("" + cell.getStringCellValue() + "");
      break;
     case Cell.CELL_TYPE_FORMULA:
      sb.append("" + cell.getCellFormula() + "");
      break;
     case Cell.CELL_TYPE_BOOLEAN:
      sb.append("" + cell.getBooleanCellValue() + "");
      break;
     case Cell.CELL_TYPE_BLANK:
      sb.append("");
      break;
     default:
      sb.append("");
      break;
    }
    sb.append(",");
   }
   if (sb.length() > 0) {
    sb.setLength(sb.length() - 1);
   }
   sb.append("\r\n");
  }

 //Output
wb.write(outputStream);
outputStream.close();
//End