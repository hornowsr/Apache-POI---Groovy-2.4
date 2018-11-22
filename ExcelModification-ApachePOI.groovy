import java.io.*;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.lang.String;
import java.util.Locale;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Variables iniciation
DataValidation dataValidation = null;
DataValidationConstraint constraint = null;
DataValidationHelper validationHelper = null;
DataValidation dataValidation_2 = null;
DataValidationConstraint constraint_2 = null;
DataValidationHelper validationHelper_2 = null;

XSSFRow row = null;

int rownum = 0;
int rownum_2 = 0;
int cellnum = 0;

Cell cell = null;

String[] fields = null;

// Create new excel file with one sheet
XSSFWorkbook wb = new XSSFWorkbook();
XSSFSheet sheet = wb.createSheet("Sheet1");
XSSFSheet sheet_2 = wb.createSheet("Sheet2");
BufferedReader reader_2 = null;

FileOutputStream fileOut = new FileOutputStream("wb.xls");

BufferedReader reader = null;

InputStream inp = null;

CellStyle styleHeader = wb.createCellStyle();
CellStyle styleUnlockedCell = wb.createCellStyle();
CellStyle styleLockedCell = wb.createCellStyle();

CellRangeAddress[] regions = null;
CellRangeAddressList addressList = null;

SheetConditionalFormatting sheetCF = null;

ConditionalFormattingRule rule_1 = null;
ConditionalFormattingRule rule_2 = null;

PatternFormatting fill2 = null;
PatternFormatting fill_1 = null;
//End


// Read input into variables
inp = new FileInputStream("wb.xls");
reader = new BufferedReader(inp);
//End

//Set protection options and lock sheet
sheet.lockDeleteColumns(true);
// Restrict deleting rows
sheet.lockDeleteRows(true);
// Restrict formatting cells
 sheet.lockFormatCells(true);
// Restrict formatting columns
sheet.lockFormatColumns(false);
// Restrict formatting rows
sheet.lockFormatRows(false);
// Restrict inserting columns
sheet.lockInsertColumns(true);
// Restrict inserting rows
sheet.lockInsertRows(true);
// Restrinct auto filter
sheet.lockAutoFilter(false);
// Lock the sheet NOTE: Without this option previous locking will NOT work
sheet.enableLocking();
// Set width for the column
Set the width (in units of 1/256th of a character width)
//End

//Set cell styles
styleHeader.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
styleHeader.setBorderBottom(BorderStyle.MEDIUM);
styleHeader.setBorderLeft(BorderStyle.MEDIUM);
styleHeader.setBorderRight(BorderStyle.MEDIUM);
styleHeader.setBorderTop(BorderStyle.MEDIUM);
styleHeader.setLocked(true);

styleUnlockedCell.setBorderBottom(BorderStyle.MEDIUM);
styleUnlockedCell.setBorderLeft(BorderStyle.MEDIUM);
styleUnlockedCell.setBorderRight(BorderStyle.MEDIUM);
styleUnlockedCell.setBorderTop(BorderStyle.MEDIUM);
styleUnlockedCell.setLocked(false);

styleLockedCell.setBorderBottom(BorderStyle.MEDIUM);
styleLockedCell.setBorderLeft(BorderStyle.MEDIUM);
styleLockedCell.setBorderRight(BorderStyle.MEDIUM);
styleLockedCell.setBorderTop(BorderStyle.MEDIUM);
styleLockedCell.setLocked(true);
//End

//Fill first sheet with data
while ((line = reader.readLine()) != null ) {
    
    row = sheet.createRow(rownum);
    rownum++;
    cellnum = 0;
    fields = line.split(",");
    for (int j = 0; j < fields.length; j++) {
        cell = row.createCell(cellnum);
        switch(j){
            case 0:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[0]);
                break;
            case 1:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[1]);
                break;
            case 2:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[2]);
                break;
            case 3:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[3]);
                break;
            case 4:
                if(fields[j]==""){
                    cell.setCellValue(-1);
                }else{
                    cell.setCellValue(fields[4]);
                }
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleUnlockedCell);
                    }
                break;
            case 5:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[5]);
                break;
            case 6:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleUnlockedCell);
                    }
                cell.setCellValue(fields[6]);
                break;
            case 7:
                if(rownum == 1){
                    cell.setCellStyle(styleHeader);
                }
                else{
                        cell.setCellStyle(styleLockedCell);
                    }
                cell.setCellValue(fields[7]);
                break;
                
        }
        cellnum++;
    }
  
    
}
//End

//
for(int z = i ; i < rownum+1 ; i++){
    regions = CellRangeAddress.valueOf("A"+i+":G"+i);
    sheetCF = sheet.getSheetConditionalFormatting();
    rule_1 = sheetCF.createConditionalFormattingRule("\$E"+i+"=-1");
    fill_1 = rule_1.createPatternFormatting();
    fill_1.setFillBackgroundColor(IndexedColors.RED.index);
    fill_1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
    
    rule_2 = sheetCF.createConditionalFormattingRule("\$E"+z+"<>-1");
    fill_2 = rule_2.createPatternFormatting();
    fill_2.setFillBackgroundColor(IndexedColors.WHITE.index);
    fill_2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
    
    sheetCF.addConditionalFormatting(regions, rule_1, rule_2);
}
//End

//Autosize collumns to fit all text inside in first sheet
for(int j = 0; j < cellnum ; j++){
    sheet.autoSizeColumn(j);
}
//End

//Set Auto Filters on
sheet.setAutoFilter(CellRangeAddress.valueOf("A1:H"+rownum));
//End

//Create second sheet with data
reader_2 = new BufferedReader(new InputStreamReader(is2));
rownum_2 = 0;
//End

//Fill second sheet with data
while ((line = reader_2.readLine()) != null){
  row = sheet_2.createRow(rownum_2);
  rownum_2++;
  cellnum = 0;
  fields = line.split(",");
  for (int j = 0; j < fields.length; j++) {
   row.createCell(cellnum).setCellValue(fields[j] as String);
   cellnum++;
}
//End
  
//Autosize collumns to fit all text inside in second sheet
}
for(int j = 0; j < cellnum ; j++){
    sheet_2.autoSizeColumn(j);
}
//End


//Set drop down list from cell range
validationHelper=new XSSFDataValidationHelper(sheet);
addressList = new  CellRangeAddressList(1,rownum-1,4,4);
constraint =validationHelper.createFormulaListConstraint("'Sheet2'!\$A\$2:\$A\$"+(rownum_2));
dataValidation = validationHelper.createValidation(constraint, addressList);
dataValidation.setSuppressDropDownArrow(true);      
sheet.addValidationData(dataValidation);
//End

//Set second drop down list from staic data
def testArray = ["CREATE/UPDATE", "DELETE"] as String[];

CellRangeAddressList addressList_2 = new  CellRangeAddressList(1,rownum-1,6,6);
constraint_2 =validationHelper.createExplicitListConstraint(testArray);
dataValidation_2 = validationHelper.createValidation(constraint_2, addressList_2);
sheet.addValidationData(dataValidation_2);
//End

//Hide second sheet
wb.setSheetHidden(1, true);
//End

//Output
wb.write(outputStream);
outputStream.close();
//End



