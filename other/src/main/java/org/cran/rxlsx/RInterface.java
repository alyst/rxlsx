package org.cran.rxlsx;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class RInterface {

    public static final double NA_DOUBLE = Double.NaN;
    public static final int NA_INT = -2147483648;
    public static final String NA_STRING = "NA";

    public static boolean isNA( double x )
    {
        return ( Double.isNaN(x) );
    }

    public static boolean isNA( int x )
    {
        return ( x == NA_INT );
    }

    public static boolean isNA( String x )
    {
        return ( x == null );
    }

   /*
    * Read a column of doubles from the sheet.  If the cell is not a Number show a NaN!
    */
   public static double[] readColDoubles(Sheet sheet, int startRowIndex, int endRowIndex, int colIndex) {
     
     int N = endRowIndex - startRowIndex + 1;
     double[] res = new double[N];
     for (int i=0; i<N; i++) {
       Row row = sheet.getRow(startRowIndex+i);
       Cell cell = row.getCell(colIndex, Row.CREATE_NULL_AS_BLANK);
 
       switch(cell.getCellType()) {
       case Cell.CELL_TYPE_STRING:
         try {
           res[i] = Double.valueOf(cell.getStringCellValue()).doubleValue();
         } catch (NumberFormatException nfe) {
           //System.out.println("NumberFormatException: " + nfe.getMessage());
           res[i] = NA_DOUBLE; 
         }
         break;
       case Cell.CELL_TYPE_NUMERIC:
         res[i] = cell.getNumericCellValue();
         break;
       case Cell.CELL_TYPE_FORMULA:
         try {
           res[i] = cell.getNumericCellValue();
         } catch (IllegalStateException e) {
           res[i] = NA_DOUBLE;  // cell.getCellFormula();
         }
         break;
       case Cell.CELL_TYPE_BOOLEAN:
         res[i] = (double) (cell.getBooleanCellValue()?1:0);
         break;
       case Cell.CELL_TYPE_ERROR:
         res[i] = NA_DOUBLE;
         break;
       default:
         res[i] = NA_DOUBLE;
       }
     }
     
     return res;
   }

   /*
    * Read a column of Strings from the sheet.  If the cell type is not a String, convert to String!
    */
   public static String[] readColStrings(Sheet sheet, int startRowIndex, int endRowIndex, int colIndex) {
     
     int N = endRowIndex - startRowIndex + 1;
     String[] res = new String[N];
     for (int i=0; i<N; i++) {
       Row row = sheet.getRow(startRowIndex+i);
       Cell cell = row.getCell(colIndex, Row.CREATE_NULL_AS_BLANK); 
       switch(cell.getCellType()) {
         case Cell.CELL_TYPE_STRING:
           res[i] = cell.getStringCellValue();
           break;
         case Cell.CELL_TYPE_NUMERIC:
           res[i] = Double.toString(cell.getNumericCellValue());
           break;
         case Cell.CELL_TYPE_FORMULA:
           try {
             res[i] = cell.getStringCellValue();
           } catch (IllegalStateException e) {
             res[i] = NA_STRING;  // cell.getCellFormula();
           }
           break;
         case Cell.CELL_TYPE_BOOLEAN:
           res[i] = new Boolean(cell.getBooleanCellValue()).toString();
           break;
         case Cell.CELL_TYPE_BLANK:
           res[i] = "";
           break;
         case Cell.CELL_TYPE_ERROR:
           res[i] = "ERROR";
           break;
         default:
           res[i] = "";
       }
     }
     
     return res;
   }
 
   
   /*
    * Read a row of Strings from the sheet
    */
   public static String[] readRowStrings(Sheet sheet, int startColIndex, int endColIndex, int rowIndex) {
     
     int N = endColIndex - startColIndex + 1;
     String[] res = new String[N];
     Row row = sheet.getRow(rowIndex);
     for (int i=0; i<N; i++) {
       Cell cell = row.getCell(startColIndex+i, Row.CREATE_NULL_AS_BLANK); 
       switch(cell.getCellType()) {
       case Cell.CELL_TYPE_STRING:
         res[i] = cell.getStringCellValue();
         break;
       case Cell.CELL_TYPE_NUMERIC:
         res[i] = Double.toString(cell.getNumericCellValue());
         break;
       case Cell.CELL_TYPE_FORMULA:
         res[i] = cell.getStringCellValue();
         break;
       case Cell.CELL_TYPE_BOOLEAN:
         res[i] = new Boolean(cell.getBooleanCellValue()).toString();
         break;
       default:
         res[i] = "";
       }
     }
     
     return res;
   }
}




/*
public String getCellValue(Cell cell, boolean keepFormulas){
  
 String out="";
 switch(cell.getCellType()) {
   case Cell.CELL_TYPE_STRING:
   out = "S\t" + cell.getRichStringCellValue().getString();
   break;
 case Cell.CELL_TYPE_NUMERIC:
   if(DateUtil.isCellDateFormatted(cell)) {
     out = "D\t" + String.valueOf(cell.getDateCellValue().getTime());
   } else {                
     out = "N\t" + String.valueOf(cell.getNumericCellValue());
   }
   break;
 case Cell.CELL_TYPE_BOOLEAN:
   out = "B\t" + String.valueOf(cell.getBooleanCellValue());
   break;
 case Cell.CELL_TYPE_FORMULA:
   if (keepFormulas){
     out = "S\t" + cell.getCellFormula();
   } else {
     try{ 
       out = "N\t"+String.valueOf(cell.getNumericCellValue());
     } catch (Exception ex){
       out = "S\t";
     }
   }
   break;
 default:
   out="S\t";
 }
  
 return out;
}
*/



/*
 * read Sheet.  It ignores missing cells.  So you need to pass the
 * cell coordinates.
 * To fix: Dates are shown as Java Dates.  
 */
/*public String readSheet(String filename, int sheetIndex, int nrows, 
   int skip, boolean keepFormulas) 
   throws IOException, InvalidFormatException {
     
   StringBuffer out = new StringBuffer();
     
   InputStream inp = new FileInputStream(filename);

   Workbook wb = WorkbookFactory.create(inp);
   Sheet sheet = wb.getSheetAt(sheetIndex);
   int lastRowIndex = sheet.getLastRowNum(); 
   if (nrows < 1)
     nrows = lastRowIndex - skip + 1;        // get them all
   
   for (Row row : sheet) {
     int rowNum = row.getRowNum();
     if (rowNum < skip)                  // skip first skip rows
       continue;
     if (rowNum >= skip + nrows)
       break;
               
     out.append(rowNum + "\t");    // write the row number first 
     for (Cell cell : row) { 
       out.append(cell.getColumnIndex()+"\t");  // write column index
       out.append(getCellValue(cell, keepFormulas));
       out.append("\t");  // elements separated by tabs
     }
     out.append("\n");    // rows separated by new lines
   }

   return out.toString();
 }*/

