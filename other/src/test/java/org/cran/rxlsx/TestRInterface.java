package org.cran.rxlsx;

import java.io.FileInputStream;
import java.io.IOException;

//import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class TestRInterface {
  @Test
  public void readXLSX() throws InvalidFormatException, IOException{
    FileInputStream in = new FileInputStream( TestRCellBlock.class.getResource("/test_import.xlsx").getFile() );
    Workbook wb = WorkbookFactory.create(in);
    Sheet sheet = wb.getSheetAt(7); 
    
    //String[] res = R.readRowStrings(sheet, 0, 10, 0);
//    String[] res = R.readColStrings(sheet, 1, 20, 1);
    double[] res = RInterface.readColDoubles(sheet, 1, 10, 4);
    //double[] res = R.readColDoubles(sheet, 0, 2006, 0);
    for (int i=0; i<res.length; i++){
      System.out.println(res[i]);
    }
  }
}
