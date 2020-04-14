package java_projects;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
XSSFSheet sh;
public Excel() throws IOException
{
FileInputStream f=new FileInputStream("C:\\Users\\asnab\\OneDrive\\Desktop\\myfile.xlsx");
XSSFWorkbook w=new XSSFWorkbook(f);
sh=w.getSheet("sheet1");
}
public double readData(int i,int j) {
Row r=sh.getRow(i);
Cell c=r.getCell(j);
double d=c.getNumericCellValue();
return d;
	
}}
