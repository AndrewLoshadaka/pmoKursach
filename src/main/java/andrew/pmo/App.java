package andrew.pmo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

/**
 * Hello world!
 *
 */
public class App {
    public static void main( String[] args ) throws IOException {
        Solve solve = new Solve();
        solve.setValueOnTable();
        App.readFromExcel("C:/Users/Andrew/Desktop/3_курс/пмо/курсач.xlsx");
        System.out.println( "Hello World!" );
    }

    public static void readFromExcel(String file) throws IOException{
        XSSFWorkbook myBook = new XSSFWorkbook(Files.newInputStream(Paths.get(file)));
        XSSFSheet sheet = myBook.getSheet("Лист1");
        XSSFRow row = sheet.getRow(0);

        if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
            String name = row.getCell(0).getStringCellValue();
            System.out.println("name : " + name);
        }

        myBook.close();
    }
}
