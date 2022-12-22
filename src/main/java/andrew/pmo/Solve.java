package andrew.pmo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Solve {
    private final XSSFWorkbook myBook = new XSSFWorkbook(Files.newInputStream(Paths.get("C:/Users/Andrew/Desktop/3_курс/пмо/курсач.xlsx")));
    private final XSSFSheet sheet = myBook.getSheet("Лист1");
    private final List<Double> list1 = new ArrayList<>();
    private final List<Double> list2 = new ArrayList<>();

    private final BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
    private double extraction;
    private double export;
    private double costs;
    XSSFRow row;
    public Solve() throws IOException {
    }

    public void setValue() throws IOException{
        System.out.println("Введите сколько добыто");
        String enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        extraction = Double.parseDouble(enter);

        System.out.println("Введите долю на экспорт");
        enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        export =
                Double.parseDouble(enter);


        System.out.println("Введите суммарные издержки");
        enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        costs =
                Double.parseDouble(enter);
        System.out.println("Введите цену");
        System.out.println("Новоросийск");
        String temp = reader.readLine();
        addOnList(list1, temp);

        System.out.println("Туапсе");
        addOnList(list1, reader.readLine());
        System.out.println("Вентспилс");
        addOnList(list1, reader.readLine());
        System.out.println("Приморск");
        addOnList(list1, reader.readLine());
        System.out.println("Одесса");
        addOnList(list1, reader.readLine());
        System.out.println("Чехия");
        addOnList(list1, reader.readLine());
        System.out.println("Словакия");
        addOnList(list1, reader.readLine());
        System.out.println("Венгрия");
        addOnList(list1, reader.readLine());
        System.out.println("Германия");
        addOnList(list1, reader.readLine());
        System.out.println("Польша");
        addOnList(list1, reader.readLine());
        System.out.println("Введите издержки");
        System.out.println("Новоросийск");
        addOnList(list2, reader.readLine());
        System.out.println("Туапсе");
        addOnList(list2, reader.readLine());
        System.out.println("Вентспилс");
        addOnList(list2, reader.readLine());
        System.out.println("Приморск");
        addOnList(list2, reader.readLine());
        System.out.println("Одесса");
        addOnList(list2, reader.readLine());
        System.out.println("Чехия");
        addOnList(list2, reader.readLine());
        System.out.println("Словакия");
        addOnList(list2, reader.readLine());
        System.out.println("Венгрия");
        addOnList(list2, reader.readLine());
        System.out.println("Германия");
        addOnList(list2, reader.readLine());
        System.out.println("Польша");
        addOnList(list2, reader.readLine());
    }

    public void setValueOnTable() throws IOException {
        setValue();
        for(int i = 0; i < 10; i++){
            row = sheet.getRow(i + 1);
            row.createCell(1).setCellValue(list1.get(i));
            row.createCell(2).setCellValue(list2.get(i));
        }

        row = sheet.getRow(1);
        row.createCell(9).setCellValue(extraction);
        row.createCell(11).setCellValue(export);
        row.createCell(13).setCellValue(costs);

        String formula =
                "((B2-C2)*F2+(B3-C3)*F3+(B4-C4)*F4+(B5-C5)*F5+(B6-C6)*F6+(B7-C7)*F7+(B8-C8)*F8+(B9-C9)*F9+(B10-C10)*F10+(B11-C11)*F11) - N2*L2*J2";
        sheet.getRow(19).createCell(1).setCellFormula(formula);

        myBook.write(Files.newOutputStream(Paths.get("C:/Users/Andrew/Desktop/3_курс/пмо/курсач.xlsx")));
        myBook.close();
    }

    public boolean isDouble(String value){
        try{
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException ignored) {
        }
        return false;
    }

    public void addOnList(List<Double> list, String value) throws IOException{
        while (!isDouble(value)) {
            System.out.println("Повторите!");
            value = reader.readLine();
        }

        if (Double.parseDouble(value) > 0)
            list.add(Double.parseDouble(value));
        else
            System.out.println("Повторите!");
    }
}
