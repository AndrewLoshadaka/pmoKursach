package xamrah.pmo;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class Solve {
    private final XSSFWorkbook myBook = new XSSFWorkbook(Files.newInputStream(Paths.get("C:/Users/Andrew/Desktop/3_курс/пмо/хамрач_хуй_курсач.xlsx")));
    private final XSSFSheet sheet = myBook.getSheet("Лист1");
    private final List<Double> list1 = new ArrayList<>();
    private final List<Double> list2 = new ArrayList<>();
    private final List<Double> list3 = new ArrayList<>();

    private final BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
    private double detailX;
    private double detailY;
    private double detailZ;

    private final double[] possibility = new double[5];
    XSSFRow row;
    public Solve() throws IOException {
    }

    public void setValue() throws IOException{
        System.out.println("Введите количество деталей типа Х");
        String enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        detailX = Double.parseDouble(enter);

        System.out.println("Введите количество деталей типа Y");
        enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        detailY = Double.parseDouble(enter);

        System.out.println("Введите количество деталей типа Z");
        enter =  reader.readLine();
        while (!isDouble(enter)){
            System.out.println("enter correct value");
            enter =  reader.readLine();
        }
        detailZ = Double.parseDouble(enter);

        for(int i = 0; i < possibility.length; i++){
            System.out.println("Воможность производства " + (i + 1) + " предприятия");
            enter =  reader.readLine();
            while (!isDouble(enter)){
                System.out.println("enter correct value");
                enter =  reader.readLine();
            }
            possibility[i] = Double.parseDouble(enter);
        }

        for(int i = 0; i < 5; i++){
            System.out.println("Введите стоимость x" + (i+1));
            addOnList(list1, reader.readLine());
            System.out.println("Введите стоимость y" + (i+1));
            addOnList(list2, reader.readLine());
            System.out.println("Введите стоимость z" + (i+1));
            addOnList(list3, reader.readLine());
        }
    }

    public void setValueOnTable() throws IOException {
        setValue();

        row = sheet.getRow(1);
        for(int i = 0; i < possibility.length; i++){
            row.createCell(i + 8).setCellValue(possibility[i]);

        }

        row.createCell(5).setCellValue(detailX);
        row.createCell(6).setCellValue(detailY);
        row.createCell(7).setCellValue(detailZ);

        for(int i = 0; i < 5; i++){
            row = sheet.getRow(i + 1);
            row.createCell(1).setCellValue(list1.get(i));
            row.createCell(2).setCellValue(list2.get(i));
            row.createCell(3).setCellValue(list3.get(i));
        }

        String formula =
                "B2*C10+C2*D10+D2*E10+B3*C11+C3*D11+D3*E11+C12*B4+C4*D12+D4*E12+B5*C13+C5*D13+D5*E13+B6*C14+C6*D14+D6*E14";
        sheet.getRow(17).createCell(2).setCellFormula(formula);

        myBook.write(Files.newOutputStream(Paths.get("C:/Users/Andrew/Desktop/3_курс/пмо/хамрач_хуй_курсач.xlsx")));
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
