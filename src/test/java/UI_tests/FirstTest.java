package UI_tests;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class FirstTest extends BaseTest {

    public static void readFromExcel(String file) throws IOException {
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheet("Проверка");
        HSSFRow row = myExcelSheet.getRow(0);

        if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
            String name = row.getCell(0).getStringCellValue();
            System.out.println("name : " + name);
        }

        if(row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
            Date birthdate = row.getCell(1).getDateCellValue();
            System.out.println("birthdate :" + birthdate);
        }

        myExcelBook.close();

    }

    @Test()
    public void checkPage() {

        ArrayList<String> list = new ArrayList<>();

        try {
            FileInputStream fis = new FileInputStream("D:\\DigitalGovernmentPractice\\gospabl.xlsx");
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            for(Row row : sheet) {
                list.add(row.getCell(1).getStringCellValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        int ListSize = list.size();

        int flag = 0;
        ArrayList<String> exGoverment = new ArrayList<>();
        ArrayList<String> exMap = new ArrayList<>();
        ArrayList<String> exAdress = new ArrayList<>();

        String linka;

        for (int i = 1; i<ListSize; i++) {
            try {
                driver.get(list.get(i));
                driver.manage().timeouts().pageLoadTimeout(30000, TimeUnit.MILLISECONDS);
            }
            catch (Exception e) {
                int stroka = i+1;
                System.out.println("ссылка в эксель неккоректа/её нет: " + stroka + " строка");
                flag=1;
            }

            if (flag != 1) {
                try {
                    driver.findElement(By.className("GovernmentCommunityBadge--tooltip"));
                } catch (Exception e) {
                    System.out.println("не нашло надпись Госорганизация:" + list.get(i));
                    exGoverment.add(list.get(i));
                }
                try {
                    driver.findElement(By.className("ymaps-2-1-79-events-pane"));
                } catch (Exception e) {
                    System.out.println("не нашло карту:" + list.get(i));
                    exMap.add(list.get(i));
                }
                try {
                    driver.findElement(By.className("address_text"));
                } catch (Exception e) {
                    try {
                        driver.findElement(By.className("group_info_row address"));
                    } catch (Exception e1) {
                        try {
                            driver.findElement(By.className("address_link"));
                        } catch (Exception e2) {
                            System.out.println("не нашло адрес:" + list.get(i));
                            exAdress.add(list.get(i));
                        }
                    }
                }
            }
            flag=0;
        }

        int exGovermentSize = exGoverment.size();
        int exMapSize = exMap.size();
        int exAdressSize = exAdress.size();

        /*try {
            System.out.println("///////////////////////////////");
            for (int i = 0; i<exAdressSize; i++) {
                System.out.println(exAdress.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Список пуст");
        }

        try {
            System.out.println("///////////////////////////////");
            for (int i = 0; i<exMapSize; i++) {
                System.out.println(exMap.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Список пуст");
        }

        try {
            System.out.println("///////////////////////////////");
            for (int i = 0; i<exGovermentSize; i++) {
                System.out.println(exGoverment.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Список пуст");
        }*/

        Workbook workbook = new XSSFWorkbook();
        Sheet SheetMap = workbook.createSheet("Нет карт");
        Sheet SheetAdress = workbook.createSheet("Нет адреса");
        Sheet SheetGover = workbook.createSheet("Нет надписи");

        try {
            for (int i = 0; i < exMapSize; i++) {
                SheetMap.createRow(i).createCell(0).setCellValue(exMap.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Что-то пошло не так");
        }

        try {
            for (int i = 0; i < exAdressSize; i++) {
                SheetAdress.createRow(i).createCell(0).setCellValue(exAdress.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Что-то пошло не так");
        }

        try {
            for (int i = 0; i < exGovermentSize; i++) {
                SheetGover.createRow(i).createCell(0).setCellValue(exGoverment.get(i));
            }
        }
        catch (Exception e) {
            System.out.println("Что-то пошло не так");
        }

        try {
            FileOutputStream fileOut = new FileOutputStream("D:\\DigitalGovernmentPractice\\endFile\\end2.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Файл создан");
        }
        catch (Exception e) {
            System.out.println("Что-то пошло не так");
        }

    }
}
