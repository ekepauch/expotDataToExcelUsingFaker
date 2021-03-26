package com.exportexcel.exportexcel;

import com.fasterxml.jackson.databind.node.ArrayNode;
import com.github.javafaker.Faker;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;



@SpringBootApplication
public class ExportExcelApplication {


    public static void main(String[] args) throws IOException {
        SpringApplication.run(ExportExcelApplication.class, args);
//This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();

        Faker faker = new Faker();
        int totalcount = 250000;
        int i = 0;
        while (i <= totalcount) {

            String payee = faker.number().numberBetween(1, 9) + "0" + faker.number().digits(4);
            String firstName = faker.address().firstName();
            String email = faker.internet().emailAddress();
            String phone = faker.number().numberBetween(7, 9) + "0" + faker.number().digits(8);
            String scheme = faker.code().asin();
            String schemeCode = faker.number().numberBetween(1, 9) + "0" + faker.number().digits(2);
            String account = faker.number().numberBetween(0, 9) + "0" + faker.number().digits(8);
            String amount = faker.commerce().price();


            data.put(String.valueOf(i), new Object[]{payee, firstName, email, phone, scheme, schemeCode, account, amount});
            i++;
        }


        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("PAYMENT Data");


        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);

            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("PAYMENT.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("COMPLETED");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


}
