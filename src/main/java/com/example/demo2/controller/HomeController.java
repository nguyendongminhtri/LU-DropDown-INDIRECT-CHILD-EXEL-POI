package com.example.demo2.controller;
import jakarta.servlet.ServletRequest;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;


@RestController
public class HomeController {
    @GetMapping("/hello")
    public String hello(ServletRequest servletRequest) throws IOException {
        try {

            XSSFWorkbook workbook;
            File file = new File("D:\\Project\\TEST\\File Import\\output3.xlsx");
            FileOutputStream fos = new FileOutputStream(file);
            workbook = new XSSFWorkbook();

            String sname = "TestSheet";
            String parentName = "PARENT";
            String childName1 = "GUJARAT";
            String childName2 = "KARNATAKA";
            String childName3 = "MAHARASHTRA";
            XSSFSheet sheet = workbook.createSheet(sname);

            Row row = null;
            Cell cell = null;
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("PARENT");

//            row = sheet.createRow(0);
            cell = row.createCell(1);
            cell.setCellValue("CHILD");

            row = sheet.createRow(9);
            cell = row.createCell(0);
            cell.setCellValue("Gujarat");
            cell = row.createCell(1);
            cell.setCellValue("Karnataka");
            cell = row.createCell(2);
            cell.setCellValue("Maharashtra");

            row = sheet.createRow(10);
            cell = row.createCell(0);
            cell.setCellValue("Ahmedabad");
            cell = row.createCell(1);
            cell.setCellValue("Rajkot");
            cell = row.createCell(2);
            cell.setCellValue("Gandhinagar");
            cell = row.createCell(3);
            cell.setCellValue("Surat");
            cell = row.createCell(4);
            cell.setCellValue("Vapi");

            row = sheet.createRow(11);
            cell = row.createCell(0);
            cell.setCellValue("Bangalore");
            cell = row.createCell(1);
            cell.setCellValue("Hasan");
            cell = row.createCell(2);
            cell.setCellValue("Mysore");
            cell = row.createCell(3);
            cell.setCellValue("Mangalore");

            row = sheet.createRow(12);
            cell = row.createCell(0);
            cell.setCellValue("Mumbai");
            cell = row.createCell(1);
            cell.setCellValue("Pune");
            cell = row.createCell(2);
            cell.setCellValue("Aurangabad");


            // 1. create named range for a single cell using areareference
            Name namedCell1 = sheet.getWorkbook().createName();
            namedCell1.setNameName(parentName);
            String reference1 = sname+"!$A$10:$C$10"; // area reference
            namedCell1.setRefersToFormula(reference1);

            Name namedCell2 = sheet.getWorkbook().createName();
            namedCell2.setNameName(childName1);
            String reference2 = sname+"!$A$11:$E$11"; // area reference
            namedCell2.setRefersToFormula(reference2);

            Name namedCell3 = sheet.getWorkbook().createName();
            namedCell3.setNameName(childName2);
            String reference3 = sname+"!$A$12:$D$12"; // area reference
            namedCell3.setRefersToFormula(reference3);

            Name namedCell4 = sheet.getWorkbook().createName();
            namedCell4.setNameName(childName3);
            String reference4 = sname+"!$A$13:$C$13"; // area reference
            namedCell4.setRefersToFormula(reference4);

            DataValidationHelper helper = null;
            DataValidationConstraint constraint = null;
            DataValidation validation = null;

            helper = sheet.getDataValidationHelper();
            constraint = helper.createFormulaListConstraint(parentName);
            validation = helper.createValidation(constraint, new CellRangeAddressList(1,1,0,0));
            sheet.addValidationData(validation);

            constraint = helper.createFormulaListConstraint("INDIRECT(UPPER($A$2))");
            validation = helper.createValidation(constraint, new CellRangeAddressList(1,1,1,1));
            sheet.addValidationData(validation);

            workbook.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
          e.getStackTrace();
        }
        return "hello World";
    }
}
