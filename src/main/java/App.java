import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class App {

    public static void main(String[] args) {

        // Criando o arquivo e uma planilha chamada "Calculadora"
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Calculadora");

        // Definindo alguns padroes de layout
        sheet.setDefaultColumnWidth(15);
        sheet.setDefaultRowHeight((short)400);

        //Carregando os produtos
//        List products = getProducts();

        int rownum = 0;
        int cellnum = 0;
        Cell cell;
        Row row;

        row = sheet.createRow(rownum++);
        cellnum = 0;
        cell = row.createCell(cellnum++);
        cell.setCellValue(1);

        cell = row.createCell(cellnum++);
        cell.setCellValue(1);

        cell = row.createCell(cellnum++);
        cell.setCellValue(1);

        cell = row.createCell(cellnum++);
        cell.setCellFormula("SUM(A1:C3)");
//        XSSFFormulaEvaluator
        HSSFFormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        formulaEvaluator.evaluateFormulaCell(cell);

        try {
            //Escrevendo o arquivo em disco
            FileOutputStream out = new FileOutputStream("arquivo.xlsx");
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Success!!");

            // Criando o arquivo e uma planilha chamada "Calculadora"
            FileInputStream file = new FileInputStream(new File("arquivo.xlsx"));
            workbook = new HSSFWorkbook(file);
            sheet = workbook.getSheet("Calculadora");
            System.out.println(sheet.getRow(0).getCell(3));


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
