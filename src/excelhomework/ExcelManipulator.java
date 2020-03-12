/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelhomework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.HashMap;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bys99
 */
public class ExcelManipulator {

    public ExcelManipulator() {
    }
    HashMap<String, Double> MyExport = new HashMap<String, Double>();
    DescriptiveStatistics geometric = new DescriptiveStatistics();
    public void export() throws FileNotFoundException, IOException {
        Path file_path = FileSystems.getDefault().getPath("ДЗ2.xlsx");
        XSSFWorkbook MyBook = new XSSFWorkbook(new FileInputStream(file_path.toString()));
        XSSFSheet MySheet = MyBook.getSheetAt(0);
        
        int rowCount = MySheet.getPhysicalNumberOfRows();
        
        XSSFRow headers = MySheet.getRow(0);
        
        for (int i = 0; i < headers.getPhysicalNumberOfCells(); i++) {
            XSSFCell header = headers.getCell(i);
            
            String ColName = header.getStringCellValue();
            
            double[] values = new double[rowCount - 1];
            int k = 0;
            for (int j = 1; j < rowCount; j++) {    
                values[k] = MySheet.getRow(j).getCell(i).getNumericCellValue();
                geometric.addValue(Math.abs(values[k]));
                k++;
            }
            MyExport.put(ColName, geometric.getGeometricMean());
            geometric.clear();
        }  
        System.err.println("Импорт выполнен");
    }
    
    public void createNewBook() throws IOException {
        XSSFWorkbook MyBook = new XSSFWorkbook();
        XSSFSheet MySheet = MyBook.createSheet("First list");
        int i =0;
        XSSFRow MyRow = MySheet.createRow(0);
        XSSFRow MySecondRow = MySheet.createRow(1);
        for(HashMap.Entry<String, Double> item : MyExport.entrySet()){        
            XSSFCell MyFirstCell = MyRow.createCell(i);
            MyFirstCell.setCellValue(item.getKey());
            XSSFCell MySecondCell = MySecondRow.createCell(i);
            MySecondCell.setCellValue(item.getValue());
            i++;
        }
        
        Path file_path = FileSystems.getDefault().getPath("Homework.xlsx");
        FileOutputStream stream = new FileOutputStream(new File(file_path.toString()));
        MyBook.write(stream);
        MyBook.close();
        
        System.err.println("Загрузка завершена");
    }
}
