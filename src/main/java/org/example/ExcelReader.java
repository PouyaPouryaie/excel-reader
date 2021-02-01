package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class ExcelReader {
    private static String FILE_LOCATION = "C:\\Users\\po.pouryaie\\Desktop\\98-update.xlsx";

    static FileInputStream file;

    static Workbook workbook;

    public static void reader() {
        try{
            file = new FileInputStream(new File(FILE_LOCATION));
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<>());
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell == null || cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType._NONE){
                    data.get(i).add(" ");
                }
                else {
                    switch (cell.getCellType()) {
                        case STRING:
                            data.get(i).add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                data.get(i).add(cell.getDateCellValue() + "");
                            } else {
                                data.get(i).add(cell.getNumericCellValue() + "");
                            }
                            break;
                        case BOOLEAN:
                            data.get(i).add(cell.getBooleanCellValue() + "");
                            break;
                        case FORMULA:
                            data.get(i).add(cell.getCellFormula() + "");
                            break;
                        default: data.get(new Integer(i)).add(" ");
                    }
                }
            }
            i++;
        }
        System.out.printf("Done");
    }


    public Map<Integer, List<String>> superReader(List<String> fileLocations, String... columnNames) {

        Map<Integer, List<String>> data = new HashMap<>();
        Map<String, Integer> indexColumn = new HashMap<>();
        FileInputStream file;
        Workbook workbook = null;
        int i = 0;
        for(String fileLocation: fileLocations){
            try{
                file = new FileInputStream(new File(fileLocation));
                workbook = new XSSFWorkbook(file);
            } catch (IOException e) {
                e.printStackTrace();
            }

            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);
            int columnSize = firstRow.getLastCellNum();
            for (Row row : sheet) {
                data.put(i, new ArrayList<>());
                if(i == 0){
                    for (int j=0; j < columnSize ; j++){
                        String[] clone = columnNames.clone();
                        Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        if(Arrays.stream(clone).anyMatch(s -> s.equals(cell.getStringCellValue().toLowerCase()))){
                            System.out.println(String.format("index %s is : %s", cell.getStringCellValue().toLowerCase(), j));
                            indexColumn.put(cell.getStringCellValue().toLowerCase(), j);
                        }
                    }
                }
                for(int j=0; j < columnSize ; j++){
                    int finalJ = j;
                    if(indexColumn.values().stream().anyMatch(s -> s == finalJ)){
                        Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // find null column and fill it
                        if(cell.getCellType() == CellType.BLANK){
                            data.get(i).add("-");
                        }
                        else{
                            data.get(i).add(cell.getStringCellValue());
                        }
                    }
                }
                i++;
            }
        }
        System.out.printf("Done");
        return data;
    }



}
