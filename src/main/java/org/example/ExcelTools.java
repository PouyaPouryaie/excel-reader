package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelTools {

    static ExcelTools excelTools = null;

    public ExcelTools() {
    }

    static ExcelTools getInstance(){
        if(excelTools == null){
            excelTools = new ExcelTools();
        }
        return excelTools;
    }

    public void reader(String fileLocation) {

        FileInputStream file;
        Workbook workbook = null;

        try{
            file = new FileInputStream(new File(fileLocation));
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
        System.out.println("Done");
    }


    public Map<Integer, List<String>> readerWithColumnSelect(List<String> fileLocations, String... columnNames) {

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
        System.out.println("Done");
        return data;
    }

    public boolean writer(Map<Integer, List<String>> data, String fileName){

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("data");
        XSSFRow header = sheet.createRow(0);

        List<String> strings = data.get(0);
        Cell headerCell;
        int index = 0;
        for(String name: strings){
            headerCell = header.createCell(index);
            headerCell.setCellValue(name);
            index++;
        }

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        Iterator<Integer> iterator = new HashSet<>(data.keySet()).iterator();
        Cell cell;
        int rowIndex = 1;
        while (iterator.hasNext()){
            Row row = sheet.createRow(rowIndex);
            index = 0;
            Integer key = iterator.next();
            if(key != 0){
                List<String> stringsData = data.get(key);
                for(String s: stringsData){
                    cell = row.createCell(index);
                    cell.setCellValue(s);
                    cell.setCellStyle(style);
                    index++;
                }
                rowIndex++;
            }
        }

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(fileLocation);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return true;
    }

    public boolean removeCurrentFile(String fileName){
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";
        try {
            Files.deleteIfExists(Paths.get(fileLocation));
        } catch (IOException e) {
            System.out.println(e.getMessage());
            return false;
        }

        return true;
    }



}
