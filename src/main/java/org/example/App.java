package org.example;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.stream.Collectors;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        ExecutorService executorService = Executors.newFixedThreadPool(1);
        ExcelTools instance = ExcelTools.getInstance();

        String MESSAGE_TEMPLATE = "بیمه گذار محترم شرکت بیمه حکمت صبا\n" +
                "\n" +
                "به اطلاع میرساند شرکت بیمه حکمت جهت جلب رضایت و حفظ سلامتی شما بیمه گذار محترم به دلیل شیوع بیماری کرونا، صدور بیمه نامه بدنه را با ارائه تخفیفات ویژه تا سقف 70درصد و به صورت اقساط بلندمدت، بدون نیاز به مراجعه حضوری عرضه میدارد.\n" +
                "\n" +
                "لطفاً جهت صدور و یا تمدید بیمه نامه خود با شماره تلفن %s تماس حاصل نمائید";

        String fileLocation = "C:\\Users\\po.pouryaie\\Desktop\\98-update.xlsx";
        String fileLocation2 = "C:\\Users\\po.pouryaie\\Desktop\\97-update.xlsx";
        String fileLocation3 = "C:\\Users\\po.pouryaie\\Desktop\\96-update.xlsx";
//        String fileLocation2 = "C:\\Users\\po.pouryaie\\Desktop\\sampleTest.xlsx";
        List<String> fileLocations = new ArrayList<>();
        fileLocations.add(fileLocation);
        fileLocations.add(fileLocation2);
        fileLocations.add(fileLocation3);


        String columnForMessageTo = "mobile";
        String columnForPhone = "telefon";
        String columnForMessage = "";
        int indexMessageTo = 0;
        int indexPhone = 0;
        int indexMessage = 2;
//        ExcelTools excelTools = new ExcelTools();
        Map<Integer, List<String>> resultReader = null;
        try {
            resultReader = executorService.submit(() -> instance.readerWithColumnSelect(fileLocations, "telefon", "mobile")).get();
        } catch (InterruptedException e) {
            e.printStackTrace();
        } catch (ExecutionException e) {
            e.printStackTrace();
        }
//        Map<Integer, List<String>> resultReader = ExcelTools.readerWithColumnSelect(fileLocations, "telefon", "mobile");
        List<String> stringsColumns = resultReader.get(0);
        for(int i=0; i < stringsColumns.size(); i++){
            if(stringsColumns.get(i).toLowerCase().equals(columnForMessageTo)){
                indexMessageTo = i;
            }
            else if(stringsColumns.get(i).toLowerCase().equals(columnForPhone)){
                indexPhone = i;
            }
        }
        Map<Integer, List<String>> convertData = new HashMap<>();
        List columnNames = new ArrayList();
        columnNames.add("mobile");
        columnNames.add("phone");
        columnNames.add("message");
        int finalIndexMessageTo = indexMessageTo;
        Map<Integer, List<String>> finalResultReader = resultReader;
        Iterator<Integer> iterator = resultReader.keySet().stream().filter(key -> !finalResultReader.get(key).get(finalIndexMessageTo).equals("-")).collect(Collectors.toSet()).iterator();
        while (iterator.hasNext()){
            Integer key = iterator.next();
            if(key != 0){
                List list = new ArrayList();
                list = resultReader.get(key);
                list.add(String.format(MESSAGE_TEMPLATE,
                        resultReader.get(key).get(indexPhone).equals("-") ? "02141395100"
                                : resultReader.get(key).get(indexPhone)));
                convertData.put(key, list);
            }
            else{
                convertData.put(0, columnNames);
            }
        }

        executorService.submit(() -> instance.writer(convertData, "sampleTest3"));

        try {
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        System.out.println("after sleep");

        executorService.submit(() -> instance.removeCurrentFile("sampleTest3"));

        try {
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        executorService.shutdown();

        System.out.println("Done Project");
    }
}
