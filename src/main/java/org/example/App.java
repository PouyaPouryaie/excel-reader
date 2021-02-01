package org.example;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        String message = "بیمه گذار محترم شرکت بیمه حکمت صبا\n" +
                "\n" +
                "به اطلاع میرساند شرکت بیمه حکمت جهت جلب رضایت و حفظ سلامتی شما بیمه گذار محترم به دلیل شیوع بیماری کرونا، صدور بیمه نامه بدنه را با ارائه تخفیفات ویژه تا سقف 70درصد و به صورت اقساط بلندمدت، بدون نیاز به مراجعه حضوری عرضه میدارد.\n" +
                "\n" +
                "لطفاً جهت صدور و یا تمدید بیمه نامه خود با شماره تلفن %s تماس حاصل نمائید";

        String fileLocation = "C:\\Users\\po.pouryaie\\Desktop\\98-update.xlsx";
        String fileLocation2 = "C:\\Users\\po.pouryaie\\Desktop\\97-update.xlsx";
//        String fileLocation2 = "C:\\Users\\po.pouryaie\\Desktop\\test.xlsx";
        List<String> fileLocations = new ArrayList<>();
        fileLocations.add(fileLocation);
        fileLocations.add(fileLocation2);


        System.out.println( "Hello World!" );
        String columnForMessageTo = "mobile";
        String columnForModify = "telefon";
        int indexMessageTo = 0;
        int indexModify = 0;
        ExcelReader excelReader = new ExcelReader();
        Map<Integer, List<String>> resultReader = excelReader.superReader(fileLocations, "telefon", "mobile");
        List<String> stringsColumns = resultReader.get(0);
        for(int i=0; i < stringsColumns.size(); i++){
            if(stringsColumns.get(i).toLowerCase().equals(columnForMessageTo)){
                indexMessageTo = i;
            }
            else if(stringsColumns.get(i).toLowerCase().equals(columnForModify)){
                indexModify = i;
            }
        }
        Map<Integer, List<String>> convertData = new HashMap<>();
        int finalIndexMessageTo = indexMessageTo;
        Iterator<Integer> iterator = resultReader.keySet().stream().filter(key -> !resultReader.get(key).get(finalIndexMessageTo).equals("-")).collect(Collectors.toSet()).iterator();
        while (iterator.hasNext()){
            Integer key = iterator.next();
            if(key != 0){
                resultReader.get(key).set(indexModify, String.format(message,
                        resultReader.get(key).get(indexModify).equals("-") ? "02141395100"
                                : resultReader.get(key).get(indexModify)));
                convertData.put(key, resultReader.get(key));
            }
        }
        System.out.println("done");
    }
}
