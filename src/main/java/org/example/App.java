package org.example;
import java.util.*;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        Pattern mobilePattern = Pattern.compile("\\d");
        ExecutorService executorService = Executors.newFixedThreadPool(1);
        ExcelTools instance = ExcelTools.getInstance();
        String MESSAGE_TEMPLATE = "بیمه گذار محترم" +
                "\n" +
                "به اطلاع میرساند شرکت بیمه حکمت صبا به دلیل شیوع بیماری کرونا و  به جهت حفظ سلامتی شما بیمه گذار محترم، صدور بیمه نامه بدنه را با ارائه تخفیفات ویژه تا سقف 70درصد و به صورت اقساط بلند مدت بدون نیاز به مراجعه حضوری عرضه میدارد.\n" +
                "\n" +
                "لطفاً در صورت تمایل به صدور بیمه نامه با شماره %s تماس حاصل نمائید" +
                "\n" +
                "با آرزوی سلامتی";
        String fileLocation96 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\96-telephone.xlsx";
        String fileLocation97 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\97-telephone.xlsx";
        String fileLocation98 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\98-telephone.xlsx";
        String fileLocation99 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\99-telephone.xlsx";
        String fileLocation982 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\98.xlsx";
        String fileLocation992 = "C:\\Users\\j.mosavian.HT\\Desktop\\sms_badaneh\\99.xlsx";
//        String x99 = "C:\\Users\\j.mosavian.HT\\Desktop\\sampleTest3.xlsx";
//        String fileLocation2 = "C:\\Users\\po.pouryaie\\Desktop\\sampleTest.xlsx";
        List<String> fileLocations = new ArrayList<>();
        fileLocations.add(fileLocation96);
        fileLocations.add(fileLocation97);
        fileLocations.add(fileLocation98);
        fileLocations.add(fileLocation99);
        fileLocations.add(fileLocation982);
        fileLocations.add(fileLocation992);
        String columnForMessageTo = "customer_mobile";
        String columnForPhone = "telephone";
        String columnForMessage = "";
        int indexMessageTo = 0;
        int indexPhone = 0;
        int indexMessage = 2;
//        ExcelTools excelTools = new ExcelTools();
        Map<Integer, List<String>> resultReader = null;
        try {
            resultReader = executorService.submit(() -> instance.readerWithColumnSelect(fileLocations, "telephone", "customer_mobile")).get();
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
        AtomicBoolean isHeader = new AtomicBoolean(true);
        Iterator<Integer> iterator = resultReader.keySet().stream().filter(key -> !finalResultReader.get(key).get(finalIndexMessageTo).equals("-"))
                .filter(key -> {
                    if (!mobilePattern.matcher(finalResultReader.get(key).get(finalIndexMessageTo)).find() && isHeader.get()) {
                        isHeader.set(false);
                        return true;
                    } else {
                        return mobilePattern.matcher(finalResultReader.get(key).get(finalIndexMessageTo)).find();
                    }
                })
                .collect(Collectors.toSet()).iterator();
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
        // remove duplicate mobile numbers
        Map<Integer, List<String>> convertData2 = new HashMap<>();
        List columnNames2 = new ArrayList();
        columnNames2.add("mobile");
        columnNames2.add("phone");
        columnNames2.add("message");
        convertData2.put(0, columnNames2);
        convertData.keySet().forEach(key -> {
            if (key == 0)
                return;
            convertData2.put(Integer.valueOf(convertData.get(key).get(0).replaceFirst("09", "")), convertData.get(key));
        });
        executorService.submit(() -> instance.writer(convertData2, "sampleTest3"));
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
