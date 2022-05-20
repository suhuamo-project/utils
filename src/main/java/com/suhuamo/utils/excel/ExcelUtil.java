package com.suhuamo.utils.excel;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItem;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author suhuamo
 * @slogan 想和喜欢的人睡在冬日的暖阳里
 * @date 2022/05/15
 * excel工具类
 */
public class ExcelUtil {

    /**
     * 读取excel中的数据,读取该文件中的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为列数,空行不读取
     * 返回类型为List<Map>，每一行数据对应为一个map,map类型为<Integer,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file excel文件，格式必须为 xlsx格式
     * @return List<Map < String, Object>>
     */
    public static List<Map<Integer, Object>> readBigExcel(MultipartFile file) throws Exception {
        //定义返回值
        List<Map<Integer, Object>> resultList = new ArrayList<Map<Integer, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            // 这里相当于读到哪个row了，就加载哪个row，所以是使用的缓存，所以在这一个for里面能遍历完所有的行，但是是每读了100(上面定义的100)个row就才开始又加载100row，但始终是在这个for里面加载的
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列数，内容
                Map<Integer, Object> paramMap = new HashMap<Integer, Object>();//定义一个map做数据接收
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可
                if (isEmpty(row)) {
                    continue;
                }
                // 如果该行有效，那么读取每一个单元格的数据
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    //获取单元格数据
                    cell = row.getCell(i);
                    // 如果该单元格结果为空，那么直接在该格放入null值
                    if (isEmpty(cell)) {
                        //将单元格值放入map中
                        paramMap.put(i + 1, null);
                    } else {
                        //将单元格值放入map中
                        paramMap.put(i + 1, cell.getStringCellValue());
                    }
                }
                //一行循环完成，将该行的数据存入list
                resultList.add(paramMap);
            }
        }
        // 返回最终结果
        return resultList;
    }


    /**
     * 读取excel中的数据,读取该文件从 headLineNum（表头行）开始的的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum) throws Exception {
        // 获取每一列的列名 属性： 列数,名称
        HashMap<Integer, String> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取数据,直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(i, row.getCell(i).getStringCellValue());
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果该行有效，那么读取每一个单元格的数据
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    //获取单元格数据
                    cell = row.getCell(i);
                    // 如果该单元格结果为空，那么直接在该格放入null值
                    if (isEmpty(cell)) {
                        //将单元格值放入map
                        paramMap.put(heads.get(i), null);
                    } else {
                        //将单元格值放入map
                        paramMap.put(heads.get(i), cell.getStringCellValue());
                    }
                }
                //一行循环完成，将该行的数据存入list
                resultList.add(paramMap);
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum, Integer startRowNum) throws Exception {
        // 获取每一列的列名 属性： 列数,名称
        HashMap<Integer, String> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取数据，直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(i, row.getCell(i).getStringCellValue());
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果当前行到了需要读取的行以内了，那么开始读取和存入数据
                if (cntRow >= startRowNum) {
                    // 如果该行有效，那么读取每一个单元格的数据
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        //获取单元格数据
                        cell = row.getCell(i);
                        // 如果该单元格结果为空，那么直接在该格放入null值
                        if (isEmpty(cell)) {
                            //将单元格值放入map
                            paramMap.put(heads.get(i), null);
                        } else {
                            //将单元格值放入map
                            paramMap.put(heads.get(i), cell.getStringCellValue());
                        }
                    }
                    //一行循环完成，将该行的数据存入list
                    resultList.add(paramMap);
                }
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的 length 行数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @param length      总共需要读取多少行，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum, Integer startRowNum, Integer length) throws Exception {
        // 获取每一列的列名，属性： 列数,名称
        HashMap<Integer, String> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果当前行已经超过需要读取的行数了，那么就直接break，返回当前已经读取的所有数据
                if (cntRow >= startRowNum + length) {
                    break;
                }
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取表头数据，直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(i, row.getCell(i).getStringCellValue());
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果当前行到了需要读取的行以内了，那么开始读取和存入数据
                if (cntRow >= startRowNum) {
                    // 如果该行有效，那么读取每一个单元格的数据
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        //获取单元格数据
                        cell = row.getCell(i);
                        // 如果该单元格结果为空，那么直接在该格放入null值
                        if (isEmpty(cell)) {
                            //将单元格值放入map
                            paramMap.put(heads.get(i), null);
                        } else {
                            //将单元格值放入map
                            paramMap.put(heads.get(i), cell.getStringCellValue());
                        }
                    }
                    //一行循环完成，将该行的数据存入list
                    resultList.add(paramMap);
                }
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }


    /**
     * 读取excel中的数据,读取该文件从 headLineNum（表头行）开始的的所有行的数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum, List<String> rowNameList) throws Exception {
        // 获取每一列的列名 属性： 名称,列数
        HashMap<String, Integer> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取表头数据，直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(row.getCell(i).getStringCellValue(), i);
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果该行有效，那么读取需要的列的信息
                for (String rowName : rowNameList) {
                    //获取单元格数据
                    cell = row.getCell(heads.get(rowName));
                    // 如果该单元格结果为空，那么直接在该格放入null值
                    if (isEmpty(cell)) {
                        //将单元格值放入map
                        paramMap.put(rowName, null);
                    } else {
                        //将单元格值放入map
                        paramMap.put(rowName, cell.getStringCellValue());
                    }
                }
                //一行循环完成，将该行的数据存入list
                resultList.add(paramMap);
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的所有行数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum, List<String> rowNameList, Integer startRowNum) throws Exception {
        // 获取每一列的列名 属性： 名称,列数
        HashMap<String, Integer> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取表头数据，直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(row.getCell(i).getStringCellValue(), i);
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果当前行到了需要读取的行以内了，那么开始读取和存入数据
                if (cntRow >= startRowNum) {
                    // 如果该行有效，那么读取需要的列的信息
                    for (String rowName : rowNameList) {
                        //获取单元格数据
                        cell = row.getCell(heads.get(rowName));
                        // 如果该单元格结果为空，那么直接在该格放入null值
                        if (isEmpty(cell)) {
                            //将单元格值放入map
                            paramMap.put(rowName, null);
                        } else {
                            //将单元格值放入map
                            paramMap.put(rowName, cell.getStringCellValue());
                        }
                    }
                    //一行循环完成，将该行的数据存入list
                    resultList.add(paramMap);
                }
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的 length 行数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @param length      总共需要读取多少行，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(MultipartFile file, Integer headLineNum, List<String> rowNameList, Integer startRowNum, Integer length) throws Exception {
        // 获取每一列的列名 属性： 名称,列数
        HashMap<String, Integer> heads = new HashMap<>();
        // 记录当前在第几行，从1开始
        int cntRow = 1;
        //定义返回值
        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        // 获取文件流
        InputStream inputStream = file.getInputStream();
        // 获取文件输入流
        try (Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(1024 * 100)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(inputStream);) { //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            Sheet sheet = wk.getSheetAt(0);// 默认读取第0个工作簿
            Cell cell = null;//定义单元格
            // 这里相当于读到哪个row了，就加载哪个row，所以是使用的缓存，所以在这一个for里面能遍历完所有的行，但是是每读了100(上面定义的100)个row就才开始又加载100row，但始终是在这个for里面加载的
            //获取当前循环的行数据(因为只缓存了部分数据，所以不能用getRow来获取)此处采用增强for循环直接获取row对象
            for (Row row : sheet) {
                // 设置当前这一行保存数据的对象，格式： 列名，内容
                Map<String, Object> paramMap = new HashMap<String, Object>();//定义一个map做数据接收
                // 如果当前行已经超过需要读取的行数了，那么就直接break，返回当前已经读取的所有数据
                if (cntRow >= startRowNum + length) {
                    break;
                }
                // 如果该行是空行，那么无法使用任何函数，直接存读取下一行即可 || 或者还未到列名那一栏
                if (isEmpty(row) || cntRow < headLineNum) {
                    // 行号++
                    cntRow++;
                    // 不进行读取表头数据，直接跳到下一行
                    continue;
                }
                // 如果到了标题行，那么就进行标题序号的对应
                if (cntRow == headLineNum) {
                    // 遍历标题行的每一个字段，并存入heads中
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        heads.put(row.getCell(i).getStringCellValue(), i);
                    }
                    // 不进行读取表头数据，直接跳到下一行
                    // 行号++
                    cntRow++;
                    // 直接跳到下一行
                    continue;
                }
                // 如果当前行到了需要读取的行以内了，那么开始读取和存入数据
                if (cntRow >= startRowNum) {
                    // 如果该行有效，那么读取需要的列的信息
                    for (String rowName : rowNameList) {
                        //获取单元格数据
                        cell = row.getCell(heads.get(rowName));
                        // 如果该单元格结果为空，那么直接在该格放入null值
                        if (isEmpty(cell)) {
                            //将单元格值放入map
                            paramMap.put(rowName, null);
                        } else {
                            //将单元格值放入map
                            paramMap.put(rowName, cell.getStringCellValue());
                        }
                    }
                    //一行循环完成，将该行的数据存入list
                    resultList.add(paramMap);
                }
                // 行号++
                cntRow++;
            }
        }
        // 返回最终结果
        return resultList;
    }


    /**
     * 读取excel中的数据,读取该文件中的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为列数,空行不读取
     * 返回类型为List<Map>，每一行数据对应为一个map,map类型为<Integer,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file excel文件，格式必须为 xlsx格式
     * @return List<Map < String, Object>>
     */
    public static List<Map<Integer, Object>> readBigExcel(File file) throws Exception {
        return readBigExcel(FileToMultipartFile(file));
    }

    /**
     * 读取excel中的数据,读取该文件从 headLineNum（表头行）开始的的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum);
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的所有行的数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum, Integer startRowNum) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum, startRowNum);
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的 length 行数据，并且读取每一列的信息,每一列的列名映射对应为字符串-真实列名,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @param length      总共需要读取多少行，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum, Integer startRowNum, Integer length) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum, startRowNum, length);
    }

    /**
     * 读取excel中的数据,读取该文件从 headLineNum（表头行）开始的的所有行的数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum, List<String> rowNameList) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum, rowNameList);
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的所有行数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum, List<String> rowNameList, Integer startRowNum) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum, rowNameList, startRowNum);
    }

    /**
     * 读取excel中的数据,读取该文件中从 startRowNum 行开始的 length 行数据，并且只读取需要的列,空行不读取
     * 返回类型未Map映射，每一行数据为一个map，map类型为<string,object>,对应属性为，列名：单元格内容
     * 格式限制：必须使用xlsx的格式，调用前需判断格式
     *
     * @param file        excel文件，格式必须为 xlsx格式
     * @param headLineNum 标题所在的行号，从1开始
     * @param rowNameList 需要查询的列信息 如 "姓名,身份证" 等
     * @param startRowNum 读取的起始行，从1开始 -即从哪一行开始读取数据
     * @param length      总共需要读取多少行，从1开始
     * @return List<Map < String, Object>>
     */
    public static List<Map<String, Object>> readBigExcel(File file, Integer headLineNum, List<String> rowNameList, Integer startRowNum, Integer length) throws Exception {
        return readBigExcel(FileToMultipartFile(file), headLineNum, rowNameList, startRowNum, length);
    }



    /**
     * 判断当前这一行是否存在
     *
     * @param row
     * @return boolean
     */
    private static boolean isEmpty(Row row) {
        // 如果为空，则返回true
        if (row == null) {
            return true;
        }
        return false;
    }

    /**
     * 判断当前这一个表格是否有数据
     *
     * @param cell
     * @return boolean
     */
    private static boolean isEmpty(Cell cell) {
        // 如果为空，则返回false
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return true;
        }
        return false;
    }

    /**
     * 将file文件转换为MultipartFile文件格式
     *
     * @param
     * @return MultipartFile
     */
    private static MultipartFile FileToMultipartFile(File file) throws IOException {
        // 需要导入commons-fileupload的包
        FileItem fileItem = new DiskFileItem(file.getName(), Files.probeContentType(file.toPath()), false, file.getName(), (int) file.length(), file.getParentFile());
        byte[] buffer = new byte[4096];
        int n;
        MultipartFile multipartFile = null;
        try (InputStream inputStream = new FileInputStream(file); OutputStream os = fileItem.getOutputStream()) {
            while ((n = inputStream.read(buffer, 0, 4096)) != -1) {
                os.write(buffer, 0, n);
            }
            //也可以用IOUtils.copy(inputStream,os);
            multipartFile = new CommonsMultipartFile(fileItem);
            System.out.println(multipartFile.getName());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return multipartFile;
    }
}
