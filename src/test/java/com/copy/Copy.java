package com.copy;

import com.bean.NeetRecord;
import com.read.Read;
import com.write.Write;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.List;
import java.util.Map;

public class Copy {
    String read_file_name = "eligible_neet_1000";
    String write_file_name = "eligible_neet_1000_formatted";
    String read_sheet_name = "Table 1";
    String write_sheet_name = "Table 1";
    Read read;
    Write write;
    List<String> header_list;
    Map<Integer, NeetRecord> allDataRecordMap;

    public void readData(){
        read = new Read();
        FileInputStream file = read.getFile(read_file_name);
        Workbook workbook = read.getWorkBook(file);
        Sheet sheet = read.getSheet(workbook, read_sheet_name);
        header_list = read.getHeaderList(sheet);
        allDataRecordMap = read.getData(sheet, header_list);
        read.closeFile();
    }

    public void writeData(){
        write = new Write();
        XSSFWorkbook workbook = write.getWorkBok();
        XSSFSheet sheet = write.getSheet(workbook, write_sheet_name);
        write.writeHeader(sheet, header_list, 0);
        write.writeData(sheet, allDataRecordMap, 1);
        write.writeSheet(workbook,write_file_name);
    }
}
