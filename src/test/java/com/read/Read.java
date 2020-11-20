package com.read;

import com.bean.NeetRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class Read {

    FileInputStream file=null;

    public FileInputStream getFile(String file_name){
        try{
            file = new FileInputStream(new File(System.getProperty("user.dir") + "\\file\\" + file_name + ".xlsx"));
        }
        catch(FileNotFoundException e){
            System.out.println(e);
        }

        return file;
    }

    public void closeFile(){
        try {
            this.file.close();
        }
        catch(IOException e){
            System.out.println(e);
        }
    }

    public Workbook getWorkBook(FileInputStream file){
        Workbook workbook=null;
        try{
            workbook = new XSSFWorkbook(file);
        }
        catch (IOException e){
            System.out.println(e);
        }
        return workbook;
    }

    public Sheet getSheet(Workbook workbook, String sheet_name){
        Sheet sheet;
        sheet = workbook.getSheet(sheet_name);
        return sheet;
    }

    public List<String> getHeaderList(Sheet sheet){
        int header_starting = 0;
        List<String> header_list = new ArrayList<String>();
        for(int i = header_starting; i<=header_starting+2; i++){
            Row row = sheet.getRow(i);
            for(int j=0; j< row.getLastCellNum(); j++){
                if (row.getCell(j).getStringCellValue().length() > 0) {
                    header_list.add(row.getCell(j).getStringCellValue());
                }
            }
        }
        return header_list;
    }

    public boolean isRowEmpty(Row row){
        boolean flag = true;
        for (int i=0; i< row.getLastCellNum(); i++){
            try {
                if (row.getCell(i).getStringCellValue().length() > 0) {
                    flag = false;
                    break;
                }
            }
            catch (IllegalStateException e){
                flag = false;
                break;
            }
        }
        return flag;
    }

    public boolean isHeader(List<String> header_list, Row row){
        boolean flag = false;
        for (int i=0; i< row.getLastCellNum(); i++){
            try {
                if (header_list.contains(row.getCell(i).getStringCellValue())) {
                    flag = true;
                    break;
                }
            }
            catch (IllegalStateException e){
                flag = false;
                break;
            }
        }
        return flag;
    }

    public Map<Integer, NeetRecord> getData(Sheet sheet, List<String> header_list){
        boolean stop = true;
        int map_key = 1;
        Map<Integer, NeetRecord> allRecordMap = new TreeMap<Integer, NeetRecord>();
        for (int i=0; i< sheet.getLastRowNum(); i++){
            Row row = sheet.getRow(i);
            if (isRowEmpty(row) || isHeader(header_list, row))
                continue;
            else{
                List<Object> data;
                NeetRecord neetRecord=null;
                data = readData(sheet, i, i + 3, header_list);
                if(data.size() == 20) {
                    neetRecord = getNeetRecordObject(data);
                    i = i+2;
                }
                else{
                    if(data.size()>=8){
                        i=i+2;
                        i = readRemainingData(sheet, i, 1, header_list, data);
                        neetRecord = getNeetRecordObject(data);
                    }
                    if(data.size() <= 7){
                        i = i+1;
                        i = readRemainingData(sheet, i, 2, header_list,data)+1;
                        neetRecord = getNeetRecordObject(data);
                    }
                }
                allRecordMap.put(map_key++, neetRecord);
            }
        }
        return allRecordMap;
    }

    public List<Object> readData(Sheet sheet, int start_row, int end_row, List<String> header_list){
        List<Object> data = new ArrayList<Object>();
        for (int i=start_row; i<end_row; i++){
            Row row = sheet.getRow(i);
            if (isRowEmpty(row) || isHeader(header_list, row))
                continue;
            else {
                Iterator<Cell> cell_ite = row.cellIterator();
                while (cell_ite.hasNext()) {
                    Cell cell = cell_ite.next();
                    if(cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.STRING) {
                        data.add(new DataFormatter().formatCellValue(cell));
                    }
                }
            }

        }
        return data;
    }

    public int readRemainingData(Sheet sheet, int start_row, int remaining_no, List<String> header_list, List<Object> data){
        while(data.size()!=20){
            Row row = sheet.getRow(start_row);
            if(isRowEmpty(row) || isHeader(header_list, row)){
                start_row++;
            }
            else{
                data.addAll(readData(sheet, start_row, start_row+remaining_no, header_list));
            }
        }
        return start_row;
    }

    public NeetRecord getNeetRecordObject(List<Object> data){
        NeetRecord neetRecord = new NeetRecord();
        neetRecord.setNeet_roll_no(data.get(0).toString());
        neetRecord.setAppln_no(data.get(1).toString());
        neetRecord.setCandidate_name(data.get(2).toString());
        neetRecord.setNationality(data.get(3).toString());
        neetRecord.setHk(data.get(4).toString());
        neetRecord.setJk(data.get(5).toString());
        neetRecord.setReligious_minority(data.get(6).toString());
        neetRecord.setNeet_score(data.get(7).toString());
        neetRecord.setIncome(data.get(8).toString());
        neetRecord.setFather_name(data.get(9).toString());
        neetRecord.setClause(data.get(10).toString());
        neetRecord.setRural(data.get(11).toString());
        neetRecord.setSpecial_category(data.get(12).toString());
        neetRecord.setLingustic_minority(data.get(13).toString());
        neetRecord.setNeet_ai_rank(data.get(14).toString());
        neetRecord.setCet_no(data.get(15).toString());
        neetRecord.setMother_name(data.get(16).toString());
        neetRecord.setCategory(data.get(17).toString());
        neetRecord.setKannada(data.get(18).toString());
        neetRecord.setNri_ward(data.get(19).toString());
        return neetRecord;


    }

}
