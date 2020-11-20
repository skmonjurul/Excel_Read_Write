package com.write;

import com.bean.NeetRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class Write {

    public XSSFWorkbook getWorkBok(){
        return new XSSFWorkbook();
    }

    public XSSFSheet getSheet(XSSFWorkbook workbook,String sheet_name){
        return workbook.createSheet(sheet_name);

    }

    public void writeHeader(XSSFSheet sheet, List<String> header_list, int row_no){
        Row row = sheet.createRow(row_no);
        int cellnum = 0;
        for(String str: header_list){
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(str);
        }
    }

    public void writeData(XSSFSheet sheet, Map<Integer, NeetRecord> allRecordsData, int row_no){
        Set<Integer> keySet = allRecordsData.keySet();
        for(Integer key: keySet){
            Row row = sheet.createRow(row_no++);
            NeetRecord neetRecord = allRecordsData.get(key);
            writeRecord(row, neetRecord);
        }

    }

    public void writeRecord(Row row, NeetRecord neetRecord){
        row.createCell(0).setCellValue(neetRecord.getNeet_roll_no());
        row.createCell(1).setCellValue(neetRecord.getAppln_no());
        row.createCell(2).setCellValue(neetRecord.getCandidate_name());
        row.createCell(3).setCellValue(neetRecord.getNationality());
        row.createCell(4).setCellValue(neetRecord.getHk());
        row.createCell(5).setCellValue(neetRecord.getJk());
        row.createCell(6).setCellValue(neetRecord.getReligious_minority());
        row.createCell(7).setCellValue(neetRecord.getNeet_score());
        row.createCell(8).setCellValue(neetRecord.getIncome());
        row.createCell(9).setCellValue(neetRecord.getFather_name());
        row.createCell(10).setCellValue(neetRecord.getClause());
        row.createCell(11).setCellValue(neetRecord.getRural());
        row.createCell(12).setCellValue(neetRecord.getSpecial_category());
        row.createCell(13).setCellValue(neetRecord.getLingustic_minority());
        row.createCell(14).setCellValue(neetRecord.getNeet_ai_rank());
        row.createCell(15).setCellValue(neetRecord.getCet_no());
        row.createCell(16).setCellValue(neetRecord.getMother_name());
        row.createCell(17).setCellValue(neetRecord.getCategory());
        row.createCell(18).setCellValue(neetRecord.getKannada());
        row.createCell(19).setCellValue(neetRecord.getNri_ward());

    }

    public void writeSheet(XSSFWorkbook workbook, String file_name){
        try{
            FileOutputStream out = new FileOutputStream(new File(System.getProperty("user.dir") + "\\file\\" + file_name + ".xlsx"));
            workbook.write(out);
            out.close();
        }
        catch(FileNotFoundException e){
            System.out.println(e);
        }
        catch(IOException e){
            System.out.println(e);
        }
    }

}
