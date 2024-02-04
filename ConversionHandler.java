package com.reve.sms.common.excelParser;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.reve.sms.common.radius.RadiusConfiguration;
import com.reve.sms.common.radius.RadiusConfigurationConstants;
import com.reve.sms.rateAndDestination.ratePlan.dto.*;
import com.reve.sms.rateAndDestination.smsCountry.SMSMccMncRepository;
import com.reve.sms.rateAndDestination.smsCountry.dto.SMSCountryDTO;
import com.reve.sms.util.DisplayFormat;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

//Singleton
public class ConversionHandler {
    private static final Logger logger = LoggerFactory.getLogger(ConversionHandler.class);
    private static ConversionHandler instance;
    private final  DataFormatter dataFormatter;

    private ConversionHandler(){
        dataFormatter = new DataFormatter();
    }

    public static ConversionHandler getInstance() {
        if (instance == null) {
            instance = new ConversionHandler();
        }
        return instance;
    }
    /**
     * This function receives a File and converts it to a ParsedFile Object.
     * The first non-null row of a sheet is considered as the header row and the rest of the rows are parsed as data rows.
     *
     * @param uploadedFile The file to be converted.
     * @return The ParsedFile object that is converted from the uploadedFile.
     * @throws IllegalArgumentException If the file is not of xls or xlsx format then this exception is thrown.
     */
    public ParsedFile parseExcelToJson(File uploadedFile) throws IllegalArgumentException{
        // hold the excel data sheet wise
        ObjectMapper mapper = new ObjectMapper();
        ObjectNode excelData = mapper.createObjectNode();
        FileInputStream fis = null;
        Workbook workbook = null;
        try {
            // Creating file input stream
            fis = new FileInputStream(uploadedFile);

            String fileName = uploadedFile.getName().toLowerCase();
            String fileExtension;

            if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {
                // creating workbook object based on excel file format
                if (fileName.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                    fileExtension = "xls";
                } else {
                    workbook = new XSSFWorkbook(fis);
                    fileExtension = "xlsx";
                }

                if (workbook.getNumberOfSheets() == 0) {
                    return null;
                }

                ArrayList<ParsedSheet> parsedSheets = new ArrayList<>();

                // Reading each sheet one by one
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    ArrayList<String> headerList = new ArrayList<>();

                    ArrayNode sheetData = mapper.createArrayNode();
                    // Reading each row of the sheet
                    boolean foundHeader = false;
                    for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                        Row row = sheet.getRow(j);
                        if (row == null) {
                            continue;
                        }
                        if (!foundHeader) {
                            // reading sheet header's name
                            for (int k = 0; k < row.getLastCellNum(); k++) {
                                headerList.add(row.getCell(k).getStringCellValue());
                            }
                            ///validate header list, if invalid then clear the current headerlist
                            foundHeader = isHeaderValid(headerList);
                            if(!foundHeader) headerList.clear();
                        } else {
                            // reading work sheet data
                            ObjectNode rowData = mapper.createObjectNode();
                            for (int k = 0; k < headerList.size(); k++) {
                                Cell cell = row.getCell(k);
                                String headerName = headerList.get(k);
                                if (cell != null) {
                                    rowData.put(headerName, dataFormatter.formatCellValue(cell));
                                } else {
                                    rowData.put(headerName, "");
                                }
                            }
                            sheetData.add(rowData);
                        }
                    }
                    ParsedSheet parsedSheet = new ParsedSheet(sheetName, headerList, sheetData);
                    parsedSheets.add(parsedSheet);
                }

                ParsedFile parsedFile = new ParsedFile(fileName, fileExtension, parsedSheets);
                return parsedFile;
            } else {
                logger.error("File format not supported.");
                throw new IllegalArgumentException("File format not supported.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }
        return null;
    }
    /**
     * This function validates a header row. A header row is valid if none of its cells are empty
     *
     * @param headerList The list of the possible header.
     * @return true if the header row is valid, false otherwise.
     */
    private boolean isHeaderValid(ArrayList<String> headerList){
        for(String headerCell: headerList){
            if(headerCell == null || headerCell.trim().equals("")) return false;
        }
        return true;
    }
}
