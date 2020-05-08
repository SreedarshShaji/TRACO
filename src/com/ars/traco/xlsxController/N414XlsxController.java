/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.traco.xlsxController;

import java.text.DateFormatSymbols;
import java.text.DecimalFormat;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.*;

import com.ars.traco.databeans.n414.Sectie;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;

/**
 *
 * @author Sreedarshs
 */
public class N414XlsxController {

    public DecimalFormat dec = new DecimalFormat("#0.00");

    public boolean handleXlsx(String inputFile, Sectie bean) {

        boolean result = false;
        String directory = inputFile.substring(0, inputFile.lastIndexOf('\\'));
        String fileName = inputFile.substring(inputFile.lastIndexOf('\\') + 1);

        System.out.println("Absolute path: " + inputFile);
        System.out.println("File Path: " + directory);
        System.out.println("Filename: " + fileName);

        String fileNameToWrite = sourceFileName(fileName);

        /* Check if file exists in the directory */
        boolean outputfileExists = (outputFileExists(directory, fileNameToWrite));
        try {
            if (outputfileExists) {
                System.out.println("File exists");
            } else {
                System.out.println("File doesnot exists");
                System.out.println("Creating a new copy of the file ");

                Path root = FileSystems.getDefault().getPath("").toAbsolutePath();
                Path sourceDirectory = Paths.get(root + "\\Resources\\N414.xlsx");
                Path targetDirectory = Paths.get(directory + "\\" + fileNameToWrite);

                System.out.println("Path : " + root + "\\Resources\\N414.xlsx");

                // copy source to target using Files Class
                Files.copy(sourceDirectory, targetDirectory);

            }

            // Write to the file
            boolean writingStatus = writingToFile(directory + "\\" + fileNameToWrite, bean);
            if (writingStatus) {
                System.out.println("File written to successfully");
                result = true;
            } else {
                System.out.println("File writing failed");
                result = false;
            }

            // writingToFile(fileNameToWrite);
        } catch (IOException ioe) {
            System.out.println("IOException : " + ioe);
        }

        return result;
    }

    public boolean checkForExcel() {

        return false;
    }

    /*
	 * Create the target file name
	 * 
     */
    public String sourceFileName(String fileName) {
        /* Extracting the date of excel file */
        String fileDate = fileName.substring(fileName.indexOf('_') + 1, fileName.lastIndexOf('.'));
        String fileYear = fileDate.substring(0, 4);
        String fileMonth = fileDate.substring(4, 6);
        // String fileDay=fileDate.substring(6,8);

        fileMonth = new DateFormatSymbols().getMonths()[Integer.parseInt(fileMonth) - 1];

        String fileNameToLocate = "TRACO_OWN_N414_Daily_Report_" + fileMonth + fileYear + ".xlsx";

        return fileNameToLocate;

    }

    /*
	 * Check if the file exists
     */
    public boolean outputFileExists(String filePath, String OutputFileName) {
        /* Checks if the file exists */
        boolean result = false;
        try {
            File tempFile = new File(filePath + '\\' + OutputFileName);
            result = tempFile.exists();
        } catch (Exception ex) {
            System.err.println(ex.getMessage());
        }

        return result;
    }

    public boolean writingToFile(String path, Sectie bean) {
        boolean result = false;

        try (InputStream inp = new FileInputStream(path)) {
            /* Locating the workbook*/
            Workbook wb = WorkbookFactory.create(inp);

            /* Gets date from the bean*/
            String date = bean.getDate();
            int day = Integer.parseInt(date.substring(date.lastIndexOf("-") + 1));
            System.out.println(day);

            /**
             * **************************************************************************************************
             */
            /**
             * ****************************************TCS UT N414 R - 1*****************************************
             */
            /**
             * **************************************************************************************************
             */

            /*Reading R1 the sheet*/
            Sheet R1sheet = wb.getSheet("TCS UT N414 R - 1");

            /* Locating the row and column */
            Iterator<Row> iterator = R1sheet.iterator();
            int rowNumber = 0;

            while (iterator.hasNext()) {
                rowNumber++;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                Cell currentCell = cellIterator.next();
                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    continue;
                } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    if ((int) currentCell.getNumericCellValue() == day) {
                        break;
                    }
                }

            }

            System.out.println("Row Count :" + rowNumber);

            Row rowToInsert = R1sheet.getRow(rowNumber - 1);
            //Cell cellToInsert = rowToInsert.getCell(1);

            /*
			 * Defining cells to populate the value
			 * */
            Cell cellR1Begin = rowToInsert.getCell(1);
            Cell cellR1End = rowToInsert.getCell(2);
            Cell cellR1Max = rowToInsert.getCell(3);
            Cell cellR1Aantal = rowToInsert.getCell(4);
            Cell cellR1Gem_kmpuur = rowToInsert.getCell(5);
            Cell cellR1Max_kmpuur = rowToInsert.getCell(6);
            Cell cellR1Total = rowToInsert.getCell(7);
            Cell cellR1Hand = rowToInsert.getCell(8);
            Cell cellR1Auto = rowToInsert.getCell(9);
            Cell cellR1Dubbele_overtredingen_pardon = rowToInsert.getCell(10);
            Cell cellR1Overig_pardon = rowToInsert.getCell(11);
            Cell cellR1OvertredingenRatio = rowToInsert.getCell(12);
            Cell cellR1Handhaafratio = rowToInsert.getCell(13);
            Cell cellR1TijdVolledigBeschikbaar_in_minuten = rowToInsert.getCell(14);
            Cell cellR1BeschikbaarheidsRatio = rowToInsert.getCell(15);
            Cell cellR1MatchRatio = rowToInsert.getCell(16);
            Cell cellR1ProductMatchRatioRegistratieratio = rowToInsert.getCell(17);
            Cell cellR1AutoRatio = rowToInsert.getCell(18);

            //Cell cellR1OvertredingenRatio= rowToInsert.getCell(11);
            /*
			Cell cellToInsert = rowToInsert.getCell(1);
             */
 /* If such a cell doesn't exist , then the cell is created
			if (cellToInsert == null)
				cellToInsert = rowToInsert.createCell(3); */
 /* Setting cell type
			cellR1Begin.setCellType(CellType.);
			cellR1End.setCellType(CellType.NUMERIC);
             */
 /* Setting cell value */
            cellR1Begin.setCellValue(bean.getR1().getpassagesType().getTotalIn());
            cellR1End.setCellValue(bean.getR1().getpassagesType().getTotalUit());
            cellR1Max.setCellFormula("IF(MAX(B" + rowNumber + ":C" + rowNumber + ")=0,\"\",MAX(B" + rowNumber + ":C" + rowNumber + "))");
            cellR1Aantal.setCellValue(bean.getR1().getMatches());

            /*Setting the decimal preciosion points
            String r1Gem_kmpuur = dec.format(Math.ceil(bean.getR1().getsnelhedenType().getGemiddeld()));
            String r1Max_kmpuur = dec.format(Math.ceil(bean.getR1().getsnelhedenType().getMax()));*/

            
            
            cellR1Gem_kmpuur.setCellValue(bean.getR1().getsnelhedenType().getGemiddeld());
            cellR1Max_kmpuur.setCellValue(bean.getR1().getsnelhedenType().getMax());

            cellR1Total.setCellValue(bean.getR1().getOvertredingenType().getOvertredingenTotaal());
            cellR1Hand.setCellValue(bean.getR1().getOvertredingenType().getHand());
            cellR1Auto.setCellValue(bean.getR1().getOvertredingenType().getAuto());
            cellR1Dubbele_overtredingen_pardon.setCellValue(bean.getR1().getOvertredingenType().getDubbeleOvertredingenPardon());
            cellR1Overig_pardon.setCellValue(bean.getR1().getOvertredingenType().getOverigPardon());

            String R1OvertredingenRatio = dec.format((bean.getR1().getPerformanceType().getOvertredingenratio()) * 100);
            cellR1OvertredingenRatio.setCellValue(R1OvertredingenRatio + "%");

            cellR1Handhaafratio.setCellValue(bean.getR1().getPerformanceType().getHandhaafratio());

            String tijdvolledigbeschiPerkbaarR1 = bean.getR1().getPerformanceType().getTijdvolledigbeschikbaar();
            int daysR1 = Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(1, tijdvolledigbeschiPerkbaarR1.indexOf('D')));
            int hoursR1 = Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(tijdvolledigbeschiPerkbaarR1.indexOf('T') + 1, tijdvolledigbeschiPerkbaarR1.indexOf('H')));
            int minutesR1 = Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(tijdvolledigbeschiPerkbaarR1.indexOf('H') + 1, tijdvolledigbeschiPerkbaarR1.indexOf('M')));

            cellR1TijdVolledigBeschikbaar_in_minuten.setCellValue(daysR1 * 24 * 60 + hoursR1 * 60 + minutesR1);

            cellR1BeschikbaarheidsRatio.setCellFormula("IF((O" + rowNumber + "/1440)=0,\"\",(O" + rowNumber + "/1440))");

            /* Tested */
            cellR1MatchRatio.setCellValue(bean.getR1().getPerformanceType().getMatchratio());

            cellR1ProductMatchRatioRegistratieratio.setCellFormula("IF((Q" + rowNumber + "*'Total Systemperformance'!$C$6)=0,\"\",(Q" + rowNumber + "*'Total Systemperformance'!$C$6))");

            cellR1AutoRatio.setCellValue((bean.getR1().getPerformanceType().getAutoratio()));

            /**
             * **************************************************************************************************
             */
            /**
             * ****************************************TCS UT N414 L - 2*****************************************
             */
            /**
             * **************************************************************************************************
             */

            /*Reading L2 the sheet*/
            Sheet L2sheet = wb.getSheet("TCS UT N414 L - 2");

            /* Locating the row and column */
            Iterator<Row> iteratorL2 = L2sheet.iterator();
            int rowNumberL2 = 0;

            while (iterator.hasNext()) {
                rowNumberL2++;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                Cell currentCell = cellIterator.next();
                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    continue;
                } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    if ((int) currentCell.getNumericCellValue() == day) {
                        break;
                    }
                }

            }

            System.out.println("L2Row Count :" + rowNumberL2);

            Row rowToInsertL2 = L2sheet.getRow(rowNumber - 1);
            //Cell cellToInsert = rowToInsert.getCell(1);

            /*
			 * Defining cells to populate the value
			 * */
            Cell cellL2Begin = rowToInsertL2.getCell(1);
            Cell cellL2End = rowToInsertL2.getCell(2);
            Cell cellL2Max = rowToInsertL2.getCell(3);
            Cell cellL2Aantal = rowToInsertL2.getCell(4);
            Cell cellL2Gem_kmpuur = rowToInsertL2.getCell(5);
            Cell cellL2Max_kmpuur = rowToInsertL2.getCell(6);
            Cell cellL2Total = rowToInsertL2.getCell(7);
            Cell cellL2Hand = rowToInsertL2.getCell(8);
            Cell cellL2Auto = rowToInsertL2.getCell(9);
            Cell cellL2Dubbele_overtredingen_pardon = rowToInsertL2.getCell(10);
            Cell cellL2Overig_pardon = rowToInsertL2.getCell(11);
            Cell cellL2OvertredingenRatio = rowToInsertL2.getCell(12);
            Cell cellL2Handhaafratio = rowToInsertL2.getCell(13);
            Cell cellL2TijdVolledigBeschikbaar_in_minuten = rowToInsertL2.getCell(14);
            Cell cellL2BeschikbaarheidsRatio = rowToInsertL2.getCell(15);
            Cell cellL2MatchRatio = rowToInsertL2.getCell(16);
            Cell cellL2ProductMatchRatioRegistratieratio = rowToInsertL2.getCell(17);
            Cell cellL2AutoRatio = rowToInsertL2.getCell(18);

            //Cell cellR1OvertredingenRatio= rowToInsert.getCell(11);
            /*
			Cell cellToInsert = rowToInsert.getCell(1);
             */
 /* If such a cell doesn't exist , then the cell is created
			if (cellToInsert == null)
				cellToInsert = rowToInsert.createCell(3); */
 /* Setting cell type
			cellR1Begin.setCellType(CellType.);
			cellR1End.setCellType(CellType.NUMERIC);
             */
 /* Setting cell value */
            cellL2Begin.setCellValue(bean.getL2().getpassagesType().getTotalIn());
            cellL2End.setCellValue(bean.getL2().getpassagesType().getTotalUit());
            cellL2Max.setCellFormula("IF(MAX(B" + rowNumber + ":C" + rowNumber + ")=0,\"\",MAX(B" + rowNumber + ":C" + rowNumber + "))");
            cellL2Aantal.setCellValue(bean.getL2().getMatches());

            /* Set decimal precison points
            String l1Gem_kmpuur = dec.format();
            String l1Max_kmpuur = dec.format(Math.ceil());
             */

            cellL2Gem_kmpuur.setCellValue(bean.getL2().getSnelheden().getGemiddeld());
            cellL2Max_kmpuur.setCellValue(bean.getL2().getSnelheden().getMax());

            cellL2Total.setCellValue(bean.getL2().getOvertredingenType().getOvertredingenTotaal());
            cellL2Hand.setCellValue(bean.getL2().getOvertredingenType().getHand());
            cellL2Auto.setCellValue(bean.getL2().getOvertredingenType().getAuto());
            cellL2Dubbele_overtredingen_pardon.setCellValue(bean.getL2().getOvertredingenType().getDubbeleOvertredingenPardon());
            cellL2Overig_pardon.setCellValue(bean.getL2().getOvertredingenType().getOverigPardon());

            Double L2OvertredingenRatio = (bean.getL2().getPerformanceType().getOvertredingenratio() * 100);
            L2OvertredingenRatio = BigDecimal.valueOf(L2OvertredingenRatio).setScale(2, RoundingMode.HALF_UP).doubleValue();
            cellL2OvertredingenRatio.setCellValue(L2OvertredingenRatio / 100);
            /* Temporary patch, needs a further look */

            cellL2Handhaafratio.setCellValue(bean.getL2().getPerformanceType().getHandhaafratio());

            String tijdvolledigbeschiPerkbaarL2 = bean.getL2().getPerformanceType().getTijdvolledigbeschikbaar();
            int daysL2 = Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(1, tijdvolledigbeschiPerkbaarL2.indexOf('D')));
            int hoursL2 = Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(tijdvolledigbeschiPerkbaarL2.indexOf('T') + 1, tijdvolledigbeschiPerkbaarL2.indexOf('H')));
            int minutesL2 = Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(tijdvolledigbeschiPerkbaarL2.indexOf('H') + 1, tijdvolledigbeschiPerkbaarL2.indexOf('M')));

            cellL2TijdVolledigBeschikbaar_in_minuten.setCellValue(daysL2 * 24 * 60 + hoursL2 * 60 + minutesL2);

            cellL2BeschikbaarheidsRatio.setCellFormula("IF((O" + rowNumber + "/1440)=0,\"\",(O" + rowNumber + "/1440))");

            Double matchRatioL2 = (bean.getL2().getPerformanceType().getMatchratio());
            matchRatioL2 = BigDecimal.valueOf(matchRatioL2).setScale(2, RoundingMode.HALF_UP).doubleValue();

            cellL2MatchRatio.setCellValue(bean.getL2().getPerformanceType().getMatchratio());

            cellL2ProductMatchRatioRegistratieratio.setCellFormula("IF((Q" + rowNumber + "*'Total Systemperformance'!$C$6)=0,\"\",(Q" + rowNumber + "*'Total Systemperformance'!$C$6))");

            cellL2AutoRatio.setCellValue((bean.getL2().getPerformanceType().getAutoratio()));

            /**
             * **************************************************************************************************
             */
            /**
             * **************************************** Total_R&L    *****************************************
             */
            /**
             * **************************************************************************************************
             */

            /*Reading L2 the sheet*/
            try {
                Sheet Total_R_Lsheet = wb.getSheet("Total_R&L");
                Iterator<Row> iteratorTotal_R_L = Total_R_Lsheet.iterator();

                System.out.println("Total_R&L Count :" + rowNumberL2);

                Row rowToInsertTotal_R_L = Total_R_Lsheet.getRow(rowNumber - 1);

                Cell cellTotalBegin = rowToInsertTotal_R_L.getCell(1);
                Cell cellTotalEnd = rowToInsertTotal_R_L.getCell(2);
                Cell cellTotalMax = rowToInsertTotal_R_L.getCell(3);
                Cell cellTotalAantal = rowToInsertTotal_R_L.getCell(4);
                Cell cellTotalGem_kmpuur = rowToInsertTotal_R_L.getCell(5);
                Cell cellTotalMax_kmpuur = rowToInsertTotal_R_L.getCell(6);
                Cell cellTotalTotal = rowToInsertTotal_R_L.getCell(7);
                Cell cellTotalHand = rowToInsertTotal_R_L.getCell(8);
                Cell cellTotalAuto = rowToInsertTotal_R_L.getCell(9);
                Cell cellTotalDubbele_overtredingen_pardon = rowToInsertTotal_R_L.getCell(10);
                Cell cellTotalOverig_pardon = rowToInsertTotal_R_L.getCell(11);
                Cell cellTotalOvertredingenRatio = rowToInsertTotal_R_L.getCell(12);
                Cell cellTotalHandhaafratio = rowToInsertTotal_R_L.getCell(13);
                Cell cellTotalTijdVolledigBeschikbaar_in_minuten = rowToInsertTotal_R_L.getCell(14);
                Cell cellTotalBeschikbaarheidsRatio = rowToInsertTotal_R_L.getCell(15);
                Cell cellTotalMatchRatio = rowToInsertTotal_R_L.getCell(16);
                Cell cellTotalProductMatchRatioRegistratieratio = rowToInsertTotal_R_L.getCell(17);
                Cell cellTotalAutoRatio = rowToInsertTotal_R_L.getCell(18);

                /*Setting the total cell value*/
                cellTotalBegin.setCellFormula("'TCS UT N414 R - 1'!B" + rowNumber + "+'TCS UT N414 L - 2'!B" + rowNumber + "");
                cellTotalEnd.setCellFormula("'TCS UT N414 R - 1'!C" + rowNumber + "+'TCS UT N414 L - 2'!C" + rowNumber + "");
                cellTotalMax.setCellFormula("MAX(B" + rowNumber + ":C" + rowNumber + ")");
                cellTotalAantal.setCellFormula("'TCS UT N414 R - 1'!E" + rowNumber + "+'TCS UT N414 L - 2'!E" + rowNumber + "");
                cellTotalGem_kmpuur.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!F" + rowNumber + ",'TCS UT N414 L - 2'!F" + rowNumber + "),\"\")");
                cellTotalMax_kmpuur.setCellFormula("MAX('TCS UT N414 R - 1'!G" + rowNumber + ",'TCS UT N414 L - 2'!G" + rowNumber + ")");
                cellTotalTotal.setCellFormula("'TCS UT N414 R - 1'!H" + rowNumber + "+'TCS UT N414 L - 2'!H" + rowNumber + "");
                cellTotalHand.setCellFormula("'TCS UT N414 R - 1'!I" + rowNumber + "+'TCS UT N414 L - 2'!I" + rowNumber + "");
                cellTotalAuto.setCellFormula("'TCS UT N414 R - 1'!J" + rowNumber + "+'TCS UT N414 L - 2'!J" + rowNumber + "");
                cellTotalDubbele_overtredingen_pardon.setCellFormula("'TCS UT N414 R - 1'!K" + rowNumber + "+'TCS UT N414 L - 2'!K" + rowNumber + "");
                cellTotalOverig_pardon.setCellFormula("'TCS UT N414 R - 1'!L" + rowNumber + "+'TCS UT N414 L - 2'!L" + rowNumber + "");
                cellTotalOvertredingenRatio.setCellFormula("IF(H" + rowNumber + "=0,\"\",AVERAGE('TCS UT N414 R - 1'!M" + rowNumber + ",'TCS UT N414 L - 2'!M" + rowNumber + "))");
                cellTotalHandhaafratio.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!N" + rowNumber + ",'TCS UT N414 L - 2'!N" + rowNumber + "),\"\")");
                cellTotalTijdVolledigBeschikbaar_in_minuten.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!O" + rowNumber + ",'TCS UT N414 L - 2'!O" + rowNumber + "),\"\")");
                cellTotalBeschikbaarheidsRatio.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!P" + rowNumber + ",'TCS UT N414 L - 2'!P" + rowNumber + "),\"\")");
                cellTotalMatchRatio.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!Q" + rowNumber + ",'TCS UT N414 L - 2'!Q" + rowNumber + "),\"\")");
                cellTotalProductMatchRatioRegistratieratio.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!R" + rowNumber + ",'TCS UT N414 L - 2'!R" + rowNumber + "),\"\")");
                cellTotalAutoRatio.setCellFormula("IFERROR(AVERAGE('TCS UT N414 R - 1'!S" + rowNumber + ",'TCS UT N414 L - 2'!S" + rowNumber + "),\"\")");
            } catch (Exception ex) {
                System.err.println("Error in the total sheet " + ex);
            }

            /* Write the output to a file */
            try (OutputStream fileOut = new FileOutputStream(path)) {
                wb.write(fileOut);
                System.out.println("Adding data to the N414 excel sheet.");
                result = true;
            } catch (Exception e) {
                System.out.println(e.getMessage());
                //JOptionPane.showMessageDialog(null, "File open in another application. Unable to write");
                result = false;
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
            result = false;
        }

        return result;
    }

}
