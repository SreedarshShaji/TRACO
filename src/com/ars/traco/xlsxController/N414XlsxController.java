/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.traco.xlsxController;

import java.text.DateFormatSymbols;
import java.text.DecimalFormat;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import com.ars.traco.databeans.n414.Sectie;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Iterator;

import javax.swing.JOptionPane;

/**
 *
 * @author Sreedarshs
 */
public class N414XlsxController {

	public boolean handleXlsx(String inputFile,Sectie bean) {

		boolean result=false;
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
			boolean writingStatus=writingToFile(directory + "\\" + fileNameToWrite,bean);
			if(writingStatus)
			{
				System.out.println("File written to successfully");
				result=true;
			}
			else
			{
				System.out.println("File writing failed");
				result=false;
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

	public boolean writingToFile(String path,Sectie bean) {
		boolean result = false;

		try (InputStream inp = new FileInputStream(path)) {
			/* Locating the workbook*/
			Workbook wb = WorkbookFactory.create(inp);
			
			/* Gets date from the bean*/ 
			String date=bean.getDate();
			int day=Integer.parseInt(date.substring(date.lastIndexOf("-")+1));
		    System.out.println(day);

/*****************************************************************************************************/
/******************************************TCS UT N414 R - 1******************************************/
/*****************************************************************************************************/		

			/*Reading R1 the sheet*/
			Sheet R1sheet = wb.getSheet("TCS UT N414 R - 1");
			
			/* Locating the row and column */
			Iterator<Row> iterator = R1sheet.iterator();
			int rowNumber=0;
			
			while (iterator.hasNext()) {
				rowNumber++;
				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				
					 Cell currentCell = cellIterator.next();
					 if (currentCell.getCellTypeEnum() == CellType.STRING) {
	                         continue;
	                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
	                        if((int)currentCell.getNumericCellValue() ==day)
	                        {
	                        	break;
	                        }
	                    }
					 
			}
			
			 System.out.println("Row Count :"+rowNumber);
			
			Row rowToInsert = R1sheet.getRow(rowNumber-1);
			//Cell cellToInsert = rowToInsert.getCell(1);
			
			/*
			 * Defining cells to populate the value
			 * */

			Cell cellR1Begin = rowToInsert.getCell(1);
			Cell cellR1End = rowToInsert.getCell(2);
			Cell cellR1Max = rowToInsert.getCell(3);
			Cell cellR1Aantal= rowToInsert.getCell(4);
			Cell cellR1Gem_kmpuur= rowToInsert.getCell(5);
			Cell cellR1Max_kmpuur= rowToInsert.getCell(6);
			Cell cellR1Total= rowToInsert.getCell(7);
			Cell cellR1Hand= rowToInsert.getCell(8);
			Cell cellR1Auto= rowToInsert.getCell(9);
			Cell cellR1Dubbele_overtredingen_pardon= rowToInsert.getCell(10);
			Cell cellR1Overig_pardon= rowToInsert.getCell(11);
			Cell cellR1OvertredingenRatio= rowToInsert.getCell(12);
			Cell cellR1Handhaafratio= rowToInsert.getCell(13);
			Cell cellR1TijdVolledigBeschikbaar_in_minuten= rowToInsert.getCell(14);
			Cell cellR1BeschikbaarheidsRatio=rowToInsert.getCell(15);
			Cell cellR1MatchRatio= rowToInsert.getCell(16);
			Cell cellR1ProductMatchRatioRegistratieratio= rowToInsert.getCell(17);
			Cell cellR1AutoRatio= rowToInsert.getCell(18);


			
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
			cellR1Max.setCellFormula("IF(MAX(B"+rowNumber+":C"+rowNumber+")=0,\"\",MAX(B"+rowNumber+":C"+rowNumber+"))");
			cellR1Aantal.setCellValue(bean.getR1().getMatches());
			cellR1Gem_kmpuur.setCellValue(bean.getR1().getsnelhedenType().getGemiddeld());
			cellR1Max_kmpuur.setCellValue(bean.getR1().getsnelhedenType().getMax());
			
			cellR1Total.setCellValue(bean.getR1().getOvertredingenType().getOvertredingenTotaal());
			cellR1Hand.setCellValue(bean.getR1().getOvertredingenType().getHand());
			cellR1Auto.setCellValue(bean.getR1().getOvertredingenType().getAuto());
			cellR1Dubbele_overtredingen_pardon.setCellValue(bean.getR1().getOvertredingenType().getDubbeleOvertredingenPardon());
			cellR1Overig_pardon.setCellValue(bean.getR1().getOvertredingenType().getOverigPardon());
			cellR1OvertredingenRatio.setCellValue((bean.getR1().getPerformanceType().getOvertredingenratio()*100)+"%");
			cellR1Handhaafratio.setCellValue((bean.getR1().getPerformanceType().getHandhaafratio()*100)+"%");
			

                        String tijdvolledigbeschiPerkbaarR1=bean.getR1().getPerformanceType().getTijdvolledigbeschikbaar();
			int daysR1=Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(1,tijdvolledigbeschiPerkbaarR1.indexOf('D')));
			int hoursR1=Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(tijdvolledigbeschiPerkbaarR1.indexOf('T')+1,tijdvolledigbeschiPerkbaarR1.indexOf('H')));
			int minutesR1=Integer.parseInt(tijdvolledigbeschiPerkbaarR1.substring(tijdvolledigbeschiPerkbaarR1.indexOf('H')+1,tijdvolledigbeschiPerkbaarR1.indexOf('M')));
			
			cellR1TijdVolledigBeschikbaar_in_minuten.setCellValue(daysR1*24*60 + hoursR1*60 + minutesR1);
			
			cellR1BeschikbaarheidsRatio.setCellFormula("IF((O"+rowNumber+"/1440)=0,\"\",(O"+rowNumber+"/1440))");
			
			Double matchRatio=(bean.getR1().getPerformanceType().getMatchratio()*100);
			matchRatio=BigDecimal.valueOf(matchRatio).setScale(2,RoundingMode.HALF_UP).doubleValue();
			cellR1MatchRatio.setCellValue((matchRatio)+"%");
			
			cellR1ProductMatchRatioRegistratieratio.setCellFormula("IF((Q"+rowNumber+"*'Total Systemperformance'!$C$6)=0,\"\",(Q"+rowNumber+"*'Total Systemperformance'!$C$6))");
			
			cellR1AutoRatio.setCellValue((bean.getR1().getPerformanceType().getAutoratio()*100)+"%");

			

/*****************************************************************************************************/
/******************************************TCS UT N414 L - 2******************************************/
/*****************************************************************************************************/		

			/*Reading R1 the sheet*/
			Sheet L2sheet = wb.getSheet("TCS UT N414 L - 2");
			
			/* Locating the row and column */
			Iterator<Row> iteratorL2 = L2sheet.iterator();
			int rowNumberL2=0;
			
			while (iterator.hasNext()) {
				rowNumberL2++;
				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				
					 Cell currentCell = cellIterator.next();
					 if (currentCell.getCellTypeEnum() == CellType.STRING) {
	                         continue;
	                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
	                        if((int)currentCell.getNumericCellValue() ==day)
	                        {
	                        	break;
	                        }
	                    }
					 
			}
			
			 System.out.println("L2Row Count :"+rowNumberL2);
			
			Row rowToInsertL2 = L2sheet.getRow(rowNumber-1);
			//Cell cellToInsert = rowToInsert.getCell(1);
			
			/*
			 * Defining cells to populate the value
			 * */

			Cell cellL2Begin = rowToInsertL2.getCell(1);
			Cell cellL2End = rowToInsertL2.getCell(2);
			Cell cellL2Max = rowToInsertL2.getCell(3);
			Cell cellL2Aantal= rowToInsertL2.getCell(4);
			Cell cellL2Gem_kmpuur= rowToInsertL2.getCell(5);
			Cell cellL2Max_kmpuur= rowToInsertL2.getCell(6);
			Cell cellL2Total= rowToInsertL2.getCell(7);
			Cell cellL2Hand= rowToInsertL2.getCell(8);
			Cell cellL2Auto= rowToInsertL2.getCell(9);
			Cell cellL2Dubbele_overtredingen_pardon= rowToInsertL2.getCell(10);
			Cell cellL2Overig_pardon= rowToInsertL2.getCell(11);
			Cell cellL2OvertredingenRatio= rowToInsertL2.getCell(12);
			Cell cellL2Handhaafratio= rowToInsertL2.getCell(13);
			Cell cellL2TijdVolledigBeschikbaar_in_minuten= rowToInsertL2.getCell(14);
			Cell cellL2BeschikbaarheidsRatio=rowToInsertL2.getCell(15);
			Cell cellL2MatchRatio= rowToInsertL2.getCell(16);
			Cell cellL2ProductMatchRatioRegistratieratio= rowToInsertL2.getCell(17);
			Cell cellL2AutoRatio= rowToInsertL2.getCell(18);


			
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
                        cellL2Max.setCellFormula("IF(MAX(B"+rowNumber+":C"+rowNumber+")=0,\"\",MAX(B"+rowNumber+":C"+rowNumber+"))");
			cellL2Aantal.setCellValue(bean.getL2().getMatches());
			cellL2Gem_kmpuur.setCellValue(bean.getL2().getSnelheden().getGemiddeld());
			cellL2Max_kmpuur.setCellValue(bean.getL2().getSnelheden().getMax());
			
			cellL2Total.setCellValue(bean.getL2().getOvertredingenType().getOvertredingenTotaal());
			cellL2Hand.setCellValue(bean.getL2().getOvertredingenType().getHand());
			cellL2Auto.setCellValue(bean.getL2().getOvertredingenType().getAuto());
			cellL2Dubbele_overtredingen_pardon.setCellValue(bean.getL2().getOvertredingenType().getDubbeleOvertredingenPardon());
			cellL2Overig_pardon.setCellValue(bean.getL2().getOvertredingenType().getOverigPardon());
			
			Double L2OvertredingenRatio=(bean.getL2().getPerformanceType().getOvertredingenratio()*100);
			L2OvertredingenRatio=BigDecimal.valueOf(L2OvertredingenRatio).setScale(2,RoundingMode.HALF_UP).doubleValue();
			cellL2OvertredingenRatio.setCellValue(L2OvertredingenRatio/100);
			/* Temporary patch, needs a further look */
			
			cellL2Handhaafratio.setCellValue((bean.getL2().getPerformanceType().getHandhaafratio()*100)+"%");
			

	    	String tijdvolledigbeschiPerkbaarL2=bean.getL2().getPerformanceType().getTijdvolledigbeschikbaar();
			int daysL2=Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(1,tijdvolledigbeschiPerkbaarL2.indexOf('D')));
			int hoursL2=Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(tijdvolledigbeschiPerkbaarL2.indexOf('T')+1,tijdvolledigbeschiPerkbaarL2.indexOf('H')));
			int minutesL2=Integer.parseInt(tijdvolledigbeschiPerkbaarL2.substring(tijdvolledigbeschiPerkbaarL2.indexOf('H')+1,tijdvolledigbeschiPerkbaarL2.indexOf('M')));
			
			cellL2TijdVolledigBeschikbaar_in_minuten.setCellValue(daysL2*24*60 + hoursL2*60 + minutesL2);
			
			cellL2BeschikbaarheidsRatio.setCellFormula("IF((O"+rowNumber+"/1440)=0,\"\",(O"+rowNumber+"/1440))");
			
			Double matchRatioL2=(bean.getL2().getPerformanceType().getMatchratio()*100);
			matchRatioL2=BigDecimal.valueOf(matchRatioL2).setScale(2,RoundingMode.HALF_UP).doubleValue();

			cellL2MatchRatio.setCellValue(matchRatioL2+"%");
			
			cellL2ProductMatchRatioRegistratieratio.setCellFormula("IF((Q"+rowNumber+"*'Total Systemperformance'!$C$6)=0,\"\",(Q"+rowNumber+"*'Total Systemperformance'!$C$6))");
			
			cellL2AutoRatio.setCellValue((bean.getL2().getPerformanceType().getAutoratio()*100)+"%");

			/* Write the output to a file */
			try (OutputStream fileOut = new FileOutputStream(path)) {
				wb.write(fileOut);
				System.out.println("Adding data to the N414 excel sheet.");
				result=true;
			} catch (Exception e) {
				System.out.println(e.getMessage());
				//JOptionPane.showMessageDialog(null, "File open in another application. Unable to write");
				result=false;
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
			result=false;
		}

		return result;
	}
	
}
