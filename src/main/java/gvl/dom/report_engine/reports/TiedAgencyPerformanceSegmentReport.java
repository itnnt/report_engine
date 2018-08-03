package gvl.dom.report_engine.reports;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.MessageFormat;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gvl.dom.report_engine.ReportExtracting;
import main.utils.MySQLConnect;
import main.utils.ResultSetToExcel;
import main.utils.XLSXReadWriteHelper;

/**
 * Hello world!
 *
 */
public class TiedAgencyPerformanceSegmentReport {
	final static Logger logger = Logger.getLogger(TiedAgencyPerformanceSegmentReport.class);
	String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_template.xlsx";
	String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31-RESULT.xlsx";


	
	public void fetchDataForCountrySheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "Country";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_man_power_by_designation(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment 
			sqlcommand = MessageFormat.format("call report_segmentation_recruitment_by_designation_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// AL recruitment KPIs  
			sqlcommand = MessageFormat.format("call report_segmentation_al_recruitment_kpis_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// Rookie Performance  
			sqlcommand = MessageFormat.format("call report_segmentation_rookie_performance_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void fetchDataForNorthSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "North";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_man_power_by_designation_north(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment 
			sqlcommand = MessageFormat.format("call report_segmentation_recruitment_by_designation_north(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// AL recruitment KPIs  
			sqlcommand = MessageFormat.format("call report_segmentation_al_recruitment_kpis_tiedagency_north(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// Rookie Performance  
			sqlcommand = MessageFormat.format("call report_segmentation_rookie_performance_tiedagency_north(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void fetchDataForSouthSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "South";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_man_power_by_designation_south(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment 
			sqlcommand = MessageFormat.format("call report_segmentation_recruitment_by_designation_south(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// AL recruitment KPIs  
			sqlcommand = MessageFormat.format("call report_segmentation_al_recruitment_kpis_tiedagency_south(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// Rookie Performance  
			sqlcommand = MessageFormat.format("call report_segmentation_rookie_performance_tiedagency_south(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void fetchDataForEndingMPStructureSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Ending MP_Structure";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_man_power_by_designation_all_group(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void fetchDataForRecruitmentStructureSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Recruitment_Structure";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_recruitment_by_designation_all_group(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForRecruitmentKPIStructureSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Recruitment KPI_Structure";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_al_recruitment_kpis_all_group_level(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForRookieMetricSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Rookie Metric";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
//			CellStyle cellStyle1 = book.createCellStyle();
//			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
//			CellStyle cellStyle3 = book.createCellStyle();
//			CreationHelper createHelper = book.getCreationHelper();
//			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
//			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_rookie_performance_tiedagency_all_group(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForGASheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "GA";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
//			CellStyle cellStyle1 = book.createCellStyle();
//			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
//			CellStyle cellStyle3 = book.createCellStyle();
//			CreationHelper createHelper = book.getCreationHelper();
//			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
//			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_ga_performance(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForRiderSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Rider";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
//			CellStyle cellStyle1 = book.createCellStyle();
//			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
//			CellStyle cellStyle3 = book.createCellStyle();
//			CreationHelper createHelper = book.getCreationHelper();
//			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
//			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_rider_attach_tiedagency_all_group(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public void fetchDataForProductMixSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 0;
		final int SECTOR_APEMIX_ROWINDEX = 4;
		final int SECTOR_APERIDERMIX_ROWINDEX = 17;
		final int SECTOR_COUNTRIDERMIX_ROWINDEX = 34;
		final String SHEET_NAME = "Product Mix";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
//			CellStyle cellStyle1 = book.createCellStyle();
//			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
//			CellStyle cellStyle3 = book.createCellStyle();
//			CreationHelper createHelper = book.getCreationHelper();
//			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
//			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// report_product_mix_tiedagency
			sqlcommand = MessageFormat.format("call report_product_mix_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APEMIX_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// report_product_mix_rider_tiedagency
			sqlcommand = MessageFormat.format("call report_product_mix_rider_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APERIDERMIX_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// report_product_mix_rider_tiedagency
			sqlcommand = MessageFormat.format("call report_product_mix_count_rider_products_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_COUNTRIDERMIX_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
			logger.error(sqlcommand);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForBDSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "BD";
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
//			CellStyle cellStyle1 = book.createCellStyle();
//			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
//			CellStyle cellStyle3 = book.createCellStyle();
//			CreationHelper createHelper = book.getCreationHelper();
//			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
//			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_team_performance(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
}
