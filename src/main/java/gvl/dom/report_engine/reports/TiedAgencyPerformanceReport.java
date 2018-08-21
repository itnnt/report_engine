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
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gvl.dom.report_engine.ReportExtractingTiedAgencyPerformanceSegmentReport;
import main.utils.MySQLConnect;
import main.utils.ResultSetToExcel;
import main.utils.XLSXReadWriteHelper;

/**
 * Hello world!
 *
 */
public class TiedAgencyPerformanceReport {
	final static Logger logger = Logger.getLogger(TiedAgencyPerformanceReport.class);
	String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_template.xls";
	String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_2018-07-31-RESULT.xls";

	public void fetchDataForGASheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "6.0 GA Performance";

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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_ga_performance(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
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

	public void updateCoverSheet(String excelTemplate, String excelReport, String inputPeriodTo) {
		final String SHEET_NAME = "Cover";
		final int asAtRowIdx = 6;
		final int asAtColIdx = 5;
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;

		// format date
		SimpleDateFormat sdfMMMyy = new SimpleDateFormat("yyyy-MM-dd");

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			// CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("dd-mm-yyyy"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// reading sheet retention data
			XSSFSheet sheet = book.getSheet(SHEET_NAME);
			XSSFCell cell = sheet.getRow(asAtRowIdx).getCell(asAtColIdx);
			cell.setCellValue(sdfMMMyy.parse(inputPeriodTo));
			cell.setCellStyle(cellStyle2);

			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataApeMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_ape_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

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
	
	public void fetchDataForDataFypMomSheet(String excelTemplate, String excelReport, String sheetName, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_fyp_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

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

	public void fetchDataForDataCasecountMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_casecount_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

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

	public void fetchDataForDataNewrecruitMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_newrecruit_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

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
	
	public void fetchDataForDataManpowerMomSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_manpower_mom";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_manpower_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
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
	
	public void fetchDataForDataActiveRatioMomSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_group_active_ratio_mom";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
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
	
	public void fetchDataForDataActiveRatioYoySheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_group_active_ratio_yoy";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_alllevels_ytd(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
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
	
	public void fetchDataForDataActiveRatioSAexclMomSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_active_ratio_sa_excl_mom";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_saexcle_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
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
	public void fetchDataForDataActiveRatioSAexclYoySheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_active_ratio_sa_excl_yoy";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_sa_excluded_alllevels_ytd(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
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
	
	public void fetchDataForDataCasesizeMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_casesize_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	
	public void fetchDataForDataCaseperactiveMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_caseperactive_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	public void fetchDataForDataApeperactiveMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_apeperactive_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	
	public void fetchDataForDataActiveMomSheet(String excelTemplate, String excelReport, String sheetname, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_active_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	
	public void fetchDataForDataMPbyDesignationSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "group_manpower_by_desc_monthly";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_dynamic_man_power_by_designation_alllevels(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	public void fetchDataForDataRecruitmentSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "recruitment_monthly";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_dynamic_recruitment_by_designation_alllevels(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	public void fetchDataForDataRecruitmentActiveALSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "active_recruit_leader_monthly";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_dynamic_activealrecruit_alllevels(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
	public void fetchDataForDataRookie90daysSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_rookie90days";
		
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
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call report_dynamic_rookie_performance_tiedagency_alllevels(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
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
