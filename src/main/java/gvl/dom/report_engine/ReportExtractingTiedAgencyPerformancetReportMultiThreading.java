package gvl.dom.report_engine;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.MessageFormat;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyPerformanceReport;
import main.utils.MySQLConnect;

public class ReportExtractingTiedAgencyPerformancetReportMultiThreading {
	final static Logger logger = Logger.getLogger(ReportExtractingTiedAgencyPerformancetReportMultiThreading.class);
	public static void main(String[] args) {
		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_template.xlsm";
		final String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_2018-07-31-RESULT.xlsm";
		
		// Ngày đầu tiên của năm liền trước năm hiện tại
		String y1 = "2017-01-01";
		
		// Ngày cuối cùng của năm liền trước năm hiện tại
		String y1End = "2017-12-31";
		
		// Ngày đầu tiên của năm hiện tại
		String y0 = "2018-01-01";
		
		// Ngày cuối cùng của năm hiện tại
		String y0End = "2018-12-31";
		
		// Ngày đầu tiên của tháng chạy report
		final String m0Start = "2018-07-01";
		
		// Ngày cuối cùng của tháng chạy report
		final String m0End = "2018-07-31";
		
		final TiedAgencyPerformanceReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceReport();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Updating data for Cover sheet");
		}
		tiedAgencyPerformanceSegmentReport.updateCoverSheet(excelTemplate, excelReport, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		Runnable r1 = new Runnable() {
			public void run() {
				if (logger.isInfoEnabled()) {
					logger.info("Fetching data for data_detail_in_segment_ape sheet");
				}
				try {
					MySQLConnect mySQLConnect = null;
					String sqlcommand = null;
					ResultSet rs = null;
					// fetch data from the database
					mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
					mySQLConnect.connect(true);
					// ending manpower
					sqlcommand = MessageFormat.format("call tiedagency_ape_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
					rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
					tiedAgencyPerformanceSegmentReport.writeDataForSheet(excelReport, excelReport, "data_detail_in_segment_ape", -1, 0, rs);
					mySQLConnect.close();
				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		};
		Thread thread1 = new Thread(r1);
		thread1.start();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		Runnable r2 = new Runnable() {
			public void run() {
				if (logger.isInfoEnabled()) {
					logger.info("Fetching data for data_detail_in_segment_manpower sheet");
				}
				try {
					MySQLConnect mySQLConnect = null;
					String sqlcommand = null;
					ResultSet rs = null;
					// fetch data from the database
					mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
					mySQLConnect.connect(true);
					// ending manpower
					sqlcommand = MessageFormat.format("call tiedagency_mp_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
					rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
					tiedAgencyPerformanceSegmentReport.writeDataForSheet(excelReport, excelReport, "data_detail_in_segment_manpower", -1, 0, rs);
					mySQLConnect.close();
				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		};
		Thread thread2 = new Thread(r2);
		thread2.start();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		Runnable r3 = new Runnable() {
			public void run() {
				if (logger.isInfoEnabled()) {
					logger.info("Fetching data for data_detail_in_segment_cscnt sheet");
				}
				try {
					MySQLConnect mySQLConnect = null;
					String sqlcommand = null;
					ResultSet rs = null;
					// fetch data from the database
					mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
					mySQLConnect.connect(true);
					// ending manpower
					sqlcommand = MessageFormat.format("call tiedagency_casecount_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
					rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
					tiedAgencyPerformanceSegmentReport.writeDataForSheet(excelReport, excelReport, "data_detail_in_segment_cscnt", -1, 0, rs);
					mySQLConnect.close();
				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		};
		Thread thread3 = new Thread(r3);
		thread3.start();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		Runnable r4 = new Runnable() {
			public void run() {
				if (logger.isInfoEnabled()) {
					logger.info("Fetching data for data_detail_in_segment_active sheet");
				}
				try {
					MySQLConnect mySQLConnect = null;
					String sqlcommand = null;
					ResultSet rs = null;
					// fetch data from the database
					mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
					mySQLConnect.connect(true);
					// ending manpower
					sqlcommand = MessageFormat.format("call tiedagency_active_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
					rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
					tiedAgencyPerformanceSegmentReport.writeDataForSheet(excelReport, excelReport, "data_detail_in_segment_active", -1, 0, rs);
					mySQLConnect.close();
				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		};
		Thread thread4 = new Thread(r4);
		thread4.start();
		/* --------------------------------------------------------------------------------- */
	}

}
