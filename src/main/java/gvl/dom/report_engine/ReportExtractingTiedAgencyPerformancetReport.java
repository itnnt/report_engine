package gvl.dom.report_engine;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyPerformanceReport;

public class ReportExtractingTiedAgencyPerformancetReport {
	final static Logger logger = Logger.getLogger(ReportExtractingTiedAgencyPerformancetReport.class);
	public static void main(String[] args) {
		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_template.xlsm";
		String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_2018-07-31-RESULT.xlsm";
		
		// Ngày đầu tiên của năm liền trước năm hiện tại
		String y1 = "2017-01-01";
		
		// Ngày đầu tiên của năm hiện tại
		String y0 = "2018-01-01";
		
		// Ngày cuối cùng của năm hiện tại
		String y0End = "2018-12-31";
		
		// Ngày đầu tiên của tháng chạy report
		String m0Start = "2018-07-01";
		
		// Ngày cuối cùng của tháng chạy report
		String m0End = "2018-07-31";
		
		TiedAgencyPerformanceReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceReport();
		
		if(logger.isInfoEnabled()){
			logger.info("Updating data for Cover sheet");
		}
		tiedAgencyPerformanceSegmentReport.updateCoverSheet(excelTemplate, excelReport, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for 6.0 GA Performance sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForGASheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_ape_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataApeMomSheet(excelReport, excelReport, "data_ape_mom", m0Start, m0End, 7);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_ape_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataApeMomSheet(excelReport, excelReport, "data_ape_yoy", y0, m0End, 7);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_fyp_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataFypMomSheet(excelReport, excelReport, "data_fyp_mom", m0Start, m0End, 8);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_fyp_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataFypMomSheet(excelReport, excelReport, "data_fyp_yoy", y0, m0End, 8);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_case_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCasecountMomSheet(excelReport, excelReport, "data_case_mom", m0Start, m0End, 9);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_case_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCasecountMomSheet(excelReport, excelReport, "data_case_yoy", y0, m0End, 9);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_newrecruit_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataNewrecruitMomSheet(excelReport, excelReport, "data_newrecruit_mom", m0Start, m0End, 10);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_newrecruit_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataNewrecruitMomSheet(excelReport, excelReport, "data_newrecruit_yoy", y0, m0End, 10);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_manpower_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataManpowerMomSheet(excelReport, excelReport, m0Start, m0End, 11);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_group_active_ratio_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveRatioMomSheet(excelReport, excelReport, m0Start, m0End, 7);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_group_active_ratio_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveRatioYoySheet(excelReport, excelReport, y0, m0End, 7);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_active_ratio_sa_excl_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveRatioSAexclMomSheet(excelReport, excelReport, m0Start, m0End, 8);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_active_ratio_sa_excl_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveRatioSAexclYoySheet(excelReport, excelReport, y0, m0End, 8);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_group_casesize_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCasesizeMomSheet(excelReport, excelReport, "data_group_casesize_mom", m0Start, m0End, 9);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_group_casesize_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCasesizeMomSheet(excelReport, excelReport, "data_group_casesize_yoy", y0, m0End, 9);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_caseperactive_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCaseperactiveMomSheet(excelReport, excelReport, "data_caseperactive_mom", m0Start, m0End, 10);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_caseperactive_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataCaseperactiveMomSheet(excelReport, excelReport, "data_caseperactive_yoy", y0, m0End, 10);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_apeperactive_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataApeperactiveMomSheet(excelReport, excelReport, "data_apeperactive_mom", m0Start, m0End, 11);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_apeperactive_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataApeperactiveMomSheet(excelReport, excelReport, "data_apeperactive_yoy", y0, m0End, 11);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_active_agents_mom sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveMomSheet(excelReport, excelReport, "data_active_agents_mom", m0Start, m0End, 12);
	
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_active_agents_yoy sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataActiveMomSheet(excelReport, excelReport, "data_active_agents_yoy", y0, m0End, 12);
	}

}
