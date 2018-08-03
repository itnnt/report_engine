package gvl.dom.report_engine;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyPerformanceSegmentReport;

public class ReportExtracting {
	final static Logger logger = Logger.getLogger(ReportExtracting.class);
	public static void main(String[] args) {
		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_template.xlsx";
		String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31-RESULT.xlsx";
		TiedAgencyPerformanceSegmentReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceSegmentReport();
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Country sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForCountrySheet(excelTemplate, excelReport, "2018-01-01", "2018-07-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for North sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForNorthSheet(excelReport, excelReport, "2018-01-01", "2018-07-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for South sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSouthSheet(excelReport, excelReport, "2018-01-01", "2018-07-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Ending MP_Structure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForEndingMPStructureSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RecruitmentStructure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRecruitmentStructureSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RecruitmentKPIStructure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRecruitmentKPIStructureSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RookieMetric sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRookieMetricSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for GA sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForGASheet(excelReport, excelReport, "2018-01-01", "2018-07-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Rider sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRiderSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for ProductMix sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixSheet(excelReport, excelReport, "2017-01-01", "2018-12-31");
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for BD sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForBDSheet(excelReport, excelReport, "2018-01-01", "2018-07-31");
	}

}
