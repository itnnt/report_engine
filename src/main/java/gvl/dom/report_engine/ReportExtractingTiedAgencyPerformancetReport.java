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
		
		// Ngày cuối cùng của năm liền trước năm hiện tại
		String y1End = "2017-12-31";
		
		// Ngày đầu tiên của năm hiện tại
		String y0 = "2018-01-01";
		
		// Ngày cuối cùng của năm hiện tại
		String y0End = "2018-12-31";
		
		// Ngày đầu tiên của tháng chạy report
		String m0Start = "2018-07-01";
		
		// Ngày cuối cùng của tháng chạy report
		String m0End = "2018-07-31";
		
		TiedAgencyPerformanceReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceReport();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Updating data for Cover sheet");
		}
		tiedAgencyPerformanceSegmentReport.updateCoverSheet(excelTemplate, excelReport, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_in_segment_ape sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForAPESegDtlSheet(excelReport, excelReport, m0Start, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_in_segment_manpower sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForMPSegDtlSheet(excelReport, excelReport, m0Start, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_in_segment_cscnt sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForCasecountSegDtlSheet(excelReport, excelReport, m0Start, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_in_segment_active sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForActiveSegDtlSheet(excelReport, excelReport, m0Start, m0End);
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_seg_chart_y0_ape sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_detail_seg_chart_y0_ape(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_seg_chart_y0_mp sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_detail_seg_chart_y0_mp(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_seg_chart_y0_case sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_detail_seg_chart_y0_case(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_detail_seg_chart_y0_act sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_detail_seg_chart_y0_act(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_manpower_monthly_total sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_group_manpower_monthly_total(excelReport, excelReport, "group_manpower_monthly_total",y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_manpower_monthly_total_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_group_manpower_monthly_total(excelReport, excelReport, "group_manpower_monthly_total_y1",y1, y1End);

		if(logger.isInfoEnabled()){
			logger.info("Fetching data for newrecruit_monthly_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_recruitment_total(excelReport, excelReport, "newrecruit_monthly_y1",y1, y1End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for newrecruit_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_recruitment_total(excelReport, excelReport, "newrecruit_monthly",y0, m0End);
	
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_ape_y0 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_ape_total(excelReport, excelReport, "data_ape_y0",y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_ape_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_ape_total(excelReport, excelReport, "data_ape_y1",y1, y1End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_active_ratio_monthly_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_ar_total(excelReport, excelReport, "group_active_ratio_monthly_y1",y1, y1End);

		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_active_ratio_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_ar_total(excelReport, excelReport, "group_active_ratio_monthly",y0, m0End);
	
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_casesize_monthly_12ms_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_casesize_total(excelReport, excelReport, "group_casesize_monthly_12ms_y1",y1, y1End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_casesize_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_casesize_total(excelReport, excelReport, "group_casesize_monthly",y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for case_per_active_monthly_12ms_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_caseperactive_total(excelReport, excelReport, "case_per_active_monthly_12ms_y1",y1, y1End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for case_per_active_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_caseperactive_total(excelReport, excelReport, "case_per_active_monthly",y0, m0End);
		
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for case_per_active_monthly_12ms_y1 sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_apeperactive_total(excelReport, excelReport, "ape_per_active_monthly_12ms_y1",y1, y1End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for case_per_active_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSheet_data_apeperactive_total(excelReport, excelReport, "ape_per_active_monthly",y0, m0End);
		
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

		if(logger.isInfoEnabled()){
			logger.info("Fetching data for group_manpower_by_desc_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataMPbyDesignationSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for recruitment_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataRecruitmentSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for active_recruit_leader_monthly sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataRecruitmentActiveALSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_rookie90days sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForDataRookie90daysSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_product_mix_ape sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixAPESheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_product_rider_mix sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixRiderSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_product_mix_total sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixTotalSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for data_product_rider_mix_count sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixCountSheet(excelReport, excelReport, y0, m0End);
		
	}

}
