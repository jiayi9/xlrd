-- old version
CREATE TABLE reports_MOE6 (
	ANALYSIS_NO varchar(255),
	EXTRACTION_DEVICE varchar(255),
	TEST_SPEC varchar(255),
	STANDARD varchar(255),
	COMPONENT_NO varchar(255),
	BATCH_NO varchar(255),
	MEDIA_SAMPLE varchar(255),
	OTHERS varchar(255),
	SAMPLE_NO varchar(255),
	SHIPPING_PACKAGE varchar(255),
	EXTRACTION_DATE varchar(255),
	FILTERING_DRYING varchar(255),
	MICROSCOPE varchar(255),
	SOFTWARE_VERSION varchar(255),
	CALIBRATION varchar(255),
	FILTER_SIZE varchar(255),
	ANALYZE_SIZE varchar(255),
	THRESHOLD_LEVEL varchar(255),
	POPULATION_DENSITY varchar(255),
	RESULT varchar(255),
	INSPECTOR varchar(255),
	DATETIME datetime,
	COMMENT varchar(255),
	REMARK varchar(255),
	AREA varchar(255),
	FILE_NAME varchar(255),
	FILE_PATH varchar(255)
);


-- new


CREATE TABLE report_moe3 (
	ANALYSIS_NO nvarchar(255),
	EXTRACTION_DEVICE nvarchar(255),
	TEST_SPEC nvarchar(255),
	STANDARD nvarchar(255),
	COMPONENT_NO nvarchar(255),
	BATCH_NO nvarchar(255),
	MEDIA_SAMPLE nvarchar(255),
	OTHERS nvarchar(255),
	SAMPLE_NO nvarchar(255),
	SHIPPING_PACKAGE nvarchar(255),
	EXTRACTION_DATE nvarchar(255),
	FILTERING_DRYING nvarchar(255),
	MICROSCOPE nvarchar(255),
	SOFTWARE_VERSION nvarchar(255),
	CALIBRATION nvarchar(255),
	FILTER_SIZE nvarchar(255),
	ANALYZE_SIZE nvarchar(255),
	THRESHOLD_LEVEL nvarchar(255),
	POPULATION_DENSITY nvarchar(255),
	RESULT nvarchar(255),
	INSPECTOR nvarchar(255),
	DATETIME datetime,
	COMMENT nvarchar(255),
	REMARK nvarchar(255),
	AREA nvarchar(255),
	FILE_NAME nvarchar(255),
	FILE_PATH nvarchar(255)
);


CREATE TABLE report_moe6 (
	ANALYSIS_NO nvarchar(255),
	EXTRACTION_DEVICE nvarchar(255),
	TEST_SPEC nvarchar(255),
	STANDARD nvarchar(255),
	COMPONENT_NO nvarchar(255),
	BATCH_NO nvarchar(255),
	MEDIA_SAMPLE nvarchar(255),
	OTHERS nvarchar(255),
	SAMPLE_NO nvarchar(255),
	SHIPPING_PACKAGE nvarchar(255),
	EXTRACTION_DATE nvarchar(255),
	FILTERING_DRYING nvarchar(255),
	MICROSCOPE nvarchar(255),
	SOFTWARE_VERSION nvarchar(255),
	CALIBRATION nvarchar(255),
	FILTER_SIZE nvarchar(255),
	ANALYZE_SIZE nvarchar(255),
	THRESHOLD_LEVEL nvarchar(255),
	POPULATION_DENSITY nvarchar(255),
	RESULT nvarchar(255),
	INSPECTOR nvarchar(255),
	DATETIME datetime,
	COMMENT nvarchar(255),
	REMARK nvarchar(255),
	AREA nvarchar(255),
	FILE_NAME nvarchar(255),
	FILE_PATH nvarchar(255)
);


create table result_moe3 (
	ANALYSIS_NO nvarchar(255),
	STANDARD nvarchar(255),
	ITEM nvarchar(255),
	SIZE nvarchar(255),
	VALUE float
);

create table result_moe6 (
	ANALYSIS_NO nvarchar(255),
	STANDARD nvarchar(255),
	ITEM nvarchar(255),
	SIZE nvarchar(255),
	VALUE float
);


CREATE VIEW reports_rbcd AS  
SELECT ANALYSIS_NO, STANDARD, COMPONENT_NO, DATETIME, RESULT, REMARK FROM report_moe3
UNION ALL  
SELECT ANALYSIS_NO, STANDARD, COMPONENT_NO, DATETIME, RESULT, REMARK FROM report_moe6;


create view results_rbcd as
select * from result_moe3
union all
select * from result_moe6;







