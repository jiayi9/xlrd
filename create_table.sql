CREATE TABLE Reports (
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



create table Results (
	ANALYSIS_NO varchar(255),
	STANDARD varchar(255),
	ITEM varchar(255),
	SIZE varchar(255),
	VALUE float
);


drop table Reports;
drop table Results;





drop view Reports_view;

create view Reports_view as
select ANALYSIS_NO, 
STANDARD,
replace(replace(component_no, '-', ''),' ', '') as COMPONENT_NO,
DATETIME,
REMARK 
from Reports
where STANDARD like '%[^P]'
;

select * from Reports_view;

select * from Reports_view
