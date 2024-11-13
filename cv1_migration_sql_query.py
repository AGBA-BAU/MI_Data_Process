source_query_1 = """
   --active_case_category_all
    select name, type, status, source, NULL RECORD_COUNT, NULL DATA_AS_OF, NULL size, NULL code, NULL message
    from [HUB_LIVE].[HUB_LIVE].[dbo].[case_category]
    where status = 'active'
"""
destination_table_1 = "[MIS_Report].[dbo].[Dwh_ApiFindActiveCaseCategoryAll_tmp]"

source_query_2 = """
   --case
    select 
	a.case_id, a.case_number, a.customer_name, a.created_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.created_date) created_date, NULL consultantEmail, a.status,
	caseT.name case_type, planT.name plan_type, companyT.name company_type, a.locked_by, principalT.name principal_type,
	a.reference_number, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.updated_date) updated_date, a.language_preference, NULL RECORD_COUNT, NULL DATA_AS_OF, NULL size, NULL code, NULL masssage
from [HUB_LIVE].[HUB_LIVE].[dbo].[case] a
left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] caseT on caseT.type = 'case' and caseT.case_category_id = a.case_type_id
left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] companyT on companyT.type = 'company' and companyT.case_category_id = a.company_id
left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] planT on planT.type = 'plan' and planT.case_category_id = a.plan_type_id
left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] principalT on principalT.type = 'principal' and principalT.case_category_id = a.principal_type_id
 
"""
destination_table_2 = "[MIS_Report].[dbo].[Dwh_ApiFindCase_tmp]"

source_query_3 = """
--case_revision
   select
	a.version, b.case_number, a.ticket_number, a.status, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.created_date) created_date, a.created_by, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.sorted_date) sorted_date, a.sorted_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.maker_date) maker_date, 
	a.maker_by, a.maker_result, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.approved_date) approved_date, a.approved_by, a.approved_result, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.approver_to_maker_date) approver_to_maker_date, a.matched_by, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.matched_date) matched_date, a.original_document_number, a.processed_document_number, a.final_document_number, 
	a.comments_for_consultant, a.comments_for_staff, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.updated_date) updated_date, NULL RECORD_COUNT, NULL DATA_AS_OF, NULL size, NULL code, NULL message
from [HUB_LIVE].[HUB_LIVE].[dbo].[case_revision] a
left join [HUB_LIVE].[HUB_LIVE].[dbo].[case] b on b.case_id = a.case_id
 
"""
destination_table_3 = "[MIS_Report].[dbo].[Dwh_ApiFindCaseRevision_tmp]"


source_query_4 = """
--case_group
select a.case_group_id, a.name, a.case_type_ids, a.plan_type_ids, a.company_ids, NULL principal_ids, NULL principal_types, 
	caseN.case_type, planN.plan_type, companyN.company_type, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.cut_off_time) cut_off_time, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.time_limit) time_limit, a.time_limit_color, a.consultant_codes, a.status, a.created_by, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.created_date) created_date, a.updated_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', a.updated_date) updated_date, NULL size, NULL code, NULL message
from [HUB_LIVE].[HUB_LIVE].[dbo].[case_group_tmp] a
left join (
	select case_group_id, '["' + STRING_AGG(c.name, '", "') + '"]' case_type
	from (
		select case_group_id, trim(SplitStr.value) [case_type_ids]
		from [HUB_LIVE].[HUB_LIVE].[dbo].[case_group_tmp] a
		CROSS APPLY STRING_SPLIT(a.case_type_ids, ',') as SplitStr 
	) a
	left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] c on c.case_category_id = a.case_type_ids
	group by case_group_id
) caseN on caseN.case_group_id = a.case_group_id
left join (
	select case_group_id, '["' + STRING_AGG(c.name, '", "') + '"]' plan_type
	from (
		select case_group_id, trim(SplitStr.value) [plan_type_ids]
		from [HUB_LIVE].[HUB_LIVE].[dbo].[case_group_tmp] a
		CROSS APPLY STRING_SPLIT(a.plan_type_ids, ',') as SplitStr 
	) a
	left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] c on c.case_category_id = a.plan_type_ids
	group by case_group_id
) planN on planN.case_group_id = a.case_group_id
left join (
	select case_group_id, '["' + STRING_AGG(c.name, '", "') + '"]' company_type
	from (
		select case_group_id, trim(SplitStr.value) [company_ids]
		from [HUB_LIVE].[HUB_LIVE].[dbo].[case_group_tmp] a
		CROSS APPLY STRING_SPLIT(a.company_ids, ',') as SplitStr 
	) a
	left join [HUB_LIVE].[HUB_LIVE].[dbo].[case_category] c on c.case_category_id = a.company_ids
	group by case_group_id
) companyN on companyN.case_group_id = a.case_group_id
"""
destination_table_4 = "[MIS_Report].[dbo].[Dwh_ApiFindCaseGroup_tmp]"


source_query_5 = """
--ticket
select t.ticket_id, NULL [counter], q.[description], l.name [location], q.name [queue], NULL [queue_description], q.status queue_status,
	t.ticket_number, t.form_amount, t.status, DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.start_date) start_date, DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.end_date) end_date, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.waiting_date) waiting_date, t.handled_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.created_date) created_date, t.created_by, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.updated_date) updated_date, t.updated_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', t.voided_date) voided_date, t.voided_by, 
	NULL RECORD_COUNT, NULL size, NULL code, NULL message
from [ET_LIVE].[ET_LIVE].[dbo].[ticket] t
left join [ET_LIVE].[ET_LIVE].[dbo].[queue] q on q.queue_id = t.queue_id
left join [ET_LIVE].[ET_LIVE].[dbo].[location] l on q.location_id = l.location_id
"""
destination_table_5 = "[MIS_Report].[dbo].[Dwh_ApiFindTicket_tmp]"


source_query_6 = """
--counter_record
select
	c.application_type, c.application_type_detail, c.application_type_remark, c.barcode, c.client_given_name, c.client_surname, c.created_by, 
	DATEDIFF_BIG(MILLISECOND, '1970-01-01', c.created_date) created_date, DATEDIFF_BIG(MILLISECOND, '1970-01-01', c.in_date) in_date, c.is_club, 
	c.is_reviewed, c.submission_channel, l.name [location], c.msv, c.number_of_error, DATEDIFF_BIG(MILLISECOND, '1970-01-01', c.out_date) out_date, 
	c.plan_name, c.principal, DATEDIFF_BIG(MILLISECOND, '1970-01-01', c.processing_date) processing_date, c.product_type, c.staff_code, c.status, 
	c.submission_version, c.updated_by, DATEDIFF_BIG(MILLISECOND, '1970-01-01', c.updated_date) updated_date, NULL size, NULL code, NULL message
from [ET_LIVE].[ET_LIVE].[dbo].[counter_record] c
left join [ET_LIVE].[ET_LIVE].[dbo].[location] l on l.location_id = c.location_id
"""
destination_table_6 = "[MIS_Report].[dbo].[Dwh_ApiFindCounterRecord_tmp]"

