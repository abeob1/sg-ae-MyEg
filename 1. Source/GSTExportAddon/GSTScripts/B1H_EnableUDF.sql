alter system alter configuration ('indexserver.ini', 'system') set ('sqlscript', 'enable_select_into_scalar_udf') = 'true' with reconfigure
;
alter system alter configuration ('indexserver.ini', 'system') set ('sqlscript', 'sudf_support_level_select_into') = 'silent' with reconfigure
;
alter system alter configuration ('indexserver.ini', 'system') set ('sqlscript', 'dynamic_sql_ddl_error_level') = 'silent' with reconfigure
;