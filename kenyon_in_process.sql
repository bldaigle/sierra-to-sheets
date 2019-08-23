SELECT
	concat('i', rm.record_num, 'a') AS "Item Record Number",
	brp.best_title AS "Title",
	irp.barcode AS "Barcode",
	(TRIM(REGEXP_REPLACE(irp.call_number,'\|.',' ','g'))) AS "Call Number",
	TO_CHAR(rm.creation_date_gmt AT TIME ZONE 'EST', 'YYYY-MM-DD') AS "Date Created",
	TO_CHAR(rm.record_last_updated_gmt AT TIME ZONE 'EST', 'YYYY-MM-DD') AS "Date Updated"
FROM
	sierra_view.item_record_property irp
INNER JOIN
	sierra_view.item_record ir ON irp.item_record_id = ir.record_id
INNER JOIN
	sierra_view.bib_record_item_record_link bil ON irp.item_record_id = bil.item_record_id
INNER JOIN
	sierra_view.bib_record_property brp ON bil.bib_record_id = brp.bib_record_id
INNER JOIN
	sierra_view.record_metadata rm ON irp.item_record_id = rm.id
WHERE
	ir.location_code LIKE 'k%' AND
	ir.item_status_code = 'p'
ORDER BY
	"Date Updated" ASC
LIMIT 1000;