SELECT
	CONCAT(order_view.record_type_code, order_view.record_num, 'a') AS "Order Record Number",
	CONCAT(bib_view.record_type_code, bib_view.record_num, 'a') AS "Bib Record Number",
	bib_record_property.best_title AS "Title",
	order_record.vendor_record_code AS "Vendor",
	TO_CHAR(record_metadata.creation_date_gmt, 'YYYY-MM-DD') AS "Created Date",
	TO_CHAR(record_metadata.record_last_updated_gmt, 'YYYY-MM-DD') AS "Updated Date"
FROM
	sierra_view.order_record
INNER JOIN
	sierra_view.order_view ON sierra_view.order_record.record_id = sierra_view.order_view.id
INNER JOIN
	sierra_view.bib_record_order_record_link ON sierra_view.order_record.record_id = sierra_view.bib_record_order_record_link.order_record_id
INNER JOIN
	sierra_view.bib_record_property ON sierra_view.bib_record_order_record_link.bib_record_id = sierra_view.bib_record_property.bib_record_id
INNER JOIN
	sierra_view.bib_view ON sierra_view.bib_record_property.bib_record_id = sierra_view.bib_view.id
INNER JOIN
	sierra_view.record_metadata ON sierra_view.order_record.record_id = sierra_view.record_metadata.id
WHERE
	order_record.accounting_unit_code_num = 2 AND
	order_record.order_status_code = 'o'
ORDER BY
	order_view.record_num ASC
LIMIT 1000;