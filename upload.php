<?php
echo move_uploaded_file( 
	$_FILES["upFile"]["tmp_name"],
	"EMA_Pending_Report.xls"
) ? "OK" : "ERROR";
?>
