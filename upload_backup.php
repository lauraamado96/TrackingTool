<?php
$today = date("Ymd_Hi");
$nom_backup = "backups/EMA_Pending_Report_".$today."_backup.xls"; 
echo move_uploaded_file( 
	$_FILES["upFile"]["tmp_name"],
	$nom_backup
) ? "OK" : "ERROR";
?>
