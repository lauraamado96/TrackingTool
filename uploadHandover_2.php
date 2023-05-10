<?php
$nom_file = "EMA_Pending_Report_Handover_2.txt"; 
echo move_uploaded_file( 
	$_FILES["upFile"]["tmp_name"],
	$nom_file
) ? "OK" : "ERROR";
?>
