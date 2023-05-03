// Función que lanza una caja de texto antes de cerrar pestaña, ventana o cargar la página para recordarte que si no exportas los datos actuales se perderán	
function closeIt() {

    return "Any string value here forces a dialog box to \n" +
        "appear before closing the window.";
}
window.onbeforeunload = closeIt;

// Función de la fecha de hoy
function getToday() {

    // Fecha completa
    var date = new Date();

    // Año
    var year = date.getFullYear();

    // Mes
    var month = date.getMonth() + 1;
    if (Number(month) < 10) {
        month = "0" + month;
    }

    // Día
    var day = date.getDate();
    if (Number(day) < 10) {
        day = "0" + day;
    }

    // Hora
    var hours = date.getHours();
    if (Number(hours) < 10) {
        hours = "0" + hours;
    }

    // Minutos
    var minutes = date.getMinutes();
    if (Number(minutes) < 10) {
        minutes = "0" + minutes;
    }

    // Dia de la semana
    var dayWeek = date.getDay();

    return [day, month, year, hours, minutes, dayWeek, date];
}

// Función de la fecha de hoy
function todayFunction() {

    // Fecha completa
    var today = getToday();
    var day = today[0],
        month = today[1],
        year = today[2],
        hours = today[3],
        minutes = today[4];

    // Fecha de hoy con formato personalizado
    var date_of_today = day + "/" + month + "/" + year + " " + hours + ":" + minutes;
    document.getElementById("current_date").innerHTML = date_of_today;

    // Lanzamos la función cada segundo para que se vayan actualizando los minutos
    setTimeout("todayFunction()", 1000);

}

// Función que carga el archivo EMA_Pending_Report.xls guardado en mcsprod
function mcsprodFile() {

    var req = new XMLHttpRequest();
    req.open("GET", "EMA_Pending_Report.xls", true);
    req.responseType = "arraybuffer";

    req.onload = function (e) {
        var data = new Uint8Array(req.response);

        // Leemos el excel de mcsprod
        var work_book = XLSX.read(data, { type: "array" });

        // Conocemos el Modo en el que vamos a actuar - Editor o Visor
        var sel_mode = document.getElementById("Mode").value;

        // Una vez seleccionemos el Modo y cargemos el excel desaparece el elemento de seleccionar Modo
        document.getElementById("select-mode").style.display = 'none';

        // Guardamos en una variable todas las Sheets
        var sheet_name = work_book.SheetNames;

        var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], { header: 1 }); 		// Sheet 1 - CMs
        var sheet_data_2 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[1]], { header: 1 }); 	// Sheet 2 - IMs
        var sheet_data_3 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[2]], { header: 1 }); 	// Sheet 3 - PMs
        var sheet_data_4 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[3]], { header: 1 });	// Sheet 4 - Blackouts

        // Sheet 1 - CMs
        if (sheet_data.length > 0) { createSheet1(sheet_data, sel_mode); }

        // Sheet 2 - IMs
        if (sheet_data_2.length > 0) { createSheet2(sheet_data_2, sel_mode); }

        // Sheet 3 - PMs
        if (sheet_data_3.length > 0) { createSheet3(sheet_data_3, sel_mode); }

        // Sheet 4 - Blackouts
        if (sheet_data_4.length > 0) { createSheet4(sheet_data_4, sel_mode); }

        // Handover
        createHandover(sel_mode);

        excel_file.value = '';

        // Una vez seleccionado el Modo Editor o Visor mostramos los botones de páginas CM/IM/PM/Blackout y desaparece el elemento del input para cargar el excel
        document.getElementById("section-buttons").style.display = 'block';
        document.getElementById("section-CM").style.display = 'block';
        document.getElementById("import-part").style.display = 'none';
    }

    req.send();

}

// Función que actualiza los datos en tiempo real
function mcsprodFileUpdate() {

    var envTable = ["myTable-prod", "myTable-xcomp", "myTable-training", "myTable-test", "myTable-dev", "myTable-perftest", "myTable-sit", "myTable-uat", "myTable-rtest", "myTable-Decommission", "myTable-ims", "myTable-pms", "myTable-blackouts"];

    for (var e = 0; e < envTable.length; e++) {
        var myTable = document.getElementById(envTable[e]);
        var num_rows = myTable.rows.length - 1;

        for (f = 0; f < num_rows; f++) {
            var tbody = myTable.getElementsByTagName("tbody")[0];
            tbody.removeChild(myTable.rows[1]);

        }

    }

    mcsprodFile();

}

// Función que exporta un backup de los datos cada hora
function backupExcel() {

    // Si estamos en Modo Editor
    if (document.getElementById("Mode").value == 1) {
        // Llamamos a la función que exporta el excel de datos 	
        exportExcel(true);

        // Lanzamos la función cada hora
        setTimeout("backupExcel()", 3600000);
    }
}

// Función que lanza todayFunction() y backupExcel() - Esta función se ejecuta cuando se carga la página 
function shuttle() {

    todayFunction();

    // Esperamos a generar el primer backup una hora
    setTimeout("backupExcel()", 3600000);
}

// Función para señalizar en que página (CM/IM/PM/BLACKOUT) nos encontramos - Marca en distinto color la página en la que nos encontramos
function pageChange(pageType) {

    if (pageType == 'CM') {
        document.getElementById("section-CM").style.display = 'block';
        document.getElementById("button-menu-cm").style = "background-color: #AA643B";
        document.getElementById("h3-cm").style = "color: #E9E1CA";
        document.getElementById("section-IM").style.display = 'none';
        document.getElementById("button-menu-im").style = "background-color: #E9E1CA";
        document.getElementById("h3-im").style = "color: #7A736E";
        document.getElementById("section-PM").style.display = 'none';
        document.getElementById("button-menu-pm").style = "background-color: #E9E1CA";
        document.getElementById("h3-pm").style = "color: #7A736E";
        document.getElementById("section-BLACKOUT").style.display = 'none';
        document.getElementById("button-menu-bo").style = "background-color: #E9E1CA";
        document.getElementById("h3-bo").style = "color: #7A736E";
        document.getElementById("section-HANDOVER").style.display = 'none';
        document.getElementById("button-menu-ho").style = "background-color: #E9E1CA";
        document.getElementById("h3-ho").style = "color: #7A736E";
    } else if (pageType == 'IM') {
        document.getElementById("section-CM").style.display = 'none';
        document.getElementById("button-menu-cm").style = "background-color: #E9E1CA";
        document.getElementById("h3-cm").style = "color: #7A736E";
        document.getElementById("section-IM").style.display = 'block';
        document.getElementById("button-menu-im").style = "background-color: #AA643B";
        document.getElementById("h3-im").style = "color: #E9E1CA";
        document.getElementById("section-PM").style.display = 'none';
        document.getElementById("button-menu-pm").style = "background-color: #E9E1CA";
        document.getElementById("h3-pm").style = "color: #7A736E";
        document.getElementById("section-BLACKOUT").style.display = 'none';
        document.getElementById("button-menu-bo").style = "background-color: #E9E1CA";
        document.getElementById("h3-bo").style = "color: #7A736E";
        document.getElementById("section-HANDOVER").style.display = 'none';
        document.getElementById("button-menu-ho").style = "background-color: #E9E1CA";
        document.getElementById("h3-ho").style = "color: #7A736E";
    } else if (pageType == 'PM') {
        document.getElementById("section-CM").style.display = 'none';
        document.getElementById("button-menu-cm").style = "background-color: #E9E1CA";
        document.getElementById("h3-cm").style = "color: #7A736E";
        document.getElementById("section-IM").style.display = 'none';
        document.getElementById("button-menu-im").style = "background-color: #E9E1CA";
        document.getElementById("h3-im").style = "color: #7A736E";
        document.getElementById("section-PM").style.display = 'block';
        document.getElementById("button-menu-pm").style = "background-color: #AA643B";
        document.getElementById("h3-pm").style = "color: #E9E1CA";
        document.getElementById("section-BLACKOUT").style.display = 'none';
        document.getElementById("button-menu-bo").style = "background-color: #E9E1CA";
        document.getElementById("h3-bo").style = "color: #7A736E";
        document.getElementById("section-HANDOVER").style.display = 'none';
        document.getElementById("button-menu-ho").style = "background-color: #E9E1CA";
        document.getElementById("h3-ho").style = "color: #7A736E";
    } else if (pageType == 'BLACKOUT') {
        document.getElementById("section-CM").style.display = 'none';
        document.getElementById("button-menu-cm").style = "background-color: #E9E1CA";
        document.getElementById("h3-cm").style = "color: #7A736E";
        document.getElementById("section-IM").style.display = 'none';
        document.getElementById("button-menu-im").style = "background-color: #E9E1CA";
        document.getElementById("h3-im").style = "color: #7A736E";
        document.getElementById("section-PM").style.display = 'none';
        document.getElementById("button-menu-pm").style = "background-color: #E9E1CA";
        document.getElementById("h3-pm").style = "color: #7A736E";
        document.getElementById("section-BLACKOUT").style.display = 'block';
        document.getElementById("button-menu-bo").style = "background-color: #AA643B";
        document.getElementById("h3-bo").style = "color: #E9E1CA";
        document.getElementById("section-HANDOVER").style.display = 'none';
        document.getElementById("button-menu-ho").style = "background-color: #E9E1CA";
        document.getElementById("h3-ho").style = "color: #7A736E";
    } else {
        document.getElementById("section-CM").style.display = 'none';
        document.getElementById("button-menu-cm").style = "background-color: #E9E1CA";
        document.getElementById("h3-cm").style = "color: #7A736E";
        document.getElementById("section-IM").style.display = 'none';
        document.getElementById("button-menu-im").style = "background-color: #E9E1CA";
        document.getElementById("h3-im").style = "color: #7A736E";
        document.getElementById("section-PM").style.display = 'none';
        document.getElementById("button-menu-pm").style = "background-color: #E9E1CA";
        document.getElementById("h3-pm").style = "color: #7A736E";
        document.getElementById("section-BLACKOUT").style.display = 'none';
        document.getElementById("button-menu-bo").style = "background-color: #E9E1CA";
        document.getElementById("h3-bo").style = "color: #7A736E";
        document.getElementById("section-HANDOVER").style.display = 'block';
        document.getElementById("button-menu-ho").style = "background-color: #AA643B";
        document.getElementById("h3-ho").style = "color: #E9E1CA";
    }

}

// Función utilizada para la creación del excel donde se guardarán todos los datos
function s2ab(s) {
    if (typeof ArrayBuffer !== 'undefined') {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    } else {
        var buf = new Array(s.length);
        for (var i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
}

// Función para guardar y subir un fichero a mcsprod
function uploadFile(file, typeFile, nomFile) {

    var myBlob = new Blob(
        [file],
        { type: typeFile }
    );

    var data = new FormData();
    data.append("upFile", myBlob);

    var xhr = new XMLHttpRequest();
    xhr.open("POST", nomFile);
    xhr.onload = function () {
        console.log(this.response);
    }

    xhr.send(data);

}

/* 
 * Función que exporta los datos en un excel con diversas páginas: Sheet 1 - CMs; Sheet 2 - IMs; Sheet 3 - PMs; Sheet 4 - Blackouts. 
 * backup_flag = true -> Se le llama desde la función backupExcel() y generamos un archivo de los datos backup
 * backup_flag = false -> Se ha dado al botón export y se genera un archivo de los datos en ese momento 
 */
function exportExcel(backup_flag) {

    // Creamos un nuevo archivo excel
    var wb = XLSX.utils.book_new();

    // Sheet 1 - CMs
    var wsName = 'CMs';
    var wsData = [];
    var envTable = ["myTable-prod", "myTable-xcomp", "myTable-training", "myTable-test", "myTable-dev", "myTable-perftest", "myTable-sit", "myTable-uat", "myTable-rtest", "myTable-Decommission"];
    var env = ["Prod", "XCOMP", "Training", "Test", "Dev", "Perf Test", "SIT", "UAT", "RTEST", "Decommission"];

    var w = 0;
    for (var e = 0; e < env.length; e++) {
        var myTable = document.getElementById(envTable[e]);
        var numRows = myTable.rows.length;
        var numCells = myTable.rows[0].cells.length - 1;

        // 	Creamos una matriz con el número de filas y celdas necesario	
        for (var c = 0; c < (numRows + 1); c++) {
            wsData[w] = new Array(numCells);
            w++;
        }

    }

    var k = 0;
    for (var e = 0; e < env.length; e++) {
        var myTable = document.getElementById(envTable[e]);
        var numRows = myTable.rows.length;
        var numCells = myTable.rows[0].cells.length - 1;

        // Guardamos en la matriz los datos  
        if (Number(numRows) > 1) {

            // En primer lugar, guardamos el entorno (Prod, XCOMP, ...)
            wsData[k][0] = env[e];
            k++;

            // Segundo, guardamos el nombre de los campos de cada celda (CM Ticket, Status, ...)
            for (var h = 0; h < numCells; h++) {
                wsData[k][h] = myTable.rows[0].cells[h].textContent;
            }
            k++;

            // Tercero, guardamos los datos de cada fila y celda de toda la tabla
            for (var r = 1; r < numRows; r++) {
                for (var g = 0; g < numCells; g++) {
                    wsData[k][g] = myTable.rows[r].cells[g].children[0].value;
                }
                k++;
            }
        }
    }

    var ws = XLSX.utils.aoa_to_sheet(wsData);

    // Guardamos en el excel -> Sheet 1 - CMs
    XLSX.utils.book_append_sheet(wb, ws, wsName);


    // Sheet 2 - IMs
    var wsName2 = 'IMs';
    var wsData2 = [];

    var myTable2 = document.getElementById("myTable-ims");
    var numRows2 = myTable2.rows.length;
    var numCells2 = myTable2.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario		
    for (var c = 0; c < numRows2; c++) {
        wsData2[c] = new Array(numCells2);
    }

    // Primero, guardamos el nombre de los campos de cada celda (IM Ticket, STASK number, ...)
    for (var h = 0; h < numCells2; h++) {
        wsData2[0][h] = myTable2.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla
    for (var r = 1; r < numRows2; r++) {
        for (var g = 0; g < numCells2; g++) {
            wsData2[r][g] = myTable2.rows[r].cells[g].children[0].value;
        }
    }

    var ws2 = XLSX.utils.aoa_to_sheet(wsData2);

    // Guardamos en el excel -> Sheet 2 - IMs
    XLSX.utils.book_append_sheet(wb, ws2, wsName2);

    // Sheet 3 - PMs
    var wsName3 = 'PMs';
    var wsData3 = [];

    var myTable3 = document.getElementById("myTable-pms");
    var numRows3 = myTable3.rows.length;
    var numCells3 = myTable3.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario						
    for (var c = 0; c < numRows3; c++) {
        wsData3[c] = new Array(numCells3);
    }

    // Primero, guardamos el nombre de los campos de cada celda (PM Ticket, STASK number, ...)
    for (var h = 0; h < numCells3; h++) {
        wsData3[0][h] = myTable3.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla		
    for (var r = 1; r < numRows3; r++) {
        for (var g = 0; g < numCells3; g++) {
            wsData3[r][g] = myTable3.rows[r].cells[g].children[0].value;
        }
    }

    var ws3 = XLSX.utils.aoa_to_sheet(wsData3);

    // Guardamos en el excel -> Sheet 3 - PMs		
    XLSX.utils.book_append_sheet(wb, ws3, wsName3);

    // Sheet 4 - Blackouts
    var wsName4 = 'Blackouts';
    var wsData4 = [];

    var myTable4 = document.getElementById("myTable-blackouts");
    var numRows4 = myTable4.rows.length;
    var numCells4 = myTable4.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario										
    for (var c = 0; c < numRows4; c++) {
        wsData4[c] = new Array(numCells4);
    }

    // Primero, guardamos el nombre de los campos de cada celda (CM Ticket, STASK number, ...)		
    for (var h = 0; h < numCells4; h++) {
        wsData4[0][h] = myTable4.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla		
    for (var r = 1; r < numRows4; r++) {
        for (var g = 0; g < numCells4; g++) {
            wsData4[r][g] = myTable4.rows[r].cells[g].children[0].value;
        }
    }

    var ws4 = XLSX.utils.aoa_to_sheet(wsData4);

    // Guardamos en el excel -> Sheet 4 - Blackouts		
    XLSX.utils.book_append_sheet(wb, ws4, wsName4);

    // Conocemos la fecha de hoy 
    var today = getToday();
    var day = today[0],
        month = today[1],
        year = today[2],
        hours = today[3];

    // Conocemos el turno en el que nos encontramos al generar el export normal
    if (Number(hours) < 7) {
        var turno = "noche";
    } else if (Number(hours) < 15) {
        var turno = "mañana";
    } else {
        var turno = "tarde";
    }

    // backup_flag = true -> creamos un backup dentro de mcsprod; backup_flag = false -> se podrá exportar bajo el nombre EMA_Pending_Report_daymonthyear_turno.xlsx y además se guardará en mcsprod
    if (backup_flag) {

        // Creamos el archivo excel y lo subimos a mcsprod
        var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
        uploadFile(s2ab(wbout), "application/vnd.ms-excel;base64", "upload_backup.php");

    } else {
        var nomFile = "EMA_Pending_Report_" + year + month + day + "_" + turno + ".xlsx";

        // Generamos un archivo excel con el nombre del valor que guarde la variable nomFile y escribimos lo que hemos ido guardando en la variable wb
        XLSX.writeFile(wb, nomFile);

        // Creamos el archivo excel y lo subimos a mcsprod
        var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
        uploadFile(s2ab(wbout), "application/vnd.ms-excel;base64", "upload.php");

    }

}


// Función que guarda los datos actuales en mcsprod
function dataSavedMachine() {

    // Creamos un nuevo archivo excel
    var wb = XLSX.utils.book_new();

    // Sheet 1 - CMs
    var wsName = 'CMs';
    var wsData = [];
    var envTable = ["myTable-prod", "myTable-xcomp", "myTable-training", "myTable-test", "myTable-dev", "myTable-perftest", "myTable-sit", "myTable-uat", "myTable-rtest", "myTable-Decommission"];
    var env = ["Prod", "XCOMP", "Training", "Test", "Dev", "Perf Test", "SIT", "UAT", "RTEST", "Decommission"];

    var w = 0;
    for (var e = 0; e < env.length; e++) {
        var myTable = document.getElementById(envTable[e]);
        var numRows = myTable.rows.length;
        var numCells = myTable.rows[0].cells.length - 1;

        // 	Creamos una matriz con el número de filas y celdas necesario	
        for (var c = 0; c < (numRows + 1); c++) {
            wsData[w] = new Array(numCells);
            w++;
        }

    }

    var k = 0;
    for (var e = 0; e < env.length; e++) {
        var myTable = document.getElementById(envTable[e]);
        var numRows = myTable.rows.length;
        var numCells = myTable.rows[0].cells.length - 1;

        // Guardamos en la matriz los datos  
        if (Number(numRows) > 1) {

            // En primer lugar, guardamos el entorno (Prod, XCOMP, ...)
            wsData[k][0] = env[e];
            k++;

            // Segundo, guardamos el nombre de los campos de cada celda (CM Ticket, Status, ...)
            for (var h = 0; h < numCells; h++) {
                wsData[k][h] = myTable.rows[0].cells[h].textContent;
            }
            k++;

            // Tercero, guardamos los datos de cada fila y celda de toda la tabla
            for (var r = 1; r < numRows; r++) {
                for (var g = 0; g < numCells; g++) {
                    wsData[k][g] = myTable.rows[r].cells[g].children[0].value;
                }
                k++;
            }
        }
    }

    var ws = XLSX.utils.aoa_to_sheet(wsData);

    // Guardamos en el excel -> Sheet 1 - CMs
    XLSX.utils.book_append_sheet(wb, ws, wsName);


    // Sheet 2 - IMs
    var wsName2 = 'IMs';
    var wsData2 = [];

    var myTable2 = document.getElementById("myTable-ims");
    var numRows2 = myTable2.rows.length;
    var numCells2 = myTable2.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario		
    for (var c = 0; c < numRows2; c++) {
        wsData2[c] = new Array(numCells2);
    }

    // Primero, guardamos el nombre de los campos de cada celda (IM Ticket, STASK number, ...)
    for (var h = 0; h < numCells2; h++) {
        wsData2[0][h] = myTable2.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla
    for (var r = 1; r < numRows2; r++) {
        for (var g = 0; g < numCells2; g++) {
            wsData2[r][g] = myTable2.rows[r].cells[g].children[0].value;
        }
    }

    var ws2 = XLSX.utils.aoa_to_sheet(wsData2);

    // Guardamos en el excel -> Sheet 2 - IMs
    XLSX.utils.book_append_sheet(wb, ws2, wsName2);

    // Sheet 3 - PMs
    var wsName3 = 'PMs';
    var wsData3 = [];

    var myTable3 = document.getElementById("myTable-pms");
    var numRows3 = myTable3.rows.length;
    var numCells3 = myTable3.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario						
    for (var c = 0; c < numRows3; c++) {
        wsData3[c] = new Array(numCells3);
    }

    // Primero, guardamos el nombre de los campos de cada celda (PM Ticket, STASK number, ...)
    for (var h = 0; h < numCells3; h++) {
        wsData3[0][h] = myTable3.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla		
    for (var r = 1; r < numRows3; r++) {
        for (var g = 0; g < numCells3; g++) {
            wsData3[r][g] = myTable3.rows[r].cells[g].children[0].value;
        }
    }

    var ws3 = XLSX.utils.aoa_to_sheet(wsData3);

    // Guardamos en el excel -> Sheet 3 - PMs		
    XLSX.utils.book_append_sheet(wb, ws3, wsName3);

    // Sheet 4 - Blackouts
    var wsName4 = 'Blackouts';
    var wsData4 = [];

    var myTable4 = document.getElementById("myTable-blackouts");
    var numRows4 = myTable4.rows.length;
    var numCells4 = myTable4.rows[0].cells.length - 1;

    // 	Creamos una matriz con el número de filas y celdas necesario										
    for (var c = 0; c < numRows4; c++) {
        wsData4[c] = new Array(numCells4);
    }

    // Primero, guardamos el nombre de los campos de cada celda (CM Ticket, STASK number, ...)		
    for (var h = 0; h < numCells4; h++) {
        wsData4[0][h] = myTable4.rows[0].cells[h].textContent;
    }

    // Segundo, guardamos los datos de cada fila y celda de toda la tabla		
    for (var r = 1; r < numRows4; r++) {
        for (var g = 0; g < numCells4; g++) {
            wsData4[r][g] = myTable4.rows[r].cells[g].children[0].value;
        }
    }

    var ws4 = XLSX.utils.aoa_to_sheet(wsData4);

    // Guardamos en el excel -> Sheet 4 - Blackouts		
    XLSX.utils.book_append_sheet(wb, ws4, wsName4);


    // Creamos el excel donde guardamos todo y lo subimos a mcsprod
    var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });

    var myBlob = new Blob(
        [s2ab(wbout)],
        { type: "application/vnd.ms-excel;base64" }
    );

    var data = new FormData();
    data.append("upFile", myBlob);

    var xhr = new XMLHttpRequest();
    xhr.open("POST", "upload.php");
    xhr.onload = function () {
        console.log(this.response);
    }

    xhr.send(data);

}

// Función que guarda lo que actualmente contiene el HO a mcsprod
function saveHandover() {

    var textoHandover_1 = document.getElementById("avisosPermanentes").value;
    var textoHandover_2 = document.getElementById("cambiosT2actions").value;

    // Guardamos y subimos a mcsprod un txt que contiene los avisos permanentes
    uploadFile(textoHandover_1, "text/plain", "uploadHandover_1.php");

    // Guardamos y subimos a mcsprod un txt que contiene las cosas pendientes
    uploadFile(textoHandover_2, "text/plain", "uploadHandover_2.php");

}

// Función que genera el HO con los datos actuales que contiene toda la herramienta (genera las cosas pendientes para el siguiente turno)
function editHandover() {

    var envTable_1 = ["myTable-prod", "myTable-xcomp", "myTable-training", "myTable-test", "myTable-dev", "myTable-perftest", "myTable-sit", "myTable-uat", "myTable-rtest", "myTable-Decommission"];
    var envTable_2 = ["myTable-ims", "myTable-pms"];
    var cmEnv = ["\n### PROD ###", "\n### NON PROD ###", "\n### DECOMMISSION ###"];

    // Chequeamos la fecha actual
    var today = getToday();
    var day = today[0],
        month = today[1],
        year = today[2],
        hours = today[3],
        dayWeek = today[5],
        date = today[6];

    var turnoTarde = false;
    var textCambiosT2actions = "";

    // Guardamos la fecha que nos interesa exportar en el Handover (la correspondiente al siguiente turno)
    if ((Number(dayWeek) == 0) || (Number(dayWeek) > 5) || ((Number(dayWeek) == 5) && (Number(hours) > 15))) { // Si hoy es domingo, sabado o viernes - turno tarde

        if (Number(dayWeek) == 5) {
            var tomorrow = new Date();
            tomorrow.setDate(date.getDate() + 1);
            day = tomorrow.getDate();
            month = tomorrow.getMonth() + 1;
            year = tomorrow.getFullYear();
            if (Number(month) < 10) {
                month = "0" + month;
            }
            if (Number(day) < 10) {
                day = "0" + day;
            }
        }

        // Añadimos al HO el fin de semana (en el que caso que haya CMs programados)
        var vuelta = 0;

        if (Number(dayWeek) == 0) {
            vuelta++;
        }

        do {

            var firstTime_1 = false;
            var firstTime_2 = false;
            var z = 0;

            // Generamos los CMs pendientes 
            for (var a = 0; a < envTable_1.length; a++) {

                if (a == 2) {
                    firstTime_2 = false;
                    z++;
                } else if (a == 9) {
                    firstTime_2 = false;
                    z++;
                }

                var myTable = document.getElementById(envTable_1[a]);
                var num_rows = myTable.rows.length;

                for (var f = 1; f < num_rows; f++) {
                    var checkDate = myTable.rows[f].cells[4].children[0].value;
                    var checkDate_2 = Number(checkDate.substring(0, 2));
                    var checkDate_4 = Number(checkDate.substring(3, 5));
                    var checkDate_6 = Number(checkDate.substring(6, 10));
                    var checkDate_8 = Number(checkDate.substring(11, 13));
                    var checkActions = myTable.rows[f].cells[7].children[0].value;

                    if ((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) {

                        if (!firstTime_1) {
                            firstTime_1 = true;
                            textCambiosT2actions = textCambiosT2actions + "\nTurno de fin de semana - " + day + "/" + month + " :redsiren:" + "\n";
                        }

                        if (!firstTime_2) {
                            firstTime_2 = true;
                            textCambiosT2actions = textCambiosT2actions + cmEnv[z] + "\n";
                        }
                        textCambiosT2actions = textCambiosT2actions + "*CM " + myTable.rows[f].cells[0].children[0].value.trim() + " | " + myTable.rows[f].cells[2].children[0].value.trim() + " | " + myTable.rows[f].cells[4].children[0].value.trim() + " | " + myTable.rows[f].cells[3].children[0].value + "\n";

                    }
                }

            }

            vuelta++;

            if (vuelta == 1) {
                var domingo = new Date();
                if (Number(dayWeek) == 5) {
                    domingo.setDate(tomorrow.getDate() + 1);
                } else {
                    domingo.setDate(date.getDate() + 1);
                }
                day = domingo.getDate();
                month = domingo.getMonth() + 1;
                year = domingo.getFullYear();
            }

        } while (vuelta < 2);

        // Generamos la fecha del proximo lunes
        var lunes = new Date();

        if (Number(dayWeek) == 5) {
            lunes.setDate(date.getDate() + 3);
        } else if (Number(dayWeek) == 6) {
            lunes.setDate(date.getDate() + 2);
        } else {
            lunes.setDate(date.getDate() + 1);
        }
        day = lunes.getDate();
        month = lunes.getMonth() + 1;
        year = lunes.getFullYear();

        turnoTarde = true;
        textCambiosT2actions = textCambiosT2actions + "\n";

    } else { // Si no es viernes turno de tarde o fin de semana, comprobamos si estamos en turno de tarde/noche para generar el HO de turno de mañana

        if (Number(hours) > 15) {	// turno de tarde
            turnoTarde = true;
            var tomorrow = new Date();
            tomorrow.setDate(date.getDate() + 1);
            day = tomorrow.getDate();
            dayWeek = tomorrow.getDay();
            month = tomorrow.getMonth() + 1;
            year = tomorrow.getFullYear();
            if (Number(month) < 10) {
                month = "0" + month;
            }
            if (Number(day) < 10) {
                day = "0" + day;
            }
        } else if (Number(hours) < 6) { // turno de noche
            turnoTarde = true;
        }

    }

    if (turnoTarde) {
        turnoHO = "mañana";
    } else {
        turnoHO = "tarde";
    }

    var dateHO = "Turno de " + turnoHO + " - " + day + "/" + month + " :redsiren:" + "\n";
    textCambiosT2actions = textCambiosT2actions + dateHO;

    var firstTime = false;
    var k = 0;

    // Generamos los CMs pendientes 
    for (var a = 0; a < envTable_1.length; a++) {

        if (a == 2) {
            firstTime = false;
            k++;
        } else if (a == 9) {
            firstTime = false;
            k++;
        }

        var myTable = document.getElementById(envTable_1[a]);
        var num_rows = myTable.rows.length;

        for (var f = 1; f < num_rows; f++) {
            var checkDate = myTable.rows[f].cells[4].children[0].value;
            var checkDate_2 = Number(checkDate.substring(0, 2));
            var checkDate_4 = Number(checkDate.substring(3, 5));
            var checkDate_6 = Number(checkDate.substring(6, 10));
            var checkDate_8 = Number(checkDate.substring(11, 13));
            var checkActions = myTable.rows[f].cells[7].children[0].value;

            if (((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) && ((!turnoTarde) || ((turnoTarde) && (Number(checkDate_8) < 15)))) {

                if (!firstTime) {
                    firstTime = true;
                    textCambiosT2actions = textCambiosT2actions + cmEnv[k] + "\n";
                }
                textCambiosT2actions = textCambiosT2actions + "*CM " + myTable.rows[f].cells[0].children[0].value.trim() + " | " + myTable.rows[f].cells[2].children[0].value.trim() + " | " + myTable.rows[f].cells[4].children[0].value.trim() + " | " + myTable.rows[f].cells[3].children[0].value + "\n";

            } else if ((checkActions.localeCompare("Tier 2 actions") == 0) && (Number(checkDate_2) == 0)) {

                if (!firstTime) {
                    firstTime = true;
                    textCambiosT2actions = textCambiosT2actions + cmEnv[k] + "\n";
                }
                textCambiosT2actions = textCambiosT2actions + "*CM " + myTable.rows[f].cells[0].children[0].value.trim() + " | " + myTable.rows[f].cells[2].children[0].value.trim() + " | " + myTable.rows[f].cells[3].children[0].value + "\n";

            }
        }

    }

    // Generamos los IMs pendientes
    var myTable = document.getElementById(envTable_2[0]);
    var num_rows = myTable.rows.length;
    var firstTime = false;

    for (var f = 1; f < num_rows; f++) {
        var checkActions = myTable.rows[f].cells[5].children[0].value;
        if (checkActions.localeCompare("Tier 2 actions") == 0) {
            if (!firstTime) {
                firstTime = true;
                textCambiosT2actions = textCambiosT2actions + "\n" + "### IMs ###" + "\n";
            }
            textCambiosT2actions = textCambiosT2actions + "*IM " + myTable.rows[f].cells[0].children[0].value.trim() + " | " + myTable.rows[f].cells[1].children[0].value.trim() + " | " + myTable.rows[f].cells[2].children[0].value + " --> " + myTable.rows[f].cells[4].children[0].value + "\n";
        }
    }

    // Generamos los PMs pendientes
    var myTable = document.getElementById(envTable_2[1]);
    var num_rows = myTable.rows.length;
    var firstTime = false;

    for (var f = 1; f < num_rows; f++) {
        var checkActions = myTable.rows[f].cells[5].children[0].value;
        if (checkActions.localeCompare("Tier 2 actions") == 0) {
            if (!firstTime) {
                firstTime = true;
                textCambiosT2actions = textCambiosT2actions + "\n" + "### PMs ###" + "\n";
            }
            textCambiosT2actions = textCambiosT2actions + "*PM " + myTable.rows[f].cells[0].children[0].value.trim() + " | " + myTable.rows[f].cells[1].children[0].value.trim() + " | " + myTable.rows[f].cells[2].children[0].value + " --> " + myTable.rows[f].cells[4].children[0].value + "\n";
        }
    }

    if (textCambiosT2actions.localeCompare(dateHO) == 0) {
        textCambiosT2actions = textCambiosT2actions + "\nN/A";
    }

    document.getElementById("cambiosT2actions").value = textCambiosT2actions;

}

// Función que recorre las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
function updateRows(table) {

    var t = table.rows.length - 1;
    var today = getToday();

    var day = today[0],
        month = today[1],
        year = today[2];

    while (t > 0) {

        var checkDate = table.rows[t].cells[4].children[0].value;
        var checkDate_2 = Number(checkDate.substring(0, 2));
        var checkDate_4 = Number(checkDate.substring(3, 5));
        var checkDate_6 = Number(checkDate.substring(6, 10));
        var checkWaiting = table.rows[t].cells[7].children[0].value;

        if ((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) {
            table.rows[t].cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
            table.rows[t].cells[4].style = "background-color: #FF9F9F;";

            if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
                table.rows[t].cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                table.rows[t].cells[7].style = "background-color: #F3BDFF";
                table.rows[t].onmouseout = function () {
                    this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";
                    this.cells[7].style = "background-color: #F3BDFF;";
                    this.cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
                    this.cells[4].style = "background-color: #FF9F9F;";
                };
                table.rows[t].onmouseover = function () {
                    this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";;
                    this.cells[7].style = "background-color: #F3BDFF;";
                    this.cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
                    this.cells[4].style = "background-color: #FF9F9F;";
                };
            } else {

                if ((t % 2) == 0) {
                    table.rows[t].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
                    table.rows[t].cells[7].style = "background-color: #ffffff;";
                    table.rows[t].onmouseout = function () {
                        this.cells[7].children[0].style = "background-color: #ffffff";
                        this.cells[7].style = "background-color: #ffffff;";
                        this.cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
                        this.cells[4].style = "background-color: #FF9F9F;";
                    };
                } else {
                    table.rows[t].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
                    table.rows[t].cells[7].style = "background-color: #f2f2f2;";
                    table.rows[t].onmouseout = function () {
                        this.cells[7].children[0].style = "background-color: #f2f2f2";
                        this.cells[7].style = "background-color: #f2f2f2;";
                        this.cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
                        this.cells[4].style = "background-color: #FF9F9F;";
                    };
                }
                table.rows[t].onmouseover = function () {
                    this.cells[7].children[0].style = "background-color: #ddd;";
                    this.cells[7].style = "background-color: #ddd;";
                    this.cells[4].children[0].style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;";
                    this.cells[4].style = "background-color: #FF9F9F;";
                };

            }
        } else {
            if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
                table.rows[t].cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                table.rows[t].cells[7].style = "background-color: #F3BDFF";

                if ((t % 2) == 0) {
                    table.rows[t].cells[4].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
                    table.rows[t].cells[4].style = "background-color: #ffffff;";
                    table.rows[t].onmouseout = function () {
                        this.cells[4].children[0].style = "background-color: #ffffff";
                        this.cells[4].style = "background-color: #ffffff;";
                        this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";
                        this.cells[7].style = "background-color: #F3BDFF;";
                    };
                } else {
                    table.rows[t].cells[4].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
                    table.rows[t].cells[4].style = "background-color: #f2f2f2;";
                    table.rows[t].onmouseout = function () {
                        this.cells[4].children[0].style = "background-color: #f2f2f2";
                        this.cells[4].style = "background-color: #f2f2f2;";
                        this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";
                        this.cells[7].style = "background-color: #F3BDFF;";
                    };
                }
                table.rows[t].onmouseover = function () {
                    this.cells[4].children[0].style = "background-color: #ddd;";
                    this.cells[4].style = "background-color: #ddd;";
                    this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";;
                    this.cells[7].style = "background-color: #F3BDFF;";
                };

            } else {

                if ((t % 2) == 0) {
                    table.rows[t].cells[4].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
                    table.rows[t].cells[4].style = "background-color: #ffffff;";
                    table.rows[t].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
                    table.rows[t].cells[7].style = "background-color: #ffffff;";
                    table.rows[t].onmouseout = function () {
                        this.cells[4].children[0].style = "background-color: #ffffff";
                        this.cells[4].style = "background-color: #ffffff;";
                        this.cells[7].children[0].style = "background-color: #ffffff";
                        this.cells[7].style = "background-color: #ffffff;";
                    };
                } else {
                    table.rows[t].cells[4].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
                    table.rows[t].cells[4].style = "background-color: #f2f2f2;";
                    table.rows[t].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
                    table.rows[t].cells[7].style = "background-color: #f2f2f2;";
                    table.rows[t].onmouseout = function () {
                        this.cells[4].children[0].style = "background-color: #f2f2f2";
                        this.cells[4].style = "background-color: #f2f2f2;";
                        this.cells[7].children[0].style = "background-color: #f2f2f2";
                        this.cells[7].style = "background-color: #f2f2f2;";
                    };
                }
                table.rows[t].onmouseover = function () {
                    this.cells[4].children[0].style = "background-color: #ddd;";
                    this.cells[4].style = "background-color: #ddd;";
                    this.cells[7].children[0].style = "background-color: #ddd;";
                    this.cells[7].style = "background-color: #ddd;";
                };

            }

        }

        t--;
    }
}

// Función que crea una nueva fila en las tablas de CMs
function addFunctionCM() {

    // Conocemos que entorno es
    var env = document.getElementById("enviroment").value;

    // Si no se ha seleccionado entorno al intentar añadir fila en la página de CMs te obliga a hacerlo
    if (env.localeCompare("Select Enviroment") == 0) {
        alert("Select an environment, please");
    } else {

        var nTable;

        // Conociendo ya el entorno, mostramos esa tabla por si estuviese oculta al no contar aún con ninguna fila
        if (env.localeCompare("Prod") == 0) {
            nTable = "myTable-prod";
            document.getElementById("section-PROD").style.display = 'block';
        } else if (env.localeCompare("XCOMP") == 0) {
            nTable = "myTable-xcomp";
            document.getElementById("section-XCOMP").style.display = 'block';
        } else if (env.localeCompare("Training") == 0) {
            nTable = "myTable-training";
            document.getElementById("section-Training").style.display = 'block';
        } else if (env.localeCompare("Test") == 0) {
            nTable = "myTable-test";
            document.getElementById("section-Test").style.display = 'block';
        } else if (env.localeCompare("Dev") == 0) {
            nTable = "myTable-dev";
            document.getElementById("section-Dev").style.display = 'block';
        } else if (env.localeCompare("Perf Test") == 0) {
            nTable = "myTable-perftest";
            document.getElementById("section-PerfTest").style.display = 'block';
        } else if (env.localeCompare("SIT") == 0) {
            nTable = "myTable-sit";
            document.getElementById("section-SIT").style.display = 'block';
        } else if (env.localeCompare("UAT") == 0) {
            nTable = "myTable-uat";
            document.getElementById("section-UAT").style.display = 'block';
        } else if (env.localeCompare("RTEST") == 0) {
            nTable = "myTable-rtest";
            document.getElementById("section-RTEST").style.display = 'block';
        } else if (env.localeCompare("Decommission") == 0) {
            nTable = "myTable-Decommission";
            document.getElementById("section-Decommission").style.display = 'block';
        }

        var table = document.getElementById(nTable);

        // Campo CM Ticket
        var input = document.createElement("input");
        input.type = "text";
        input.className = "cm";
        input.value = document.getElementById("cm_id").value;

        // Campo Status
        var select2 = document.createElement("select");
        select2.id = "status";
        select2.className = "status";
        var array_status = ["Select Status", "Draft", "Pending approval", "Pending scheduling", "Scheduled", "Completed", "WIP"];
        for (var i = 0; i < array_status.length; i++) {
            var optionSt = document.createElement("option");
            optionSt.value = array_status[i];
            optionSt.text = array_status[i];
            select2.appendChild(optionSt);
        }
        select2.value = document.getElementById("status").value;
        var select_value = select2.value;
        if (select_value.localeCompare("Draft") == 0) {
            select2.style = "background-color: yellow; color: black;";
        } else if (select_value.localeCompare("Pending approval") == 0) {
            select2.style = "background-color: #81D0E5; color: black;";
        } else if (select_value.localeCompare("Pending scheduling") == 0) {
            select2.style = "background-color: #00C5FF; color: black;";
        } else if (select_value.localeCompare("Scheduled") == 0) {
            select2.style = "background-color: #9F01A9; color: white;";
        } else if (select_value.localeCompare("Completed") == 0) {
            select2.style = "background-color: red; color: white;";
        } else if (select_value.localeCompare("WIP") == 0) {
            select2.style = "background-color: #81F07A; color: black;";
        }

        // Campo STASK number
        var input3 = document.createElement("input");
        input3.type = "text";
        input3.className = "stask";
        input3.value = document.getElementById("stask_id").value;

        // Campo Summary
        var textarea4 = document.createElement("textarea");
        textarea4.type = "text";
        textarea4.className = "summary";
        textarea4.onkeyup = function () { textAreaAdjust(this) };
        textarea4.style = "overflow:hidden";
        textarea4.value = document.getElementById("summary").value;

        // Campo Scheduled Date Madrid time
        var input5 = document.createElement("input");
        input5.type = "text";
        input5.id = "date-order";
        input5.className = "date";
        input5.value = document.getElementById("date").value;

        // Campo Implementer Name
        var input6 = document.createElement("input");
        input6.type = "text";
        input6.className = "implementer";
        input6.value = document.getElementById("imp_name").value;

        // Campo Observations
        var textarea7 = document.createElement("textarea");
        textarea7.type = "text";
        textarea7.className = "observations";
        textarea7.onkeyup = function () { textAreaAdjust(this) };
        textarea7.style = "overflow:hidden";
        textarea7.value = document.getElementById("obs").value;

        // Campo Waiting for
        var input8 = document.createElement("input");
        input8.type = "text";
        input8.className = "waiting";
        input8.setAttribute("list", "waiting_for_list");
        var datalist_input8 = document.createElement("datalist");
        datalist_input8.id = "waiting_for_list";
        var wfstatus = ["Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
        wfstatus.forEach(function (item) {
            var option = document.createElement('option');
            option.value = item;
            datalist_input8.appendChild(option);
        });
        input8.value = document.getElementById("waiting_for").value;

        /* 
         * Si aún no existe ninguna fila de datos en esa tabla, lo añade en la primera fila (El valor es 1 porque existe la fila con el nombre de los campos (CM Ticket, Status, ...))
         * Sino compara el campo "Scheduled Date Madrid time" de la fila que queremos introducir con los de las que ya están en la tabla para saber en que posición debe añadirla
         * Las tablas están ordenadas de forma ascendente, si el campo "Scheduled Date Madrid time" está vacío la fila se colocará al final de la tabla
         */
        if (table.rows.length == 1) {
            var row = table.insertRow(1);
        } else {
            var i = 1;
            var end = false;
            var row;
            // Conocemos el campo "Scheduled Date Madrid time" de la fila que queremos introducir - A partir de ahora Fila-Y
            var y = document.getElementById("date").value;

            // Si el campo "Scheduled Date Madrid time" está vacío la Fila-Y se colocará al final de la tabla 
            if (y.localeCompare("") == 0) {
                row = table.insertRow(-1);
            } else {
                // Conocemos el día del campo "Scheduled Date Madrid time" de la Fila-Y
                var y_2 = Number(y.substring(0, 2));
                // Conocemos el mes del campo "Scheduled Date Madrid time" de la Fila-Y
                var y_4 = Number(y.substring(3, 5));
                // Conocemos el año del campo "Scheduled Date Madrid time" de la Fila-Y
                var y_6 = Number(y.substring(6, 10));
                // Conocemos la hora del campo "Scheduled Date Madrid time" de la Fila-Y
                var y_8 = Number(y.substring(11, 13));
                // Conocemos los minutos del campo "Scheduled Date Madrid time" de Fila-Y
                var y_10 = Number(y.substring(14, 16));

                // Recorremos todas las filas ya existentes
                while (end == false) {
                    // Campo "Scheduled Date Madrid time" de una fila que ya existe -  A partir de ahora Fila-X
                    var x = table.rows[i].cells[4].children[0].value;

                    // Conocemos el día del campo "Scheduled Date Madrid time" de Fila-X
                    var x_2 = Number(x.substring(0, 2));
                    // Conocemos el mes del campo "Scheduled Date Madrid time" de Fila-X
                    var x_4 = Number(x.substring(3, 5));
                    // Conocemos el año del campo "Scheduled Date Madrid time" de Fila-X
                    var x_6 = Number(x.substring(6, 10));
                    // Conocemos la hora del campo "Scheduled Date Madrid time" de Fila-X
                    var x_8 = Number(x.substring(11, 13));
                    // Conocemos los minutos del campo "Scheduled Date Madrid time" de Fila-X
                    var x_10 = Number(x.substring(14, 16));

                    /* 
                     * Si el campo "Scheduled Date Madrid time" de Fila-X está vacío
                     * Si el año de la Fila-Y es igual que el de Fila-X y el mes de Fila-Y es menor que el de Fila-X
                     * Si el año de la Fila-Y es menor que el de Fila-X
                     * Si el año y mes de la Fila-Y es igual que el de Fila-X y el día de Fila-Y es menor que el de Fila-X
                     * Si es el mismo día, mes y año de la Fila-Y y Fila-X y la hora de la Fila-Y es menor que la de Fila-X
                     * Si es el mismo día, mes, año y hora de la Fila-Y y Fila-X y los minutos de la Fila-Y es menor que la de Fila-X 
                     * Se añade la Fila-Y justo encima de la Fila-X
                     */
                    if ((x.localeCompare("") == 0) || ((Number(y_6) == Number(x_6)) && (Number(y_4) < Number(x_4))) || (Number(y_6) < Number(x_6)) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) < Number(x_2))) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) == Number(x_2)) && (Number(y_8) < Number(x_8))) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) == Number(x_2)) && (Number(y_8) == Number(x_8)) && (Number(y_10) < Number(x_10)))) {
                        end = true;
                        row = document.createElement("tr");
                        var tbody = table.getElementsByTagName("tbody")[0];
                        tbody.insertBefore(row, table.rows[i]);
                    } else if (i == (table.rows.length - 1)) {
                        // Si ya estamos comparando con la última fila existente de la tabla y no se ha cumplido lo anterior, añadimos la Fila-Y al final de la tabla
                        end = true;
                        row = table.insertRow(-1);
                    }

                    // Siguiente fila existente	
                    i++;
                }
            }

        }

        var cell1 = row.insertCell(0);
        var cell2 = row.insertCell(1);
        var cell3 = row.insertCell(2);
        var cell4 = row.insertCell(3);
        var cell5 = row.insertCell(4);
        var cell6 = row.insertCell(5);
        var cell7 = row.insertCell(6);
        var cell8 = row.insertCell(7);
        var cell9 = row.insertCell(8);

        var checkDate = input5.value;
        var checkDate_2 = Number(checkDate.substring(0, 2));
        var checkDate_4 = Number(checkDate.substring(3, 5));
        var checkDate_6 = Number(checkDate.substring(6, 10));
        var today = getToday();
        var day = today[0],
            month = today[1],
            year = today[2];

        if ((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) {
            input5.style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;"; // Si es el día de hoy aparecerá en rojo
            cell5.style = "background-color: #FF9F9F";
        }

        var checkWaiting = input8.value;
        if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
            input8.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
            cell8.style = "background-color: #F3BDFF";
        }

        cell1.appendChild(input);
        cell2.appendChild(select2);
        cell3.appendChild(input3);
        cell4.appendChild(textarea4);
        cell5.appendChild(input5);
        cell6.appendChild(input6);
        cell7.appendChild(textarea7);
        cell8.appendChild(input8);
        cell8.appendChild(datalist_input8);
        cell1.className = "cm";
        cell1.addEventListener('change', updateValueCM);
        cell2.className = "status";
        cell2.addEventListener('change', updateValueStatus);
        cell3.className = "stask";
        cell4.className = "summary";
        cell5.className = "date";
        cell5.id = "date-container";
        cell5.addEventListener('change', updateValue);
        cell6.className = "implementer";
        cell7.className = "obs";
        cell8.className = "waiting";
        cell8.addEventListener('change', updateValueWaiting);

        // Campo Delete
        var input9 = document.createElement("input");
        input9.type = "button";
        input9.className = "button";
        input9.value = "Delete";
        input9.id = "del_id";

        input9.onclick = function () {

            if (confirm('Are you sure you want to delete?')) {
                // Delete!
                var nTable = table.id;
                var fila = this.parentNode.parentNode;
                var tbody = table.getElementsByTagName("tbody")[0];
                tbody.removeChild(fila);

                // Guardamos los datos actuales en mcsprod
                dataSavedMachine();

                // Si al borrar una fila de datos nos quedamos con la tabla vacía, se oculta la tabla
                if (table.rows.length == 1) {
                    if (nTable.localeCompare("myTable-prod") == 0) {
                        document.getElementById("section-PROD").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-xcomp") == 0) {
                        document.getElementById("section-XCOMP").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-training") == 0) {
                        document.getElementById("section-Training").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-test") == 0) {
                        document.getElementById("section-Test").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-dev") == 0) {
                        document.getElementById("section-Dev").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-perftest") == 0) {
                        document.getElementById("section-PerfTest").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-sit") == 0) {
                        document.getElementById("section-SIT").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-uat") == 0) {
                        document.getElementById("section-UAT").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-rtest") == 0) {
                        document.getElementById("section-RTEST").style.display = 'none';
                    } else if (nTable.localeCompare("myTable-Decommission") == 0) {
                        document.getElementById("section-Decommission").style.display = 'none';
                    }
                }

                // Recorremos las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
                updateRows(table);

            }

        }

        cell9.appendChild(input9);
        cell9.className = "del";

        document.getElementById("cm_id").value = "";
        document.getElementById("status").value = "Select Status";
        document.getElementById("enviroment").value = "Select Enviroment";
        document.getElementById("stask_id").value = "";
        document.getElementById("summary").value = "";
        document.getElementById("date").value = "";
        document.getElementById("imp_name").value = "";
        document.getElementById("obs").value = "";
        document.getElementById("waiting_for").value = "";

        // Recorremos las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
        updateRows(table);
    }

    // Guardamos los datos actuales en mcsprod
    dataSavedMachine();
}

// Función que crea una nueva fila en la tabla de IMs		
function addFunctionIM() {

    // Mostramos la tabla por si estuviese oculta al no contar aún con ninguna fila
    document.getElementById("section-IM-table").style.display = 'block';

    var table = document.getElementById("myTable-ims");

    // Campo IM Ticket
    var input = document.createElement("input");
    input.type = "text";
    input.className = "im";
    input.value = document.getElementById("im_id").value;

    // Campo STASK number		
    var input2 = document.createElement("input");
    input2.type = "text";
    input2.className = "stask";
    input2.value = document.getElementById("im_stask_id").value;

    // Campo Summary
    var textarea3 = document.createElement("textarea");
    textarea3.type = "text";
    textarea3.className = "summary";
    textarea3.onkeyup = function () { textAreaAdjust(this) };
    textarea3.style = "overflow:hidden";
    textarea3.value = document.getElementById("im_summary").value;

    // Campo Implementer Name
    var input4 = document.createElement("input");
    input4.type = "text";
    input4.className = "implementer";
    input4.value = document.getElementById("im_imp_name").value;

    // Campo Observations
    var textarea5 = document.createElement("textarea");
    textarea5.type = "text";
    textarea5.className = "observations";
    textarea5.onkeyup = function () { textAreaAdjust(this) };
    textarea5.style = "overflow:hidden";
    textarea5.value = document.getElementById("im_obs").value;

    // Campo Waiting for
    var input6 = document.createElement("input");
    input6.type = "text";
    input6.className = "waiting";
    input6.setAttribute("list", "im_waiting_for_list");
    var datalist_input6 = document.createElement("datalist");
    datalist_input6.id = "im_waiting_for_list";
    var wfstatus = ["WIP", "Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
    wfstatus.forEach(function (item) {
        var option = document.createElement('option');
        option.value = item;
        datalist_input6.appendChild(option);
    });
    input6.value = document.getElementById("im_waiting_for").value;

    // Insertamos la fila al final de la tabla
    var row = table.insertRow(-1);

    var cell1 = row.insertCell(0);
    var cell2 = row.insertCell(1);
    var cell3 = row.insertCell(2);
    var cell4 = row.insertCell(3);
    var cell5 = row.insertCell(4);
    var cell6 = row.insertCell(5);
    var cell7 = row.insertCell(6);

    var checkWaiting = input6.value;
    if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
        input6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
        cell6.style = "background-color: #F3BDFF";
    }

    cell1.appendChild(input);
    cell2.appendChild(input2);
    cell3.appendChild(textarea3);
    cell4.appendChild(input4);
    cell5.appendChild(textarea5);
    cell6.appendChild(input6);
    cell6.appendChild(datalist_input6);
    cell1.className = "im";
    cell2.className = "stask";
    cell3.className = "summary";
    cell4.className = "implementer";
    cell5.className = "obs";
    cell6.className = "waiting";
    cell6.addEventListener('change', updateValueWaitingIMPM);

    // Campo Delete
    var input7 = document.createElement("input");
    input7.type = "button";
    input7.className = "button";
    input7.value = "Delete";
    input7.id = "del_id";
    input7.onclick = function () {
        if (confirm('Are you sure you want to delete?')) {
            // Delete!
            var fila = this.parentNode.parentNode;
            var tbody = table.getElementsByTagName("tbody")[0];
            tbody.removeChild(fila);

            // Guardamos los datos actuales en mcsprod
            dataSavedMachine();

            // Si al borrar una fila de datos nos quedamos con la tabla vacía, se oculta la tabla
            if (table.rows.length == 1) {
                document.getElementById("section-IM-table").style.display = 'none';
            }

        }
    }

    cell7.appendChild(input7);
    cell7.className = "del";

    document.getElementById("im_id").value = "";
    document.getElementById("im_stask_id").value = "";
    document.getElementById("im_summary").value = "";
    document.getElementById("im_imp_name").value = "";
    document.getElementById("im_obs").value = "";
    document.getElementById("im_waiting_for").value = "";

    // Guardamos los datos actuales en mcsprod
    dataSavedMachine();
}

// Función que crea una nueva fila en la tabla de PMs		
function addFunctionPM() {

    // Mostramos la tabla por si estuviese oculta al no contar aún con ninguna fila		
    document.getElementById("section-PM-table").style.display = 'block';

    var table = document.getElementById("myTable-pms");

    // Campo PM Ticket
    var input = document.createElement("input");
    input.type = "text";
    input.className = "pm";
    input.value = document.getElementById("pm_id").value;

    // Campo STASK number			
    var input2 = document.createElement("input");
    input2.type = "text";
    input2.className = "stask";
    input2.value = document.getElementById("pm_stask_id").value;

    // Campo Summary
    var textarea3 = document.createElement("textarea");
    textarea3.type = "text";
    textarea3.className = "summary";
    textarea3.onkeyup = function () { textAreaAdjust(this) };
    textarea3.style = "overflow:hidden";
    textarea3.value = document.getElementById("pm_summary").value;

    // Campo Implementer Name
    var input4 = document.createElement("input");
    input4.type = "text";
    input4.className = "implementer";
    input4.value = document.getElementById("pm_imp_name").value;

    // Campo Observations
    var textarea5 = document.createElement("textarea");
    textarea5.type = "text";
    textarea5.className = "observations";
    textarea5.onkeyup = function () { textAreaAdjust(this) };
    textarea5.style = "overflow:hidden";
    textarea5.value = document.getElementById("pm_obs").value;

    // Campo Waiting for
    var input6 = document.createElement("input");
    input6.type = "text";
    input6.className = "waiting";
    input6.setAttribute("list", "pm_waiting_for_list");
    var datalist_input6 = document.createElement("datalist");
    datalist_input6.id = "pm_waiting_for_list";
    var wfstatus = ["WIP", "Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
    wfstatus.forEach(function (item) {
        var option = document.createElement('option');
        option.value = item;
        datalist_input6.appendChild(option);
    });
    input6.value = document.getElementById("pm_waiting_for").value;

    // Insertamos la fila al final de la tabla
    var row = table.insertRow(-1);

    var cell1 = row.insertCell(0);
    var cell2 = row.insertCell(1);
    var cell3 = row.insertCell(2);
    var cell4 = row.insertCell(3);
    var cell5 = row.insertCell(4);
    var cell6 = row.insertCell(5);
    var cell7 = row.insertCell(6);

    var checkWaiting = input6.value;
    if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
        input6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
        cell6.style = "background-color: #F3BDFF";
    }

    cell1.appendChild(input);
    cell2.appendChild(input2);
    cell3.appendChild(textarea3);
    cell4.appendChild(input4);
    cell5.appendChild(textarea5);
    cell6.appendChild(input6);
    cell6.appendChild(datalist_input6);
    cell1.className = "pm";
    cell2.className = "stask";
    cell3.className = "summary";
    cell4.className = "implementer";
    cell5.className = "obs";
    cell6.className = "waiting";
    cell6.addEventListener('change', updateValueWaitingIMPM);

    // Campo Delete
    var input7 = document.createElement("input");
    input7.type = "button";
    input7.className = "button";
    input7.value = "Delete";
    input7.id = "del_id";
    input7.onclick = function () {
        if (confirm('Are you sure you want to delete?')) {
            // Delete!
            var fila = this.parentNode.parentNode;
            var tbody = table.getElementsByTagName("tbody")[0];
            tbody.removeChild(fila);

            // Guardamos los datos actuales en mcsprod
            dataSavedMachine();

            // Si al borrar una fila de datos nos quedamos con la tabla vacía, se oculta la tabla
            if (table.rows.length == 1) {
                document.getElementById("section-PM-table").style.display = 'none';
            }
        }
    }

    cell7.appendChild(input7);
    cell7.className = "del";

    document.getElementById("pm_id").value = "";
    document.getElementById("pm_stask_id").value = "";
    document.getElementById("pm_summary").value = "";
    document.getElementById("pm_imp_name").value = "";
    document.getElementById("pm_obs").value = "";
    document.getElementById("pm_waiting_for").value = "";

    // Guardamos los datos actuales en mcsprod
    dataSavedMachine();
}

// Función que crea una nueva fila en la tabla de Blackouts
function addFunctionBLACKOUT() {

    // Mostramos la tabla por si estuviese oculta al no contar aún con ninguna fila		
    document.getElementById("section-BLACKOUT-table").style.display = 'block';

    var table = document.getElementById("myTable-blackouts");

    // Campo CM Ticket
    var input = document.createElement("input");
    input.type = "text";
    input.className = "cm";
    input.value = document.getElementById("blackout_cm_id").value;

    // Campo STASK number		
    var input2 = document.createElement("input");
    input2.type = "text";
    input2.className = "stask";
    input2.value = document.getElementById("blackout_stask_id").value;

    // Campo IR number
    var input3 = document.createElement("input");
    input3.type = "text";
    input3.className = "ir";
    input3.value = document.getElementById("ir_id").value;

    // Campo Summary
    var textarea4 = document.createElement("textarea");
    textarea4.type = "text";
    textarea4.className = "summary";
    textarea4.onkeyup = function () { textAreaAdjust(this) };
    textarea4.style = "overflow:hidden";
    textarea4.value = document.getElementById("blackout_summary").value;

    // Insertamos la fila al final de la tabla
    var row = table.insertRow(-1);

    var cell1 = row.insertCell(0);
    var cell2 = row.insertCell(1);
    var cell3 = row.insertCell(2);
    var cell4 = row.insertCell(3);
    var cell5 = row.insertCell(4);

    cell1.appendChild(input);
    cell2.appendChild(input2);
    cell3.appendChild(input3);
    cell4.appendChild(textarea4);
    cell1.className = "cm";
    cell2.className = "stask";
    cell3.className = "ir";
    cell4.className = "summary";

    // Campo Delete
    var input5 = document.createElement("input");
    input5.type = "button";
    input5.className = "button";
    input5.value = "Delete";
    input5.id = "del_id";
    input5.onclick = function () {
        if (confirm('Are you sure you want to delete?')) {
            // Delete!
            var fila = this.parentNode.parentNode;
            var tbody = table.getElementsByTagName("tbody")[0];
            tbody.removeChild(fila);

            // Guardamos los datos actuales en mcsprod
            dataSavedMachine();

            // Si al borrar una fila de datos nos quedamos con la tabla vacía, se oculta la tabla
            if (table.rows.length == 1) {
                document.getElementById("section-BLACKOUT-table").style.display = 'none';
            }
        }
    }

    cell5.appendChild(input5);
    cell5.className = "del";

    document.getElementById("blackout_cm_id").value = "";
    document.getElementById("blackout_stask_id").value = "";
    document.getElementById("ir_id").value = "";
    document.getElementById("blackout_summary").value = "";

    // Guardamos los datos actuales en mcsprod
    dataSavedMachine();
}


/* 
 * Función que dependiendo de si estamos en Modo Visor o Editor muestra u oculta elementos
 * Hasta que no se selecciona el Modo Visor o Editor no se muestra el botón de input para cargar el excel 
 * Solamente en Modo Editor se muestran los elementos para añadir linea y el botón de exportar el excel 
 */
function displayDivDemo(id1, id2, id3, id4, id5, id6, id7, id8, id9, id10, elementValue) {
    document.getElementById(id1).style.display = ((elementValue.value == 1) || (elementValue.value == 2)) ? 'block' : 'none';
    document.getElementById(id2).style.display = elementValue.value == 1 ? 'block' : 'none';
    document.getElementById(id3).style.display = elementValue.value == 1 ? 'block' : 'none';
    document.getElementById(id4).style.display = elementValue.value == 1 ? 'block' : 'none';
    document.getElementById(id5).style.display = elementValue.value == 1 ? 'block' : 'none';
    document.getElementById(id6).style.display = elementValue.value == 1 ? 'inline-block' : 'none';
    document.getElementById(id7).style.display = elementValue.value == 1 ? 'inline-block' : 'none';
    document.getElementById(id8).style.display = elementValue.value == 2 ? 'inline-block' : 'none';
    document.getElementById(id9).style.display = elementValue.value == 1 ? 'inline-block' : 'none';
    document.getElementById(id10).style.display = elementValue.value == 1 ? 'inline-block' : 'none';
}

// Función para leer los datos del Sheet 1 - CMs del excel cargado a través del input
function createSheet1(sheet_data, sel_mode) {

    // Si contiene algo
    if (sheet_data.length > 0) {

        // Contador de filas
        var a = 0;

        // Recorremos las filas del excel una a una - Mientras el contador de filas sea menor que las filas existentes en el excel
        while (a < sheet_data.length) {

            // Conocemos el entorno (Prod, XCOMP, ...)
            var table_name = sheet_data[a][0];

            // Conociendo ya el entorno, mostramos esa tabla porque por defecto está oculta al no contar con ninguna fila
            if (table_name.localeCompare("Prod") == 0) {
                var get_table = document.getElementById("myTable-prod");
                document.getElementById("section-PROD").style.display = 'block';
            } else if (table_name.localeCompare("XCOMP") == 0) {
                var get_table = document.getElementById("myTable-xcomp");
                document.getElementById("section-XCOMP").style.display = 'block';
            } else if (table_name.localeCompare("Training") == 0) {
                var get_table = document.getElementById("myTable-training");
                document.getElementById("section-Training").style.display = 'block';
            } else if (table_name.localeCompare("Test") == 0) {
                var get_table = document.getElementById("myTable-test");
                document.getElementById("section-Test").style.display = 'block';
            } else if (table_name.localeCompare("Dev") == 0) {
                var get_table = document.getElementById("myTable-dev");
                document.getElementById("section-Dev").style.display = 'block';
            } else if (table_name.localeCompare("Perf Test") == 0) {
                var get_table = document.getElementById("myTable-perftest");
                document.getElementById("section-PerfTest").style.display = 'block';
            } else if (table_name.localeCompare("SIT") == 0) {
                var get_table = document.getElementById("myTable-sit");
                document.getElementById("section-SIT").style.display = 'block';
            } else if (table_name.localeCompare("UAT") == 0) {
                var get_table = document.getElementById("myTable-uat");
                document.getElementById("section-UAT").style.display = 'block';
            } else if (table_name.localeCompare("RTEST") == 0) {
                var get_table = document.getElementById("myTable-rtest");
                document.getElementById("section-RTEST").style.display = 'block';
            } else if (table_name.localeCompare("Decommission") == 0) {
                var get_table = document.getElementById("myTable-Decommission");
                document.getElementById("section-Decommission").style.display = 'block';
            }

            // Nos saltamos la fila de campos (CM Ticket, Status, ...) y pasamos a la de datos
            a = a + 2;

            // Mientras el contador de filas sea menor que las filas existentes en el excel y no sea la fila que indica el entorno (Prod, XCOMP, ...)
            while ((a < sheet_data.length) && (sheet_data[a].length != 1)) {

                // Introducimos la fila al final de la tabla
                var row_table = get_table.insertRow(-1);
                var cell1 = row_table.insertCell(0);
                var cell2 = row_table.insertCell(1);
                var cell3 = row_table.insertCell(2);
                var cell4 = row_table.insertCell(3);
                var cell5 = row_table.insertCell(4);
                var cell6 = row_table.insertCell(5);
                var cell7 = row_table.insertCell(6);
                var cell8 = row_table.insertCell(7);

                // Modo Editor
                if (Number(sel_mode) == 1) {

                    // Campo CM Ticket
                    var input = document.createElement("input");
                    input.type = "text";
                    input.className = "cm";
                    input.value = sheet_data[a][0];

                    // Campo Status		
                    var select2 = document.createElement("select");
                    select2.id = "status";
                    select2.className = "status";
                    var array_status = ["Select Status", "Draft", "Pending approval", "Pending scheduling", "Scheduled", "Completed", "WIP"];
                    for (var i = 0; i < array_status.length; i++) {
                        var optionSt = document.createElement("option");
                        optionSt.value = array_status[i];
                        optionSt.text = array_status[i];
                        select2.appendChild(optionSt);
                    }
                    select2.value = sheet_data[a][1];
                    if (sheet_data[a][1].localeCompare("Draft") == 0) {
                        select2.style = "background-color: yellow; color: black;";
                    } else if (sheet_data[a][1].localeCompare("Pending approval") == 0) {
                        select2.style = "background-color: #81D0E5; color: black;";
                    } else if (sheet_data[a][1].localeCompare("Pending scheduling") == 0) {
                        select2.style = "background-color: #00C5FF; color: black;";
                    } else if (sheet_data[a][1].localeCompare("Scheduled") == 0) {
                        select2.style = "background-color: #9F01A9; color: white;";
                    } else if (sheet_data[a][1].localeCompare("Completed") == 0) {
                        select2.style = "background-color: red; color: white;";
                    } else if (sheet_data[a][1].localeCompare("WIP") == 0) {
                        select2.style = "background-color: #81F07A; color: black;";
                    }

                    // Campo STASK number		
                    var input3 = document.createElement("input");
                    input3.type = "text";
                    input3.className = "stask";
                    input3.value = sheet_data[a][2];

                    // Campo Summary
                    var textarea4 = document.createElement("textarea");
                    textarea4.type = "text";
                    textarea4.className = "summary";
                    textarea4.onkeyup = function () { textAreaAdjust(this) };
                    textarea4.style = "overflow:hidden";
                    textarea4.value = sheet_data[a][3];

                    // Campo Scheduled Date Madrid time	
                    var input5 = document.createElement("input");
                    input5.type = "text";
                    input5.id = "date-order";
                    input5.className = "date";
                    input5.value = sheet_data[a][4];

                    var checkDate = sheet_data[a][4];
                    var checkDate_2 = Number(checkDate.substring(0, 2));
                    var checkDate_4 = Number(checkDate.substring(3, 5));
                    var checkDate_6 = Number(checkDate.substring(6, 10));
                    var today = getToday();
                    var day = today[0],
                        month = today[1],
                        year = today[2];

                    if ((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) {
                        input5.style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;"; // Si es el día de hoy aparecerá en rojo
                        cell5.style = "background-color: #FF9F9F";
                    }

                    // Campo Implementer Name
                    var input6 = document.createElement("input");
                    input6.type = "text";
                    input6.className = "implementer";
                    input6.value = sheet_data[a][5];

                    // Campo Observations	
                    var textarea7 = document.createElement("textarea");
                    textarea7.type = "text";
                    textarea7.className = "observations";
                    textarea7.onkeyup = function () { textAreaAdjust(this) };
                    textarea7.style = "overflow:hidden";
                    textarea7.value = sheet_data[a][6];

                    // Campo Waiting for
                    var input8 = document.createElement("input");
                    input8.type = "text";
                    input8.className = "waiting";
                    input8.setAttribute("list", "waiting_for_list");
                    var datalist_input8 = document.createElement("datalist");
                    datalist_input8.id = "waiting_for_list";
                    var wfstatus = ["Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
                    wfstatus.forEach(function (item) {
                        var option = document.createElement('option');
                        option.value = item;
                        datalist_input8.appendChild(option);
                    });
                    input8.value = sheet_data[a][7];

                    if (sheet_data[a][7].localeCompare("Tier 2 actions") == 0) {
                        input8.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                        cell8.style = "background-color: #F3BDFF";
                    }

                    cell1.appendChild(input);
                    cell2.appendChild(select2);
                    cell3.appendChild(input3);
                    cell4.appendChild(textarea4);
                    cell5.appendChild(input5);
                    cell6.appendChild(input6);
                    cell7.appendChild(textarea7);
                    cell8.appendChild(input8);
                    cell8.appendChild(datalist_input8);
                    cell1.className = "cm";
                    cell1.addEventListener('change', updateValueCM);
                    cell2.className = "status";
                    cell2.addEventListener('change', updateValueStatus);
                    cell3.className = "stask";
                    cell4.className = "summary";
                    cell5.className = "date";
                    cell5.id = "date-container";
                    cell5.addEventListener('change', updateValue);
                    cell6.className = "implementer";
                    cell7.className = "obs";
                    cell8.className = "waiting";
                    cell8.addEventListener('change', updateValueWaiting);

                    var cell9 = row_table.insertCell(8);

                    // Campo Delete
                    var input9 = createInputDel(get_table);

                    cell9.appendChild(input9);
                    cell9.className = "del";

                } else { // Modo Visor

                    // Campo CM Ticket
                    var text = document.createElement("h5");
                    text.className = "cm";
                    text.textContent = sheet_data[a][0];

                    // Campo Status	
                    var text2 = document.createElement("h5");
                    text2.className = "status";
                    text2.textContent = sheet_data[a][1];
                    if (sheet_data[a][1].localeCompare("Draft") == 0) {
                        text2.style = "background-color: yellow; color: black;";
                        cell2.style = "background-color: yellow;";
                    } else if (sheet_data[a][1].localeCompare("Pending approval") == 0) {
                        text2.style = "background-color: #81D0E5; color: black;";
                        cell2.style = "background-color: #81D0E5;";
                    } else if (sheet_data[a][1].localeCompare("Pending scheduling") == 0) {
                        text2.style = "background-color: #00C5FF; color: black;";
                        cell2.style = "background-color: #00C5FF;";
                    } else if (sheet_data[a][1].localeCompare("Scheduled") == 0) {
                        text2.style = "background-color: #9F01A9; color: white;";
                        cell2.style = "background-color: #9F01A9;";
                    } else if (sheet_data[a][1].localeCompare("Completed") == 0) {
                        text2.style = "background-color: red; color: white;";
                        cell2.style = "background-color: red;";
                    } else if (sheet_data[a][1].localeCompare("WIP") == 0) {
                        text2.style = "background-color: #81F07A; color: black;";
                        cell2.style = "background-color: #81F07A;";
                    }

                    // Campo STASK number	
                    var text3 = document.createElement("h5");
                    text3.className = "stask";
                    text3.textContent = sheet_data[a][2];

                    // Campo Summary
                    var text4 = document.createElement("h5");
                    text4.className = "summary";
                    text4.textContent = sheet_data[a][3];

                    // Campo Scheduled Date Madrid time	
                    var text5 = document.createElement("h5");
                    text5.className = "date";
                    text5.textContent = sheet_data[a][4];

                    var checkDate = sheet_data[a][4];
                    var checkDate_2 = Number(checkDate.substring(0, 2));
                    var checkDate_4 = Number(checkDate.substring(3, 5));
                    var checkDate_6 = Number(checkDate.substring(6, 10));
                    var today = getToday();
                    var day = today[0],
                        month = today[1],
                        year = today[2];

                    if ((Number(checkDate_2) == Number(day)) && (Number(checkDate_4) == Number(month)) && (Number(checkDate_6) == Number(year))) {
                        text5.style = "color: #C20000; font-weight: bold; background-color: #FF9F9F;"; // Si es el día de hoy aparecerá en rojo
                        cell5.style = "background-color: #FF9F9F";
                    }


                    // Campo Implementer Name					
                    var text6 = document.createElement("h5");
                    text6.className = "implementer";
                    text6.textContent = sheet_data[a][5];

                    // Campo Observations
                    var text7 = document.createElement("h5");
                    text7.className = "observations";
                    text7.textContent = sheet_data[a][6];

                    // Campo Waiting for
                    var text8 = document.createElement("h5");
                    text8.className = "waiting";
                    text8.textContent = sheet_data[a][7];

                    if (sheet_data[a][7].localeCompare("Tier 2 actions") == 0) {
                        text8.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                        cell8.style = "background-color: #F3BDFF";
                    }

                    cell1.appendChild(text);
                    cell2.appendChild(text2);
                    cell3.appendChild(text3);
                    cell4.appendChild(text4);
                    cell5.appendChild(text5);
                    cell6.appendChild(text6);
                    cell7.appendChild(text7);
                    cell8.appendChild(text8);
                    cell1.className = "cm";
                    cell2.className = "status";
                    cell3.className = "stask";
                    cell4.className = "summary";
                    cell5.className = "date";
                    cell6.className = "implementer";
                    cell7.className = "obs";
                    cell8.className = "waiting";

                }

                // Siguiente fila
                a++;

            }

        }

    }

}

// Función para leer los datos del Sheet 2 - IMs del excel cargado a través del input	
function createSheet2(sheet_data_2, sel_mode) {

    // Si contiene algo
    if (sheet_data_2.length > 0) {

        // Si tienes más filas que la del nombre de los campos (IM Ticket, STASK number, ...)) -> mostramos esa tabla porque por defecto está oculta al no contar con ninguna fila 
        if (sheet_data_2.length > 1) { document.getElementById("section-IM-table").style.display = 'block'; }

        // Contador de filas
        var b = 1;

        // Recorremos las filas del excel una a una - Mientras el contador de filas sea menor que las filas existentes en el excel
        while (b < sheet_data_2.length) {

            var get_table = document.getElementById("myTable-ims");

            // Introducimos la fila al final de la tabla
            var row_table = get_table.insertRow(-1);
            var cell1 = row_table.insertCell(0);
            var cell2 = row_table.insertCell(1);
            var cell3 = row_table.insertCell(2);
            var cell4 = row_table.insertCell(3);
            var cell5 = row_table.insertCell(4);
            var cell6 = row_table.insertCell(5);

            // Modo Editor
            if (Number(sel_mode) == 1) {

                // Campo IM Ticket
                var input = document.createElement("input");
                input.type = "text";
                input.className = "im";
                input.value = sheet_data_2[b][0];

                // Campo STASK number	
                var input2 = document.createElement("input");
                input2.type = "text";
                input2.className = "stask";
                input2.value = sheet_data_2[b][1];

                // Campo Summary		
                var textarea3 = document.createElement("textarea");
                textarea3.type = "text";
                textarea3.className = "summary";
                textarea3.onkeyup = function () { textAreaAdjust(this) };
                textarea3.style = "overflow:hidden";
                textarea3.value = sheet_data_2[b][2];

                // Campo Implementer Name
                var input4 = document.createElement("input");
                input4.type = "text";
                input4.className = "implementer";
                input4.value = sheet_data_2[b][3];

                // Campo Observations
                var textarea5 = document.createElement("textarea");
                textarea5.type = "text";
                textarea5.className = "observations";
                textarea5.onkeyup = function () { textAreaAdjust(this) };
                textarea5.style = "overflow:hidden";
                textarea5.value = sheet_data_2[b][4];

                // Campo Waiting for
                var input6 = document.createElement("input");
                input6.type = "text";
                input6.className = "waiting";
                input6.setAttribute("list", "im_waiting_for_list");
                var datalist_input6 = document.createElement("datalist");
                datalist_input6.id = "im_waiting_for_list";
                var wfstatus = ["WIP", "Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
                wfstatus.forEach(function (item) {
                    var option = document.createElement('option');
                    option.value = item;
                    datalist_input6.appendChild(option);
                });
                input6.value = sheet_data_2[b][5];

                if (sheet_data_2[b][5].localeCompare("Tier 2 actions") == 0) {
                    input6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                    cell6.style = "background-color: #F3BDFF";
                }

                cell1.appendChild(input);
                cell2.appendChild(input2);
                cell3.appendChild(textarea3);
                cell4.appendChild(input4);
                cell5.appendChild(textarea5);
                cell6.appendChild(input6);
                cell6.appendChild(datalist_input6);
                cell1.className = "im";
                cell2.className = "stask";
                cell3.className = "summary";
                cell4.className = "implementer";
                cell5.className = "obs";
                cell6.className = "waiting";
                cell6.addEventListener('change', updateValueWaitingIMPM);

                var cell7 = row_table.insertCell(6);

                // Campo Delete
                var input7 = createInputDel(get_table);

                cell7.appendChild(input7);
                cell7.className = "del";

            } else { // Modo Visor

                // Campo IM Ticket
                var text = document.createElement("h5");
                text.className = "im";
                text.textContent = sheet_data_2[b][0];

                // Campo STASK number	
                var text2 = document.createElement("h5");
                text2.className = "stask";
                text2.textContent = sheet_data_2[b][1];

                // Campo Summary
                var text3 = document.createElement("h5");
                text3.className = "summary";
                text3.textContent = sheet_data_2[b][2];

                // Campo Implementer Name
                var text4 = document.createElement("h5");
                text4.className = "implementer";
                text4.textContent = sheet_data_2[b][3];

                // Campo Observations
                var text5 = document.createElement("h5");
                text5.className = "observations";
                text5.textContent = sheet_data_2[b][4];

                // Campo Waiting for
                var text6 = document.createElement("h5");
                text6.className = "waiting";
                text6.textContent = sheet_data_2[b][5];

                if (sheet_data_2[b][5].localeCompare("Tier 2 actions") == 0) {
                    text6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                    cell6.style = "background-color: #F3BDFF";
                }

                cell1.appendChild(text);
                cell2.appendChild(text2);
                cell3.appendChild(text3);
                cell4.appendChild(text4);
                cell5.appendChild(text5);
                cell6.appendChild(text6);
                cell1.className = "im";
                cell2.className = "stask";
                cell3.className = "summary";
                cell4.className = "implementer";
                cell5.className = "obs";
                cell6.className = "waiting";

            }

            // Siguiente fila
            b++;

        }

    }

}

// Función para leer los datos del Sheet 3 - PMs del excel cargado a través del input	
function createSheet3(sheet_data_3, sel_mode) {

    // Si contiene algo
    if (sheet_data_3.length > 0) {

        // Si tienes más filas que la del nombre de los campos (IM Ticket, STASK number, ...)) -> mostramos esa tabla porque por defecto está oculta al no contar con ninguna fila 
        if (sheet_data_3.length > 1) { document.getElementById("section-PM-table").style.display = 'block'; }

        // Contador de filas
        var c = 1;

        // Recorremos las filas del excel una a una - Mientras el contador de filas sea menor que las filas existentes en el excel			
        while (c < sheet_data_3.length) {

            var get_table = document.getElementById("myTable-pms");

            // Introducimos la fila al final de la tabla
            var row_table = get_table.insertRow(-1);
            var cell1 = row_table.insertCell(0);
            var cell2 = row_table.insertCell(1);
            var cell3 = row_table.insertCell(2);
            var cell4 = row_table.insertCell(3);
            var cell5 = row_table.insertCell(4);
            var cell6 = row_table.insertCell(5);

            // Modo Editor
            if (Number(sel_mode) == 1) {

                // Campo PM Ticket
                var input = document.createElement("input");
                input.type = "text";
                input.className = "pm";
                input.value = sheet_data_3[c][0];

                // Campo STASK number	
                var input2 = document.createElement("input");
                input2.type = "text";
                input2.className = "stask";
                input2.value = sheet_data_3[c][1];

                // Campo Summary
                var textarea3 = document.createElement("textarea");
                textarea3.type = "text";
                textarea3.className = "summary";
                textarea3.onkeyup = function () { textAreaAdjust(this) };
                textarea3.style = "overflow:hidden";
                textarea3.value = sheet_data_3[c][2];

                // Campo Implementer Name
                var input4 = document.createElement("input");
                input4.type = "text";
                input4.className = "implementer";
                input4.value = sheet_data_3[c][3];

                // Campo Observations
                var textarea5 = document.createElement("textarea");
                textarea5.type = "text";
                textarea5.className = "observations";
                textarea5.onkeyup = function () { textAreaAdjust(this) };
                textarea5.style = "overflow:hidden";
                textarea5.value = sheet_data_3[c][4];

                // Campo Waiting for
                var input6 = document.createElement("input");
                input6.type = "text";
                input6.className = "waiting";
                input6.setAttribute("list", "pm_waiting_for_list");
                var datalist_input6 = document.createElement("datalist");
                datalist_input6.id = "pm_waiting_for_list";
                var wfstatus = ["WIP", "Tier 2 actions", "Customer feedback", "PSEs feedback", "T3DBA feedback", "T3FMW feedback", "T3SYS feedback"];
                wfstatus.forEach(function (item) {
                    var option = document.createElement('option');
                    option.value = item;
                    datalist_input6.appendChild(option);
                });
                input6.value = sheet_data_3[c][5];

                if (sheet_data_3[c][5].localeCompare("Tier 2 actions") == 0) {
                    input6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                    cell6.style = "background-color: #F3BDFF";
                }

                cell1.appendChild(input);
                cell2.appendChild(input2);
                cell3.appendChild(textarea3);
                cell4.appendChild(input4);
                cell5.appendChild(textarea5);
                cell6.appendChild(input6);
                cell6.appendChild(datalist_input6);
                cell1.className = "pm";
                cell2.className = "stask";
                cell3.className = "summary";
                cell4.className = "implementer";
                cell5.className = "obs";
                cell6.className = "waiting";
                cell6.addEventListener('change', updateValueWaitingIMPM);

                var cell7 = row_table.insertCell(6);

                // Campo Delete
                var input7 = createInputDel(get_table);

                cell7.appendChild(input7);
                cell7.className = "del";

            } else { // Modo Visor

                // Campo PM Ticket
                var text = document.createElement("h5");
                text.className = "pm";
                text.textContent = sheet_data_3[c][0];

                // Campo STASK number	
                var text2 = document.createElement("h5");
                text2.className = "stask";
                text2.textContent = sheet_data_3[c][1];

                // Campo Summary
                var text3 = document.createElement("h5");
                text3.className = "summary";
                text3.textContent = sheet_data_3[c][2];

                // Campo Implementer Name
                var text4 = document.createElement("h5");
                text4.className = "implementer";
                text4.textContent = sheet_data_3[c][3];

                // Campo Observations
                var text5 = document.createElement("h5");
                text5.className = "observations";
                text5.textContent = sheet_data_3[c][4];

                // Campo Waiting for
                var text6 = document.createElement("h5");
                text6.className = "waiting";
                text6.textContent = sheet_data_3[c][5];

                if (sheet_data_3[c][5].localeCompare("Tier 2 actions") == 0) {
                    text6.style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
                    cell6.style = "background-color: #F3BDFF";
                }

                cell1.appendChild(text);
                cell2.appendChild(text2);
                cell3.appendChild(text3);
                cell4.appendChild(text4);
                cell5.appendChild(text5);
                cell6.appendChild(text6);
                cell1.className = "pm";
                cell2.className = "stask";
                cell3.className = "summary";
                cell4.className = "implementer";
                cell5.className = "obs";
                cell6.className = "waiting";

            }

            // Siguiente fila
            c++;

        }

    }

}

// Función para leer los datos del Sheet 4 - Blackouts del excel cargado a través del input		
function createSheet4(sheet_data_4, sel_mode) {

    // Si contiene algo
    if (sheet_data_4.length > 0) {

        // Si tienes más filas que la del nombre de los campos (IM Ticket, STASK number, ...)) -> mostramos esa tabla porque por defecto está oculta al no contar con ninguna fila 			
        if (sheet_data_4.length > 1) { document.getElementById("section-BLACKOUT-table").style.display = 'block'; }

        // Contador de filas
        var d = 1;

        // Recorremos las filas del excel una a una - Mientras el contador de filas sea menor que las filas existentes en el excel			
        while (d < sheet_data_4.length) {

            var get_table = document.getElementById("myTable-blackouts");

            // Introducimos la fila al final de la tabla				
            var row_table = get_table.insertRow(-1);
            var cell1 = row_table.insertCell(0);
            var cell2 = row_table.insertCell(1);
            var cell3 = row_table.insertCell(2);
            var cell4 = row_table.insertCell(3);

            // Modo Editor
            if (Number(sel_mode) == 1) {

                // Campo CM Ticket
                var input = document.createElement("input");
                input.type = "text";
                input.className = "cm";
                input.value = sheet_data_4[d][0];

                // Campo STASK number
                var input2 = document.createElement("input");
                input2.type = "text";
                input2.className = "stask";
                input2.value = sheet_data_4[d][1];

                // Campo IR number	
                var input3 = document.createElement("input");
                input3.type = "text";
                input3.className = "ir";
                input3.value = sheet_data_4[d][2];

                // Campo Summary
                var textarea4 = document.createElement("textarea");
                textarea4.type = "text";
                textarea4.className = "summary";
                textarea4.onkeyup = function () { textAreaAdjust(this) };
                textarea4.style = "overflow:hidden";
                textarea4.value = sheet_data_4[d][3];


                cell1.appendChild(input);
                cell2.appendChild(input2);
                cell3.appendChild(input3);
                cell4.appendChild(textarea4);
                cell1.className = "cm";
                cell2.className = "stask";
                cell3.className = "ir";
                cell4.className = "summary";

                var cell5 = row_table.insertCell(4);

                // Campo Delete
                var input5 = createInputDel(get_table);

                cell5.appendChild(input5);
                cell5.className = "del";

            } else { // Modo Visor

                // Campo CM Ticket
                var text = document.createElement("h5");
                text.className = "cm";
                text.textContent = sheet_data_4[d][0];

                // Campo STASK number
                var text2 = document.createElement("h5");
                text2.className = "stask";
                text2.textContent = sheet_data_4[d][1];

                // Campo IR number
                var text3 = document.createElement("h5");
                text3.className = "ir";
                text3.textContent = sheet_data_4[d][2];

                // Campo Summary
                var text4 = document.createElement("h5");
                text4.className = "summary";
                text4.textContent = sheet_data_4[d][3];


                cell1.appendChild(text);
                cell2.appendChild(text2);
                cell3.appendChild(text3);
                cell4.appendChild(text4);
                cell1.className = "cm";
                cell2.className = "stask";
                cell3.className = "ir";
                cell4.className = "summary";

            }

            // Siguiente fila
            d++;

        }

    }

}

// Funcion que crea el boton Delete 	 	
function createInputDel(table) {

    var input = document.createElement("input");
    input.type = "button";
    input.className = "button";
    input.value = "Delete";
    input.id = "del_id";

    input.onclick = function () {

        if (confirm('Are you sure you want to delete?')) {
            // Delete!
            var nTable = table.id;
            var fila = this.parentNode.parentNode;
            var tbody = table.getElementsByTagName("tbody")[0];
            tbody.removeChild(fila);

            // Guardamos los datos actuales en mcsprod
            dataSavedMachine();

            // Si al borrar una fila de datos nos quedamos con la tabla vacía, se oculta la tabla
            if (table.rows.length == 1) {
                if (nTable.localeCompare("myTable-prod") == 0) {
                    document.getElementById("section-PROD").style.display = 'none';
                } else if (nTable.localeCompare("myTable-xcomp") == 0) {
                    document.getElementById("section-XCOMP").style.display = 'none';
                } else if (nTable.localeCompare("myTable-training") == 0) {
                    document.getElementById("section-Training").style.display = 'none';
                } else if (nTable.localeCompare("myTable-test") == 0) {
                    document.getElementById("section-Test").style.display = 'none';
                } else if (nTable.localeCompare("myTable-dev") == 0) {
                    document.getElementById("section-Dev").style.display = 'none';
                } else if (nTable.localeCompare("myTable-perftest") == 0) {
                    document.getElementById("section-PerfTest").style.display = 'none';
                } else if (nTable.localeCompare("myTable-sit") == 0) {
                    document.getElementById("section-SIT").style.display = 'none';
                } else if (nTable.localeCompare("myTable-uat") == 0) {
                    document.getElementById("section-UAT").style.display = 'none';
                } else if (nTable.localeCompare("myTable-rtest") == 0) {
                    document.getElementById("section-RTEST").style.display = 'none';
                } else if (nTable.localeCompare("myTable-Decommission") == 0) {
                    document.getElementById("section-Decommission").style.display = 'none';
                } else if (nTable.localeCompare("myTable-ims") == 0) {
                    document.getElementById("section-IM-table").style.display = 'none';
                } else if (nTable.localeCompare("myTable-pms") == 0) {
                    document.getElementById("section-PM-table").style.display = 'none';
                } else if (nTable.localeCompare("myTable-blackouts") == 0) {
                    document.getElementById("section-BLACKOUT-table").style.display = 'none';
                }
            }

            if ((nTable.localeCompare("myTable-ims") != 0) && (nTable.localeCompare("myTable-pms") != 0) && (nTable.localeCompare("myTable-blackouts") != 0)) {

                // Recorremos las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
                updateRows(table);

            }
        }

    }
    return input;
}

// Función que carga el HO que actualmente está guardado en mcsprod 
function createHandover(sel_mode) {

    // Carga los avisos permanentes
    var req = new XMLHttpRequest();
    req.open("GET", "EMA_Pending_Report_Handover_1.txt", true);

    req.onreadystatechange = function (e) {
        document.getElementById("avisosPermanentes").value = req.responseText;
        document.getElementById("avisosPermanentes").onkeyup = function () { textAreaAdjust(this) };
    }

    req.send();

    // Carga las cosas pendientes
    var req2 = new XMLHttpRequest();
    req2.open("GET", "EMA_Pending_Report_Handover_2.txt", true);

    req2.onreadystatechange = function (e) {
        document.getElementById("cambiosT2actions").value = req2.responseText;
        document.getElementById("cambiosT2actions").onkeyup = function () { textAreaAdjust(this) };
    }

    req2.send();

    // Si estamos en Modo Visor no deja que el HO pueda editarse
    if (Number(sel_mode) == 2) {
        document.getElementById("avisosPermanentes").readOnly = "true";
        document.getElementById("cambiosT2actions").readOnly = "true";
    }
}

// Funcion que comprueba si la fecha programada es hoy al editar el input de Scheduled Date Madrid time
function updateValue(e) {

    // Ordenamos las filas de la tabla 
    var num_row = this.parentNode.rowIndex;
    var y = e.target.value;
    var current_row = e.target.parentNode.parentNode;
    var table_id = e.target.parentNode.parentNode.parentNode.parentNode.id;
    var table = document.getElementById(table_id);
    var tbody = table.getElementsByTagName("tbody")[0];
    var q = 1;

    if (table.rows.length - 1 > 1) {
        if (y === "") {
            tbody.insertBefore(current_row, table.rows[table.rows.length - 1]);
            tbody.insertBefore(table.rows[table.rows.length - 1], table.rows[table.rows.length - 2]);
            q = table.rows.length;
        } else {
            var end = false;

            // Conocemos el día del campo "Scheduled Date Madrid time" de la Fila-Y
            var y_2 = Number(y.substring(0, 2));
            // Conocemos el mes del campo "Scheduled Date Madrid time" de la Fila-Y
            var y_4 = Number(y.substring(3, 5));
            // Conocemos el año del campo "Scheduled Date Madrid time" de la Fila-Y
            var y_6 = Number(y.substring(6, 10));
            // Conocemos la hora del campo "Scheduled Date Madrid time" de la Fila-Y
            var y_8 = Number(y.substring(11, 13));
            // Conocemos los minutos del campo "Scheduled Date Madrid time" de Fila-Y
            var y_10 = Number(y.substring(14, 16));

            // Recorremos todas las filas para ver en que posicion debe ir la fila editada
            while (end == false) {
                if (q != num_row) {

                    // Campo "Scheduled Date Madrid time" de una fila que ya existe -  A partir de la primera fila
                    var x = table.rows[q].cells[4].children[0].value;

                    // Conocemos el día del campo "Scheduled Date Madrid time" de Fila-X
                    var x_2 = Number(x.substring(0, 2));
                    // Conocemos el mes del campo "Scheduled Date Madrid time" de Fila-X
                    var x_4 = Number(x.substring(3, 5));
                    // Conocemos el año del campo "Scheduled Date Madrid time" de Fila-X
                    var x_6 = Number(x.substring(6, 10));
                    // Conocemos la hora del campo "Scheduled Date Madrid time" de Fila-X
                    var x_8 = Number(x.substring(11, 13));
                    // Conocemos los minutos del campo "Scheduled Date Madrid time" de Fila-X
                    var x_10 = Number(x.substring(14, 16));

                    /* 
                     * Si el campo "Scheduled Date Madrid time" de Fila-X está vacío
                     * Si el año de la Fila-Y es igual que el de Fila-X y el mes de Fila-Y es menor que el de Fila-X
                     * Si el año de la Fila-Y es menor que el de Fila-X
                     * Si el año y mes de la Fila-Y es igual que el de Fila-X y el día de Fila-Y es menor que el de Fila-X
                     * Si es el mismo día, mes y año de la Fila-Y y Fila-X y la hora de la Fila-Y es menor que la de Fila-X
                     * Si es el mismo día, mes, año y hora de la Fila-Y y Fila-X y los minutos de la Fila-Y es menor que la de Fila-X 
                     * Se añade la Fila-Y justo encima de la Fila-X 
                     */
                    if ((x.localeCompare("") == 0) || ((Number(y_6) == Number(x_6)) && (Number(y_4) < Number(x_4))) || (Number(y_6) < Number(x_6)) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) < Number(x_2))) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) == Number(x_2)) && (Number(y_8) < Number(x_8))) || ((Number(y_6) == Number(x_6)) && (Number(y_4) == Number(x_4)) && (Number(y_2) == Number(x_2)) && (Number(y_8) == Number(x_8)) && (Number(y_10) < Number(x_10)))) {
                        end = true;
                        tbody.insertBefore(current_row, table.rows[q]);
                    } else if (q == (table.rows.length - 1)) {
                        // Si ya estamos comparando con la última fila existente de la tabla y no se ha cumplido lo anterior, añadimos la Fila-Y al final de la tabla
                        end = true;
                        tbody.insertBefore(current_row, table.rows[table.rows.length - 1]);
                        tbody.insertBefore(table.rows[table.rows.length - 1], table.rows[table.rows.length]);
                    }
                }
                // Siguiente fila existente	
                q++;
            }
        }

        // Recorremos las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
        updateRows(table);

    }
}

// Funcion que comprueba el CM al editar el input del campo CM Ticket
function updateValueCM(e) {

    var num_row = this.parentNode.rowIndex;
    var table_id = e.target.parentNode.parentNode.parentNode.parentNode.id;
    var table = document.getElementById(table_id);
    var checkCm = table.rows[num_row].cells[0].children[0].value;
    var checkInp = table.rows[num_row].cells[0].children[0].className;
    var checkStyle = table.rows[num_row].cells[0].children[0].style.textDecoration;

    if ((checkCm.indexOf("*") >= 0) && (checkInp.localeCompare("cm") == 0) && (checkStyle.localeCompare("line-through") != 0)) {
        table.rows[num_row].cells[0].children[0].style = "text-decoration: line-through;";
        checkCm = checkCm.replace("*", "");
        table.rows[num_row].cells[0].children[0].value = checkCm;
        var input = document.createElement("input");
        input.type = "text";
        input.className = "cm";
        input.value = "";
        table.rows[num_row].cells[0].append(input);
    }

}

// Funcion que comprueba al editar el campo Status
function updateValueStatus(e) {

    var num_row = this.parentNode.rowIndex;
    var table_id = e.target.parentNode.parentNode.parentNode.parentNode.id;
    var table = document.getElementById(table_id);

    var checkStatus = table.rows[num_row].cells[1].children[0].value;
    if (checkStatus.localeCompare("Draft") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: yellow; color: black;";
    } else if (checkStatus.localeCompare("Pending approval") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: #81D0E5; color: black;";
    } else if (checkStatus.localeCompare("Pending scheduling") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: #00C5FF; color: black;";
    } else if (checkStatus.localeCompare("Scheduled") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: #9F01A9; color: white;";
    } else if (checkStatus.localeCompare("Completed") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: red; color: white;";
    } else if (checkStatus.localeCompare("WIP") == 0) {
        table.rows[num_row].cells[1].children[0].style = "background-color: #81F07A; color: black;";
    }

}

// Funcion que comprueba si el campo Waiting for es "Tier 2 actions" al editar dicho campo
function updateValueWaiting(e) {

    var num_row = this.parentNode.rowIndex;
    var table_id = e.target.parentNode.parentNode.parentNode.parentNode.id;
    var table = document.getElementById(table_id);

    var checkWaiting = table.rows[num_row].cells[7].children[0].value;
    if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
        table.rows[num_row].cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
        table.rows[num_row].cells[7].style = "background-color: #F3BDFF";
        table.rows[num_row].onmouseout = function () {
            this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";
            this.cells[7].style = "background-color: #F3BDFF;";
        };
        table.rows[num_row].onmouseover = function () {
            this.cells[7].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";;
            this.cells[7].style = "background-color: #F3BDFF;";
        };
    } else {

        if ((num_row % 2) == 0) {
            table.rows[num_row].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
            table.rows[num_row].cells[7].style = "background-color: #ffffff;";
            table.rows[num_row].onmouseout = function () {
                this.cells[7].children[0].style = "background-color: #ffffff";
                this.cells[7].style = "background-color: #ffffff;";
            };
        } else {
            table.rows[num_row].cells[7].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
            table.rows[num_row].cells[7].style = "background-color: #f2f2f2;";
            table.rows[num_row].onmouseout = function () {
                this.cells[7].children[0].style = "background-color: #f2f2f2";
                this.cells[7].style = "background-color: #f2f2f2;";
            };
        }
        table.rows[num_row].onmouseover = function () {
            this.cells[7].children[0].style = "background-color: #ddd;";
            this.cells[7].style = "background-color: #ddd;";
        };

    }

    // Recorremos las filas para que los colores de la columna Scheduled Date Madrid time y Waiting for sean los adecuados
    updateRows(table);

}

// Funcion que comprueba si el campo Waiting for es "Tier 2 actions" al editar dicho campo
function updateValueWaitingIMPM(e) {

    var num_row = this.parentNode.rowIndex;
    var table_id = e.target.parentNode.parentNode.parentNode.parentNode.id;
    var table = document.getElementById(table_id);

    var checkWaiting = table.rows[num_row].cells[5].children[0].value;
    if (checkWaiting.localeCompare("Tier 2 actions") == 0) {
        table.rows[num_row].cells[5].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;"; // Si se esperan acciones nuestras aparecerá en morado
        table.rows[num_row].cells[5].style = "background-color: #F3BDFF";
        table.rows[num_row].onmouseout = function () {
            this.cells[5].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";
            this.cells[5].style = "background-color: #F3BDFF;";
        };
        table.rows[num_row].onmouseover = function () {
            this.cells[5].children[0].style = "color: #762189; font-weight: bold; background-color: #F3BDFF;";;
            this.cells[5].style = "background-color: #F3BDFF;";
        };
    } else {

        if ((num_row % 2) == 0) {
            table.rows[num_row].cells[5].children[0].style = "color: black; font-weight: normal; background-color: #ffffff";
            table.rows[num_row].cells[5].style = "background-color: #ffffff;";
            table.rows[num_row].onmouseout = function () {
                this.cells[5].children[0].style = "background-color: #ffffff";
                this.cells[5].style = "background-color: #ffffff;";
            };
        } else {
            table.rows[num_row].cells[5].children[0].style = "color: black; font-weight: normal; background-color: #f2f2f2";
            table.rows[num_row].cells[5].style = "background-color: #f2f2f2;";
            table.rows[num_row].onmouseout = function () {
                this.cells[5].children[0].style = "background-color: #f2f2f2";
                this.cells[5].style = "background-color: #f2f2f2;";
            };
        }
        table.rows[num_row].onmouseover = function () {
            this.cells[5].children[0].style = "background-color: #ddd;";
            this.cells[5].style = "background-color: #ddd;";
        };

    }

}

// Función que permite ajustar los textarea al texto
function textAreaAdjust(element) {
    element.style.height = "1px";
    element.style.height = (25 + element.scrollHeight) + "px";
}

// Función que valida las credenciales
function validateUser() {

    var user = document.getElementById('username').value;
    var clavemd5 = md5(document.getElementById('password').value);

    if (user == "acsoper" && clavemd5 == "6fe4be3b91842f4581809d232b219aa5") {
        document.getElementById("select-mode").style.display = 'block';
        document.getElementById("validate-credentials").style.display = 'none';
    } else {
        alert('Incorrect username or password. Try again')
    }
}
