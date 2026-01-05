
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Filtro de Ventas PREMIUM - Con Fechas</title>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 30px; background-color: #f4f7f6; color: #333; }
        .container { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); max-width: 1200px; margin: auto; }
        .upload-section { border: 3px dashed #3498db; padding: 20px; text-align: center; border-radius: 10px; cursor: pointer; background: #f0faff; margin-bottom: 20px; }
        
        /* Estilos de los Filtros */
        .filter-bar { background: #ebf2f7; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-group { display: flex; flex-direction: column; gap: 5px; }
        input[type="date"], select { padding: 8px; border-radius: 5px; border: 1px solid #ccc; }
        
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px; border: 1px solid #ddd; text-align: left; font-size: 13px; }
        th { background-color: #2c3e50; color: white; cursor: pointer; }
        th:hover { background-color: #34495e; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        
        .stats { font-weight: bold; color: #1a5276; margin-bottom: 10px; }
        .btn-download { background: #27ae60; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; }
        .btn-filter { background: #3498db; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; align-self: flex-end; }
    </style>
</head>
<body>

<div class="container">
    <h2>Extractor de Ventas PREMIUM</h2>
    
    <div class="upload-section" onclick="document.getElementById('fileInput').click()">
        <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" style="display:none">
        <strong>📁 Paso 1: Sube tu archivo VentasConsignaVSE</strong>
        <p>Haz clic aquí para cargar</p>
    </div>

    <div id="resultArea" style="display:none;">
        <div class="filter-bar">
            <div class="filter-group">
                <label>Desde:</label>
                <input type="date" id="dateStart">
            </div>
            <div class="filter-group">
                <label>Hasta:</label>
                <input type="date" id="dateEnd">
            </div>
            <div class="filter-group">
                <label>Orden de Fecha:</label>
                <select id="sortOrder">
                    <option value="asc">Más antiguas primero</option>
                    <option value="desc">Más recientes primero</option>
                </select>
            </div>
            <button class="btn-filter" onclick="aplicarFiltros()">Aplicar Filtros</button>
            <button id="downloadBtn" class="btn-download" style="margin-left: auto;">Descargar Excel</button>
        </div>

        <div class="stats" id="rowCount"></div>
        <table id="ventasTable">
            <thead>
                <tr>
                    <th>Fecha de Captura ↕</th>
                    <th>Cuenta AVS</th>
                    <th>Orden</th>
                    <th>Nombre del Cliente</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>
</div>

<script>
    let datosOriginales = [];
    let datosFiltrados = [];

    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', raw: false });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            // Procesar y limpiar datos
            datosOriginales = jsonData.filter(row => {
                const tipo = obtenerValor(row, ['Tipo de Venta', 'TIPO']);
                return String(tipo).trim().toUpperCase() === 'ACTIVA - PREMIUM SIN EQUIPO';
            }).map(row => {
                const fechaTexto = obtenerValor(row, ['Fecha de Captura', 'FECHA']);
                return {
                    fechaStr: fechaTexto,
                    fechaObjeto: parsearFecha(fechaTexto),
                    cuenta: obtenerValor(row, ['Cuenta AVS', 'CUENTA']),
                    orden: obtenerValor(row, ['Orden', 'ORDEN']),
                    cliente: obtenerValor(row, ['Cliente', 'NOMBRE'])
                };
            });

            aplicarFiltros();
        };
        reader.readAsArrayBuffer(file);
    });

    function obtenerValor(obj, nombres) {
        for (let n of nombres) {
            for (let clave in obj) {
                if (clave.trim().toUpperCase() === n.toUpperCase()) return obj[clave];
            }
        }
        return "N/A";
    }

    // Convierte texto DD/MM/AAAA a un objeto de fecha real para poder comparar/ordenar
    function parsearFecha(str) {
        if(!str || str === "N/A") return new Date(0);
        const partes = str.split('/');
        if(partes.length === 3) {
            return new Date(partes[2], partes[1] - 1, partes[0]);
        }
        return new Date(str); 
    }

    function aplicarFiltros() {
        const inicio = document.getElementById('dateStart').value;
        const fin = document.getElementById('dateEnd').value;
        const orden = document.getElementById('sortOrder').value;

        datosFiltrados = [...datosOriginales];

        // Filtro de rango
        if (inicio) {
            const dInicio = new Date(inicio);
            datosFiltrados = datosFiltrados.filter(d => d.fechaObjeto >= dInicio);
        }
        if (fin) {
            const dFin = new Date(fin);
            datosFiltrados = datosFiltrados.filter(d => d.fechaObjeto <= dFin);
        }

        // Ordenar
        datosFiltrados.sort((a, b) => {
            return orden === 'asc' 
                ? a.fechaObjeto - b.fechaObjeto 
                : b.fechaObjeto - a.fechaObjeto;
        });

        renderizar();
    }

    function renderizar() {
        const tableBody = document.querySelector('#ventasTable tbody');
        tableBody.innerHTML = '';
        document.getElementById('resultArea').style.display = 'block';
        document.getElementById('rowCount').innerText = `Mostrando ${datosFiltrados.length} registros.`;

        datosFiltrados.forEach(d => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${d.fechaStr}</td>
                <td>${d.cuenta}</td>
                <td>${d.orden}</td>
                <td>${d.cliente}</td>
            `;
            tableBody.appendChild(tr);
        });
    }

    document.getElementById('downloadBtn').addEventListener('click', function() {
        const paraExportar = datosFiltrados.map(d => ({
            "Fecha de Captura": d.fechaStr,
            "Cuenta AVS": d.cuenta,
            "Orden": d.orden,
            "Cliente": d.cliente
        }));
        const ws = XLSX.utils.json_to_sheet(paraExportar);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Filtrado");
        XLSX.writeFile(wb, "Reporte_Premium_Filtrado.xlsx");
    });
</script>

</body>
</html>
