let data = {
    sheets: {},
    resultados: []
};

let charts = {};

let hojasModelo = {
    costoMD: null,
    costoMOD: null,
    proceso1: null,
    proceso2: null,
    proceso3: null
};

function leerExcel() {
    const archivo = document.getElementById("excelFile").files[0];

    if (!archivo) {
        alert("Seleccione un archivo Excel.");
        return;
    }

    const extension = archivo.name.split(".").pop().toLowerCase();

    if (!["xlsx", "xls", "xlsm", "csv"].includes(extension)) {
        alert("Formato no permitido. Suba un archivo Excel válido.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const dataExcel = new Uint8Array(e.target.result);
            const workbook = XLSX.read(dataExcel, { type: "array" });

            data.sheets = {};

            workbook.SheetNames.forEach(nombreHoja => {
                data.sheets[nombreHoja.trim()] = XLSX.utils.sheet_to_json(
                    workbook.Sheets[nombreHoja],
                    {
                        header: 1,
                        defval: "",
                        raw: false
                    }
                );
            });

            detectarHojasModelo();
            calcularCostos();

            alert("Archivo Excel cargado correctamente.");
        } catch (error) {
            console.error("Error al procesar Excel:", error);
            alert("El archivo se cargó, pero hubo un problema al procesarlo.");
        }
    };

    reader.readAsArrayBuffer(archivo);
}

function normalizarTexto(texto) {
    return String(texto || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[_\-\s]+/g, "")
        .trim();
}

function buscarHojaPorNombre(posiblesNombres) {
    const hojas = Object.keys(data.sheets);

    for (let hoja of hojas) {
        let hojaNormalizada = normalizarTexto(hoja);

        for (let nombre of posiblesNombres) {
            if (hojaNormalizada.includes(normalizarTexto(nombre))) {
                return hoja;
            }
        }
    }

    return null;
}

function detectarHojasModelo() {
    hojasModelo.costoMD = buscarHojaPorNombre([
        "Costo MD",
        "MD",
        "Materia Directa",
        "Costo Materia Directa"
    ]);

    hojasModelo.costoMOD = buscarHojaPorNombre([
        "Costo MOD",
        "MOD",
        "Mano de Obra",
        "Costo Mano de Obra"
    ]);

    hojasModelo.proceso1 = buscarHojaPorNombre([
        "Proceso 1",
        "Proceso1"
    ]);

    hojasModelo.proceso2 = buscarHojaPorNombre([
        "Proceso 2",
        "Proceso2"
    ]);

    hojasModelo.proceso3 = buscarHojaPorNombre([
        "Proceso 3",
        "Proceso3"
    ]);

    console.log("Hojas detectadas:", hojasModelo);
}

function celda(hojaNombre, fila, columna) {
    const mapa = {
        "Costo MD": hojasModelo.costoMD,
        "Costo MOD": hojasModelo.costoMOD,
        "Proceso 1": hojasModelo.proceso1,
        "Proceso 2": hojasModelo.proceso2,
        "Proceso 3": hojasModelo.proceso3
    };

    const hojaReal = mapa[hojaNombre] || hojaNombre;

    return data.sheets[hojaReal]?.[fila]?.[columna] ?? "";
}

function numero(valor) {
    if (valor === null || valor === undefined || valor === "") return 0;

    if (typeof valor === "string") {
        valor = valor
            .replace("S/", "")
            .replace("s/", "")
            .replace(/,/g, "")
            .replace("%", "")
            .trim();
    }

    return Number(valor) || 0;
}

function obtenerObjetosCosto() {
    let objetos = [];

    for (let col = 3; col <= 7; col++) {
        let ot = celda("Costo MD", 3, col);
        let producto = celda("Costo MD", 4, col);
        let cantidad = numero(celda("Costo MD", 5, col));

        if (ot || producto || cantidad > 0) {
            objetos.push({
                ot: ot || `OT-${col - 2}`,
                producto: producto || `Objeto de costo ${col - 2}`,
                cliente: "No definido",
                periodo: "Modelo ABC",
                cantidad: cantidad || 1,
                colMD: col,
                index: col - 3,
                md: 0,
                mod: 0,
                cif: 0,
                costoTotal: 0,
                costoUnitario: 0,
                ventas: 0,
                utilidad: 0,
                rentabilidad: 0,
                actividades: {}
            });
        }
    }

    return objetos;
}

function calcularCMD(objetos) {
    objetos.forEach(obj => {
        obj.md = numero(celda("Costo MD", 25, obj.colMD));
    });
}

function calcularCMOD(objetos) {
    objetos.forEach(obj => {
        obj.mod = numero(celda("Costo MOD", 19, obj.colMD));
    });
}

function calcularCostoElementoACentro() {
    let costoCentroActividad = {};

    for (let col = 5; col <= 12; col++) {
        let centro = celda("Proceso 1", 37, col);

        if (centro) {
            costoCentroActividad[centro] = 0;
        }
    }

    for (let fila = 52; fila <= 61; fila++) {
        let costoElemento = numero(celda("Proceso 1", fila, 2));
        let filaDriver = fila - 14;
        let totalDriver = numero(celda("Proceso 1", filaDriver, 13));

        for (let col = 5; col <= 12; col++) {
            let centro = celda("Proceso 1", 37, col);
            let driver = numero(celda("Proceso 1", filaDriver, col));

            if (centro && totalDriver > 0) {
                costoCentroActividad[centro] += costoElemento * (driver / totalDriver);
            }
        }
    }

    return costoCentroActividad;
}

function calcularCostoCentroAActividad(costoCentroActividad) {
    let costoActividad = {};

    for (let col = 5; col <= 16; col++) {
        let actividad = celda("Proceso 2", 33, col);

        if (actividad) {
            costoActividad[actividad] = 0;
        }
    }

    for (let fila = 34; fila <= 41; fila++) {
        let centro = celda("Proceso 2", fila, 1);
        let costoCentro = costoCentroActividad[centro] || 0;
        let totalDriver = numero(celda("Proceso 2", fila, 17));

        for (let col = 5; col <= 16; col++) {
            let actividad = celda("Proceso 2", 33, col);
            let driver = numero(celda("Proceso 2", fila, col));

            if (actividad && totalDriver > 0) {
                costoActividad[actividad] += costoCentro * (driver / totalDriver);
            }
        }
    }

    return costoActividad;
}

function calcularCostoActividadAObjeto(costoActividad, objetos) {
    objetos.forEach(obj => {
        obj.cif = 0;
        obj.actividades = {};
    });

    for (let fila = 44; fila <= 55; fila++) {
        let actividad = celda("Proceso 3", fila, 1);
        let costoAct = costoActividad[actividad] || 0;
        let totalDriver = numero(celda("Proceso 3", fila, 10));

        objetos.forEach((obj, index) => {
            let col = 5 + index;
            let driver = numero(celda("Proceso 3", fila, col));
            let costoAsignado = 0;

            if (actividad && totalDriver > 0) {
                costoAsignado = costoAct * (driver / totalDriver);
            }

            obj.actividades[actividad] = costoAsignado;
            obj.cif += costoAsignado;
        });
    }
}

function obtenerResultados() {
    let objetos = obtenerObjetosCosto();

    if (objetos.length === 0) {
        return [];
    }

    calcularCMD(objetos);
    calcularCMOD(objetos);

    let costoCentroActividad = calcularCostoElementoACentro();
    let costoActividad = calcularCostoCentroAActividad(costoCentroActividad);

    calcularCostoActividadAObjeto(costoActividad, objetos);

    objetos.forEach(obj => {
        obj.costoTotal = obj.md + obj.mod + obj.cif;
        obj.costoUnitario = obj.cantidad > 0 ? obj.costoTotal / obj.cantidad : 0;

        obj.ventas = obj.costoTotal * 1.3;
        obj.utilidad = obj.ventas - obj.costoTotal;
        obj.rentabilidad = obj.ventas > 0 ? (obj.utilidad / obj.ventas) * 100 : 0;
    });

    data.resultados = objetos;
    return objetos;
}

function calcularCostos() {
    const resultados = obtenerResultados();

    if (resultados.length === 0) {
        alert("El Excel fue cargado, pero no se encontraron objetos de costo.");
        return;
    }

    llenarTablaCostos(resultados);
    actualizarTarjetas(resultados);

    llenarTablaInventario(resultados);
    llenarTablaABC();

    generarDashboards(resultados);
    llenarTablaObjetosCosto(resultados);
    llenarTablaRelevancia(resultados);
    generarPropuestasMejora(resultados);
}

function llenarTablaCostos(resultados) {
    const tbody = document.querySelector("#tablaCostos tbody");
    tbody.innerHTML = "";

    resultados.forEach(r => {
        const claseRentabilidad = r.rentabilidad >= 0 ? "positivo" : "negativo";

        tbody.innerHTML += `
            <tr>
                <td>${r.ot}</td>
                <td>${r.producto}</td>
                <td>${r.cliente}</td>
                <td>${r.periodo}</td>
                <td>${r.cantidad.toLocaleString("es-PE")}</td>
                <td>S/ ${r.md.toFixed(2)}</td>
                <td>S/ ${r.mod.toFixed(2)}</td>
                <td>S/ ${r.cif.toFixed(2)}</td>
                <td>S/ ${r.costoTotal.toFixed(2)}</td>
                <td>S/ ${r.costoUnitario.toFixed(2)}</td>
                <td>S/ ${r.ventas.toFixed(2)}</td>
                <td>S/ ${r.utilidad.toFixed(2)}</td>
                <td class="${claseRentabilidad}">${r.rentabilidad.toFixed(2)}%</td>
            </tr>
        `;
    });
}

function llenarTablaInventario(resultados) {
    const tbody = document.querySelector("#tablaInventario tbody");

    if (!tbody) return;

    tbody.innerHTML = "";

    resultados.forEach(obj => {
        let stockInicial = obj.cantidad;
        let costoUnitario = obj.cantidad > 0 ? obj.md / obj.cantidad : 0;
        let entradas = obj.cantidad * 0.20;
        let salidas = obj.cantidad * 0.75;

        let valorInicial = stockInicial * costoUnitario;
        let valorEntradas = entradas * costoUnitario;

        let costoPromedio = (stockInicial + entradas) > 0
            ? (valorInicial + valorEntradas) / (stockInicial + entradas)
            : 0;

        let stockFinal = stockInicial + entradas - salidas;
        let valorFinal = stockFinal * costoPromedio;

        tbody.innerHTML += `
            <tr>
                <td>${obj.ot}</td>
                <td>${obj.producto}</td>
                <td>${stockInicial.toLocaleString("es-PE")}</td>
                <td>S/ ${costoUnitario.toFixed(2)}</td>
                <td>${entradas.toFixed(0)}</td>
                <td>${salidas.toFixed(0)}</td>
                <td>S/ ${costoPromedio.toFixed(2)}</td>
                <td>${stockFinal.toFixed(0)}</td>
                <td>S/ ${valorFinal.toFixed(2)}</td>
            </tr>
        `;
    });
}

function llenarTablaABC() {
    const tbody = document.querySelector("#tablaABC tbody");

    if (!tbody) return;

    tbody.innerHTML = "";

    let costoCentroActividad = calcularCostoElementoACentro();
    let costoActividad = calcularCostoCentroAActividad(costoCentroActividad);

    Object.entries(costoActividad).forEach(([actividad, costoTotal]) => {
        let volumenDriver = buscarVolumenDriverActividad(actividad);
        let tasaABC = volumenDriver > 0 ? costoTotal / volumenDriver : 0;

        tbody.innerHTML += `
            <tr>
                <td>${actividad}</td>
                <td>Driver ABC</td>
                <td>S/ ${costoTotal.toFixed(2)}</td>
                <td>${volumenDriver.toLocaleString("es-PE")}</td>
                <td>S/ ${tasaABC.toFixed(2)}</td>
            </tr>
        `;
    });
}

function buscarVolumenDriverActividad(actividadBuscada) {
    for (let fila = 44; fila <= 55; fila++) {
        let actividad = celda("Proceso 3", fila, 1);

        if (String(actividad).trim() === String(actividadBuscada).trim()) {
            return numero(celda("Proceso 3", fila, 10));
        }
    }

    return 0;
}

function actualizarTarjetas(resultados) {
    let totalProduccion = resultados.reduce((sum, r) => sum + r.costoTotal, 0);
    let totalVentas = resultados.reduce((sum, r) => sum + r.ventas, 0);
    let totalUtilidad = resultados.reduce((sum, r) => sum + r.utilidad, 0);

    let costoPromedio = resultados.length > 0 ? totalProduccion / resultados.length : 0;
    let rentabilidadGeneral = totalVentas > 0 ? (totalUtilidad / totalVentas) * 100 : 0;

    document.getElementById("totalProduccion").textContent = `S/ ${totalProduccion.toFixed(2)}`;
    document.getElementById("costoPromedio").textContent = `S/ ${costoPromedio.toFixed(2)}`;
    document.getElementById("ventasTotales").textContent = `S/ ${totalVentas.toFixed(2)}`;
    document.getElementById("rentabilidad").textContent = `${rentabilidadGeneral.toFixed(2)}%`;
}

function llenarTablaObjetosCosto(resultados) {
    const tbody = document.querySelector("#tablaObjetosCosto tbody");
    tbody.innerHTML = "";

    resultados.forEach(obj => {
        let clase = obj.rentabilidad >= 0 ? "positivo" : "negativo";

        tbody.innerHTML += `
            <tr>
                <td>${obj.producto}</td>
                <td>${obj.ot}</td>
                <td>${obj.cantidad.toLocaleString("es-PE")}</td>
                <td>S/ ${obj.costoTotal.toFixed(2)}</td>
                <td>S/ ${obj.ventas.toFixed(2)}</td>
                <td>S/ ${obj.utilidad.toFixed(2)}</td>
                <td class="${clase}">${obj.rentabilidad.toFixed(2)}%</td>
            </tr>
        `;
    });
}

function llenarTablaRelevancia(resultados) {
    const tbody = document.querySelector("#tablaRelevancia tbody");
    tbody.innerHTML = "";

    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);
    let totalCostos = totalMD + totalMOD + totalCIF;

    let estructura = [
        {
            componente: "Materia Directa",
            monto: totalMD,
            participacion: totalCostos > 0 ? (totalMD / totalCostos) * 100 : 0
        },
        {
            componente: "Mano de Obra Directa",
            monto: totalMOD,
            participacion: totalCostos > 0 ? (totalMOD / totalCostos) * 100 : 0
        },
        {
            componente: "CIF ABC",
            monto: totalCIF,
            participacion: totalCostos > 0 ? (totalCIF / totalCostos) * 100 : 0
        }
    ];

    estructura.sort((a, b) => b.participacion - a.participacion);

    estructura.forEach(item => {
        let interpretacion = "";

        if (item.participacion >= 50) {
            interpretacion = "Componente crítico. Requiere control prioritario.";
        } else if (item.participacion >= 25) {
            interpretacion = "Componente relevante. Debe monitorearse.";
        } else {
            interpretacion = "Componente secundario. Puede optimizarse.";
        }

        tbody.innerHTML += `
            <tr>
                <td>${item.componente}</td>
                <td>S/ ${item.monto.toFixed(2)}</td>
                <td>${item.participacion.toFixed(2)}%</td>
                <td>${interpretacion}</td>
            </tr>
        `;
    });
}

function generarPropuestasMejora(resultados) {
    const contenedor = document.getElementById("propuestasMejora");
    contenedor.innerHTML = "";

    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);
    let totalCostos = resultados.reduce((sum, r) => sum + r.costoTotal, 0);
    let totalVentas = resultados.reduce((sum, r) => sum + r.ventas, 0);
    let totalUtilidad = resultados.reduce((sum, r) => sum + r.utilidad, 0);

    let rentabilidadGeneral = totalVentas > 0 ? (totalUtilidad / totalVentas) * 100 : 0;
    let porcentajeMD = totalCostos > 0 ? (totalMD / totalCostos) * 100 : 0;
    let porcentajeMOD = totalCostos > 0 ? (totalMOD / totalCostos) * 100 : 0;
    let porcentajeCIF = totalCostos > 0 ? (totalCIF / totalCostos) * 100 : 0;

    let propuestas = [];

    if (porcentajeMD >= 50) {
        propuestas.push({
            tipo: "alerta",
            titulo: "Optimización de materia prima",
            problema: `La materia directa representa el ${porcentajeMD.toFixed(2)}% del costo total.`,
            propuesta: "Renegociar precios con proveedores, reducir mermas y mejorar el rendimiento de materiales.",
            impacto: "Reducción del costo unitario y mejora del margen bruto."
        });
    }

    if (porcentajeMOD >= 25) {
        propuestas.push({
            tipo: "alerta",
            titulo: "Mejora de productividad laboral",
            problema: `La mano de obra directa representa el ${porcentajeMOD.toFixed(2)}% del costo total.`,
            propuesta: "Estandarizar tiempos, balancear cargas de trabajo y reducir tiempos muertos.",
            impacto: "Menor costo laboral por unidad producida."
        });
    }

    if (porcentajeCIF >= 20) {
        propuestas.push({
            tipo: "alerta",
            titulo: "Control de CIF mediante ABC",
            problema: `Los CIF ABC representan el ${porcentajeCIF.toFixed(2)}% del costo total.`,
            propuesta: "Revisar centros de actividad, actividades y drivers que generan mayor consumo de recursos.",
            impacto: "Mejor asignación de costos indirectos y reducción de actividades sin valor agregado."
        });
    }

    let objetoMasCostoso = [...resultados].sort((a, b) => b.costoTotal - a.costoTotal)[0];

    if (objetoMasCostoso) {
        propuestas.push({
            tipo: "oportunidad",
            titulo: "Objeto de costo más relevante",
            problema: `El objeto de costo más representativo es ${objetoMasCostoso.producto}.`,
            propuesta: "Analizar su consumo de materiales, mano de obra y actividades ABC.",
            impacto: "Mayor control sobre el producto con mayor peso económico."
        });
    }

    if (rentabilidadGeneral < 20) {
        propuestas.push({
            tipo: "alerta",
            titulo: "Rentabilidad general baja",
            problema: `La rentabilidad global es de ${rentabilidadGeneral.toFixed(2)}%.`,
            propuesta: "Revisar precios, eficiencia productiva y actividades indirectas.",
            impacto: "Mejora del margen operativo."
        });
    } else {
        propuestas.push({
            tipo: "oportunidad",
            titulo: "Rentabilidad global favorable",
            problema: `La rentabilidad general es de ${rentabilidadGeneral.toFixed(2)}%.`,
            propuesta: "Mantener control sobre los productos rentables y replicar sus condiciones productivas.",
            impacto: "Consolidación de márgenes."
        });
    }

    propuestas.forEach(p => {
        contenedor.innerHTML += `
            <div class="propuesta-card ${p.tipo}">
                <h3>${p.titulo}</h3>
                <p><strong>Diagnóstico:</strong> ${p.problema}</p>
                <p><strong>Propuesta:</strong> ${p.propuesta}</p>
                <p><strong>Impacto esperado:</strong> ${p.impacto}</p>
            </div>
        `;
    });
}

function filtrarPorPeriodo() {
    calcularCostos();
}

function generarDashboards(resultados) {
    destruirGraficos();

    generarGraficoEstructuraCostos(resultados);
    generarGraficoCostoOT(resultados);
    generarGraficoRentabilidad(resultados);
    generarGraficoPeriodo(resultados);
    generarGraficoVentasCostos(resultados);
    generarGraficoPareto(resultados);
}

function destruirGraficos() {
    Object.values(charts).forEach(chart => {
        if (chart) chart.destroy();
    });

    charts = {};
}

function opcionesGrafico() {
    return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
            mode: "index",
            intersect: false
        },
        plugins: {
            legend: {
                position: "bottom",
                labels: {
                    boxWidth: 12,
                    font: {
                        size: 11
                    }
                }
            },
            tooltip: {
                enabled: true,
                backgroundColor: "#0f172a",
                titleFont: {
                    size: 13
                },
                bodyFont: {
                    size: 12
                },
                callbacks: {
                    label: function(context) {
                        let label = context.dataset.label || context.label || "";
                        let value = context.raw || 0;

                        if (label.includes("%") || label.includes("Rentabilidad")) {
                            return `${label}: ${Number(value).toFixed(2)}%`;
                        }

                        return `${label}: S/ ${Number(value).toLocaleString("es-PE", {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        })}`;
                    }
                }
            }
        }
    };
}

function generarGraficoEstructuraCostos(resultados) {
    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);

    const ctx = document.getElementById("graficoEstructuraCostos");

    charts.estructura = new Chart(ctx, {
        type: "doughnut",
        data: {
            labels: ["MD", "MOD", "CIF ABC"],
            datasets: [{
                data: [totalMD, totalMOD, totalCIF],
                hoverOffset: 12
            }]
        },
        options: {
            ...opcionesGrafico(),
            cutout: "58%"
        }
    });
}

function generarGraficoCostoOT(resultados) {
    const ctx = document.getElementById("graficoCostoOT");

    charts.costoOT = new Chart(ctx, {
        type: "bar",
        data: {
            labels: resultados.map(r => r.ot),
            datasets: [{
                label: "Costo total",
                data: resultados.map(r => r.costoTotal),
                borderWidth: 1
            }]
        },
        options: {
            ...opcionesGrafico(),
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => "S/ " + Number(value).toLocaleString("es-PE")
                    }
                },
                x: {
                    ticks: {
                        font: {
                            size: 10
                        }
                    }
                }
            }
        }
    });
}

function generarGraficoRentabilidad(resultados) {
    const ctx = document.getElementById("graficoRentabilidad");

    charts.rentabilidad = new Chart(ctx, {
        type: "bar",
        data: {
            labels: resultados.map(r => r.ot),
            datasets: [{
                label: "Rentabilidad",
                data: resultados.map(r => r.rentabilidad),
                borderWidth: 1
            }]
        },
        options: {
            ...opcionesGrafico(),
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => value + "%"
                    }
                }
            }
        }
    });
}

function generarGraficoPeriodo(resultados) {
    const ctx = document.getElementById("graficoPeriodo");

    charts.periodo = new Chart(ctx, {
        type: "line",
        data: {
            labels: resultados.map(r => r.ot),
            datasets: [{
                label: "Costo total",
                data: resultados.map(r => r.costoTotal),
                tension: 0.35,
                pointRadius: 4,
                pointHoverRadius: 7
            }]
        },
        options: {
            ...opcionesGrafico(),
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => "S/ " + Number(value).toLocaleString("es-PE")
                    }
                }
            }
        }
    });
}

function generarGraficoVentasCostos(resultados) {
    const ctx = document.getElementById("graficoVentasCostos");

    charts.ventasCostos = new Chart(ctx, {
        type: "bar",
        data: {
            labels: resultados.map(r => r.ot),
            datasets: [
                {
                    label: "Ventas estimadas",
                    data: resultados.map(r => r.ventas),
                    borderWidth: 1
                },
                {
                    label: "Costos",
                    data: resultados.map(r => r.costoTotal),
                    borderWidth: 1
                }
            ]
        },
        options: {
            ...opcionesGrafico(),
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => "S/ " + Number(value).toLocaleString("es-PE")
                    }
                }
            }
        }
    });
}

function generarGraficoPareto(resultados) {
    const ordenados = [...resultados].sort((a, b) => b.costoTotal - a.costoTotal);

    let totalCostos = ordenados.reduce((sum, r) => sum + r.costoTotal, 0);
    let acumulado = 0;

    let porcentajesAcumulados = ordenados.map(r => {
        acumulado += r.costoTotal;
        return totalCostos > 0 ? (acumulado / totalCostos) * 100 : 0;
    });

    const ctx = document.getElementById("graficoPareto");

    charts.pareto = new Chart(ctx, {
        type: "bar",
        data: {
            labels: ordenados.map(r => r.ot),
            datasets: [
                {
                    label: "Costo total",
                    data: ordenados.map(r => r.costoTotal),
                    yAxisID: "y",
                    borderWidth: 1
                },
                {
                    label: "% acumulado",
                    data: porcentajesAcumulados,
                    type: "line",
                    yAxisID: "y1",
                    tension: 0.35,
                    pointRadius: 4,
                    pointHoverRadius: 7
                }
            ]
        },
        options: {
            ...opcionesGrafico(),
            scales: {
                y: {
                    beginAtZero: true,
                    position: "left",
                    ticks: {
                        callback: value => "S/ " + Number(value).toLocaleString("es-PE")
                    }
                },
                y1: {
                    beginAtZero: true,
                    max: 100,
                    position: "right",
                    grid: {
                        drawOnChartArea: false
                    },
                    ticks: {
                        callback: value => value + "%"
                    }
                }
            }
        }
    });
}
function descargarInformeWord() {
    const resultados = data.resultados;

    if (!resultados || resultados.length === 0) {
        alert("Primero debe importar el Excel y calcular los costos.");
        return;
    }

    let totalProduccion = resultados.reduce((sum, r) => sum + r.costoTotal, 0);
    let totalVentas = resultados.reduce((sum, r) => sum + r.ventas, 0);
    let totalUtilidad = resultados.reduce((sum, r) => sum + r.utilidad, 0);
    let rentabilidadGeneral = totalVentas > 0 ? (totalUtilidad / totalVentas) * 100 : 0;

    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);
    let totalCostos = totalMD + totalMOD + totalCIF;

    let fecha = new Date().toLocaleDateString("es-PE");

    let contenido = `
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                    color: #1f2937;
                    line-height: 1.5;
                }

                h1 {
                    color: #1e3a8a;
                    text-align: center;
                }

                h2 {
                    color: #1e3a8a;
                    border-bottom: 2px solid #1e3a8a;
                    padding-bottom: 5px;
                }

                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                    font-size: 12px;
                }

                th {
                    background: #1e3a8a;
                    color: white;
                    padding: 8px;
                    border: 1px solid #ddd;
                }

                td {
                    padding: 8px;
                    border: 1px solid #ddd;
                    text-align: center;
                }

                .resumen {
                    background: #f3f6fb;
                    padding: 12px;
                    border-left: 5px solid #2563eb;
                    margin-bottom: 20px;
                }

                .positivo {
                    color: green;
                    font-weight: bold;
                }

                .negativo {
                    color: red;
                    font-weight: bold;
                }
            </style>
        </head>

        <body>
            <h1>Informe de Resultados del Sistema de Costeo ABC</h1>

            <p><strong>Fecha de emisión:</strong> ${fecha}</p>
            <p><strong>Modelo aplicado:</strong> Costeo ABC Basado en Actividades </p>

            <h2>1. Resumen Ejecutivo</h2>

            <div class="resumen">
                <p><strong>Costo total de producción:</strong> S/ ${totalProduccion.toFixed(2)}</p>
                <p><strong>Ventas estimadas:</strong> S/ ${totalVentas.toFixed(2)}</p>
                <p><strong>Utilidad total:</strong> S/ ${totalUtilidad.toFixed(2)}</p>
                <p><strong>Rentabilidad general:</strong> ${rentabilidadGeneral.toFixed(2)}%</p>
            </div>

            <h2>2. Hoja de Costos por Objeto de Costo</h2>

            <table>
                <thead>
                    <tr>
                        <th>OT</th>
                        <th>Objeto de Costo</th>
                        <th>Cantidad</th>
                        <th>MD</th>
                        <th>MOD</th>
                        <th>CIF ABC</th>
                        <th>Costo Total</th>
                        <th>Costo Unitario</th>
                        <th>Ventas</th>
                        <th>Utilidad</th>
                        <th>Rentabilidad</th>
                    </tr>
                </thead>
                <tbody>
                    ${resultados.map(r => `
                        <tr>
                            <td>${r.ot}</td>
                            <td>${r.producto}</td>
                            <td>${r.cantidad.toLocaleString("es-PE")}</td>
                            <td>S/ ${r.md.toFixed(2)}</td>
                            <td>S/ ${r.mod.toFixed(2)}</td>
                            <td>S/ ${r.cif.toFixed(2)}</td>
                            <td>S/ ${r.costoTotal.toFixed(2)}</td>
                            <td>S/ ${r.costoUnitario.toFixed(2)}</td>
                            <td>S/ ${r.ventas.toFixed(2)}</td>
                            <td>S/ ${r.utilidad.toFixed(2)}</td>
                            <td>${r.rentabilidad.toFixed(2)}%</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>

            <h2>3. Análisis de Relevancia de la Estructura de Costos</h2>

            <table>
                <thead>
                    <tr>
                        <th>Componente</th>
                        <th>Monto</th>
                        <th>Participación</th>
                        <th>Interpretación</th>
                    </tr>
                </thead>
                <tbody>
                    ${generarFilasRelevanciaWord(totalMD, totalMOD, totalCIF, totalCostos)}
                </tbody>
            </table>

            <h2>4. Análisis de Objetos de Costo</h2>

            <p>
                El sistema identifica como objetos de costo principales las órdenes de trabajo y los productos asociados.
                Cada objeto de costo acumula materia directa, mano de obra directa y costos indirectos asignados mediante el modelo ABC.
            </p>

            <table>
                <thead>
                    <tr>
                        <th>Objeto de Costo</th>
                        <th>OT</th>
                        <th>Costo Total</th>
                        <th>Ventas</th>
                        <th>Utilidad</th>
                        <th>Rentabilidad</th>
                    </tr>
                </thead>
                <tbody>
                    ${resultados.map(r => `
                        <tr>
                            <td>${r.producto}</td>
                            <td>${r.ot}</td>
                            <td>S/ ${r.costoTotal.toFixed(2)}</td>
                            <td>S/ ${r.ventas.toFixed(2)}</td>
                            <td>S/ ${r.utilidad.toFixed(2)}</td>
                            <td>${r.rentabilidad.toFixed(2)}%</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>

            <h2>5. Propuestas de Mejora</h2>

            ${generarPropuestasWord(resultados)}

            <h2>6. Conclusión</h2>

            <p>
                El sistema de costeo ABC permite identificar con mayor precisión la estructura de costos de cada objeto de costo,
                asignando los costos indirectos en función del consumo real de actividades. Esto permite tomar mejores decisiones
                sobre precios, eficiencia operativa, reducción de costos y mejora de la rentabilidad empresarial.
            </p>
        </body>
        </html>
    `;

    let blob = new Blob(["\ufeff", contenido], {
        type: "application/msword"
    });

    let enlace = document.createElement("a");
    enlace.href = URL.createObjectURL(blob);
    enlace.download = "Informe_Costeo_ABC.doc";
    document.body.appendChild(enlace);
    enlace.click();
    document.body.removeChild(enlace);
}
function generarFilasRelevanciaWord(totalMD, totalMOD, totalCIF, totalCostos) {
    let estructura = [
        {
            componente: "Materia Directa",
            monto: totalMD,
            participacion: totalCostos > 0 ? (totalMD / totalCostos) * 100 : 0
        },
        {
            componente: "Mano de Obra Directa",
            monto: totalMOD,
            participacion: totalCostos > 0 ? (totalMOD / totalCostos) * 100 : 0
        },
        {
            componente: "CIF ABC",
            monto: totalCIF,
            participacion: totalCostos > 0 ? (totalCIF / totalCostos) * 100 : 0
        }
    ];

    estructura.sort((a, b) => b.participacion - a.participacion);

    return estructura.map(item => {
        let interpretacion = "";

        if (item.participacion >= 50) {
            interpretacion = "Componente crítico. Requiere control prioritario.";
        } else if (item.participacion >= 25) {
            interpretacion = "Componente relevante. Debe monitorearse.";
        } else {
            interpretacion = "Componente secundario. Puede optimizarse.";
        }

        return `
            <tr>
                <td>${item.componente}</td>
                <td>S/ ${item.monto.toFixed(2)}</td>
                <td>${item.participacion.toFixed(2)}%</td>
                <td>${interpretacion}</td>
            </tr>
        `;
    }).join("");
}
function generarPropuestasWord(resultados) {
    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);
    let totalCostos = resultados.reduce((sum, r) => sum + r.costoTotal, 0);
    let totalVentas = resultados.reduce((sum, r) => sum + r.ventas, 0);
    let totalUtilidad = resultados.reduce((sum, r) => sum + r.utilidad, 0);

    let rentabilidadGeneral = totalVentas > 0 ? (totalUtilidad / totalVentas) * 100 : 0;
    let porcentajeMD = totalCostos > 0 ? (totalMD / totalCostos) * 100 : 0;
    let porcentajeMOD = totalCostos > 0 ? (totalMOD / totalCostos) * 100 : 0;
    let porcentajeCIF = totalCostos > 0 ? (totalCIF / totalCostos) * 100 : 0;

    let propuestas = "";

    if (porcentajeMD >= 50) {
        propuestas += `
            <p><strong>Optimización de materia prima:</strong>
            La materia directa representa el ${porcentajeMD.toFixed(2)}% del costo total.
            Se recomienda renegociar precios con proveedores, reducir mermas y mejorar el rendimiento de materiales.
            El impacto esperado es la reducción del costo unitario y mejora del margen bruto.</p>
        `;
    }

    if (porcentajeMOD >= 25) {
        propuestas += `
            <p><strong>Mejora de productividad laboral:</strong>
            La mano de obra directa representa el ${porcentajeMOD.toFixed(2)}% del costo total.
            Se recomienda estandarizar tiempos, balancear cargas de trabajo y reducir tiempos muertos.
            El impacto esperado es una reducción del costo laboral por unidad producida.</p>
        `;
    }

    if (porcentajeCIF >= 20) {
        propuestas += `
            <p><strong>Control de CIF mediante ABC:</strong>
            Los CIF ABC representan el ${porcentajeCIF.toFixed(2)}% del costo total.
            Se recomienda revisar centros de actividad, actividades y drivers que generan mayor consumo de recursos.
            El impacto esperado es una mejor asignación de costos indirectos y reducción de actividades sin valor agregado.</p>
        `;
    }

    let objetoMasCostoso = [...resultados].sort((a, b) => b.costoTotal - a.costoTotal)[0];

    if (objetoMasCostoso) {
        propuestas += `
            <p><strong>Objeto de costo más relevante:</strong>
            El objeto de costo con mayor peso económico es ${objetoMasCostoso.producto}, asociado a la OT ${objetoMasCostoso.ot}.
            Se recomienda analizar su consumo de materiales, mano de obra y actividades ABC.
            El impacto esperado es un mayor control sobre el producto que más influye en el costo total.</p>
        `;
    }

    if (rentabilidadGeneral < 20) {
        propuestas += `
            <p><strong>Rentabilidad general baja:</strong>
            La rentabilidad global es de ${rentabilidadGeneral.toFixed(2)}%.
            Se recomienda revisar precios, eficiencia productiva y actividades indirectas.
            El impacto esperado es la mejora del margen operativo.</p>
        `;
    } else {
        propuestas += `
            <p><strong>Rentabilidad global favorable:</strong>
            La rentabilidad general es de ${rentabilidadGeneral.toFixed(2)}%.
            Se recomienda mantener control sobre los productos rentables y replicar sus condiciones productivas.
            El impacto esperado es la consolidación de márgenes.</p>
        `;
    }

    return propuestas;
}