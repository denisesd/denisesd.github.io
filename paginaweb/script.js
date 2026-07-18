let data = {
    sheets: {},
    resultados: []
};

let charts = {};

document.addEventListener("DOMContentLoaded", () => {
    const header = document.querySelector(".header");
    const headerText = document.querySelector(".header-text");
    const introScreen = document.querySelector(".intro-screen");
    const contactForm = document.querySelector(".contact-form");
    const skipIntro = new URLSearchParams(window.location.search).get("skipIntro") === "1";

    if (introScreen) {
        requestAnimationFrame(() => {
            introScreen.classList.add("intro-ready");
        });
    }

    if (header) {
        const showMainScreen = () => {
            header.classList.add("header-logo-visible");
            headerText?.classList.add("header-text-hidden");
            introScreen?.classList.add("intro-hidden");
        };

        if (skipIntro) {
            showMainScreen();
        } else {
            setTimeout(showMainScreen, 4000);
        }
    }

    if (contactForm) {
        contactForm.addEventListener("submit", async (event) => {
            event.preventDefault();

            if (!contactForm.checkValidity()) {
                contactForm.reportValidity();
                return;
            }

            const submitButton = contactForm.querySelector(".form-submit");
            const originalText = submitButton?.textContent;

            if (submitButton) {
                submitButton.disabled = true;
                submitButton.textContent = "ENVIANDO...";
            }

            try {
                const response = await fetch(contactForm.action, {
                    method: "POST",
                    body: new FormData(contactForm),
                    headers: {
                        Accept: "application/json"
                    }
                });

                if (!response.ok) {
                    throw new Error("No se pudo enviar el formulario.");
                }

                alert("Mensaje enviado correctamente. Gracias por contactarnos.");
                contactForm.reset();
            } catch (error) {
                alert("No se pudo enviar el mensaje. Por favor intente nuevamente.");
            } finally {
                if (submitButton) {
                    submitButton.disabled = false;
                    submitButton.textContent = originalText;
                }
            }
        });
    }
});

function leerExcel() {
    const archivo = document.getElementById("excelFile").files[0];

    if (!archivo) {
        alert("Seleccione un archivo Excel.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const dataExcel = new Uint8Array(e.target.result);
        const workbook = XLSX.read(dataExcel, { type: "array" });

        data.sheets = {};

        workbook.SheetNames.forEach(nombreHoja => {
            data.sheets[nombreHoja] = XLSX.utils.sheet_to_json(
                workbook.Sheets[nombreHoja],
                { header: 1, defval: null }
            );
        });

        calcularCostos();
        alert("Modelo ABC importado correctamente.");
    };

    reader.readAsArrayBuffer(archivo);
}

function celda(hoja, fila, columna) {
    return data.sheets[hoja]?.[fila]?.[columna] ?? null;
}

function numero(valor) {
    if (valor === null || valor === undefined || valor === "") return 0;

    if (typeof valor === "string") {
        valor = valor.replace("S/", "").replace(",", "").trim();
    }

    return Number(valor) || 0;
}

function obtenerObjetosCosto() {
    let objetos = [];

    for (let col = 3; col <= 7; col++) {
        let ot = celda("Costo MD", 3, col);
        let producto = celda("Costo MD", 4, col);
        let cantidad = numero(celda("Costo MD", 5, col));

        if (ot || producto) {
            objetos.push({
                ot: ot || `OT-${col - 2}`,
                producto: producto || "Objeto de costo",
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
        alert("No se encontraron objetos de costo en el modelo.");
        return;
    }

    llenarTablaCostos(resultados);
    actualizarTarjetas(resultados);
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
            propuesta: "Revisar los centros de actividad, actividades y drivers que generan mayor consumo de recursos.",
            impacto: "Mejor asignación de costos indirectos y reducción de actividades sin valor agregado."
        });
    }

    let objetoMasCostoso = [...resultados].sort((a, b) => b.costoTotal - a.costoTotal)[0];

    if (objetoMasCostoso) {
        propuestas.push({
            tipo: "oportunidad",
            titulo: "Priorización del objeto de costo más relevante",
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
            propuesta: "Revisar precios, costos estándar, eficiencia productiva y actividades indirectas.",
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

function generarGraficoEstructuraCostos(resultados) {
    let totalMD = resultados.reduce((sum, r) => sum + r.md, 0);
    let totalMOD = resultados.reduce((sum, r) => sum + r.mod, 0);
    let totalCIF = resultados.reduce((sum, r) => sum + r.cif, 0);

    const ctx = document.getElementById("graficoEstructuraCostos");

    charts.estructura = new Chart(ctx, {
        type: "doughnut",
        data: {
            labels: ["Materia Directa", "Mano de Obra Directa", "CIF ABC"],
            datasets: [{
                data: [totalMD, totalMOD, totalCIF]
            }]
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
                data: resultados.map(r => r.costoTotal)
            }]
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
                label: "Rentabilidad %",
                data: resultados.map(r => r.rentabilidad)
            }]
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
                label: "Costo total por objeto de costo",
                data: resultados.map(r => r.costoTotal),
                tension: 0.3
            }]
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
                    data: resultados.map(r => r.ventas)
                },
                {
                    label: "Costos",
                    data: resultados.map(r => r.costoTotal)
                }
            ]
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
                    yAxisID: "y"
                },
                {
                    label: "% acumulado",
                    data: porcentajesAcumulados,
                    type: "line",
                    yAxisID: "y1",
                    tension: 0.3
                }
            ]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true,
                    position: "left"
                },
                y1: {
                    beginAtZero: true,
                    max: 100,
                    position: "right"
                }
            }
        }
    });
}
