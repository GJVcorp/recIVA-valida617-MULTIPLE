pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";

(function () {
    const MAX_FILES = 12;
    const WEBAPP_URL = "https://script.google.com/macros/s/AKfycby8x7DvA1V1Ncf0LTkYFseiaVxlqy5UPgOXU81ljIZRlhfk_b_Lk6j9VKPZRNwiJ6eQ/exec";
    const VIDEO_TUTORIAL_URL = "https://www.youtube.com/embed/jcpT1lllJgw?autoplay=1&loop=1&playlist=jcpT1lllJgw";
    const VIDEO_WINDOW_FEATURES = "width=800,height=600,toolbar=0,location=0,menubar=0,scrollbars=0,resizable=0";

    const CREDITO_FIELDS = [
        ["419", "TOTAL VENTAS Y OTRAS OPERACIONES"],
        ["429", "IVA EN VENTAS Y OTRAS OPERACIONES"],
        ["485", "IMPUESTO A LIQUIDAR EN EL PROXIMO MES"],
        ["499", "TOTAL IMPUESTO A LIQUIDAR EN ESTE MES"],
        ["519", "TOTAL ADQUISICIONES Y PAGOS"],
        ["529", "IVA EN ADQUISICIONES Y PAGOS"],
        ["601", "Impuesto causado"],
        ["602", "Credito tributario aplicable en este periodo"],
        ["603", "Compensacion de IVA por ventas con medio electronico"],
        ["604", "Compensacion de IVA por ventas en zonas afectadas"],
        ["605", "Saldo credito tributario del mes anterior"],
        ["606", "Por adquisiciones e importaciones"],
        ["607", "Por retenciones en la fuente de IVA"],
        ["608", "Por compensacion de IVA por ventas con medio electronico"],
        ["609", "Por compensacion de IVA por ventas en zonas afectadas"],
        ["622", "IVA devuelto o descontado por transacciones con adultos mayores o discapacidad"],
        ["610", "Ajuste por IVA devuelto por adquisiciones con medio electronico"],
        ["611", "Ajuste por IVA devuelto en adquisiciones en zonas afectadas"],
        ["612", "Ajuste por IVA devuelto o rechazado (adquisiciones en importaciones)"],
        ["613", "Ajuste por IVA devuelto o rechazado (retenciones en la fuente de IVA)"],
        ["614", "Ajuste por IVA devuelto por instituciones publicas"],
        ["615", "Saldo credito tributario para el proximo mes"],
        ["617", "Por retenciones en la fuente de IVA que le han sido efectuadas"],
        ["618", "Por compensacion de IVA por ventas efectuadas con medio electronico"],
        ["619", "Por compensacion de IVA por ventas efectuadas en zonas afectadas"],
        ["620", "SUBTOTAL A PAGAR"],
        ["621", "IVA PRESUNTIVO DE SALAS DE JUEGO Y OTROS JUEGOS DE AZAR"],
        ["699", "TOTAL IMPUESTO A PAGAR POR PERCEPCION Y RETENCIONES EFECTUADAS EN VENTAS"],
        ["721", "RETENCION DE IVA DEL 10%"],
        ["723", "RETENCION DE IVA DEL 20%"],
        ["725", "RETENCION DE IVA DEL 30%"],
        ["727", "RETENCION DE IVA DEL 50%"],
        ["729", "RETENCION DE IVA DEL 70%"],
        ["731", "RETENCION DE IVA DEL 100%"],
        ["799", "TOTAL IMPUESTO RETENIDO"]
    ];

    const RESUMEN_ROWS = [
        ["Archivo", (item) => item.fileName],
        ["Periodo Fiscal", (item) => item.contribuyente.periodoFiscal || "-"],
        ["RUC", (item) => item.contribuyente.ruc || "-"],
        ["Razon Social", (item) => item.contribuyente.razonSocial || "-"],
        ["Fecha Recaudacion", (item) => item.declaracion.fechaRecaudacion || "-"],
        ["Casillero 601", (item) => getMonto(item, "601")],
        ["Casillero 602", (item) => getMonto(item, "602")],
        ["Casillero 605", (item) => getMonto(item, "605")],
        ["Casillero 606", (item) => getMonto(item, "606")],
        ["Casillero 607", (item) => getMonto(item, "607")],
        ["Casillero 609", (item) => getMonto(item, "609")],
        ["Casillero 615", (item) => getMonto(item, "615")],
        ["Casillero 617", (item) => getMonto(item, "617")],
        ["Casillero 618", (item) => getMonto(item, "618")],
        ["Casillero 619", (item) => getMonto(item, "619")],
        ["Casillero 620", (item) => getMonto(item, "620")],
        ["Casillero 699", (item) => getMonto(item, "699")]
    ];

    const RESUMEN_ROWS_EXTRA = [
        ["Casillero 419", (item) => getMonto(item, "419")],
        ["Casillero 429", (item) => getMonto(item, "429")],
        ["Casillero 485", (item) => getMonto(item, "485")],
        ["Casillero 499", (item) => getMonto(item, "499")],
        ["Casillero 519", (item) => getMonto(item, "519")],
        ["Casillero 529", (item) => getMonto(item, "529")],
        ["Casillero 721", (item) => getMonto(item, "721")],
        ["Casillero 723", (item) => getMonto(item, "723")],
        ["Casillero 725", (item) => getMonto(item, "725")],
        ["Casillero 727", (item) => getMonto(item, "727")],
        ["Casillero 729", (item) => getMonto(item, "729")],
        ["Casillero 731", (item) => getMonto(item, "731")],
        ["Casillero 799", (item) => getMonto(item, "799")],
    ];

    const RESUMEN_GROUPS = [
        { label: "INGRESOS", rows: RESUMEN_ROWS_EXTRA.slice(0, 2) },
        { label: "LIQUIDACION DEL IVA EN EL MES", rows: RESUMEN_ROWS_EXTRA.slice(2, 4) },
        { label: "COSTOS / ADQUISICIONES", rows: RESUMEN_ROWS_EXTRA.slice(4, 6) },
        { label: "RESUMEN IMPOSITIVO", rows: RESUMEN_ROWS.slice(5) },
        { label: "RETENCIONES", rows: RESUMEN_ROWS_EXTRA.slice(6) }
    ];

    const dom = {
        fileInput: document.getElementById("fileInput"),
        processButton: resetNode("processButton"),
        clearButton: resetNode("clearButton"),
        printButton: resetNode("printButton"),
        exportExcelButton: document.getElementById("exportExcelButton"),
        analisisToggle: document.getElementById("toggleAnalisisDetalle"),
        datosContribuyenteDiv: document.getElementById("datosContribuyente"),
        datosDeclaracionDiv: document.getElementById("datosDeclaracion"),
        detalleArchivo: document.getElementById("detalleArchivo"),
        resultado: document.getElementById("resultado"),
        tituloContacto: document.getElementById("tituloContacto"),
        accionesDiv: document.querySelector(".datos-declaracion-actions"),
        panelDerecho: document.querySelector(".declaracion-col.derecha"),
        resumenSection: document.getElementById("reporteResumenSection"),
        resumenTitle: document.getElementById("toggleResumen"),
        resumenMeta: document.getElementById("reporteResumenMeta"),
        resumenTableWrapper: document.getElementById("resumenTableWrapper"),
        resumenTableHead: document.querySelector("#resumenTable thead"),
        resumenTableBody: document.querySelector("#resumenTable tbody"),
        modal: document.getElementById("modalSuscripcion"),
        closeModal: document.getElementById("closeModal"),
        btnCancelar: document.getElementById("btnCancelar"),
        btnContinuarRegistro: document.getElementById("btnContinuarRegistro"),
        modalVideoTutorial: document.getElementById("modalVideoTutorial"),
        modalConfirmacion: document.getElementById("modalConfirmacion"),
        btnCerrarConfirmacion: document.getElementById("btnCerrarConfirmacion"),
        contenidoConfirmacion: document.getElementById("contenidoConfirmacion"),
        modalTituloPlan: document.getElementById("modalTituloPlan"),
        modalRUC: document.getElementById("modalRUC"),
        modalRazonSocial: document.getElementById("modalRazonSocial"),
        modalPeriodo: document.getElementById("modalPeriodo"),
        modalSaldo: document.getElementById("modalSaldo"),
        modalCaracteristicas: document.getElementById("modalCaracteristicas"),
        formRegistro: document.getElementById("formRegistro"),
        inputEmail: document.getElementById("inputEmail"),
        inputTelefono: document.getElementById("inputTelefono"),
        inputDireccion: document.getElementById("inputDireccion"),
        checkboxContratar: document.getElementById("checkboxContratar")
    };

    const state = {
        declaraciones: [],
        seleccionadaId: null,
        analisisDetallado: true,
        datosPlanActual: getEmptyPlanData()
    };

    dom.fileInput.setAttribute("multiple", "multiple");

    function resetNode(id) {
        const oldNode = document.getElementById(id);
        const newNode = oldNode.cloneNode(true);
        oldNode.parentNode.replaceChild(newNode, oldNode);
        return newNode;
    }

    function getEmptyPlanData() {
        return {
            plan: "",
            saldo: 0,
            ruc: "",
            razonSocial: "",
            periodoFiscal: "",
            tipoServicio: "recuperaIVA",
            fileName: ""
        };
    }

    function normalizeText(value) {
        return String(value || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toUpperCase();
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function parseAmount(value) {
        const normalized = String(value || "0").replace(/,/g, "");
        const parsed = Number.parseFloat(normalized);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function formatAmount(value) {
        return Number(value || 0).toFixed(2);
    }

    function buildFlexiblePattern(label) {
        return label
            .replace(/[aáàäâ]/gi, "[aáàäâAÁÀÄÂ]")
            .replace(/[eéèëê]/gi, "[eéèëêEÉÈËÊ]")
            .replace(/[iíìïî]/gi, "[iíìïîIÍÌÏÎ]")
            .replace(/[oóòöô]/gi, "[oóòöôOÓÒÖÔ]")
            .replace(/[uúùüû]/gi, "[uúùüûUÚÙÜÛ]")
            .replace(/n/gi, "[nñNÑ]")
            .replace(/\s+/g, "\\s+");
    }

    function getFieldValue(pattern, nextPattern, text) {
        const regex = new RegExp(`${pattern}\\s*:?\\s*(.*?)\\s*(?=${nextPattern}|$)`, "is");
        const match = text.match(regex);
        return match ? match[1].trim() : "";
    }

    function determinePlan(saldo) {
        if (saldo <= 0) {
            return { nombre: "SIN SALDO", imagen: "", texto: "No tienes saldo de retenciones de IVA que se pueda recuperar", estado: "SIN SALDO" };
        }
        if (saldo > 2500) {
            return { nombre: "ASESORIA", imagen: "img/plan-asesoria.jpg", texto: "Te recomendamos contactarnos para un plan ASESORIA", estado: "VALIDA" };
        }
        if (saldo <= 400) {
            return { nombre: "REGISTRO", imagen: "img/registro.jpg", texto: "Te recomendamos la licencia REGISTRO", estado: "VALIDA" };
        }
        if (saldo <= 1200) {
            return { nombre: "BASICA", imagen: "img/basica.jpg", texto: "Te recomendamos la licencia BASICA", estado: "VALIDA" };
        }
        return { nombre: "PREMIUM", imagen: "img/premium.jpg", texto: "Te recomendamos la licencia PREMIUM", estado: "VALIDA" };
    }

    function getPeriodoInfo(periodoFiscal) {
        const source = normalizeText(periodoFiscal);
        const monthLabels = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
        const monthNameMap = {
            ENE: 1, ENERO: 1, FEB: 2, FEBRERO: 2, MAR: 3, MARZO: 3, ABR: 4, ABRIL: 4,
            MAY: 5, MAYO: 5, JUN: 6, JUNIO: 6, JUL: 7, JULIO: 7, AGO: 8, AGOSTO: 8,
            SEP: 9, SEPT: 9, SEPTIEMBRE: 9, OCT: 10, OCTUBRE: 10, NOV: 11, NOVIEMBRE: 11,
            DIC: 12, DICIEMBRE: 12
        };

        let month = 0;
        let year = 0;
        let label = periodoFiscal || "SIN PERIODO";
        let sortKey = 0;
        let match = source.match(/\b(PRIMER|PRIMERO|1ER|1RO)\s+SEMESTRE\s+(20\d{2})\b/);

        if (match) {
            year = Number(match[2]);
            month = 6;
            label = `PRIMER SEMESTRE ${year}`;
            sortKey = (year * 100) + 6.5;
        }

        if (!month || !year) {
            match = source.match(/\b(SEGUNDO|2DO|2DA)\s+SEMESTRE\s+(20\d{2})\b/);
            if (match) {
                year = Number(match[2]);
                month = 12;
                label = `SEGUNDO SEMESTRE ${year}`;
                sortKey = (year * 100) + 12.5;
            }
        }

        if (!month || !year) {
            match = source.match(/\b(20\d{2})(0[1-9]|1[0-2])\b/);
            if (match) {
                year = Number(match[1]);
                month = Number(match[2]);
            }
        }

        if (!month || !year) {
            match = source.match(/\b(0?[1-9]|1[0-2])\D+(20\d{2})\b/);
            if (match) {
                month = Number(match[1]);
                year = Number(match[2]);
            }
        }

        if (!month || !year) {
            match = source.match(/\b(20\d{2})\D+(0?[1-9]|1[0-2])\b/);
            if (match) {
                year = Number(match[1]);
                month = Number(match[2]);
            }
        }

        if (!month || !year) {
            match = source.match(/\b([A-Z]+)\s+(20\d{2})\b/);
            if (match && monthNameMap[match[1]]) {
                month = monthNameMap[match[1]];
                year = Number(match[2]);
            }
        }

        if (!month || !year) {
            match = source.match(/\b(20\d{2})\s+([A-Z]+)\b/);
            if (match && monthNameMap[match[2]]) {
                year = Number(match[1]);
                month = monthNameMap[match[2]];
            }
        }

        if (!month || !year) {
            return { month: 0, year: 0, sortKey: 0, label };
        }

        if (!sortKey) {
            sortKey = (year * 100) + month;
            label = `${monthLabels[month - 1]}-${String(year).slice(-2)}`;
        }

        return { month, year, sortKey, label };
    }

    function getMonto(item, codigo) {
        const row = (item.credito || []).find((entry) => entry.codigo === codigo);
        return row ? formatAmount(row.monto) : "0.00";
    }

    function getVisibleDeclarations() {
        return state.declaraciones
            .filter((item) => item.valid)
            .sort((a, b) => (b.periodoInfo?.sortKey || 0) - (a.periodoInfo?.sortKey || 0));
    }

    function getResumenRows() {
        if (!state.analisisDetallado) {
            return RESUMEN_ROWS;
        }

        return [
            ...RESUMEN_ROWS.slice(0, 5),
            ...RESUMEN_GROUPS.flatMap(({ label, rows }) => [
                [label, () => ""],
                ...rows
            ])
        ];
    }

    function getResumenTitleHtml() {
        const title = state.analisisDetallado
            ? "ANALISIS MENSUAL COMPLETO acumulado"
            : "ANALISIS MENSUAL acumulado";
        return `${title}<span class="click-text">(clic para revisar)</span>`;
    }

    function getResumenRowClass(label) {
        if (!state.analisisDetallado) {
            return "";
        }

        for (const group of RESUMEN_GROUPS) {
            if (label === group.label) {
                return "resumen-grupo-inicio resumen-grupo-titulo";
            }
            const rowIndex = group.rows.findIndex(([rowLabel]) => rowLabel === label);
            if (rowIndex >= 0) {
                return rowIndex === group.rows.length - 1
                    ? "resumen-grupo-fin"
                    : "resumen-grupo-cuerpo";
            }
        }

        return "";
    }

    function getCreditoGroup(codigo) {
        const rowLabel = `Casillero ${codigo}`;
        const group = RESUMEN_GROUPS.find(({ rows }) => rows.some(([label]) => label === rowLabel));
        return group?.label || "RESUMEN IMPOSITIVO";
    }

    function applyFrequentRucValidation(declaraciones) {
        const frequency = new Map();

        declaraciones.forEach((item) => {
            if (!item.valid) {
                return;
            }

            const ruc = String(item.contribuyente?.ruc || "").trim();
            if (!ruc) {
                return;
            }

            if (!frequency.has(ruc)) {
                frequency.set(ruc, { count: 0, maxSortKey: item.periodoInfo?.sortKey || 0 });
            }

            const info = frequency.get(ruc);
            info.count += 1;
            info.maxSortKey = Math.max(info.maxSortKey, item.periodoInfo?.sortKey || 0);
        });

        if (!frequency.size) {
            return declaraciones;
        }

        let frequentRuc = "";
        let bestCount = -1;
        let bestSortKey = -1;

        frequency.forEach((info, ruc) => {
            if (info.count > bestCount || (info.count === bestCount && info.maxSortKey > bestSortKey)) {
                frequentRuc = ruc;
                bestCount = info.count;
                bestSortKey = info.maxSortKey;
            }
        });

        return declaraciones.map((item) => {
            if (!item.valid) {
                return item;
            }

            const ruc = String(item.contribuyente?.ruc || "").trim();
            if (!ruc || ruc === frequentRuc) {
                return item;
            }

            return {
                ...item,
                valid: false,
                estado: "INVALIDA",
                error: `RUC distinto al mas frecuente del lote (${frequentRuc}).`,
                saldo617: 0,
                plan: "NO APLICA",
                planInfo: {
                    nombre: "NO APLICA",
                    imagen: "",
                    texto: "",
                    estado: "INVALIDA"
                }
            };
        });
    }

    function extractContribuyente(text) {
        const obligacionTributaria = getFieldValue(
            buildFlexiblePattern("Obligacion Tributaria"),
            buildFlexiblePattern("Identificacion"),
            text
        );

        const normalizedObligacion = normalizeText(obligacionTributaria);
        if (!normalizedObligacion.includes("DECLARACION") || !normalizedObligacion.includes("IVA")) {
            return { valid: false, error: "El archivo no corresponde a una declaracion de IVA valida." };
        }

        const values = {
            obligacionTributaria,
            identificacion: getFieldValue(
                buildFlexiblePattern("Identificacion"),
                buildFlexiblePattern("Razon Social"),
                text
            ),
            razonSocial: getFieldValue(
                buildFlexiblePattern("Razon Social"),
                buildFlexiblePattern("Periodo Fiscal"),
                text
            ),
            periodoFiscal: getFieldValue(
                buildFlexiblePattern("Periodo Fiscal"),
                buildFlexiblePattern("Tipo Declaracion"),
                text
            ),
            tipoDeclaracion: getFieldValue(
                buildFlexiblePattern("Tipo Declaracion"),
                buildFlexiblePattern("Formulario Sustituye"),
                text
            ),
            formularioSustituye: getFieldValue(
                buildFlexiblePattern("Formulario Sustituye"),
                "TARIFA\\s+VARIABLE",
                text
            )
        };

        if (normalizeText(values.tipoDeclaracion) === "ORIGINAL") {
            values.formularioSustituye = "";
        }

        return {
            valid: true,
            values,
            ruc: values.identificacion,
            razonSocial: values.razonSocial,
            periodoFiscal: values.periodoFiscal,
            tipoDeclaracion: values.tipoDeclaracion
        };
    }

    function extractDeclaracion(text) {
        const source = normalizeText(text);
        const regex = /CODIGO VERIFICADOR\s+NUMERO SERIAL\s+FECHA RECAUDACION\s+(\S+)\s+(\d+)\s+(\d{2}-\d{2}-\d{4})/i;
        const match = source.match(regex);

        return {
            codigoVerificador: match ? match[1] : "No encontrado",
            numeroSerial: match ? match[2] : "No encontrado",
            fechaRecaudacion: match ? match[3] : "No encontrado"
        };
    }

    function extractCredito(fullText) {
        return CREDITO_FIELDS.map(([codigo, descripcion]) => {
            const regex = new RegExp(`(?:^|\\s)${codigo}\\s+([\\d,.]+)(?=\\s|$)`);
            const match = fullText.match(regex);
            const montoTexto = match ? match[1] : "0.00";

            return {
                codigo,
                descripcion,
                montoTexto,
                monto: parseAmount(montoTexto)
            };
        });
    }

    function buildDeclaracion(fileName, fullText, firstPageText) {
        const contribuyente = extractContribuyente(fullText);

        if (!contribuyente.valid) {
            return {
                id: `${Date.now()}-${Math.random()}`,
                fileName,
                valid: false,
                estado: "INVALIDA",
                error: contribuyente.error,
                saldo617: 0,
                plan: "NO APLICA",
                planInfo: {
                    nombre: "NO APLICA",
                    imagen: "",
                    texto: "",
                    estado: "INVALIDA"
                },
                periodoInfo: { month: 0, year: 0, sortKey: 0, label: fileName },
                contribuyente: {},
                declaracion: {},
                credito: []
            };
        }

        const declaracion = extractDeclaracion(firstPageText);
        const credito = extractCredito(fullText);
        const saldo617 = credito.find((row) => row.codigo === "617")?.monto || 0;
        const planInfo = determinePlan(saldo617);

        return {
            id: `${Date.now()}-${Math.random()}`,
            fileName,
            valid: true,
            estado: planInfo.estado,
            error: "",
            saldo617,
            plan: planInfo.nombre,
            planInfo,
            periodoInfo: getPeriodoInfo(contribuyente.periodoFiscal),
            contribuyente,
            declaracion,
            credito
        };
    }

    async function readArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => resolve(event.target.result);
            reader.onerror = () => reject(new Error(`No se pudo leer el archivo ${file.name}`));
            reader.readAsArrayBuffer(file);
        });
    }

    async function extractPdfTexts(pdf) {
        let fullText = "";

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
            const page = await pdf.getPage(pageNumber);
            const textContent = await page.getTextContent();
            fullText += `${textContent.items.map((item) => item.str).join(" ")}\n`;
        }

        const firstPage = await pdf.getPage(1);
        const firstPageContent = await firstPage.getTextContent();
        const firstPageText = firstPageContent.items.map((item) => item.str).join(" ");

        return { fullText, firstPageText };
    }

    async function procesarArchivo(file) {
        const fileBytes = await readArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument(new Uint8Array(fileBytes)).promise;
        const texts = await extractPdfTexts(pdf);
        return buildDeclaracion(file.name, texts.fullText, texts.firstPageText);
    }

    function renderResumen() {
        const columnas = getVisibleDeclarations();
        const resumenRows = getResumenRows();

        dom.resumenTitle.innerHTML = getResumenTitleHtml();

        dom.resumenTableHead.innerHTML = `
            <tr>
                <th>Concepto</th>
                ${columnas.map((item) => `<th>${escapeHtml(item.periodoInfo?.label || item.fileName)}</th>`).join("")}
            </tr>
        `;

        dom.resumenTableBody.innerHTML = resumenRows.map(([label, resolver]) => `
            <tr class="${getResumenRowClass(label)}">
                <td><strong>${escapeHtml(label)}</strong></td>
                ${columnas.map((item) => `<td class="row-ok">${escapeHtml(resolver(item))}</td>`).join("")}
            </tr>
        `).join("");

        const validas = columnas.length;
        const invalidas = state.declaraciones.length - validas;
        const total609 = columnas.reduce((sum, item) => sum + parseAmount(getMonto(item, "609")), 0);
        const reciente = columnas[0] || null;
        const valor617Reciente = reciente ? formatAmount(reciente.saldo617) : "0.00";

        dom.resumenMeta.textContent = `${state.declaraciones.length} meses cargados | ${validas} validos | ${invalidas} invalidos | Total de retenciones recibidas: USD ${formatAmount(total609)} | Casillero 617 actual: USD ${valor617Reciente}`;
        dom.resumenSection.classList.toggle("is-visible", columnas.length > 0);
        dom.exportExcelButton.classList.toggle("is-visible", columnas.length > 0);
        dom.resumenTableWrapper.classList.remove("is-hidden");
    }

    function renderContribuyente(contribuyente) {
        const rows = [
            ["Obligacion Tributaria", contribuyente.values?.obligacionTributaria || ""],
            ["Identificacion", contribuyente.values?.identificacion || ""],
            ["Razon Social", contribuyente.values?.razonSocial || ""],
            ["Periodo Fiscal", contribuyente.values?.periodoFiscal || ""],
            ["Tipo Declaracion", contribuyente.values?.tipoDeclaracion || ""],
            ["Formulario Sustituye", contribuyente.values?.formularioSustituye || ""]
        ];

        dom.datosContribuyenteDiv.innerHTML = `
            <table class="datos-contribuyente-table">
                <thead>
                    <tr>
                        <th>Campo</th>
                        <th>Valor</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows.map(([label, value]) => `
                        <tr>
                            <td><strong>${escapeHtml(label)}</strong></td>
                            <td>${escapeHtml(value)}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        `;
    }

    function renderDatosDeclaracion(declaracion) {
        dom.datosDeclaracionDiv.innerHTML = `
            <table class="datos-declaracion-table">
                <thead>
                    <tr>
                        <th>CODIGO VERIFICADOR</th>
                        <th>NUMERO SERIAL</th>
                        <th>FECHA RECAUDACION</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>${escapeHtml(declaracion.codigoVerificador || "")}</td>
                        <td>${escapeHtml(declaracion.numeroSerial || "")}</td>
                        <td>${escapeHtml(declaracion.fechaRecaudacion || "")}</td>
                    </tr>
                </tbody>
            </table>
        `;
    }

    function resetDetalle() {
        dom.detalleArchivo.textContent = "";
        dom.datosContribuyenteDiv.innerHTML = "";
        dom.datosDeclaracionDiv.innerHTML = "";
        dom.resultado.textContent = "";
        dom.resultado.classList.remove("error", "success", "resultado");
        dom.resultado.classList.add("is-hidden");
        dom.tituloContacto.classList.add("is-hidden");
        dom.accionesDiv.innerHTML = "";
        dom.accionesDiv.classList.add("is-hidden");
        dom.panelDerecho.innerHTML = "";
    }

    function openVideoTutorial() {
        window.open(VIDEO_TUTORIAL_URL, "VideoTutorial", VIDEO_WINDOW_FEATURES);
    }

    function renderAcciones(item) {
        dom.accionesDiv.innerHTML = "";
        dom.accionesDiv.classList.add("is-hidden");
        dom.panelDerecho.innerHTML = "";
        dom.resultado.classList.add("is-hidden");
        dom.tituloContacto.classList.add("is-hidden");

        if (!item.valid) {
            dom.resultado.textContent = item.error || "Declaracion invalida.";
            dom.resultado.classList.add("resultado", "error");
            dom.resultado.classList.remove("success");
            dom.resultado.classList.remove("is-hidden");
            return;
        }

        if (item.saldo617 <= 0) {
            dom.resultado.textContent = "No tienes saldo de retenciones de IVA que se pueda recuperar";
            dom.resultado.classList.add("resultado", "success");
            dom.resultado.classList.remove("error");
            dom.resultado.classList.remove("is-hidden");
            dom.tituloContacto.innerHTML = 'Quieres saber mas de nosotros - <a href="https://www.youtube.com/@rentabilidadECUADOR" target="_blank" class="contact-link">@rentabilidadECUADOR</a>';
            dom.tituloContacto.classList.remove("is-hidden");
            return;
        }

        dom.resultado.textContent = `EXCELENTE, tienes saldo de retenciones de IVA a recuperar por USD ${formatAmount(item.saldo617)}`;
        dom.resultado.classList.add("resultado", "success");
        dom.resultado.classList.remove("error");
        dom.resultado.classList.remove("is-hidden");
        dom.tituloContacto.textContent = "Asegura tu recuperacion - CONTACTANOS";
        dom.tituloContacto.classList.remove("is-hidden");
        dom.accionesDiv.classList.remove("is-hidden");

        if (item.plan !== "ASESORIA") {
            const videoButton = document.createElement("button");
            videoButton.className = "verificaCT617-button small-button";
            videoButton.textContent = "VIDEO TUTORIAL";
            videoButton.addEventListener("click", openVideoTutorial);
            dom.accionesDiv.appendChild(videoButton);
        }

        const registroButton = document.createElement("button");
        registroButton.className = "verificaCT617-button";
        registroButton.textContent = item.plan === "ASESORIA" ? "ASESORIA recuperaIVA" : "REGISTRARSE";
        registroButton.addEventListener("click", () => abrirModal(item));
        dom.accionesDiv.appendChild(registroButton);

        const texto = document.createElement("p");
        texto.className = "plan-recomendado-text";
        texto.textContent = item.planInfo?.texto || "";
        dom.panelDerecho.appendChild(texto);

        if (item.planInfo?.imagen) {
            const imagen = document.createElement("img");
            imagen.className = "plan-recomendado-image";
            imagen.src = item.planInfo.imagen;
            imagen.alt = `Imagen del plan ${item.plan.toLowerCase()}`;
            dom.panelDerecho.appendChild(imagen);
        }
    }

    function seleccionarDeclaracion(id) {
        state.seleccionadaId = id;
        const item = state.declaraciones.find((row) => row.id === id);
        if (!item) {
            return;
        }

        dom.detalleArchivo.textContent = `Declaracion mas reciente: ${item.fileName}`;

        if (!item.valid) {
            dom.datosContribuyenteDiv.innerHTML = "";
            dom.datosDeclaracionDiv.innerHTML = "";
            renderAcciones(item);
            return;
        }

        renderContribuyente(item.contribuyente);
        renderDatosDeclaracion(item.declaracion);
        renderAcciones(item);
    }

    function clearAll() {
        state.declaraciones = [];
        state.seleccionadaId = null;
        state.analisisDetallado = true;
        state.datosPlanActual = getEmptyPlanData();
        dom.fileInput.value = "";
        dom.resumenTableHead.innerHTML = "";
        dom.resumenTableBody.innerHTML = "";
        dom.resumenMeta.textContent = "";
        dom.resumenSection.classList.remove("is-visible");
        dom.resumenTableWrapper.classList.remove("is-hidden");
        dom.exportExcelButton.classList.remove("is-visible");
        dom.resumenTitle.innerHTML = getResumenTitleHtml();
        dom.analisisToggle.classList.remove("is-active");
        dom.analisisToggle.title = "Clic para ampliar o reducir casilleros del analisis";
        resetDetalle();
    }

    function formatPeriodoForFile(periodoInfo) {
        const year = Number(periodoInfo?.year || 0);
        const month = Number(periodoInfo?.month || 0);
        if (!year || !month) {
            return "SINPERIODO";
        }
        return `${year}${String(month).padStart(2, "0")}`;
    }

    function applyResumenSheetPresentation(worksheet, totalColumns) {
        const dataColumns = Math.max(1, totalColumns - 1);
        const firstColumnWidth = 20;
        const otherColumnWidth = 14;
        const tokens = getComputedStyle(document.documentElement);
        const navy = tokens.getPropertyValue("--gjv-navy").trim().replace(/^#/, "").toUpperCase();
        const band = tokens.getPropertyValue("--gjv-banda").trim().replace(/^#/, "").toUpperCase();

        worksheet["!cols"] = [
            { wch: firstColumnWidth },
            ...Array.from({ length: dataColumns }, () => ({ wch: otherColumnWidth }))
        ];

        worksheet["!rows"] = Array.from({ length: 6 }, () => ({ hpt: 13 }));

        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
        for (let row = range.s.r; row <= range.e.r; row += 1) {
            for (let col = range.s.c; col <= range.e.c; col += 1) {
                const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellRef];
                if (!cell) {
                    continue;
                }

                cell.s = cell.s || {};
                cell.s.alignment = {
                    vertical: "center",
                    horizontal: col === 0 ? "left" : "center",
                    wrapText: false
                };
                cell.s.font = {
                    name: "Arial",
                    sz: row >= 1 && row <= 5 ? 8 : 9,
                    bold: row === 0 || col === 0
                };

                const rowLabel = String(worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v || "");
                const groupClass = getResumenRowClass(rowLabel);
                if (groupClass) {
                    const border = { style: "medium", color: { rgb: navy } };
                    cell.s.border = {
                        left: border,
                        right: border,
                        ...(groupClass.includes("resumen-grupo-inicio") ? { top: border } : {}),
                        ...(groupClass.includes("resumen-grupo-fin") ? { bottom: border } : {})
                    };
                    if (groupClass.includes("resumen-grupo-titulo")) {
                        cell.s.fill = { patternType: "solid", fgColor: { rgb: band } };
                        cell.s.font.bold = true;
                        cell.s.font.color = { rgb: navy };
                    }
                }
            }
        }
    }

    function exportarExcel() {
        if (!state.declaraciones.length) {
            alert("Primero procesa al menos una declaracion.");
            return;
        }

        if (typeof XLSX === "undefined") {
            alert("La libreria de Excel no esta disponible.");
            return;
        }

        const columnas = getVisibleDeclarations();
        if (!columnas.length) {
            alert("No hay declaraciones validas para exportar.");
            return;
        }

        const resumenAoA = [
            ["Concepto", ...columnas.map((item) => item.periodoInfo?.label || item.fileName)],
            ...getResumenRows().map(([label, resolver]) => [label, ...columnas.map((item) => {
                const value = resolver(item);
                return /^Casillero /.test(label) ? Number(value) : value;
            })])
        ];

        const creditoData = columnas.flatMap((item, index) => item.credito.map((row) => ({
            N: index + 1,
            ARCHIVO: item.fileName,
            RUC: item.contribuyente.ruc || "",
            RAZON_SOCIAL: item.contribuyente.razonSocial || "",
            GRUPO: getCreditoGroup(row.codigo),
            CASILLERO: row.codigo,
            DESCRIPCION: row.descripcion,
            MONTO: row.monto
        })));

        const resumenSheet = XLSX.utils.aoa_to_sheet(resumenAoA);
        applyResumenSheetPresentation(resumenSheet, resumenAoA[0].length);

        const rucLote = columnas[0]?.contribuyente?.ruc || "SINRUC";
        const periodosAsc = [...columnas].sort((a, b) => (a.periodoInfo?.sortKey || 0) - (b.periodoInfo?.sortKey || 0));
        const periodoDesde = formatPeriodoForFile(periodosAsc[0]?.periodoInfo);
        const periodoHasta = formatPeriodoForFile(periodosAsc[periodosAsc.length - 1]?.periodoInfo);
        const prefijoArchivo = state.analisisDetallado
            ? "reporteCOMPLETO-recuperaIVA"
            : "reporte-recuperaIVA";
        const nombreArchivo = `${prefijoArchivo} ${rucLote} ${periodoDesde}-${periodoHasta}.xlsx`;

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, resumenSheet, "Resumen");
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(creditoData), "Credito");
        XLSX.writeFile(workbook, nombreArchivo);
    }

    function getCaracteristicasPlan(nombrePlan, saldo) {
        if (nombrePlan === "REGISTRO") {
            return [
                { icon: "1", texto: "Video tutorial (como RECUPERAR IVA)" },
                { icon: "2", texto: "Uso programa unirTXT" },
                { icon: "3", texto: "Uso programa unirPDF" },
                { icon: "4", texto: "Crea archivo VENTAS (formato XLS)" },
                { icon: "5", texto: "Crea archivo RETENCIONES IVA (formatos TXT y XLS)" },
                { icon: "6", texto: "Crea archivo saldos MESES" },
                { icon: "7", texto: "Crea SOLICITUD (formatos TXT y PDF)" },
                { icon: "8", texto: "Crea CERTIFICACION (formato PDF) + video tutorial" },
                { icon: "9", texto: "Tiempo estimado preparacion: 1 hora" },
                { icon: "10", texto: `Recuperacion hasta $${saldo.toFixed(2)}` },
                { icon: "11", texto: "Licencia: 30 dias" }
            ];
        }

        if (nombrePlan === "BASICA") {
            return [
                { icon: "1", texto: "Video tutorial (como RECUPERAR IVA)" },
                { icon: "2", texto: "Uso programa unirTXT" },
                { icon: "3", texto: "Uso programa unirPDF" },
                { icon: "4", texto: "Crea archivo VENTAS (formato XLS)" },
                { icon: "5", texto: "Crea archivo RETENCIONES IVA (formatos TXT y XLS)" },
                { icon: "6", texto: "Crea archivo saldos MESES" },
                { icon: "7", texto: "Crea SOLICITUD (formatos TXT y PDF)" },
                { icon: "8", texto: "Crea CERTIFICACION (formato PDF) + video tutorial" },
                { icon: "9", texto: "Tiempo estimado preparacion: menos de 30 minutos" },
                { icon: "10", texto: `Recuperacion hasta $${saldo.toFixed(2)}` },
                { icon: "11", texto: "Licencia: 30 dias" }
            ];
        }

        if (nombrePlan === "PREMIUM") {
            return [
                { icon: "1", texto: "Video tutorial (como RECUPERAR IVA)" },
                { icon: "2", texto: "Uso programa unirTXT" },
                { icon: "3", texto: "Uso programa unirPDF" },
                { icon: "4", texto: "Crea archivo VENTAS (formato XLS)" },
                { icon: "5", texto: "Crea archivo RETENCIONES IVA (formatos TXT, XLS y rtaXLS)" },
                { icon: "6", texto: "Crea archivo saldos MESES" },
                { icon: "7", texto: "Crea SOLICITUD (formatos TXT y PDF)" },
                { icon: "8", texto: "Crea CERTIFICACION (formatos TXT, PDF) + video tutorial" },
                { icon: "9", texto: "Tiempo estimado preparacion: 10 minutos" },
                { icon: "10", texto: `Recuperacion hasta $${saldo.toFixed(2)}` },
                { icon: "11", texto: "Licencia: 60 dias" },
                { icon: "12", texto: "Video tutorial (errores mas frecuentes al RECUPERAR IVA)" }
            ];
        }

        if (nombrePlan === "ASESORIA") {
            return [
                { icon: "1", texto: "Todo lo incluido en PREMIUM" },
                { icon: "2", texto: "Contacto directo por WhatsApp" },
                { icon: "3", texto: "Asesor senior dedicado" },
                { icon: "4", texto: "Gestion prioritaria" },
                { icon: "5", texto: "Preparacion de toda documentacion" },
                { icon: "6", texto: "Seguimiento basico del tramite" },
                { icon: "7", texto: "Alertas en tiempo real" },
                { icon: "8", texto: `Recuperacion hasta $${saldo.toFixed(2)}` },
                { icon: "9", texto: "Plan personalizado segun necesidades" }
            ];
        }

        return [];
    }

    function mostrarModalConfirmacion(datosPlan, email, telefono, deseaContratar) {
        const esPlanAsesoria = datosPlan.plan === "ASESORIA";

        dom.contenidoConfirmacion.innerHTML = `
            <div class="dato-item"><strong>RUC:</strong> ${escapeHtml(datosPlan.ruc)}</div>
            <div class="dato-item"><strong>Razon Social:</strong> ${escapeHtml(datosPlan.razonSocial)}</div>
            <div class="dato-item"><strong>Servicio:</strong> ${escapeHtml(datosPlan.tipoServicio)}</div>
            <div class="dato-item"><strong>Plan:</strong> ${escapeHtml(datosPlan.plan)}</div>
            <div class="dato-item"><strong>Email:</strong> ${escapeHtml(email)}</div>
            <div class="dato-item"><strong>Telefono:</strong> ${escapeHtml(telefono)}</div>
            <div class="dato-item"><strong>Desea contratar:</strong> ${escapeHtml(deseaContratar)}</div>
            <div class="correo-confirmacion">
                <strong class="correo-confirmacion-titulo">Email enviado</strong>
                <p class="correo-confirmacion-detalle">Hemos enviado un correo ${esPlanAsesoria ? "informativo" : "con los datos bancarios"} a:<br><strong>${escapeHtml(email)}</strong></p>
            </div>
            <div class="separador">
                <h3>Lee esto atentamente</h3>
                <p class="advertencia-intro">Si no ves el email en 5 minutos:</p>
                <div class="advertencia-item">Revisa tu carpeta de <strong>SPAM</strong></div>
                <div class="advertencia-item">Revisa <strong>CORREO NO DESEADO</strong></div>
                <div class="advertencia-item">Revisa carpeta <strong>PROMOCIONES</strong></div>
                <p class="advertencia-seguimiento">${esPlanAsesoria ? "El email contiene informacion sobre tu asesoria personalizada." : "El email contiene los datos para activar tu licencia."}</p>
            </div>
            <div class="contacto-confirmacion">
                <p>Si no encuentras el email, contactanos indicando tu RUC:</p>
                <p><strong>WhatsApp: +593 96 367 5173</strong></p>
            </div>
            <p class="nota-confirmacion">Pronto nos pondremos en contacto contigo.</p>
        `;

        dom.modalConfirmacion.classList.add("is-open");
    }

    function abrirModal(item) {
        state.datosPlanActual = {
            plan: item.plan,
            saldo: item.saldo617,
            ruc: item.contribuyente.ruc || "",
            razonSocial: item.contribuyente.razonSocial || "",
            periodoFiscal: item.contribuyente.periodoFiscal || "",
            tipoServicio: "recuperaIVA",
            fileName: item.fileName
        };

        const planModal = item.plan === "BASICA" ? "BASICA" : item.plan === "ASESORIA" ? "ASESORIA" : item.plan;

        dom.modalTituloPlan.textContent = `Plan ${planModal}`;
        dom.modalRUC.textContent = state.datosPlanActual.ruc || "No disponible";
        dom.modalRazonSocial.textContent = state.datosPlanActual.razonSocial || "No disponible";
        dom.modalPeriodo.textContent = state.datosPlanActual.periodoFiscal || "No disponible";
        dom.modalSaldo.textContent = `$${formatAmount(state.datosPlanActual.saldo)}`;
        dom.modalCaracteristicas.innerHTML = "";

        getCaracteristicasPlan(planModal, item.saldo617).forEach((caracteristica) => {
            const div = document.createElement("div");
            div.className = "caracteristica-item";
            div.innerHTML = `
                <span class="caracteristica-icon">${escapeHtml(caracteristica.icon)}</span>
                <span class="caracteristica-texto">${escapeHtml(caracteristica.texto)}</span>
            `;
            dom.modalCaracteristicas.appendChild(div);
        });

        dom.btnContinuarRegistro.textContent = item.plan === "ASESORIA"
            ? "Registrar y Contactar por WhatsApp"
            : "Continuar con el Registro";

        dom.modal.classList.add("is-open");
    }

    async function registrarDesdeModal() {
        if (!dom.formRegistro.checkValidity()) {
            dom.formRegistro.reportValidity();
            return;
        }

        const email = dom.inputEmail.value.trim();
        const telefono = dom.inputTelefono.value.trim();
        const direccion = dom.inputDireccion.value.trim();
        const deseaContratar = dom.checkboxContratar.checked ? "SI" : "NO";

        let versionConPrecio = `version ${state.datosPlanActual.plan}`;
        if (state.datosPlanActual.plan === "REGISTRO") {
            versionConPrecio += " (USD 9)";
        } else if (state.datosPlanActual.plan === "BASICA") {
            versionConPrecio += " (USD 29)";
        } else if (state.datosPlanActual.plan === "PREMIUM") {
            versionConPrecio += " (USD 49)";
        } else if (state.datosPlanActual.plan === "ASESORIA") {
            versionConPrecio += " (Cotizacion Personalizada)";
        }

        let clave = "";
        if (state.datosPlanActual.plan !== "ASESORIA" && state.datosPlanActual.ruc.length >= 9) {
            clave = state.datosPlanActual.ruc.substring(3, 9);
        }

        const payload = {
            ruc: state.datosPlanActual.ruc,
            razonSocial: state.datosPlanActual.razonSocial,
            direccion,
            email,
            telefono: `'${telefono}`,
            deseaContratar,
            numeroIngreso: "",
            tipoServicio: state.datosPlanActual.tipoServicio,
            versionContratar: versionConPrecio,
            clave: `'${clave}`,
            clase: state.datosPlanActual.plan,
            caduca: ""
        };

        dom.btnContinuarRegistro.disabled = true;
        dom.btnContinuarRegistro.textContent = "Guardando...";

        try {
            await fetch(WEBAPP_URL, {
                method: "POST",
                mode: "no-cors",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            mostrarModalConfirmacion(state.datosPlanActual, email, telefono, deseaContratar);

            if (state.datosPlanActual.plan === "ASESORIA") {
                let mensaje = "Hola, estoy interesado en la ASESORIA para recuperar mi IVA%0A";
                mensaje += `Servicio: ${state.datosPlanActual.tipoServicio}%0A`;
                mensaje += `${state.datosPlanActual.ruc} | ${state.datosPlanActual.razonSocial} | ${state.datosPlanActual.periodoFiscal}%0A`;
                mensaje += `Saldo de CT a recuperar: USD ${formatAmount(state.datosPlanActual.saldo)}%0A`;
                mensaje += `Email: ${email}%0A`;
                mensaje += `Telefono: ${telefono}%0A`;
                mensaje += "Por favor, contactenme para mas informacion";
                setTimeout(() => window.open(`https://wa.me/593963675173?text=${mensaje}`, "_blank"), 1000);
            }

            dom.formRegistro.reset();
            dom.modal.classList.remove("is-open");
        } catch (error) {
            console.error("Error al enviar registro:", error);
            alert("Hubo un problema al guardar tu registro. Intenta nuevamente o contactanos por WhatsApp.");
        } finally {
            dom.btnContinuarRegistro.disabled = false;
            dom.btnContinuarRegistro.textContent = "REGISTRARSE";
        }
    }

    async function procesarSeleccion() {
        const files = Array.from(dom.fileInput.files || []);

        if (!files.length) {
            alert("Por favor, selecciona al menos un archivo PDF antes de procesar.");
            return;
        }

        if (files.length > MAX_FILES) {
            alert(`Solo se permiten hasta ${MAX_FILES} declaraciones por lote.`);
            return;
        }

        clearAll();
        dom.processButton.disabled = true;
        dom.processButton.textContent = "Procesando...";

        try {
            const resultados = [];

            for (const file of files) {
                try {
                    resultados.push(await procesarArchivo(file));
                } catch (error) {
                    resultados.push({
                        id: `${Date.now()}-${Math.random()}`,
                        fileName: file.name,
                        valid: false,
                        estado: "ERROR",
                        error: error.message || "No se pudo procesar el archivo.",
                        saldo617: 0,
                        plan: "NO APLICA",
                        planInfo: {
                            nombre: "NO APLICA",
                            imagen: "",
                            texto: "",
                            estado: "INVALIDA"
                        },
                        periodoInfo: { month: 0, year: 0, sortKey: 0, label: file.name },
                        contribuyente: {},
                        declaracion: {},
                        credito: []
                    });
                }
            }

            state.declaraciones = applyFrequentRucValidation(resultados)
                .sort((a, b) => (a.periodoInfo?.sortKey || 0) - (b.periodoInfo?.sortKey || 0));

            renderResumen();

            const reciente = getVisibleDeclarations()[0] || state.declaraciones[state.declaraciones.length - 1];
            if (reciente) {
                seleccionarDeclaracion(reciente.id);
            }
        } finally {
            dom.processButton.disabled = false;
            dom.processButton.textContent = "Procesar";
        }
    }

    function imprimir() {
        document.body.classList.add("imprimir");
        window.print();
        document.body.classList.remove("imprimir");
    }

    function toggleResumen() {
        dom.resumenTableWrapper.classList.toggle("is-hidden");
    }

    function toggleAnalisisDetallado() {
        state.analisisDetallado = !state.analisisDetallado;
        dom.analisisToggle.classList.toggle("is-active", state.analisisDetallado);
        dom.analisisToggle.title = state.analisisDetallado
            ? "Modo ampliado activo: clic para volver al analisis base"
            : "Clic para ampliar o reducir casilleros del analisis";

        if (state.declaraciones.length) {
            renderResumen();
        }
    }

    function closeModal() {
        dom.modal.classList.remove("is-open");
    }

    function closeConfirmacion() {
        dom.modalConfirmacion.classList.remove("is-open");
    }

    dom.processButton.addEventListener("click", procesarSeleccion);
    dom.clearButton.addEventListener("click", clearAll);
    dom.printButton.addEventListener("click", imprimir);
    dom.exportExcelButton.addEventListener("click", exportarExcel);
    dom.resumenTitle.addEventListener("click", toggleResumen);
    dom.analisisToggle.addEventListener("click", toggleAnalisisDetallado);
    dom.modalVideoTutorial.addEventListener("click", openVideoTutorial);
    dom.btnContinuarRegistro.addEventListener("click", registrarDesdeModal);
    dom.closeModal.addEventListener("click", closeModal);
    dom.btnCancelar.addEventListener("click", closeModal);
    dom.btnCerrarConfirmacion.addEventListener("click", closeConfirmacion);

    window.addEventListener("click", (event) => {
        if (event.target === dom.modal) {
            closeModal();
        }
        if (event.target === dom.modalConfirmacion) {
            closeConfirmacion();
        }
    });

    clearAll();
})();
