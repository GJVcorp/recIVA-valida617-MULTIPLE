# GJVcorp rI00 recIVA valida617 MULTIPLE

Sistema automatizado de validacion, analisis mensual acumulado, exportacion a Excel y registro comercial para recuperacion de IVA.

---

## Descripcion General

`recuperaIVA valida617 MULTIPLE` es un sistema web para contribuyentes ecuatorianos que permite:
- validar declaraciones de IVA en PDF
- procesar hasta 12 declaraciones en un mismo lote
- mostrar en pantalla solo la declaracion mas reciente
- consolidar los meses en una matriz comparativa
- exportar el informe consolidado a Excel
- ampliar el analisis con casilleros adicionales bajo demanda
- registrar planes de servicio segun el saldo a recuperar

### Objetivo Principal
Simplificar la revision del credito tributario de IVA y la toma de decision comercial sobre el servicio a contratar.

---

## Caracteristicas Principales

### Validacion de PDF
- Extraccion de datos del Formulario 104 mediante PDF.js.
- Lectura de casilleros clave del credito tributario.
- Deteccion del casillero 617 para el saldo a recuperar.
- Validacion por RUC mas frecuente del lote: las declaraciones con otro RUC se marcan como invalidas y no entran al analisis ni al Excel.

### Multicarga
- Acepta hasta 12 PDFs en una sola corrida.
- Ordena las declaraciones por periodo fiscal.
- Mantiene en pantalla la declaracion mas reciente.
- No depende del nombre del archivo para definir el periodo mas reciente.

### Analisis Multimes
- Genera una tabla comparativa con columnas por mes.
- Resume casilleros clave como 601, 602, 605, 606, 607, 609, 615, 617, 618, 619, 620 y 699.
- En modo ampliado agrega 419, 429, 519, 529 y 801.
- El modo ampliado se activa o desactiva haciendo clic sobre `recuperaIVA` en el encabezado.
- Sirve como base directa para la exportacion a Excel.
- La primera columna queda fija en pantalla para facilitar la revision horizontal.

### Registro Comercial
- Modal de registro integrado.
- Envio a Google Sheets.
- Confirmacion visual posterior al registro.
- Flujo especial para ASESORIA via WhatsApp.

---

## Flujo de Usuario

1. El usuario selecciona entre 1 y 12 declaraciones PDF.
2. El sistema procesa el lote y ordena los periodos.
3. En pantalla se muestra solo la declaracion mas reciente.
4. En la tabla inferior se presenta el analisis comparativo de todos los meses validos del lote.
5. Si hace clic en `recuperaIVA`, puede alternar entre analisis base y analisis mensual completo.
6. El usuario puede exportar el analisis a Excel o registrar el plan recomendado del mes mas reciente.

---

## Estructura de Archivos

```text
GJVcorp rI00 recIVA valida617 MULTIPLE/
  index.html
  JavaScript.js
  styles.css
  00 backup recIVA valida617 MULTIPLE.BAT
  00 CHANGELOG recIVA valida617 MULTIPLE.md
  GJVcorp rI00 recIVA valida617 MULTIPLE.md
  img/
  00 webAPP google/
```

---

## Instalacion y Configuracion

### Requisitos
- Navegador moderno
- Conexion a Internet
- Cuenta de Google para Apps Script y Sheets

### Proyecto

```bash
git clone [URL_DEL_REPOSITORIO]
cd "GJVcorp rI00 recIVA valida617 MULTIPLE"
```

### Google Apps Script
1. Crear un proyecto en `script.google.com`.
2. Copiar `00 webAPP google/GoogleAppsScript-WebApp.gs`.
3. Publicar como Web App.
4. Pegar la URL publicada en `JavaScript.js` en la constante `WEBAPP_URL`.

### Servidor Local

```bash
python -m http.server 8000
```

Abrir `http://localhost:8000`.

---

## Exportacion a Excel

La exportacion genera:
- hoja `Resumen`: matriz multimes con columnas por periodo
- hoja `Credito`: detalle por casillero y por archivo procesado

### Nombre del Archivo
- Modo base: `reporte-recuperaIVA RUC YYYYMM-YYYYMM.xlsx`
- Modo completo: `reporteCOMPLETO-recuperaIVA RUC YYYYMM-YYYYMM.xlsx`

### Presentacion
- La hoja `Resumen` reduce el tamano de fuente en el bloque superior para mejorar legibilidad.
- La primera columna se exporta con mayor ancho para conservar los conceptos visibles.

---

## Soporte

- WhatsApp: +593 96 367 5173
- Email: impuestosrecuperadosEC@gmail.com
- YouTube: `@rentabilidadECUADOR`

---

## Proyecto

- Empresa: GJVcorp
- Proyecto: `rI00 recIVA valida617 MULTIPLE`
- Version documental: `v82`
