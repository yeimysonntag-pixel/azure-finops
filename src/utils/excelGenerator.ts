import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

export interface CsvRow {
  invoiceSectionName?: string;
  customerName?: string;
  meterId?: string;
  chargeType?: string;
  meterName?: string;
  meterCategory?: string;
  meterSubCategory?: string;
  serviceFamily?: string;
  product?: string;
  payGPrice?: string | number;
  date?: string;
  unitOfMeasure?: string;
  quantity?: string | number;
  billingPeriodEndDate?: string;
  billingPeriodStartDate?: string;
  provider?: string;
  customerTenantId?: string;
  invoiceId?: string;
  subscriptionName?: string;
  resourceGroupName?: string;
  PaygCostInBillingCurrency?: string | number;
  [key: string]: any;
}

export interface CustomerGroup {
  name: string;
  country: string;
  totalCost: number;
  records: CsvRow[];
  categoryTotals: Record<string, number>;
}

const borderThinBlack: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: '000000' } },
  left: { style: 'thin', color: { argb: '000000' } },
  bottom: { style: 'thin', color: { argb: '000000' } },
  right: { style: 'thin', color: { argb: '000000' } }
};

const borderTotal: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: '000000' } },
  left: { style: 'thin', color: { argb: '000000' } },
  bottom: { style: 'double', color: { argb: '000000' } },
  right: { style: 'thin', color: { argb: '000000' } }
};

export function detectCountry(text: string): string {
  if (!text) return 'Chile';
  const str = String(text).toUpperCase();
  if (str.includes('COLOMBIA') || str.includes('_CO') || str.includes(' CO ')) return 'Colombia';
  if (str.includes('PERU') || str.includes('PERÚ') || str.includes('_PE')) return 'Perú';
  if (str.includes('MEXICO') || str.includes('MÉXICO') || str.includes('_MX')) return 'México';
  if (str.includes('CHILE') || str.includes('SENAME') || str.includes('_CL') || str.includes('CL')) return 'Chile';
  return 'Chile';
}

export function detectMonthFromRows(rows: CsvRow[], defaultPeriodCode: string = ''): string {
  const months = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  if (!rows || rows.length === 0) return 'Septiembre';

  for (const row of rows.slice(0, 50)) {
    let rawDate: any = null;

    for (const key of Object.keys(row)) {
      const lowerKey = key.toLowerCase();
      if (
        lowerKey.includes('billingperiodenddate') ||
        lowerKey.includes('billingperiodstartdate') ||
        lowerKey.includes('periodo') ||
        lowerKey.includes('date')
      ) {
        if (row[key]) {
          rawDate = row[key];
          break;
        }
      }
    }

    if (!rawDate && defaultPeriodCode) {
      rawDate = defaultPeriodCode;
    }

    if (rawDate) {
      const strDate = String(rawDate).trim();

      const matchUS = strDate.match(/^(\d{1,2})[-/]\d{1,2}[-/](\d{4})/);
      if (matchUS) {
        const m = parseInt(matchUS[1], 10) - 1;
        if (m >= 0 && m < 12) return months[m];
      }

      const matchIso = strDate.match(/^(\d{4})[-/](\d{1,2})[-/]/);
      if (matchIso) {
        const m = parseInt(matchIso[2], 10) - 1;
        if (m >= 0 && m < 12) return months[m];
      }

      const cleanDigits = strDate.replace(/\D/g, '');
      if (cleanDigits.length >= 6) {
        const m = parseInt(cleanDigits.substring(4, 6), 10) - 1;
        if (m >= 0 && m < 12) return months[m];
      }
    }
  }

  return 'Septiembre';
}

// RENDERIZADO DEL GRÁFICO
async function createCategoryChartBase64(categoryTotals: Record<string, number>, monthName: string): Promise<{ base64: string; width: number; height: number }> {
  const entries = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);
  const itemCount = entries.length;

  const calculatedHeight = Math.max(380, 100 + itemCount * 32);
  const canvas = document.createElement('canvas');
  canvas.width = 800;
  canvas.height = calculatedHeight;
  const ctx = canvas.getContext('2d');
  
  if (!ctx) return { base64: '', width: 650, height: 380 };

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  
  ctx.strokeStyle = '#E2E8F0';
  ctx.lineWidth = 1;
  ctx.strokeRect(0, 0, canvas.width, canvas.height);

  ctx.fillStyle = '#1F4E78';
  ctx.font = 'bold 18px "Plus Jakarta Sans", Calibri, sans-serif';
  ctx.fillText(`Distribución de Consumos - ${monthName}`, 25, 38);

  if (itemCount === 0) return { base64: canvas.toDataURL('image/png'), width: 680, height: calculatedHeight / 1.1 };

  const maxVal = Math.max(...entries.map(e => e[1])) || 1;
  
  let maxLabelLength = 0;
  entries.forEach(([cat]) => {
    if (cat.length > maxLabelLength) maxLabelLength = cat.length;
  });
  
  const startX = Math.min(280, Math.max(200, maxLabelLength * 7.5));
  const startY = 70;
  const barHeight = 22;
  const maxBarWidth = canvas.width - startX - 130;

  entries.forEach(([cat, val], idx) => {
    const y = startY + idx * 32;

    ctx.fillStyle = '#334155';
    ctx.font = '500 12px "Plus Jakarta Sans", Calibri, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(cat, startX - 12, y + 15);

    const barWidth = Math.max(8, (val / maxVal) * maxBarWidth);
    const gradient = ctx.createLinearGradient(startX, 0, startX + barWidth, 0);
    gradient.addColorStop(0, '#0284C7');
    gradient.addColorStop(1, '#0369A1');

    ctx.fillStyle = gradient;
    ctx.beginPath();
    ctx.roundRect(startX, y, barWidth, barHeight, 4);
    ctx.fill();

    ctx.fillStyle = '#0F172A';
    ctx.font = 'bold 11px "Plus Jakarta Sans", Calibri, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(`$${val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, startX + barWidth + 10, y + 15);
  });

  return {
    base64: canvas.toDataURL('image/png'),
    width: 680,
    height: Math.round(calculatedHeight / 1.1)
  };
}

// 1. HOJA CAPA
function buildCapaSheet(workbook: ExcelJS.Workbook): void {
  const ws = workbook.addWorksheet('Capa');
  ws.views = [{ showGridLines: true }];

  ws.getRow(3).height = 28;
  ws.getRow(4).height = 28;
  ws.getRow(5).height = 28;
  ws.getRow(6).height = 28;

  ws.mergeCells('B3:C6');
  const logoCell = ws.getCell('B3');
  logoCell.value = 'TIVIT';
  logoCell.font = { name: 'Calibri', size: 26, bold: true, color: { argb: 'C00000' } };
  logoCell.alignment = { vertical: 'middle', horizontal: 'center' };

  ws.mergeCells('D3:F4');
  const formCell = ws.getCell('D3');
  formCell.value = 'FORMULARIO';
  formCell.font = { name: 'Calibri', size: 18, bold: true };
  formCell.alignment = { vertical: 'middle', horizontal: 'center' };

  ws.mergeCells('D5:F6');
  const titleCell = ws.getCell('D5');
  titleCell.value = 'Formulario Consumo\nServicios Cloud';
  titleCell.font = { name: 'Calibri', size: 15, bold: true };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  ws.getCell('G3').value = 'Código';
  ws.getCell('H3').value = 'Revisión';
  ws.getCell('G4').value = 'CCH-FOR-003-018';
  ws.getCell('H4').value = 1;

  ws.getCell('G5').value = 'Clasificación de la Información';
  ws.getCell('H5').value = 'Página';
  ws.getCell('G6').value = 'Interno';
  ws.getCell('H6').value = 'Ver pie de página';

  for (let r = 3; r <= 6; r++) {
    for (let col = 2; col <= 8; col++) {
      const cell = ws.getCell(r, col);
      cell.border = borderThinBlack;
      if (!cell.alignment) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  }

  ['G3', 'H3', 'G5', 'H5'].forEach(key => {
    const cell = ws.getCell(key);
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'A6A6A6' } };
  });

  ['G4', 'H4', 'G6', 'H6'].forEach(key => {
    const cell = ws.getCell(key);
    cell.font = { name: 'Calibri', size: 11, bold: false };
  });

  ws.getColumn('A').width = 4;
  ws.getColumn('B').width = 14;
  ws.getColumn('C').width = 14;
  ws.getColumn('D').width = 28;
  ws.getColumn('E').width = 32;
  ws.getColumn('F').width = 18;
  ws.getColumn('G').width = 30;
  ws.getColumn('H').width = 20;
}

// 2. EXPORTACIÓN INDIVIDUAL
export async function exportSingleCustomerExcel(
  customer: CustomerGroup,
  periodCode: string,
  selectedCostKey: string
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const monthName = detectMonthFromRows(customer.records, periodCode);

  buildCapaSheet(workbook);

  // -------------------------------------------------------------
  // HOJA 1: DATA
  // -------------------------------------------------------------
  const dataSheet = workbook.addWorksheet('DATA');
  dataSheet.views = [{ showGridLines: true }];

  const headersData = [
    'Periodo', 'Cloud Provider', 'invoiceSectionName', 'meterId', 'chargeType',
    'meterName', 'meterCategory', 'meterSubCategory', 'serviceFamily', 'product',
    'payGPrice', 'date', 'unitOfMeasure', 'Quantity ', 'Total Cost'
  ];

  const headerRowData = dataSheet.addRow(headersData);
  headerRowData.height = 26;
  headerRowData.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = borderThinBlack;
  });

  customer.records.forEach((r) => {
    const costVal = parseFloat(String(r[selectedCostKey] || r.paygCostInUsd || r.costInUsd || 0));
    const payGVal = parseFloat(String(r.payGPrice || 0));
    const qtyVal = parseFloat(String(r.quantity || 0));

    const row = dataSheet.addRow([
      monthName,
      r.provider || 'Azure',
      r.invoiceSectionName || customer.name,
      r.meterId || '',
      r.chargeType || 'Usage',
      r.meterName || '',
      r.meterCategory || '',
      r.meterSubCategory || '',
      r.serviceFamily || '',
      r.product || '',
      payGVal,
      r.date || '',
      r.unitOfMeasure || '',
      qtyVal,
      costVal
    ]);

    row.eachCell((cell, colNumber) => {
      cell.border = borderThinBlack;
      cell.font = { name: 'Calibri', size: 10 };
      if (colNumber === 11) cell.numFmt = '#,##0.0000';
      if (colNumber === 14) cell.numFmt = '#,##0.00';
      if (colNumber === 15) cell.numFmt = '$#,##0.00';
    });
  });

  dataSheet.columns.forEach((column) => {
    let maxLen = 12;
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    column.width = Math.min(maxLen + 4, 45);
  });

  // -------------------------------------------------------------
  // HOJA 2: RESUMEN (TABLA DINÁMICA DE CATEGORÍAS + GRÁFICO)
  // -------------------------------------------------------------
  const summarySheet = workbook.addWorksheet('RESUMEN');
  summarySheet.views = [{ showGridLines: true }];

  summarySheet.getCell('B2').value = `Resumen de Consumos - Mes de ${monthName}`;
  summarySheet.getCell('B2').font = { name: 'Calibri', size: 13, bold: true, color: { argb: '1F4E78' } };

  const sumHeader1 = summarySheet.getCell('B3');
  sumHeader1.value = 'meterCategory';
  const sumHeader2 = summarySheet.getCell('C3');
  sumHeader2.value = 'Cost';

  [sumHeader1, sumHeader2].forEach((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = borderThinBlack;
  });

  let rIdx = 4;
  let maxCatLength = 15;
  const sortedCategories = Object.entries(customer.categoryTotals).sort((a, b) => b[1] - a[1]);

  sortedCategories.forEach(([cat, total]) => {
    if (cat.length > maxCatLength) maxCatLength = cat.length;

    const cCell = summarySheet.getCell(`B${rIdx}`);
    cCell.value = cat;
    cCell.border = borderThinBlack;

    const vCell = summarySheet.getCell(`C${rIdx}`);
    vCell.value = Number(total.toFixed(2));
    vCell.numFmt = '$#,##0.00';
    vCell.border = borderThinBlack;
    rIdx++;
  });

  const totLabel = summarySheet.getCell(`B${rIdx}`);
  totLabel.value = 'Total general';
  totLabel.font = { name: 'Calibri', size: 11, bold: true };
  totLabel.border = borderTotal;

  const totVal = summarySheet.getCell(`C${rIdx}`);
  totVal.value = { formula: `SUM(C4:C${rIdx - 1})`, result: customer.totalCost };
  totVal.font = { name: 'Calibri', size: 11, bold: true };
  totVal.numFmt = '$#,##0.00';
  totVal.border = borderTotal;

  summarySheet.getColumn('B').width = Math.max(30, maxCatLength + 5);
  summarySheet.getColumn('C').width = 22;

  // Insertar Gráfico
  const chartConfig = await createCategoryChartBase64(customer.categoryTotals, monthName);
  if (chartConfig.base64) {
    const imageId = workbook.addImage({
      base64: chartConfig.base64,
      extension: 'png',
    });

    summarySheet.addImage(imageId, {
      tl: { col: 4, row: 2 },
      ext: { width: chartConfig.width, height: chartConfig.height }
    });
  }

  // -------------------------------------------------------------
  // HOJA 3: DETALLE_RESUMEN (TABLA DE 8 CAMPOS)
  // -------------------------------------------------------------
  const detailSheet = workbook.addWorksheet('DETALLE_RESUMEN');
  detailSheet.views = [{ showGridLines: true }];

  const headersDetalle = [
    'PROCESO',
    'invoiceId',
    'subscriptionName',
    'invoiceSectionName',
    'chargeType',
    'resourceGroupName',
    'serviceFamily',
    'PaygCostInBillingCurrency'
  ];

  const headerRowDetalle = detailSheet.addRow(headersDetalle);
  headerRowDetalle.height = 26;
  headerRowDetalle.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = borderThinBlack;
  });

  let totalDetalleCost = 0;

  customer.records.forEach((r) => {
    const costVal = parseFloat(
      String(r.PaygCostInBillingCurrency || r[selectedCostKey] || r.paygCostInUsd || r.costInUsd || 0)
    );
    totalDetalleCost += costVal;

    const row = detailSheet.addRow([
      monthName,
      r.invoiceId || '',
      r.subscriptionName || '',
      r.invoiceSectionName || customer.name,
      r.chargeType || 'Usage',
      r.resourceGroupName || '',
      r.serviceFamily || '',
      costVal
    ]);

    row.eachCell((cell, colNumber) => {
      cell.border = borderThinBlack;
      cell.font = { name: 'Calibri', size: 10 };
      if (colNumber === 8) {
        cell.numFmt = '$#,##0.00';
      }
    });
  });

  const totalRowDetalle = detailSheet.addRow([
    'Total general', '', '', '', '', '', '',
    totalDetalleCost
  ]);

  totalRowDetalle.eachCell((cell, colNumber) => {
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.border = borderTotal;
    if (colNumber === 8) {
      cell.numFmt = '$#,##0.00';
    }
  });

  detailSheet.columns.forEach((column) => {
    let maxLen = 14;
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    column.width = Math.min(maxLen + 4, 45);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const safeName = customer.name.replace(/[/\\?%*:|"<>]/g, '_').trim();
  saveAs(new Blob([buffer]), `${monthName.toUpperCase()}_CONSUMOS_AZURE_${safeName}.xlsx`);
}

// 3. EXPORTACIÓN ZIP MULTIPAÍS
export async function exportAllCustomersZip(
  customersList: CustomerGroup[],
  periodCode: string,
  selectedCostKey: string
): Promise<void> {
  const zip = new JSZip();

  for (const customer of customersList) {
    const workbook = new ExcelJS.Workbook();
    const monthName = detectMonthFromRows(customer.records, periodCode);

    buildCapaSheet(workbook);

    // HOJA 1: DATA
    const dataSheet = workbook.addWorksheet('DATA');
    dataSheet.views = [{ showGridLines: true }];

    const headersData = [
      'Periodo', 'Cloud Provider', 'invoiceSectionName', 'meterId', 'chargeType',
      'meterName', 'meterCategory', 'meterSubCategory', 'serviceFamily', 'product',
      'payGPrice', 'date', 'unitOfMeasure', 'Quantity ', 'Total Cost'
    ];

    const headerRowData = dataSheet.addRow(headersData);
    headerRowData.height = 26;
    headerRowData.eachCell((cell) => {
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderThinBlack;
    });

    customer.records.forEach((r) => {
      const costVal = parseFloat(String(r[selectedCostKey] || r.paygCostInUsd || r.costInUsd || 0));
      const row = dataSheet.addRow([
        monthName,
        r.provider || 'Azure',
        r.invoiceSectionName || customer.name,
        r.meterId || '',
        r.chargeType || 'Usage',
        r.meterName || '',
        r.meterCategory || '',
        r.meterSubCategory || '',
        r.serviceFamily || '',
        r.product || '',
        parseFloat(String(r.payGPrice || 0)),
        r.date || '',
        r.unitOfMeasure || '',
        parseFloat(String(r.quantity || 0)),
        costVal
      ]);

      row.eachCell((cell, colNumber) => {
        cell.border = borderThinBlack;
        cell.font = { name: 'Calibri', size: 10 };
        if (colNumber === 11) cell.numFmt = '#,##0.0000';
        if (colNumber === 14) cell.numFmt = '#,##0.00';
        if (colNumber === 15) cell.numFmt = '$#,##0.00';
      });
    });

    dataSheet.columns.forEach((column) => {
      let maxLen = 12;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = len;
      });
      column.width = Math.min(maxLen + 4, 45);
    });

    // HOJA 2: RESUMEN
    const summarySheet = workbook.addWorksheet('RESUMEN');
    summarySheet.views = [{ showGridLines: true }];

    summarySheet.getCell('B2').value = `Resumen de Consumos - Mes de ${monthName}`;
    summarySheet.getCell('B2').font = { name: 'Calibri', size: 13, bold: true, color: { argb: '1F4E78' } };

    const sH1 = summarySheet.getCell('B3');
    sH1.value = 'meterCategory';
    const sH2 = summarySheet.getCell('C3');
    sH2.value = 'Cost';

    [sH1, sH2].forEach((cell) => {
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderThinBlack;
    });

    let rIdx = 4;
    let maxCatLength = 15;
    const sortedCategories = Object.entries(customer.categoryTotals).sort((a, b) => b[1] - a[1]);

    sortedCategories.forEach(([cat, total]) => {
      if (cat.length > maxCatLength) maxCatLength = cat.length;

      const cCell = summarySheet.getCell(`B${rIdx}`);
      cCell.value = cat;
      cCell.border = borderThinBlack;

      const vCell = summarySheet.getCell(`C${rIdx}`);
      vCell.value = Number(total.toFixed(2));
      vCell.numFmt = '$#,##0.00';
      vCell.border = borderThinBlack;
      rIdx++;
    });

    const totLabel = summarySheet.getCell(`B${rIdx}`);
    totLabel.value = 'Total general';
    totLabel.font = { name: 'Calibri', size: 11, bold: true };
    totLabel.border = borderTotal;

    const totVal = summarySheet.getCell(`C${rIdx}`);
    totVal.value = { formula: `SUM(C4:C${rIdx - 1})`, result: customer.totalCost };
    totVal.font = { name: 'Calibri', size: 11, bold: true };
    totVal.numFmt = '$#,##0.00';
    totVal.border = borderTotal;

    summarySheet.getColumn('B').width = Math.max(30, maxCatLength + 5);
    summarySheet.getColumn('C').width = 22;

    const chartConfig = await createCategoryChartBase64(customer.categoryTotals, monthName);
    if (chartConfig.base64) {
      const imageId = workbook.addImage({
        base64: chartConfig.base64,
        extension: 'png',
      });

      summarySheet.addImage(imageId, {
        tl: { col: 4, row: 2 },
        ext: { width: chartConfig.width, height: chartConfig.height }
      });
    }

    // HOJA 3: DETALLE_RESUMEN
    const detailSheet = workbook.addWorksheet('DETALLE_RESUMEN');
    detailSheet.views = [{ showGridLines: true }];

    const headersDetalle = [
      'PROCESO',
      'invoiceId',
      'subscriptionName',
      'invoiceSectionName',
      'chargeType',
      'resourceGroupName',
      'serviceFamily',
      'PaygCostInBillingCurrency'
    ];

    const headerRowDetalle = detailSheet.addRow(headersDetalle);
    headerRowDetalle.height = 26;
    headerRowDetalle.eachCell((cell) => {
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderThinBlack;
    });

    let totalDetalleCost = 0;

    customer.records.forEach((r) => {
      const costVal = parseFloat(
        String(r.PaygCostInBillingCurrency || r[selectedCostKey] || r.paygCostInUsd || r.costInUsd || 0)
      );
      totalDetalleCost += costVal;

      const row = detailSheet.addRow([
        monthName,
        r.invoiceId || '',
        r.subscriptionName || '',
        r.invoiceSectionName || customer.name,
        r.chargeType || 'Usage',
        r.resourceGroupName || '',
        r.serviceFamily || '',
        costVal
      ]);

      row.eachCell((cell, colNumber) => {
        cell.border = borderThinBlack;
        cell.font = { name: 'Calibri', size: 10 };
        if (colNumber === 8) {
          cell.numFmt = '$#,##0.00';
        }
      });
    });

    const totalRowDetalle = detailSheet.addRow([
      'Total general', '', '', '', '', '', '',
      totalDetalleCost
    ]);

    totalRowDetalle.eachCell((cell, colNumber) => {
      cell.font = { name: 'Calibri', size: 11, bold: true };
      cell.border = borderTotal;
      if (colNumber === 8) {
        cell.numFmt = '$#,##0.00';
      }
    });

    detailSheet.columns.forEach((column) => {
      let maxLen = 14;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = len;
      });
      column.width = Math.min(maxLen + 4, 45);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const safeName = customer.name.replace(/[/\\?%*:|"<>]/g, '_').trim();
    const countryFolder = zip.folder(customer.country);
    countryFolder?.file(`${monthName.toUpperCase()}_CONSUMOS_AZURE_${safeName}.xlsx`, buffer);
  }

  const sampleMonth = detectMonthFromRows(customersList[0]?.records || [], periodCode);
  const zipContent = await zip.generateAsync({ type: 'blob' });
  saveAs(zipContent, `${sampleMonth.toUpperCase()}_CONSUMOS_AZURE_TODOS_LOS_PAISES.zip`);
}