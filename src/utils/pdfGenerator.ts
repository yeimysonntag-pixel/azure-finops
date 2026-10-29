import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { CustomerGroup } from './excelGenerator';
import { detectMonthFromRows } from './excelGenerator';

export async function exportCustomerPdf(
  customer: CustomerGroup,
  periodCode: string
): Promise<void> {
  const monthName = detectMonthFromRows(customer.records, periodCode);
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });

  const pageWidth = doc.internal.pageSize.getWidth();

  // 1. ENCABEZADO CON LOGO DE TIVIT ALMAVIVA GROUP
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(22);
  doc.setTextColor(192, 0, 0); // Rojo TIVIT
  doc.text('TIVIT', pageWidth - 20, 20, { align: 'right' });

  doc.setFontSize(8);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(26, 54, 93);
  doc.text('ALMAVIVA GROUP', pageWidth - 20, 25, { align: 'right' });

  // 2. METADATOS DEL REPORTE
  doc.setFontSize(18);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(15, 23, 42);
  doc.text('Reporte Ejecutivo de Consumo Cloud', 20, 22);

  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(100, 116, 139);
  doc.text('Plataforma Azure FinOps Manager', 20, 28);

  doc.setDrawColor(226, 232, 240);
  doc.setLineWidth(0.5);
  doc.line(20, 32, pageWidth - 20, 32);

  // 3. TARJETA DATOS CLIENTE Y PERIODO
  doc.setFillColor(248, 250, 252);
  doc.roundedRect(20, 36, pageWidth - 40, 24, 3, 3, 'F');

  doc.setFontSize(9);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('CLIENTE:', 25, 45);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(customer.name, 45, 45);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('PERIODO:', 25, 53);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(monthName, 45, 53);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('PROVEEDOR:', pageWidth / 2 + 10, 45);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text('Microsoft Azure', pageWidth / 2 + 35, 45);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('COSTO TOTAL:', pageWidth / 2 + 10, 53);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(2, 132, 199);
  doc.text(`$${customer.totalCost.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD`, pageWidth / 2 + 35, 53);

  // 4. PREPARACIÓN DE TABLA ORDENADA DE MAYOR A MENOR
  const sortedCategories = Object.entries(customer.categoryTotals)
    .sort((a, b) => b[1] - a[1]);

  const tableRows = sortedCategories.map(([cat, total]) => [
    cat,
    `$ ${total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  ]);

  tableRows.push([
    'Total general',
    `$ ${customer.totalCost.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  ]);

  // 5. GENERACIÓN DE TABLA CON AUTOTABLE
  autoTable(doc, {
    startY: 66,
    head: [['meterCategory', 'Costo']],
    body: tableRows,
    theme: 'grid',
    headStyles: {
      fillColor: [31, 78, 120],
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      halign: 'left',
      fontSize: 10
    },
    bodyStyles: {
      textColor: [51, 65, 85],
      fontSize: 9.5
    },
    columnStyles: {
      0: { cellWidth: 'auto', fontStyle: 'normal' },
      1: { cellWidth: 50, halign: 'right', fontStyle: 'bold' }
    },
    margin: { left: 20, right: 20 },
    didParseCell: (data) => {
      if (data.row.index === tableRows.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [241, 245, 249];
        data.cell.styles.textColor = [15, 23, 42];
      }
    }
  });

  // 6. PIE DE PÁGINA PROFESIONAL
  const totalPages = (doc as any).internal.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    const pageHeight = doc.internal.pageSize.getHeight();

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(148, 163, 184);

    doc.text('Documento Confidencial - TIVIT FinOps Portal', 20, pageHeight - 10);
    doc.text(`Página ${i} de ${totalPages}`, pageWidth - 20, pageHeight - 10, { align: 'right' });
  }

  const safeName = customer.name.replace(/[/\\?%*:|"<>]/g, '_').trim();
  doc.save(`${monthName.toUpperCase()}_CONSUMOS_AZURE_${safeName}.pdf`);
}