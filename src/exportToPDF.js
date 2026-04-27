import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas';

export async function exportToPDF(
  allData, region,
  chartBarRef, chartPieRef, chartHomeRef, chartMissingRef,
  setPdfLoading
) {
  setPdfLoading(true);
  try {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 30;
    const usableWidth = pageWidth - margin * 2;

    // ── Page header helper ──
    const addPageHeader = (title, subtitle) => {
      doc.setFillColor(102, 126, 234);
      doc.rect(0, 0, pageWidth, 40, 'F');
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text(title, pageWidth / 2, 26, { align: 'center' });
      if (subtitle) {
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text(subtitle, pageWidth - margin, 26, { align: 'right' });
      }
      doc.setTextColor(0, 0, 0);
    };

    // ── Wait for charts to render ──
    await new Promise(r => setTimeout(r, 1200));

    // ── Capture canvas helper ──
    const captureCanvas = async (ref) => {
      if (!ref?.current) return null;
      await new Promise(r => setTimeout(r, 800));
      return await html2canvas(ref.current, {
        scale: 1.5,
        useCORS: true,
        backgroundColor: '#ffffff',
        width: ref.current.scrollWidth,
        height: ref.current.scrollHeight,
        scrollX: -99999,
        scrollY: -99999,
        windowWidth: ref.current.scrollWidth,
        windowHeight: ref.current.scrollHeight,
        x: 0,
        y: 0
      });
    };

    // ── Capture all 4 charts ──
    const [canvasBar, canvasPie, canvasHome, canvasMissing] = await Promise.all([
      captureCanvas(chartBarRef),
      captureCanvas(chartPieRef),
      captureCanvas(chartHomeRef),
      captureCanvas(chartMissingRef)
    ]);

    // ── Page 1: Dashboard 2×2 grid ──
    addPageHeader('Rate Analysis Dashboard', new Date().toLocaleDateString());

    const headerH = 40, gap = 10;
    const cellW = (usableWidth - gap) / 2;
    const cellH = (pageHeight - headerH - margin - gap) / 2;

    const positions = [
      { x: margin, y: headerH + gap, label: 'Monthly Burn Comparison' },
      { x: margin + cellW + gap, y: headerH + gap, label: 'Classification Distribution' },
      { x: margin, y: headerH + gap + cellH + gap, label: 'Missing Classifications' },
      { x: margin + cellW + gap, y: headerH + gap + cellH + gap, label: 'Resource Flags' },
    ];
    const canvases = [canvasBar, canvasPie, canvasMissing, canvasHome];

    canvases.forEach((canvas, i) => {
      if (!canvas) return;
      const { x, y, label } = positions[i];
      doc.setDrawColor(200, 200, 200);
      doc.rect(x, y, cellW, cellH);
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(80, 80, 80);
      doc.text(label, x + cellW / 2, y + 10, { align: 'center' });
      const imgPad = 14;
      const imgW = cellW - imgPad;
      const imgH = (canvas.height / canvas.width) * imgW;
      const finalH = Math.min(imgH, cellH - imgPad);
      doc.addImage(canvas.toDataURL('image/png'), 'PNG', x + imgPad / 2, y + imgPad, imgW, finalH);
    });

    // ── Build table data from allData ──
    const headers = [
      'ESA ID', 'ESA Desc', 'VZ TQ ID', 'VZ TQ Desc', 'POC', 'Emp ID', 'Name',
      'Location', 'Country', 'ACT/PCT', 'Skill Set', 'VZ Level', 'Classification',
      'Key', 'Designation', 'Service Line', 'Timesheet Hrs', 'Rate ₹/hr', 'Rate $/hr',
      'Projected $', 'Actual $', 'Variance', 'Jan-26', 'Feb-26', 'Mar-26'
    ];
    const fieldKeys = [
      'esaId', 'esaDesc', 'vzTqId', 'vzTqDesc', 'poc', 'empId', 'name',
      'location', 'country', 'actPct', 'skillSet', 'verizonLevel', 'classification',
      'key', 'designation', 'serviceLine', 'timesheetHrs', 'rateInr', 'rateUsd',
      'projectedRate', 'actualRate', 'variance', 'jan26', 'feb26', 'mar26'
    ];

    const sheetData = allData.map(row => fieldKeys.map(k => row[k] ?? ''));

    // ── Column width calculation ──
    const colWidths = headers.map((header, colIdx) => {
      const maxLen = sheetData.reduce(
        (max, row) => Math.max(max, String(row[colIdx] || '').length),
        header.length
      );
      return Math.min(Math.max(maxLen * 5.5, 40), 160);
    });
    const scale = usableWidth / colWidths.reduce((a, b) => a + b, 0);
    const scaledWidths = colWidths.map(w => w * scale);

    // ── Data table page helper ──
    const addDataTable = (label, rows) => {
      doc.addPage();
      addPageHeader('Verizon Home & Marketing Burnsheet', label);
      doc.setFontSize(8);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(80, 80, 80);
      doc.text(
        `Records: ${rows.length}   |   Region: ${region}   |   Date: ${new Date().toLocaleDateString()}`,
        margin, 52
      );
      doc.setTextColor(0, 0, 0);
      autoTable(doc, {
        head: [headers],
        body: rows,
        startY: 60,
        styles: {
          fontSize: 6,
          cellPadding: { top: 2, right: 3, bottom: 2, left: 3 },
          overflow: 'ellipsize',
          halign: 'left',
          valign: 'middle',
          lineColor: [220, 220, 220],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: [102, 126, 234],
          textColor: 255,
          fontStyle: 'bold',
          fontSize: 6.5
        },
        alternateRowStyles: { fillColor: [248, 248, 248] },
        columnStyles: Object.fromEntries(scaledWidths.map((w, idx) => [idx, { cellWidth: w }])),
        margin: { left: margin, right: margin },
        tableWidth: usableWidth
      });
    };

    // ── Filter and generate table pages ──
    const countryIndex = fieldKeys.indexOf('country');
    const indiaRows = sheetData.filter(row => String(row[countryIndex]).toLowerCase() === 'india');
    const usaRows = sheetData.filter(row => String(row[countryIndex]).toLowerCase() === 'usa');
    addDataTable('India', indiaRows);
    addDataTable('USA', usaRows);

    // ── Save ──
    doc.save(`burnsheet-${new Date().toISOString().slice(0, 10)}.pdf`);
  } finally {
    setPdfLoading(false);
  }
}
