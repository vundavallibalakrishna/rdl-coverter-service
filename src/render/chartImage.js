import PDFDocument from 'pdfkit';
import { drawChart } from './chart.js';
import { rasterizePdf } from './raster.js';

function collectDocument(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
  });
}

// Renders a single chart to a PNG buffer by drawing it onto a one-page, chart-sized PDF and
// rasterizing it with Poppler. Used to embed charts into the otherwise-native editable DOCX.
export async function renderChartPng(chart, data, config, tempDir, context, index) {
  const width = Math.max(1, chart.width);
  const height = Math.max(1, chart.height);
  const doc = new PDFDocument({ autoFirstPage: false, compress: true });
  const completion = collectDocument(doc);
  doc.addPage({ size: [width, height], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
  drawChart(doc, config, chart, data, 0, 0, width, height, context);
  doc.end();
  const pdfBuffer = await completion;
  const pages = await rasterizePdf(pdfBuffer, config, tempDir, `chart-${index}`, { dpi: 150, singleFile: true });
  return pages[0] || null;
}
