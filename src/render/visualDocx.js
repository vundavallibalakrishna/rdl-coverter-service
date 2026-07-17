import {
  Document, HorizontalPositionRelativeFrom, ImageRun, PageOrientation, Packer, PageBreak, Paragraph,
  TextWrappingType, VerticalPositionRelativeFrom,
} from 'docx';
import { pointsToDisplayPixels, pointsToTwips } from '../units.js';
import { ServiceError } from '../errors.js';
import { renderPdf } from './pdf.js';
import { rasterizePdf } from './raster.js';

export async function renderVisualDocx(model, request, config, tempDir) {
  const pdf = await renderPdf(model, request, config);
  const pages = await rasterizePdf(pdf.buffer, config, tempDir, 'visual-page', { dpi: 300 });
  if (pages.length !== pdf.pageCount) throw new ServiceError('RENDER_FAILED', 'PDF rasterization produced an unexpected page count');
  const children = [];
  for (const [index, page] of pages.entries()) {
    const data = page.data;
    children.push(new Paragraph({
      spacing: { before: 0, after: 0, line: 1 },
      children: [
        new ImageRun({
        data,
        type: 'png',
        transformation: { width: pointsToDisplayPixels(model.page.width), height: pointsToDisplayPixels(model.page.height) },
        floating: {
          horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
          verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
          allowOverlap: true,
          behindDocument: true,
          layoutInCell: false,
          wrap: { type: TextWrappingType.NONE },
        },
      })],
    }));
    if (index < pages.length - 1) children.push(new Paragraph({ spacing: { before: 0, after: 0, line: 1 }, children: [new PageBreak()] }));
  }
  // The docx library swaps width/height for landscape, so pass pre-rotation dimensions.
  const pageSize = model.page.width > model.page.height
    ? { width: pointsToTwips(model.page.height), height: pointsToTwips(model.page.width), orientation: PageOrientation.LANDSCAPE }
    : { width: pointsToTwips(model.page.width), height: pointsToTwips(model.page.height), orientation: PageOrientation.PORTRAIT };
  const document = new Document({
    creator: 'RDL Converter Service',
    title: request.outputFileName || model.name,
    sections: [{
      properties: { page: { size: pageSize, margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0 } } },
      children,
    }],
  });
  const buffer = await Packer.toBuffer(document);
  return { buffer, pageCount: pdf.pageCount, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', extension: 'docx' };
}
