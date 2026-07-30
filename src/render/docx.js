// DOCX_EDITABLE is intentionally a Windows Word page-locked renderer. The former continuous/reflowable
// implementation was removed: it estimated pagination independently from PDF and therefore could not
// guarantee the canonical page count, fragment borders, repeated headers, or page-dependent bands.
export { renderPagedEditableDocx as renderEditableDocx } from './pagedDocx.js';
