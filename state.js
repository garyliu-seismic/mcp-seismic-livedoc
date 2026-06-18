'use strict';

// ── Global state ───────────────────────────────────────────────────────────────
/** @type {Array<{id:string, kind:'scalar'|'table'|'computed'|'system', name:string, group:string, [key:string]:any}>} */
let variables = [];

const expandedIds = new Set(); // IDs of table/system rows currently expanded
let searchQuery   = '';
let editMode      = 'add';    // 'add' | 'edit' | 'dynamic' | 'dynimage' | 'dynchart'
let editKind      = 'scalar';
let editId        = null;

// Tracks the name of the currently selected shape in the slide (updated on selection change)
let selectedTableName   = null;
let selectedShapeName   = null;   // any non-indicator shape — used by dynimage drawer
let selectedChartName   = null;   // any shape selected while dynchart drawer is open
let currentSlideId      = null;   // Tracks current slide; used to detect slide switches
let pendingImageDataUrl = null;   // file uploaded in the dynimage drawer; survives form rebuilds

// ── Constants ──────────────────────────────────────────────────────────────────
const XML_NS    = 'http://schemas.livedoc.seismic.com/poc-variables/v2';
const XML_NS_V1 = 'http://schemas.livedoc.seismic.com/poc-variables/v1';

const TYPE_LABELS = { STRING: 'ABC', NUMBER: '123', DATE: 'DT', BOOLEAN: 'T/F' };

// Predefined system variable blueprints
const SYSTEM_PRESETS = {
  TOCEntries: {
    columns:     ['EntryName', 'EntryLevel', 'PageNumber', 'Index'],
    columnTypes: ['STRING', 'NUMBER', 'NUMBER', 'NUMBER'],
    description: 'Table of contents entries'
  }
};
