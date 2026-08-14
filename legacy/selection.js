'use strict';

// ── Selection tracking ─────────────────────────────────────────────────────────
// Listens for DocumentSelectionChanged and syncs drawer state when the user
// clicks a different shape or navigates to a different slide.

function initSelectionHandler() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSlideSelectionChanged,
    function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('SelectionChanged handler registration failed:', result.error && result.error.message);
      }
    }
  );
}

async function onSlideSelectionChanged() {
  try {
    var newSlideId   = null;
    var foundTable   = null;
    var foundShape   = null;
    var dynConfig    = null;
    var imgConfig    = null;
    var chartConfig  = null;

    await PowerPoint.run(async function (context) {
      // Detect current slide ID so we can tell when the user switches slides
      try {
        var selSlides = context.presentation.getSelectedSlides();
        selSlides.load('items/id');
        await context.sync();
        newSlideId = selSlides.items.length ? selSlides.items[0].id : null;
      } catch (_) {}

      // Detect selected shapes
      try {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load('items/name,items/type');
        await context.sync();
        foundTable = shapes.items.find(function (s) {
          return s.type === 'Table' ||
                 (typeof PowerPoint !== 'undefined' &&
                  typeof PowerPoint.ShapeType !== 'undefined' &&
                  s.type === PowerPoint.ShapeType.table);
        }) || null;
        foundShape = shapes.items.find(function (s) {
          return !s.name.startsWith('__LIVEDOC_IND_');
        }) || null;
      } catch (_) {}

      // Load dynamic table config if table drawer is open
      if (editMode === 'dynamic' && foundTable) {
        try {
          foundTable.tags.load('items/key,items/value');
          await context.sync();
          var dynTag = foundTable.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_TABLE'; });
          dynConfig = dynTag ? JSON.parse(dynTag.value) : null;
        } catch (_) {}
      }

      // Load dynamic image config if image drawer is open
      if (editMode === 'dynimage' && foundShape) {
        try {
          foundShape.tags.load('items/key,items/value');
          await context.sync();
          var imgTag = foundShape.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_IMAGE'; });
          imgConfig = imgTag ? JSON.parse(imgTag.value) : null;
        } catch (_) {}
      }

      // Load dynamic chart config if chart drawer is open
      if (editMode === 'dynchart' && foundShape) {
        try {
          foundShape.tags.load('items/key,items/value');
          await context.sync();
          var chartTag = foundShape.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_CHART'; });
          chartConfig = chartTag ? JSON.parse(chartTag.value) : null;
        } catch (_) {}
      }
    });

    var slideChanged  = (newSlideId !== null && newSlideId !== currentSlideId);
    if (newSlideId !== null) currentSlideId = newSlideId;
    selectedTableName = foundTable ? foundTable.name : null;
    selectedShapeName = foundShape ? foundShape.name : null;
    if (editMode === 'dynchart') selectedChartName = foundShape ? foundShape.name : null;

    if (editMode === 'dynamic') {
      if (slideChanged) {
        await refreshDynamicDrawer();
      } else if (selectedTableName) {
        syncDynamicDrawerDropdown(selectedTableName);
        syncDynamicDrawerFromConfig(dynConfig);
      }
    } else if (editMode === 'dynimage') {
      if (slideChanged) {
        await refreshDynamicImageDrawer();
      } else if (selectedShapeName) {
        syncDynamicImageDrawerDropdown(selectedShapeName);
        syncDynamicImageDrawerFromConfig(imgConfig);
      }
    } else if (editMode === 'dynchart') {
      if (slideChanged) {
        await refreshDynamicChartDrawer();
      } else if (selectedChartName) {
        syncDynamicChartDrawerDropdown(selectedChartName);
        if (chartConfig) syncDynamicChartDrawerFromConfig(chartConfig);
      }
    }
  } catch (_) {
    selectedTableName = null;
    selectedShapeName = null;
  }
}
