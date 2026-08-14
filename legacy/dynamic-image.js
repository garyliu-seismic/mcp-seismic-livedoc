'use strict';

// ── Dynamic Image Settings ─────────────────────────────────────────────────────
//
// A slide shape is tagged as a dynamic image placeholder by storing a JSON config
// in shape.tags under LIVEDOC_DYN_IMAGE: { fitMode, sourceUrl }.
// At preview time the placeholder <p:sp> is replaced by a <p:pic> element with
// the image embedded in the zip, fitted to the original shape's EMU bounds.
// A small companion indicator shape (__LIVEDOC_IND_<name>) is added to the slide
// to show the 📷 badge; these are stripped from all slides during preview.

async function openDynamicImageDrawer() {
  editMode = 'dynimage';
  pendingImageDataUrl = null;
  openDrawer('Dynamic Image Settings');
  document.getElementById('drawerForm').innerHTML =
    '<p class="form-hint" style="padding:8px 0">Loading shapes from current slide…</p>';
  document.getElementById('drawerSaveBtn').textContent = 'Save Config';

  var shapes;
  try {
    shapes = await loadCurrentSlideAllShapes();
  } catch (e) {
    document.getElementById('drawerForm').innerHTML =
      '<p class="form-hint" style="color:#c00">Could not load shapes: ' + h(e.message) + '</p>';
    return;
  }
  buildDynamicImageDrawerForm(shapes);
}

async function loadCurrentSlideAllShapes() {
  var shapes = [];
  await PowerPoint.run(async function (context) {
    var slide;
    try {
      var sel = context.presentation.getSelectedSlides();
      sel.load('items/id');
      await context.sync();
      slide = sel.items[0];
      currentSlideId = slide ? slide.id : null;
    } catch (_) {
      slide = context.presentation.slides.getItemAt(0);
    }
    slide.shapes.load('items/name,items/type');
    await context.sync();
    slide.shapes.items.forEach(function (s) {
      if (!s.name.startsWith('__LIVEDOC_IND_')) {
        shapes.push({ name: s.name, type: s.type });
      }
    });
  });
  return shapes;
}

function buildDynamicImageDrawerForm(shapes) {
  var form = document.getElementById('drawerForm');
  if (!shapes.length) {
    form.innerHTML = '<p class="form-hint">No shapes found on the current slide.<br>Add a shape to use as a placeholder first.</p>';
    return;
  }

  var shapeOpts = shapes.map(function (s) {
    var sel = (selectedShapeName && s.name === selectedShapeName) ? ' selected' : '';
    return '<option value="' + h(s.name) + '"' + sel + '>' + h(s.name) + ' (' + h(s.type) + ')</option>';
  }).join('');

  var fileStatusMsg = pendingImageDataUrl
    ? '<span style="color:#1a6b3a">✓ File loaded — ready to save</span>'
    : '';

  form.innerHTML =
    '<div class="form-row"><label>Placeholder Shape</label>' +
    '<select id="df-img-shape" class="field">' + shapeOpts + '</select></div>' +
    '<div class="form-row"><label>Fit Mode</label>' +
    '<select id="df-img-fit" class="field">' +
    '<option value="fitInside">Fit Inside (keep aspect ratio)</option>' +
    '<option value="fitWidth">Fit Width (keep aspect ratio)</option>' +
    '<option value="fitHeight">Fit Height (keep aspect ratio)</option>' +
    '<option value="stretch">Stretch (fill bounds, ignore aspect ratio)</option>' +
    '</select></div>' +
    '<div class="form-row form-col"><label>Image URL</label>' +
    '<input id="df-img-url" class="field" type="url" placeholder="https://example.com/photo.jpg"/></div>' +
    '<div class="form-row form-col"><label>— or upload local file —</label>' +
    '<input id="df-img-file" type="file" accept="image/*" class="field" onchange="onImageFileSelected(this)"/>' +
    '<div id="df-img-file-status" class="form-hint" style="margin-top:2px">' + fileStatusMsg + '</div>' +
    '</div>' +
    '<p class="form-hint">At preview the placeholder is replaced by the image, fitted to its bounds. ' +
    'A 📷 indicator is added to the slide to show the binding.</p>';
}

async function refreshDynamicImageDrawer() {
  var form = document.getElementById('drawerForm');
  if (!form) return;
  form.innerHTML = '<p class="form-hint" style="padding:8px 0">Loading shapes from current slide…</p>';
  try {
    var shapes = await loadCurrentSlideAllShapes();
    buildDynamicImageDrawerForm(shapes);
  } catch (e) {
    form.innerHTML = '<p class="form-hint" style="color:#c00">Could not load shapes: ' + h(e.message) + '</p>';
  }
}

function syncDynamicImageDrawerDropdown(shapeName) {
  var sel = document.getElementById('df-img-shape');
  if (!sel || !shapeName) return;
  for (var i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === shapeName) { sel.selectedIndex = i; break; }
  }
}

function syncDynamicImageDrawerFromConfig(config) {
  var fitSel = document.getElementById('df-img-fit');
  var urlEl  = document.getElementById('df-img-url');
  if (!fitSel || !urlEl) return;
  if (config) {
    for (var i = 0; i < fitSel.options.length; i++) {
      if (fitSel.options[i].value === config.fitMode) { fitSel.selectedIndex = i; break; }
    }
    // Don't populate URL field with data: URLs (base64 is too large to show)
    urlEl.value = (config.sourceUrl && !config.sourceUrl.startsWith('data:')) ? config.sourceUrl : '';
  }
}

async function saveDynamicImageConfig() {
  var shapeName = (document.getElementById('df-img-shape') || {}).value || '';
  var fitMode   = (document.getElementById('df-img-fit')   || {}).value || 'fitInside';
  var url       = ((document.getElementById('df-img-url')  || {}).value || '').trim();
  var fileInput = document.getElementById('df-img-file');

  if (!shapeName) { showStatus('Select a placeholder shape.', 'error'); return; }
  var hasFile = fileInput && fileInput.files && fileInput.files.length > 0;
  if (!url && !hasFile && !pendingImageDataUrl) {
    showStatus('Provide an image URL or upload a local file.', 'error'); return;
  }

  var sourceUrl = url;
  if (!sourceUrl && hasFile) {
    try { sourceUrl = await readFileAsDataUrl(fileInput.files[0]); }
    catch (e) { showStatus('File read failed: ' + e.message, 'error'); return; }
  }
  if (!sourceUrl && pendingImageDataUrl) {
    sourceUrl = pendingImageDataUrl;
  }
  pendingImageDataUrl = null; // consumed

  var config = { fitMode: fitMode, sourceUrl: sourceUrl };

  try {
    await PowerPoint.run(async function (context) {
      var slide;
      try {
        var sel = context.presentation.getSelectedSlides();
        sel.load('items');
        await context.sync();
        slide = sel.items[0];
      } catch (_) {
        slide = context.presentation.slides.getItemAt(0);
      }

      slide.shapes.load('items/name,items/left,items/top,items/width,items/height');
      await context.sync();

      var shape = slide.shapes.items.find(function (s) { return s.name === shapeName; });
      if (!shape) { showStatus('Shape "' + shapeName + '" not found on current slide.', 'error'); return; }

      var left   = shape.left;
      var top    = shape.top;
      var width  = shape.width;
      var height = shape.height;

      shape.delete();
      var staleInd = slide.shapes.items.find(function (s) { return s.name === '__LIVEDOC_IND_' + shapeName; });
      if (staleInd) staleInd.delete();
      await context.sync();

      var label = '📷  Dynamic Image\n' + shapeName;
      var ph = slide.shapes.addTextBox(label, { left: left, top: top, width: width, height: height });
      await context.sync();

      ph.name = shapeName;
      ph.fill.setSolidColor('DCE8F8');
      ph.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
      ph.textFrame.textRange.font.size  = 14;
      ph.textFrame.textRange.font.color = '3A5A8A';
      ph.textFrame.textRange.font.bold  = false;
      try {
        ph.textFrame.textRange.paragraphFormat.horizontalAlignment =
          PowerPoint.ParagraphHorizontalAlignment.center;
        ph.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      } catch (_) {}

      ph.tags.add('LIVEDOC_DYN_IMAGE', JSON.stringify(config));
      await context.sync();

      closeDrawer();
      showStatus('"' + shapeName + '" placeholder ready (' + fitMode + '). Run Preview to replace with the real image.', 'success');
    });
  } catch (e) {
    showStatus('Failed to configure placeholder: ' + e.message, 'error');
  }
}

function readFileAsDataUrl(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload  = function (e) { resolve(e.target.result); };
    reader.onerror = function ()  { reject(new Error('File read failed')); };
    reader.readAsDataURL(file);
  });
}

// Called by the file input onchange — reads the file immediately into
// pendingImageDataUrl so the data survives if the drawer form is rebuilt.
function onImageFileSelected(input) {
  var statusEl = document.getElementById('df-img-file-status');
  if (!input.files || !input.files.length) {
    pendingImageDataUrl = null;
    if (statusEl) statusEl.innerHTML = '';
    return;
  }
  var file = input.files[0];
  if (statusEl) statusEl.innerHTML = '<span style="color:#888">Reading file…</span>';
  readFileAsDataUrl(file).then(function (dataUrl) {
    pendingImageDataUrl = dataUrl;
    if (statusEl) statusEl.innerHTML = '<span style="color:#1a6b3a">✓ Loaded: ' + h(file.name) + '</span>';
  }).catch(function () {
    pendingImageDataUrl = null;
    if (statusEl) statusEl.innerHTML = '<span style="color:#c00">File read failed.</span>';
  });
}

// ── Dynamic Image — scan & expand ─────────────────────────────────────────────

async function scanDynamicImages() {
  var result = {};

  await PowerPoint.run(async function (context) {
    var slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    slides.items.forEach(function (slide) {
      slide.shapes.load('items/name');
    });
    await context.sync();

    var shapesBySlide = [];
    slides.items.forEach(function (slide, si) {
      var nonInd = slide.shapes.items.filter(function (s) {
        return !s.name.startsWith('__LIVEDOC_IND_');
      });
      if (nonInd.length) shapesBySlide.push({ slideIndex: si + 1, shapes: nonInd });
    });

    if (!shapesBySlide.length) return;

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) { s.tags.load('items/key,items/value'); });
    });
    await context.sync();

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) {
        var imgTag = s.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_IMAGE'; });
        if (!imgTag) return;
        try {
          var config = JSON.parse(imgTag.value);
          if (!result[entry.slideIndex]) result[entry.slideIndex] = [];
          result[entry.slideIndex].push(Object.assign({ shapeName: s.name }, config));
        } catch (_) {}
      });
    });
  });

  return result;
}

// Expand dynamic images in a single slide XML string.
// Also strips __LIVEDOC_IND_* indicator shapes from all slides (always called).
// Modifies zip in-place (adds ppt/media/* and updates rels).
// Returns the modified slide XML string.
async function expandDynamicImagesInXml(zip, xml, slideIndex, configs) {
  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

  var parser   = new DOMParser();
  var slideDoc = parser.parseFromString(xml, 'application/xml');
  if (slideDoc.getElementsByTagName('parseerror').length) return xml;

  // Always strip indicator shapes
  var spEls = Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'sp'));
  spEls.forEach(function (sp) {
    var nvSpPr = sp.getElementsByTagNameNS(NS_P, 'nvSpPr')[0];
    if (!nvSpPr) return;
    var cNvPr = findChildByLocalName(nvSpPr, 'cNvPr');
    if (cNvPr && (cNvPr.getAttribute('name') || '').startsWith('__LIVEDOC_IND_')) {
      sp.parentNode.removeChild(sp);
    }
  });

  if (!configs || !configs.length) {
    var ser = new XMLSerializer();
    return ser.serializeToString(slideDoc).replace(/ xmlns=""/g, '');
  }

  // Load / create the slide rels XML
  var relsPath = 'ppt/slides/_rels/slide' + slideIndex + '.xml.rels';
  var relsXml  = zip.files[relsPath]
    ? await zip.files[relsPath].async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' + NS_REL + '"></Relationships>';

  var relsDoc = parser.parseFromString(relsXml, 'application/xml');

  var mediaSeq = 0;
  var changed  = false;

  for (var ci = 0; ci < configs.length; ci++) {
    var cfg = configs[ci];

    var bounds = findShapeBoundsInXml(slideDoc, cfg.shapeName, NS_P, NS_A);
    if (!bounds) {
      console.warn('[DynImage] Shape "' + cfg.shapeName + '" not found in slide XML.');
      continue;
    }

    var imgData;
    try { imgData = await fetchImageData(cfg.sourceUrl); }
    catch (e) {
      console.warn('[DynImage] Could not fetch image for "' + cfg.shapeName + '":', e.message);
      continue;
    }

    var dims = { w: null, h: null };
    if (cfg.fitMode !== 'stretch') {
      try { dims = await getImageDimensions(imgData.bytes, imgData.mimeType); } catch (_) {}
    }

    var ext      = (imgData.mimeType.split('/')[1] || 'png').replace('jpeg', 'jpg');
    var mediaName = 'livedoc_dyn_s' + slideIndex + '_' + (++mediaSeq) + '.' + ext;
    zip.file('ppt/media/' + mediaName, imgData.bytes);

    await ensureContentType(zip, ext, imgData.mimeType);

    var rId = 'rIdLDDyn' + slideIndex + '_' + mediaSeq;
    var relEl = relsDoc.createElementNS(NS_REL, 'Relationship');
    relEl.setAttribute('Id',     rId);
    relEl.setAttribute('Type',   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    relEl.setAttribute('Target', '../media/' + mediaName);
    relsDoc.documentElement.appendChild(relEl);

    var fr = computeFillRect(dims.w, dims.h, bounds.cx, bounds.cy, cfg.fitMode);

    var picXmlStr =
      '<p:pic' +
      ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:nvPicPr>' +
      '<p:cNvPr id="' + (parseInt(bounds.shapeId, 10) + 5000) + '" name="DynImg_' + h(cfg.shapeName) + '"/>' +
      '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>' +
      '<p:nvPr/>' +
      '</p:nvPicPr>' +
      '<p:blipFill>' +
      '<a:blip r:embed="' + rId + '"/>' +
      '<a:stretch><a:fillRect' +
        (fr.l ? ' l="' + fr.l + '"' : '') +
        (fr.t ? ' t="' + fr.t + '"' : '') +
        (fr.r ? ' r="' + fr.r + '"' : '') +
        (fr.b ? ' b="' + fr.b + '"' : '') +
      '/></a:stretch>' +
      '</p:blipFill>' +
      '<p:spPr>' +
      '<a:xfrm><a:off x="' + bounds.x + '" y="' + bounds.y + '"/>' +
      '<a:ext cx="' + bounds.cx + '" cy="' + bounds.cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '</p:spPr>' +
      '</p:pic>';

    var wrapDoc  = parser.parseFromString('<root>' + picXmlStr + '</root>', 'application/xml');
    var picEl    = wrapDoc.documentElement.firstElementChild;
    var imported = slideDoc.importNode(picEl, true);

    bounds.spEl.parentNode.insertBefore(imported, bounds.spEl);
    bounds.spEl.parentNode.removeChild(bounds.spEl);
    changed = true;
  }

  if (changed) {
    var serializer = new XMLSerializer();
    zip.file(relsPath, serializer.serializeToString(relsDoc).replace(/ xmlns=""/g, ''));
  }

  var ser2 = new XMLSerializer();
  return ser2.serializeToString(slideDoc).replace(/ xmlns=""/g, '');
}

// ── Dynamic Image helpers ──────────────────────────────────────────────────────

function findShapeBoundsInXml(slideDoc, shapeName, NS_P, NS_A) {
  var candidates = Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'sp'))
    .concat(Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'pic')));

  for (var i = 0; i < candidates.length; i++) {
    var el    = candidates[i];
    var nvPr  = el.getElementsByTagNameNS(NS_P, 'nvSpPr')[0] ||
                el.getElementsByTagNameNS(NS_P, 'nvPicPr')[0];
    if (!nvPr) continue;
    var cNvPr = findChildByLocalName(nvPr, 'cNvPr');
    if (!cNvPr || cNvPr.getAttribute('name') !== shapeName) continue;

    var spPr = el.getElementsByTagNameNS(NS_P, 'spPr')[0];
    if (!spPr) continue;
    var xfrm = spPr.getElementsByTagNameNS(NS_A, 'xfrm')[0];
    if (!xfrm) continue;
    var off  = xfrm.getElementsByTagNameNS(NS_A, 'off')[0];
    var ext  = xfrm.getElementsByTagNameNS(NS_A, 'ext')[0];
    if (!off || !ext) continue;

    return {
      spEl:    el,
      shapeId: cNvPr.getAttribute('id') || '99',
      x:   parseInt(off.getAttribute('x')  || '0', 10),
      y:   parseInt(off.getAttribute('y')  || '0', 10),
      cx:  parseInt(ext.getAttribute('cx') || '0', 10),
      cy:  parseInt(ext.getAttribute('cy') || '0', 10),
    };
  }
  return null;
}

// Returns fillRect offset values (1000ths-of-percent, 100%=100000) for the
// four edges, so the image is letterboxed/pillarboxed inside the shape bounds.
function computeFillRect(iw, ih, cx, cy, fitMode) {
  var l = 0, t = 0, r = 0, b = 0;
  if (fitMode === 'stretch' || !iw || !ih || !cx || !cy) return { l: l, t: t, r: r, b: b };

  var shapeR = cx / cy;
  var imageR = iw / ih;

  if (fitMode === 'fitInside') {
    if (imageR > shapeR) {
      var scaledH = ih * cx / iw;
      var emptyV  = cy - scaledH;
      t = b = Math.max(0, Math.round(emptyV / 2 / cy * 100000));
    } else {
      var scaledW = iw * cy / ih;
      var emptyH  = cx - scaledW;
      l = r = Math.max(0, Math.round(emptyH / 2 / cx * 100000));
    }
  } else if (fitMode === 'fitWidth') {
    var scaledH2 = ih * cx / iw;
    if (scaledH2 < cy) {
      var emptyV2 = cy - scaledH2;
      t = b = Math.max(0, Math.round(emptyV2 / 2 / cy * 100000));
    }
  } else if (fitMode === 'fitHeight') {
    var scaledW2 = iw * cy / ih;
    if (scaledW2 < cx) {
      var emptyH2 = cx - scaledW2;
      l = r = Math.max(0, Math.round(emptyH2 / 2 / cx * 100000));
    }
  }

  return { l: l, t: t, r: r, b: b };
}

// Fetch image from URL or decode a data: URI. Returns { bytes: Uint8Array, mimeType }.
async function fetchImageData(sourceUrl) {
  if (sourceUrl.startsWith('data:')) {
    var m = sourceUrl.match(/^data:([^;]+);base64,(.+)$/);
    if (!m) throw new Error('Invalid data URL');
    var bin = atob(m[2]);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return { bytes: arr, mimeType: m[1] };
  }
  var resp = await fetch(sourceUrl);
  if (!resp.ok) throw new Error('HTTP ' + resp.status + ' fetching image');
  var mime = (resp.headers.get('content-type') || 'image/png').split(';')[0].trim();
  var ab   = await resp.arrayBuffer();
  return { bytes: new Uint8Array(ab), mimeType: mime };
}

// Register the image extension in [Content_Types].xml if not already present.
async function ensureContentType(zip, ext, mimeType) {
  var ctPath = '[Content_Types].xml';
  if (!zip.files[ctPath]) return;
  var ct = await zip.files[ctPath].async('string');
  if (ct.includes('Extension="' + ext + '"')) return;

  var mime = mimeType || ('image/' + ext);
  if (ext === 'jpg') mime = 'image/jpeg';

  var entry = '<Default Extension="' + ext + '" ContentType="' + mime + '"/>';
  ct = ct.replace('</Types>', entry + '</Types>');
  zip.file(ctPath, ct);
}

// Load image bytes into an <img> element to get natural dimensions.
function getImageDimensions(bytes, mimeType) {
  return new Promise(function (resolve) {
    var blob = new Blob([bytes], { type: mimeType });
    var url  = URL.createObjectURL(blob);
    var img  = new Image();
    img.onload  = function () { URL.revokeObjectURL(url); resolve({ w: img.naturalWidth,  h: img.naturalHeight }); };
    img.onerror = function () { URL.revokeObjectURL(url); resolve({ w: null, h: null }); };
    img.src = url;
  });
}
