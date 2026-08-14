'use strict';

// ── Office init ────────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('app').style.display     = 'flex';
    initSearch();
    initAddMenu();
    loadFromCustomXml();
    initSelectionHandler();
  } else {
    document.getElementById('loading').innerHTML =
      '<p style="color:#c00;padding:20px">This add-in requires PowerPoint.</p>';
  }
});

// ── Search ─────────────────────────────────────────────────────────────────────
function initSearch() {
  document.getElementById('searchInput').addEventListener('input', function () {
    searchQuery = this.value.trim().toLowerCase();
    renderTree();
  });
}

// ── Add-menu dropdown ──────────────────────────────────────────────────────────
function initAddMenu() {
  const btn  = document.getElementById('addBtn');
  const menu = document.getElementById('addMenu');
  btn.addEventListener('click', function (e) {
    e.stopPropagation();
    menu.classList.toggle('hidden');
  });
  document.addEventListener('click', function () {
    menu.classList.add('hidden');
  });
}
