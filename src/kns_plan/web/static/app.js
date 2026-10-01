// Editable grids: add and remove rows, paste blocks of cells from Excel, warn about unsaved changes.
(function () {
  "use strict";

  function addRow(gridName) {
    var table = document.querySelector('table[data-grid="' + gridName + '"]');
    var tpl = document.querySelector('template[data-row-for="' + gridName + '"]');
    if (!table || !tpl) return null;
    var i = parseInt(table.dataset.next || "0", 10);
    table.dataset.next = String(i + 1);
    var html = tpl.innerHTML.split("__i__").join(String(i));
    var tbody = table.tBodies[0];
    tbody.insertAdjacentHTML("beforeend", html);
    return tbody.rows[tbody.rows.length - 1];
  }

  function cellsOf(row) {
    return Array.prototype.filter.call(row.querySelectorAll("input:not([type=hidden]):not([type=checkbox]), select"),
      function (el) { return !el.closest(".del"); });
  }

  document.addEventListener("click", function (e) {
    var add = e.target.closest("[data-add]");
    if (add) {
      var row = addRow(add.dataset.add);
      if (row && cellsOf(row)[0]) cellsOf(row)[0].focus();
      markDirty(add);
      return;
    }
    var remove = e.target.closest("[data-remove]");
    if (remove) {
      markDirty(remove);
      remove.closest("tr").remove();
    }
  });

  // Paste a block copied from Excel (tab-separated, one line per row) starting at the focused cell.
  document.addEventListener("paste", function (e) {
    var target = e.target;
    if (!target.closest || !target.closest("table.grid")) return;
    var text = (e.clipboardData || window.clipboardData).getData("text");
    if (!text || (text.indexOf("\t") < 0 && text.indexOf("\n") < 0)) return;
    e.preventDefault();
    var lines = text.replace(/\r/g, "").replace(/\n$/, "").split("\n");
    var table = target.closest("table.grid");
    var row = target.closest("tr");
    var col = cellsOf(row).indexOf(target);
    lines.forEach(function (line, k) {
      if (k > 0) {
        row = row.nextElementSibling || addRow(table.dataset.grid);
        if (!row) return;
      }
      var cells = cellsOf(row);
      line.split("\t").forEach(function (value, j) {
        var el = cells[col + j];
        if (!el) return;
        if (el.tagName === "SELECT") {
          var v = value.trim().toLowerCase();
          for (var o = 0; o < el.options.length; o++) {
            if (el.options[o].value.toLowerCase() === v || el.options[o].text.toLowerCase() === v) { el.selectedIndex = o; break; }
          }
        } else {
          el.value = value.trim();
        }
      });
    });
    markDirty(table);
  });

  // Unsaved-changes warning.
  var dirty = false;
  function markDirty(el) { if (el.closest && el.closest("form[data-dirty]")) dirty = true; }
  document.addEventListener("input", function (e) { markDirty(e.target); });
  document.addEventListener("change", function (e) { markDirty(e.target); });
  document.addEventListener("submit", function () { dirty = false; });
  window.addEventListener("beforeunload", function (e) {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });
})();
