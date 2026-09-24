/* DipSeva UI helpers — no dependencies */
(function () {
  'use strict';

  function inr(n) { return '₹' + Number(n || 0).toLocaleString('en-IN', { maximumFractionDigits: 0 }); }
  function num(s) { var v = parseFloat(String(s).replace(/[^0-9.\-]/g, '')); return isNaN(v) ? -Infinity : v; }

  /* 1. Sortable tables: add data-sort="num|text" to <th> */
  function makeSortable(table) {
    var head = table.tHead && table.tHead.rows[table.tHead.rows.length - 1];
    if (!head) return;
    var body = table.tBodies[0];
    if (!body) return;
    Array.prototype.forEach.call(head.cells, function (th, i) {
      if (!th.dataset.sort) return;
      th.classList.add('sortable');
      th.addEventListener('click', function () {
        var dir = th.dataset.dir === 'asc' ? 'desc' : 'asc';
        Array.prototype.forEach.call(head.cells, function (o) { o.dataset.dir = ''; o.classList.remove('asc', 'desc'); });
        th.dataset.dir = dir; th.classList.add(dir);
        var rows = Array.prototype.slice.call(body.rows);
        var type = th.dataset.sort;
        rows.sort(function (a, b) {
          var x = a.cells[i] ? a.cells[i].dataset.v || a.cells[i].textContent.trim() : '';
          var y = b.cells[i] ? b.cells[i].dataset.v || b.cells[i].textContent.trim() : '';
          var r = type === 'num' ? num(x) - num(y) : String(x).localeCompare(String(y));
          return dir === 'asc' ? r : -r;
        });
        rows.forEach(function (r) { body.appendChild(r); });
      });
    });
  }

  /* 2. Live filter boxes: <input data-filter="#tableId"> */
  function makeFilters() {
    document.querySelectorAll('[data-filter]').forEach(function (box) {
      var table = document.querySelector(box.dataset.filter);
      if (!table || !table.tBodies[0]) return;
      box.addEventListener('input', function () {
        var q = box.value.trim().toLowerCase(), shown = 0;
        Array.prototype.forEach.call(table.tBodies[0].rows, function (r) {
          var hit = r.textContent.toLowerCase().indexOf(q) > -1;
          r.hidden = !hit;
          if (hit) shown++;
        });
        var badge = document.querySelector('[data-count="' + box.dataset.filter + '"]');
        if (badge) badge.textContent = shown;
        var empty = document.querySelector('[data-empty="' + box.dataset.filter + '"]');
        if (empty) empty.hidden = shown !== 0;
      });
    });
  }

  /* 3. Whole-row links: <tr data-href="..."> — ignores clicks on real controls */
  function makeRowLinks() {
    document.addEventListener('click', function (e) {
      var tr = e.target.closest('tr[data-href]');
      if (!tr) return;
      if (e.target.closest('a,button,input,select,label')) return;
      if (window.getSelection().toString()) return;
      window.location.href = tr.dataset.href;
    });
    document.querySelectorAll('tr[data-href]').forEach(function (tr) { tr.classList.add('rowlink'); });
  }

  /* 4. Toasts auto-dismiss */
  function toasts() {
    document.querySelectorAll('.toast').forEach(function (t, i) {
      t.style.animationDelay = (i * 60) + 'ms';
      setTimeout(function () { t.classList.add('bye'); setTimeout(function () { t.remove(); }, 400); }, 5200);
      t.addEventListener('click', function (e) { if (e.target.tagName !== 'BUTTON') t.classList.remove('bye'); });
    });
  }

  /* 5. Keyboard: Alt+1..5 jump, / focus filter, Esc clear */
  function keys() {
    var map = { 1: '/', 2: '/entry', 3: '/customers', 4: '/stock', 5: '/calendar' };
    document.addEventListener('keydown', function (e) {
      if (e.altKey && map[e.key]) { e.preventDefault(); window.location.href = map[e.key]; return; }
      if (e.key === '/' && !/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName)) {
        var box = document.querySelector('[data-filter]');
        if (box) { e.preventDefault(); box.focus(); box.select(); }
      }
      if (e.key === 'Escape') {
        var box = document.querySelector('[data-filter]');
        if (box && box.value) { box.value = ''; box.dispatchEvent(new Event('input')); box.blur(); }
      }
    });
  }

  /* 6. Copy to clipboard with feedback */
  function copiers() {
    document.querySelectorAll('[data-copy]').forEach(function (b) {
      b.addEventListener('click', function () {
        var src = document.querySelector(b.dataset.copy);
        if (!src) return;
        navigator.clipboard.writeText(src.textContent.trim()).then(function () {
          var old = b.textContent; b.textContent = 'Copied ✓'; b.classList.add('done');
          setTimeout(function () { b.textContent = old; b.classList.remove('done'); }, 1600);
        });
      });
    });
  }

  /* 7. Select contents on focus so a typed amount replaces the old value (any device) */
  function polish() {
    document.querySelectorAll('input[type=number],input[inputmode=decimal],input[inputmode=numeric]').forEach(function (i) {
      i.addEventListener('focus', function () { i.select(); });
    });
    document.querySelectorAll('table').forEach(function (t) {
      if (!t.closest('.rtable')) {
        var w = document.createElement('div');
        w.className = 'rtable';
        t.parentNode.insertBefore(w, t);
        w.appendChild(t);
      }
      makeSortable(t);
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    polish(); makeFilters(); makeRowLinks(); toasts(); keys(); copiers();
  });
})();
