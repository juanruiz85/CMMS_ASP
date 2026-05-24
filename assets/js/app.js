/* =============================================================================
   CMMS - JavaScript Global (app.js)
   Lógica de UI: sidebar, modales, toasts, AJAX, charts helpers
   ============================================================================= */

'use strict';

// ─── Estado Global ────────────────────────────────────────────────────────────
const CMMS = {
  version: '1.0.0',
  charts: {},
  toastQueue: [],
  currentModal: null,
};

// ─── DOMContentLoaded ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function () {
  CMMS.Sidebar.init();
  CMMS.Dropdown.init();
  CMMS.Table.initSort();
  CMMS.Form.initValidation();
  CMMS.Notifications.initDropdown();
  CMMS.initSearch();
  CMMS.initConfirmDelete();
  CMMS.initFlashMessages();
  CMMS.Topbar.setActive();
});

// ─── Sidebar ──────────────────────────────────────────────────────────────────
CMMS.Sidebar = {
  el: null,
  overlay: null,

  init() {
    this.el = document.querySelector('.sidebar');
    if (!this.el) return;

    // Botón de toggle móvil
    const toggleBtn = document.querySelector('.topbar-toggle');
    if (toggleBtn) {
      toggleBtn.addEventListener('click', () => this.toggle());
    }

    // Overlay para cerrar en móvil
    this.overlay = document.createElement('div');
    this.overlay.className = 'sidebar-overlay';
    this.overlay.style.cssText = `
      display:none; position:fixed; inset:0; background:rgba(0,0,0,0.5);
      z-index:99; backdrop-filter:blur(2px);
    `;
    this.overlay.addEventListener('click', () => this.close());
    document.body.appendChild(this.overlay);

    // Guardar estado del sidebar (colapsado/expandido)
    const isCollapsed = localStorage.getItem('sidebar_collapsed') === 'true';
    if (isCollapsed && window.innerWidth > 768) {
      this.el.classList.add('collapsed');
    }

    // Submenus
    document.querySelectorAll('.nav-item[data-submenu]').forEach(item => {
      item.addEventListener('click', (e) => {
        const target = document.getElementById(item.dataset.submenu);
        if (target) {
          target.classList.toggle('open');
          e.stopPropagation();
        }
      });
    });
  },

  toggle() {
    if (window.innerWidth <= 768) {
      this.el.classList.toggle('open');
      this.overlay.style.display = this.el.classList.contains('open') ? 'block' : 'none';
    } else {
      this.el.classList.toggle('collapsed');
      localStorage.setItem('sidebar_collapsed', this.el.classList.contains('collapsed'));
    }
  },

  close() {
    this.el.classList.remove('open');
    if (this.overlay) this.overlay.style.display = 'none';
  }
};

// ─── Topbar ───────────────────────────────────────────────────────────────────
CMMS.Topbar = {
  setActive() {
    const currentPath = window.location.pathname.toLowerCase();
    document.querySelectorAll('.nav-item[data-path]').forEach(item => {
      const path = item.dataset.path.toLowerCase();
      if (currentPath.includes(path)) {
        item.classList.add('active');
        // Expandir submenu padre si existe
        const submenu = item.closest('.nav-submenu');
        if (submenu) submenu.classList.add('open');
      }
    });
  }
};

// ─── Toast Notifications ──────────────────────────────────────────────────────
CMMS.Toast = {
  container: null,

  init() {
    if (!this.container) {
      this.container = document.getElementById('toast-container');
      if (!this.container) {
        this.container = document.createElement('div');
        this.container.id = 'toast-container';
        document.body.appendChild(this.container);
      }
    }
  },

  show(message, type = 'info', duration = 4000) {
    this.init();
    const icons = {
      success: '✓',
      error:   '✕',
      warning: '⚠',
      info:    'ℹ'
    };

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      <span style="font-size:16px;flex-shrink:0">${icons[type] || icons.info}</span>
      <span style="flex:1">${CMMS.escapeHtml(message)}</span>
      <button onclick="this.parentElement.remove()" style="background:none;border:none;cursor:pointer;color:currentColor;font-size:16px;padding:0;line-height:1;opacity:0.7">×</button>
    `;

    this.container.appendChild(toast);

    // Auto-remover
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.add('removing');
        setTimeout(() => toast.remove(), 300);
      }
    }, duration);

    return toast;
  },

  success: (msg, dur) => CMMS.Toast.show(msg, 'success', dur),
  error:   (msg, dur) => CMMS.Toast.show(msg, 'error',   dur),
  warning: (msg, dur) => CMMS.Toast.show(msg, 'warning', dur),
  info:    (msg, dur) => CMMS.Toast.show(msg, 'info',    dur),
};

// ─── Modal ────────────────────────────────────────────────────────────────────
CMMS.Modal = {
  open(id) {
    const overlay = document.getElementById(id);
    if (!overlay) return;
    overlay.classList.add('open');
    CMMS.currentModal = id;
    document.body.style.overflow = 'hidden';

    // Cerrar con Escape
    const keyHandler = (e) => {
      if (e.key === 'Escape') {
        this.close(id);
        document.removeEventListener('keydown', keyHandler);
      }
    };
    document.addEventListener('keydown', keyHandler);

    // Cerrar al clicar fuera
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) this.close(id);
    });
  },

  close(id) {
    const overlay = document.getElementById(id || CMMS.currentModal);
    if (!overlay) return;
    overlay.classList.remove('open');
    CMMS.currentModal = null;
    document.body.style.overflow = '';
  }
};

// Funciones globales para usar desde HTML
function openModal(id)  { CMMS.Modal.open(id);  }
function closeModal(id) { CMMS.Modal.close(id); }

// ─── Dropdown ─────────────────────────────────────────────────────────────────
CMMS.Dropdown = {
  init() {
    document.addEventListener('click', (e) => {
      const trigger = e.target.closest('[data-dropdown]');
      if (trigger) {
        e.stopPropagation();
        const menuId = trigger.dataset.dropdown;
        const menu   = document.getElementById(menuId);
        if (menu) {
          // Cerrar otros dropdowns
          document.querySelectorAll('.dropdown-menu.open').forEach(m => {
            if (m.id !== menuId) m.classList.remove('open');
          });
          menu.classList.toggle('open');
        }
      } else {
        // Cerrar todos
        document.querySelectorAll('.dropdown-menu.open').forEach(m => m.classList.remove('open'));
      }
    });
  }
};

// ─── Tabs ─────────────────────────────────────────────────────────────────────
function initTabs(groupName) {
  document.querySelectorAll(`[data-tab-group="${groupName}"]`).forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;

      // Botones
      document.querySelectorAll(`[data-tab-group="${groupName}"]`)
        .forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Paneles
      document.querySelectorAll(`[data-panel-group="${groupName}"]`)
        .forEach(p => p.classList.remove('active'));
      const panel = document.querySelector(`[data-panel-group="${groupName}"][data-panel="${target}"]`);
      if (panel) panel.classList.add('active');
    });
  });
}

// ─── Table Sort ───────────────────────────────────────────────────────────────
CMMS.Table = {
  initSort() {
    document.querySelectorAll('.table .sortable').forEach(th => {
      th.addEventListener('click', () => {
        const table   = th.closest('.table');
        const colIdx  = Array.from(th.parentElement.children).indexOf(th);
        const tbody   = table.querySelector('tbody');
        const rows    = Array.from(tbody.querySelectorAll('tr'));
        const isAsc   = th.classList.contains('sorted-asc');

        // Reset headers
        table.querySelectorAll('th').forEach(h => h.classList.remove('sorted-asc', 'sorted-desc'));

        // Ordenar
        rows.sort((a, b) => {
          const aVal = a.cells[colIdx]?.textContent.trim() || '';
          const bVal = b.cells[colIdx]?.textContent.trim() || '';
          const numA = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
          const numB = parseFloat(bVal.replace(/[^0-9.-]/g, ''));
          if (!isNaN(numA) && !isNaN(numB)) {
            return isAsc ? numA - numB : numB - numA;
          }
          return isAsc ? aVal.localeCompare(bVal, 'es') : bVal.localeCompare(aVal, 'es');
        });

        rows.forEach(r => tbody.appendChild(r));
        th.classList.add(isAsc ? 'sorted-desc' : 'sorted-asc');
      });
    });
  }
};

// ─── Form Validation ──────────────────────────────────────────────────────────
CMMS.Form = {
  initValidation() {
    document.querySelectorAll('form[data-validate]').forEach(form => {
      form.addEventListener('submit', (e) => {
        if (!this.validate(form)) {
          e.preventDefault();
        }
      });
    });
  },

  validate(form) {
    let valid = true;
    form.querySelectorAll('[required]').forEach(field => {
      this.clearError(field);
      if (!field.value.trim()) {
        this.showError(field, 'Este campo es requerido');
        valid = false;
      }
    });
    form.querySelectorAll('[data-type="email"]').forEach(field => {
      if (field.value && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(field.value)) {
        this.showError(field, 'Correo electrónico inválido');
        valid = false;
      }
    });
    return valid;
  },

  showError(field, message) {
    field.classList.add('is-invalid');
    const err = document.createElement('div');
    err.className = 'form-error cmms-field-err';
    err.textContent = message;
    field.parentElement.appendChild(err);
  },

  clearError(field) {
    field.classList.remove('is-invalid');
    field.parentElement.querySelectorAll('.cmms-field-err').forEach(e => e.remove());
  }
};

// ─── Search / Filter ──────────────────────────────────────────────────────────
CMMS.initSearch = function () {
  const searchInput = document.getElementById('tableSearch');
  if (!searchInput) return;

  searchInput.addEventListener('input', debounce(function () {
    const term  = this.value.toLowerCase().trim();
    const table = document.querySelector('.table tbody');
    if (!table) return;
    table.querySelectorAll('tr').forEach(row => {
      const text = row.textContent.toLowerCase();
      row.style.display = text.includes(term) ? '' : 'none';
    });
  }, 200));
};

// ─── Confirm Delete ───────────────────────────────────────────────────────────
CMMS.initConfirmDelete = function () {
  document.querySelectorAll('[data-confirm]').forEach(el => {
    el.addEventListener('click', function (e) {
      const msg = this.dataset.confirm || '¿Está seguro que desea eliminar este registro?';
      if (!confirm(msg)) {
        e.preventDefault();
        e.stopPropagation();
      }
    });
  });
};

// ─── Flash Messages auto-dismiss ─────────────────────────────────────────────
CMMS.initFlashMessages = function () {
  document.querySelectorAll('.alert').forEach(alert => {
    setTimeout(() => {
      if (alert.parentElement) {
        alert.style.opacity = '0';
        alert.style.transition = 'opacity 0.5s ease';
        setTimeout(() => alert.remove(), 500);
      }
    }, 5000);
  });
};

// ─── AJAX Helper ─────────────────────────────────────────────────────────────
CMMS.ajax = function (url, options = {}) {
  const defaults = {
    method:  'POST',
    headers: { 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest' },
  };
  const config = Object.assign({}, defaults, options);
  if (config.data && typeof config.data === 'object' && !(config.data instanceof FormData)) {
    config.body = JSON.stringify(config.data);
  } else if (config.data instanceof FormData) {
    config.body = config.data;
    delete config.headers['Content-Type'];
  }

  return fetch(url, config)
    .then(res => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.json();
    })
    .catch(err => {
      CMMS.Toast.error('Error de comunicación con el servidor');
      throw err;
    });
};

// Enviar formulario vía AJAX
CMMS.submitForm = function (formId, successCallback) {
  const form = document.getElementById(formId);
  if (!form) return;

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    const submitBtn = form.querySelector('[type="submit"]');
    const originalText = submitBtn ? submitBtn.textContent : '';

    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="spinner" style="width:16px;height:16px;border-width:2px"></span>';
    }

    const formData = new FormData(form);

    CMMS.ajax(form.action || window.location.href, {
      method: 'POST',
      data:   formData
    }).then(res => {
      if (res.success) {
        CMMS.Toast.success(res.message || 'Guardado correctamente');
        if (successCallback) successCallback(res);
      } else {
        CMMS.Toast.error(res.message || 'Error al procesar la solicitud');
      }
    }).finally(() => {
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = originalText;
      }
    });
  });
};

// ─── Notifications ────────────────────────────────────────────────────────────
CMMS.Notifications = {
  initDropdown() {
    const bell = document.getElementById('notifBell');
    if (!bell) return;

    bell.addEventListener('click', () => {
      const dropdown = document.getElementById('notifDropdown');
      if (dropdown) dropdown.classList.toggle('open');
    });
  },

  markAllRead() {
    CMMS.ajax('/CMMS/api/index.asp?action=notif_read_all', { method: 'POST' })
      .then(() => {
        document.querySelectorAll('.notif-item.unread').forEach(i => i.classList.remove('unread'));
        const dot = document.querySelector('.topbar-btn .badge-dot');
        if (dot) dot.style.display = 'none';
      });
  }
};

// ─── Chart.js Helpers ─────────────────────────────────────────────────────────
CMMS.Charts = {
  defaultColors: [
    '#6366f1', '#22d3ee', '#10b981', '#f59e0b',
    '#ef4444', '#8b5cf6', '#06b6d4', '#84cc16'
  ],

  defaults: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: '#94a3b8',
          font: { family: 'Inter', size: 12 },
          padding: 16,
        }
      },
      tooltip: {
        backgroundColor: '#0f1e38',
        borderColor: 'rgba(255,255,255,0.1)',
        borderWidth: 1,
        titleColor: '#e2e8f0',
        bodyColor: '#94a3b8',
        padding: 12,
        cornerRadius: 8,
      }
    },
    scales: {
      x: {
        ticks:  { color: '#4b5e78', font: { family: 'Inter', size: 11 } },
        grid:   { color: 'rgba(255,255,255,0.04)', drawBorder: false },
        border: { display: false }
      },
      y: {
        ticks:    { color: '#4b5e78', font: { family: 'Inter', size: 11 } },
        grid:     { color: 'rgba(255,255,255,0.04)', drawBorder: false },
        border:   { display: false },
        beginAtZero: true
      }
    }
  },

  donut(canvasId, labels, data, options = {}) {
    const ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    if (CMMS.charts[canvasId]) CMMS.charts[canvasId].destroy();

    CMMS.charts[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: this.defaultColors,
          borderColor: 'transparent',
          borderWidth: 0,
          hoverOffset: 8,
        }]
      },
      options: Object.assign({
        plugins: {
          legend: { position: 'bottom', labels: { color: '#94a3b8', padding: 16, font: { family: 'Inter', size: 12 } } },
          tooltip: this.defaults.plugins.tooltip,
        },
        cutout: '70%',
        maintainAspectRatio: false,
        responsive: true,
      }, options)
    });

    return CMMS.charts[canvasId];
  },

  bar(canvasId, labels, datasets, options = {}) {
    const ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    if (CMMS.charts[canvasId]) CMMS.charts[canvasId].destroy();

    CMMS.charts[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: { labels, datasets },
      options: Object.assign({}, this.defaults, options)
    });

    return CMMS.charts[canvasId];
  },

  line(canvasId, labels, datasets, options = {}) {
    const ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    if (CMMS.charts[canvasId]) CMMS.charts[canvasId].destroy();

    CMMS.charts[canvasId] = new Chart(ctx, {
      type: 'line',
      data: { labels, datasets },
      options: Object.assign({}, this.defaults, options)
    });

    return CMMS.charts[canvasId];
  }
};

// ─── Utilities ────────────────────────────────────────────────────────────────
function debounce(fn, delay) {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

CMMS.escapeHtml = function (str) {
  if (typeof str !== 'string') return '';
  return str.replace(/[&<>"']/g, m => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[m]));
};

CMMS.formatDate = function (dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  return isNaN(d) ? dateStr : d.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
};

CMMS.formatMoney = function (amount) {
  if (amount === null || amount === undefined || amount === '') return '$0.00';
  return '$' + parseFloat(amount).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

// Copiar al portapapeles
CMMS.copyToClipboard = function (text) {
  navigator.clipboard.writeText(text)
    .then(() => CMMS.Toast.success('Copiado al portapapeles'))
    .catch(() => CMMS.Toast.error('No se pudo copiar'));
};

// Exportar tabla a CSV
CMMS.exportTableCSV = function (tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  let csv = [];
  table.querySelectorAll('tr').forEach(row => {
    const cols = [];
    row.querySelectorAll('th, td').forEach(col => {
      let text = col.textContent.trim().replace(/"/g, '""');
      cols.push('"' + text + '"');
    });
    csv.push(cols.join(','));
  });
  const blob    = new Blob(['\uFEFF' + csv.join('\n')], { type: 'text/csv;charset=utf-8' });
  const url     = URL.createObjectURL(blob);
  const a       = document.createElement('a');
  a.href        = url;
  a.download    = (filename || 'export') + '.csv';
  a.click();
  URL.revokeObjectURL(url);
};

// Confirmar acción con UI custom
CMMS.confirm = function (message, onConfirm, type = 'danger') {
  return window.confirm(message) ? onConfirm() : null;
};

// Actualizar URL sin reload
CMMS.updateURL = function (params) {
  const url = new URL(window.location.href);
  Object.entries(params).forEach(([k, v]) => {
    if (v === null || v === '') {
      url.searchParams.delete(k);
    } else {
      url.searchParams.set(k, v);
    }
  });
  history.pushState({}, '', url.toString());
};

// Animación de contadores para stats
CMMS.animateCounters = function () {
  document.querySelectorAll('[data-counter]').forEach(el => {
    const target = parseInt(el.dataset.counter, 10);
    const duration = 1200;
    const start    = Date.now();
    const update = () => {
      const elapsed = Date.now() - start;
      const progress = Math.min(elapsed / duration, 1);
      const eased    = 1 - Math.pow(1 - progress, 3);
      el.textContent = Math.floor(eased * target).toLocaleString('es-MX');
      if (progress < 1) requestAnimationFrame(update);
    };
    requestAnimationFrame(update);
  });
};

// Ejecutar al cargar si hay contadores
document.addEventListener('DOMContentLoaded', CMMS.animateCounters);
