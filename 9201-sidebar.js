/**
 * Shared sidebar component.
 *
 * Usage:
 *   <aside class="sidebar" id="sidebar" data-active="KEY"></aside>
 */
(function () {
  'use strict';

  var NAV_USER_ITEMS = [
    { key: 'profil', href: 'index.html', label: 'Profil Saya' },
    { key: 'pegawai-terbaik', href: 'pegawai-terbaik.html', label: 'Pegawai Terbaik' },
    { key: 'surat-tugas', href: 'surat-tugas.html', label: 'Minta Surat Tugas' },
  ];

  var NAV_ADMIN_ITEMS = [
    { key: 'admin-surat', href: 'admin-surat-tugas.html', label: 'Surat Tugas' },
    { key: 'kepegawaian', href: 'admin-kepegawaian.html', label: 'Kepegawaian' },
    { key: 'predikat-kinerja', href: 'admin-predikat-kinerja.html', label: 'Predikat Kinerja' },
    { key: 'eotq', href: 'admin-eotq.html', label: 'EoTQ' },
    { key: 'kamus-pok', href: 'manajemen-kamus-pok.html', label: 'Kamus POK' },
    { key: 'manajemen-mitra', href: 'manajemen-mitra.html', label: 'Manajemen Mitra' },
    { key: 'manajemen-pengguna', href: 'manajemen-pengguna.html', label: 'Manajemen Pengguna' },
  ];

  var SIDEBAR_CSS = `
  :where(.sidebar){
    position:relative;
    background:linear-gradient(180deg,#ffffff 0%,#fbfaf7 100%)!important;
    border-right:1px solid rgba(13,35,64,.09)!important;
    box-shadow:8px 0 30px rgba(13,35,64,.035);
    padding:16px 12px!important;
    gap:4px;
    transition:width .24s ease,min-width .24s ease,padding .24s ease,transform .24s ease,opacity .18s ease;
  }
  :where(.sidebar.collapsed){
    width:0!important;min-width:0!important;padding-left:0!important;padding-right:0!important;
    border-right:0!important;box-shadow:none!important;overflow:hidden!important;
  }
  :where(.sidebar.collapsed.open){
    width:var(--sidebar-w,260px)!important;min-width:var(--sidebar-w,260px)!important;
    padding:16px 12px!important;border-right:1px solid rgba(13,35,64,.09)!important;
    box-shadow:8px 0 30px rgba(13,35,64,.035)!important;overflow-y:auto!important;opacity:1;
  }
  .sidebar-toggle{
    position:fixed;top:62px;left:calc(var(--sidebar-w,260px) - 14px);
    width:28px;height:28px;border-radius:50%;
    background:var(--white,#fff);border:1.5px solid var(--border,#e2ddd6);
    display:flex;align-items:center;justify-content:center;
    cursor:pointer;z-index:260;color:var(--navy,#0d2340);
    transition:left .25s ease,background .15s,transform .15s;
    box-shadow:0 2px 8px rgba(13,35,64,.08);
    font-size:13px;font-weight:700;line-height:1;padding:0;
  }
  .sidebar-toggle:hover{background:var(--bg,#f5f4f0);transform:scale(1.05)}
  body.sidebar-hidden .sidebar-toggle,
  body.sidebar-collapsed .sidebar-toggle{left:6px}
  .toggle-icon{display:inline-block;transition:transform .25s ease}
  body.sidebar-hidden .sidebar-toggle .toggle-icon,
  body.sidebar-collapsed .sidebar-toggle .toggle-icon{transform:rotate(180deg)}
  :where(.sidebar-section-label),:where(.nav-group-label),:where(.nav-item-icon){display:none!important}
  :where(.nav-item){
    min-height:42px;padding:0 14px!important;border-left:0!important;border-radius:10px;
    color:#586174!important;font-size:13px!important;font-weight:700!important;letter-spacing:.01em;
    background:transparent!important;position:relative;gap:0!important;
  }
  :where(.nav-item)::before{
    content:'';width:7px;height:7px;border-radius:999px;background:#d8d3c8;margin-right:11px;flex:0 0 auto;
    box-shadow:0 0 0 4px rgba(216,211,200,.2);
  }
  :where(.nav-item:hover){background:#f4f1e8!important;color:var(--navy,#0d2340)!important}
  :where(.nav-item.active){
    background:var(--navy,#0d2340)!important;color:#fff!important;box-shadow:0 10px 22px rgba(13,35,64,.18);
  }
  :where(.nav-item.active)::before{background:var(--gold,#c8a84b);box-shadow:0 0 0 4px rgba(200,168,75,.24)}
  .nav-admin-shell{margin-top:6px;padding-top:6px;border-top:1px solid rgba(13,35,64,.08)}
  .nav-collapse-toggle{
    width:100%;min-height:42px;border:0;border-radius:10px;background:#f7f5ef;color:var(--navy,#0d2340);
    font:700 13px/1.3 'Plus Jakarta Sans',sans-serif;display:flex;align-items:center;justify-content:space-between;
    padding:0 12px 0 14px;cursor:pointer;transition:background .18s,box-shadow .18s,color .18s;
  }
  .nav-collapse-toggle:hover{background:#f0ecdf}
  .nav-collapse-toggle.open{background:#fff7df;box-shadow:inset 0 0 0 1px rgba(200,168,75,.35)}
  .nav-collapse-toggle::before{
    content:'';width:7px;height:7px;border-radius:999px;background:var(--gold,#c8a84b);margin-right:11px;
    box-shadow:0 0 0 4px rgba(200,168,75,.18);
  }
  .nav-collapse-label{flex:1;text-align:left}
  .nav-collapse-caret{
    width:28px;height:28px;border:1px solid rgba(13,35,64,.14);border-radius:999px;background:#fff;
    color:var(--navy,#0d2340);display:inline-flex;align-items:center;justify-content:center;
    transition:transform .18s ease,background .18s ease,border-color .18s ease;flex:0 0 auto;
  }
  .nav-collapse-caret::before{
    content:'';width:8px;height:8px;border-right:2px solid currentColor;border-bottom:2px solid currentColor;
    transform:rotate(45deg) translateY(-2px);display:block;
  }
  .nav-collapse-toggle.open .nav-collapse-caret{transform:rotate(180deg);border-color:rgba(200,168,75,.55);background:#fffdf5}
  .nav-submenu{display:grid;grid-template-rows:0fr;transition:grid-template-rows .22s ease}
  .nav-submenu.open{grid-template-rows:1fr}
  .nav-submenu-inner{overflow:hidden;padding-top:4px}
  .nav-submenu .nav-item{min-height:38px;margin-left:10px;padding-left:12px!important;font-size:12.5px!important}
  :where(.sidebar-footer){
    border-top:1px solid rgba(13,35,64,.08)!important;margin:14px 2px 0!important;padding:14px 8px 0!important;
    font-size:10.5px!important;line-height:1.65!important;color:#8790a1!important;
  }
  :where(.sidebar-footer strong){display:none!important}
  @media(max-width:600px){
    .sidebar-toggle{left:14px;top:62px}
    :where(.sidebar){padding:14px 12px!important}
    :where(.sidebar.collapsed){width:var(--sidebar-w,260px)!important;opacity:1;padding:14px 12px!important}
  }
  `;

  function escAttr(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function injectCSS() {
    if (document.getElementById('9201-sidebar-clean-css')) return;
    var style = document.createElement('style');
    style.id = '9201-sidebar-clean-css';
    style.textContent = SIDEBAR_CSS;
    document.head.appendChild(style);
  }

  function buildItem(item, activeKey) {
    var isActive = item.key === activeKey;
    return '<a class="nav-item' + (isActive ? ' active' : '') + '" href="' + escAttr(item.href) + '">' +
      escAttr(item.label) +
      '</a>';
  }

  function buildSidebar(activeKey) {
    var html = '';

    NAV_USER_ITEMS.forEach(function (item) {
      html += buildItem(item, activeKey);
    });

    var adminOpen = NAV_ADMIN_ITEMS.some(function (item) { return item.key === activeKey; });
    html += '<div id="admin-nav" class="nav-admin-shell" style="display:none">';
    html += '<button class="nav-collapse-toggle' + (adminOpen ? ' open' : '') + '" type="button" onclick="Sidebar9201.toggleAdmin()">' +
      '<span class="nav-collapse-label">Administrasi</span><span class="nav-collapse-caret" aria-hidden="true"></span></button>';
    html += '<div class="nav-submenu' + (adminOpen ? ' open' : '') + '" id="admin-submenu"><div class="nav-submenu-inner">';
    NAV_ADMIN_ITEMS.forEach(function (item) {
      html += buildItem(item, activeKey);
    });
    html += '</div></div></div>';

    html += '<div class="sidebar-footer">' +
      'BPS Kabupaten Raja Ampat' +
      '<div id="admin-footer-badge" style="display:none;margin-top:6px">' +
      '<span class="admin-badge" id="sidebar-admin-badge">Administrator</span>' +
      '</div>' +
      '</div>';

    return html;
  }

  function render() {
    injectCSS();
    var aside = document.getElementById('sidebar');
    if (!aside) return;
    var activeKey = aside.getAttribute('data-active') || '';
    aside.innerHTML = buildSidebar(activeKey);
    ensureToggleButton();
    try {
      var persisted = localStorage.getItem('sidebar_collapsed_v1') === '1' || localStorage.getItem('nova_sidebar_hidden') === '1';
      if (persisted && !window.matchMedia('(max-width:600px)').matches) {
        aside.classList.add('collapsed');
        document.body.classList.add('sidebar-hidden');
        document.body.classList.add('sidebar-collapsed');
      }
    } catch (_) {}
    updateToggle();
  }

  function ensureToggleButton() {
    var btn = document.getElementById('sidebar-toggle-shared');
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'sidebar-toggle-shared';
      btn.className = 'sidebar-toggle';
      btn.type = 'button';
      btn.onclick = toggleSidebar;
      document.body.appendChild(btn);
    }
    btn.innerHTML = '<span class="toggle-icon">&lsaquo;</span>';
  }

  function updateToggle() {
    var btn = document.querySelector('.sidebar-toggle');
    var aside = document.getElementById('sidebar');
    if (!btn || !aside) return;
    var hidden = (aside.classList.contains('collapsed') || document.body.classList.contains('sidebar-collapsed')) && !window.matchMedia('(max-width:600px)').matches;
    btn.innerHTML = '<span class="toggle-icon">&lsaquo;</span>';
    btn.setAttribute('aria-label', hidden ? 'Tampilkan menu' : 'Sembunyikan menu');
  }

  function toggleAdmin() {
    var btn = document.querySelector('.nav-collapse-toggle');
    var menu = document.getElementById('admin-submenu');
    if (!btn || !menu) return;
    var willOpen = !menu.classList.contains('open');
    menu.classList.toggle('open', willOpen);
    btn.classList.toggle('open', willOpen);
  }

  function toggleSidebar() {
    var aside = document.getElementById('sidebar');
    if (!aside) return;
    if (window.matchMedia('(max-width:600px)').matches) {
      aside.classList.toggle('open');
      return;
    }
    var hidden = !aside.classList.contains('collapsed');
    aside.classList.toggle('collapsed', hidden);
    document.body.classList.toggle('sidebar-hidden', hidden);
    document.body.classList.toggle('sidebar-collapsed', hidden);
    try {
      localStorage.setItem('sidebar_collapsed_v1', hidden ? '1' : '0');
      localStorage.setItem('nova_sidebar_hidden', hidden ? '1' : '0');
    } catch (_) {}
    updateToggle();
  }

  if (document.getElementById('sidebar')) {
    render();
  } else if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', render);
  }

  window.Sidebar9201 = { render: render, toggleAdmin: toggleAdmin, toggleSidebar: toggleSidebar };
})();
