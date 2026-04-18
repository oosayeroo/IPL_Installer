/* global L */
(() => {
  const DEFAULT_PORTS = [
    { key: 'goole_dock', name: 'Goole Docks', center: [53.69984, -0.872207], zoom: 16 },
    { key: 'blacktoft', name: 'Blacktoft', center: [53.707031, -0.726981], zoom: 16 },
    { key: 'king_george_dock', name: 'King George Dock', center: [53.74229, -0.26751], zoom: 16 }
  ];

  const state = {
    map: null,
    ports: loadPorts(),
    portKey: 'goole_dock',
    elements: [],
    selectedId: null,
    mode: 'move',
    touchMode: false,
    history: [],
    historyIndex: -1,
    mapLayer: null,
    pickPortPoint: false,
    saveTimer: null,
    ignoreMapClicksUntil: 0,
    filters: {
      ship: true,
      warehouse: true,
      label: true,
      measurement: true,
      bollard: true
    }
  };

  const ui = {
    topBar: document.getElementById('topBar'),
    touchToggle: document.getElementById('touchToggle'),
    touchBar: document.getElementById('touchBar'),
    touchActionsBtn: document.getElementById('touchActionsBtn'),
    touchActionsMenu: document.getElementById('touchActionsMenu'),
    touchLockBtn: document.getElementById('touchLockBtn'),
    touchDeleteBtn: document.getElementById('touchDeleteBtn'),
    zoomInBtn: document.getElementById('zoomInBtn'),
    zoomOutBtn: document.getElementById('zoomOutBtn'),
    portsBody: document.getElementById('portsBody'),
    togglePortsBtn: document.getElementById('togglePortsBtn'),
    portSelect: document.getElementById('portSelect'),
    newPortBtn: document.getElementById('newPortBtn'),
    newPortForm: document.getElementById('newPortForm'),
    newPortName: document.getElementById('newPortName'),
    newPortCoords: document.getElementById('newPortCoords'),
    pickPortPointBtn: document.getElementById('pickPortPointBtn'),
    createPortBtn: document.getElementById('createPortBtn'),
    transformHint: document.getElementById('transformHint'),
    filterBtn: document.getElementById('filterBtn'),
    filterMenu: document.getElementById('filterMenu'),
    fltShip: document.getElementById('fltShip'),
    fltWarehouse: document.getElementById('fltWarehouse'),
    fltLabel: document.getElementById('fltLabel'),
    fltMeasurement: document.getElementById('fltMeasurement'),
    fltBollard: document.getElementById('fltBollard'),
    transformStatus: document.getElementById('transformStatus'),
    detailsPanel: document.getElementById('detailsPanel'),
    detailsDeleteBtn: document.getElementById('detailsDeleteBtn'),
    detailsLockBtn: document.getElementById('detailsLockBtn'),
    lblName: document.getElementById('lblName'),
    fldName: document.getElementById('fldName'),
    lblLength: document.getElementById('lblLength'),
    fldLength: document.getElementById('fldLength'),
    fldWidth: document.getElementById('fldWidth'),
    fldColor: document.getElementById('fldColor'),
    fldArrival: document.getElementById('fldArrival'),
    fldText: document.getElementById('fldText'),
    fldMapDistance: document.getElementById('fldMapDistance'),
    fldMeasure: document.getElementById('fldMeasure'),
    toggleVesselsBtn: document.getElementById('toggleVesselsBtn'),
    vesselsBody: document.getElementById('vesselsBody'),
    mapMenu: document.getElementById('mapMenu'),
    elementMenu: document.getElementById('elementMenu')
  };

  let hostMsgId = 1;
  const hostPending = new Map();

  window.__hostReceive = (msg) => {
    const slot = hostPending.get(msg.id);
    if (!slot) return;
    hostPending.delete(msg.id);
    if (msg.ok) slot.resolve(msg.payload);
    else slot.reject(new Error(msg.error || 'Unknown host error'));
  };

  function hostRequest(type, payload) {
    if (!(window.chrome && window.chrome.webview)) {
      return Promise.resolve(type === 'loadPortData' ? { elements: [], meta: {} } : {});
    }
    return new Promise((resolve, reject) => {
      const id = String(hostMsgId++);
      hostPending.set(id, { resolve, reject });
      window.chrome.webview.postMessage({ id, type, payload });
    });
  }

  function loadPorts() {
    try {
      const raw = localStorage.getItem('ipl_ports');
      if (!raw) return DEFAULT_PORTS.slice();
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed) || parsed.length === 0) return DEFAULT_PORTS.slice();
      return parsed;
    } catch {
      return DEFAULT_PORTS.slice();
    }
  }

  function savePorts() {
    localStorage.setItem('ipl_ports', JSON.stringify(state.ports));
  }

  function slug(s) {
    return s.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'port';
  }

  function today() {
    return new Date().toISOString().slice(0, 10);
  }

  function round2(n) {
    return Math.round(Number(n) * 100) / 100;
  }

  function fmt2(n) {
    const x = Number(n);
    return Number.isFinite(x) ? x.toFixed(2) : '';
  }

  function id() {
    return Math.random().toString(36).slice(2, 10);
  }

  function deepCopy(v) {
    return JSON.parse(JSON.stringify(v));
  }

  function selected() {
    return state.elements.find((x) => x.id === state.selectedId) || null;
  }

  function deleteSelectedElement() {
    const s = selected();
    if (!s) return;
    state.elements = state.elements.filter((x) => x.id !== s.id);
    state.selectedId = null;
    pushHistory();
    renderAll();
    persist();
  }

  function isTypeVisible(type) {
    return state.filters[type] !== false;
  }

  function pushHistory() {
    const snap = deepCopy(state.elements);
    state.history = state.history.slice(0, state.historyIndex + 1);
    state.history.push(snap);
    if (state.history.length > 100) state.history.shift();
    state.historyIndex = state.history.length - 1;
  }

  function applyHistory(newIndex) {
    if (newIndex < 0 || newIndex >= state.history.length) return;
    state.historyIndex = newIndex;
    state.elements = deepCopy(state.history[newIndex]);
    if (!state.elements.some((e) => e.id === state.selectedId)) state.selectedId = null;
    renderAll();
    persist();
  }

  function undo() { applyHistory(state.historyIndex - 1); }
  function redo() { applyHistory(state.historyIndex + 1); }

  function metersToLatLng(baseLat, baseLng, dx, dy) {
    const dLat = dy / 111320;
    const dLng = dx / (111320 * Math.cos((baseLat * Math.PI) / 180));
    return { lat: baseLat + dLat, lng: baseLng + dLng };
  }

  function localToLatLng(origin, x, y, deg) {
    const r = (deg * Math.PI) / 180;
    const rx = x * Math.cos(r) - y * Math.sin(r);
    const ry = x * Math.sin(r) + y * Math.cos(r);
    return metersToLatLng(origin.lat, origin.lng, rx, ry);
  }

  function distMeters(a, b) {
    return state.map.distance([a.lat, a.lng], [b.lat, b.lng]);
  }

  function addShip(latlng) {
    const item = {
      id: id(),
      type: 'ship',
      lat: latlng.lat,
      lng: latlng.lng,
      rotation: 0,
      length: 90,
      width: 15,
      color: '#4da6ff',
      name: 'New Vessel',
      arrival: today(),
      locked: false
    };
    state.elements.push(item);
    state.selectedId = item.id;
    pushHistory();
    renderAll();
    persist();
  }

  function addWarehouse(latlng) {
    const item = {
      id: id(),
      type: 'warehouse',
      lat: latlng.lat,
      lng: latlng.lng,
      rotation: 0,
      width: 70,
      height: 35,
      color: '#2b8b57',
      name: 'Warehouse',
      locked: false
    };
    state.elements.push(item);
    state.selectedId = item.id;
    pushHistory();
    renderAll();
    persist();
  }

  function addLabel(latlng) {
    const item = {
      id: id(),
      type: 'label',
      lat: latlng.lat,
      lng: latlng.lng,
      color: '#ffffff',
      text: 'Label',
      locked: false
    };
    state.elements.push(item);
    state.selectedId = item.id;
    pushHistory();
    renderAll();
    persist();
  }

  function addMeasurement(latlng) {
    const item = {
      id: id(),
      type: 'measurement',
      a: { lat: latlng.lat, lng: latlng.lng },
      b: { lat: latlng.lat + 0.0002, lng: latlng.lng + 0.0002 },
      color: '#f2bd4b',
      customMeters: null,
      locked: false
    };
    state.elements.push(item);
    state.selectedId = item.id;
    pushHistory();
    renderAll();
    persist();
  }

  function addBollard(latlng) {
    const item = {
      id: id(),
      type: 'bollard',
      lat: latlng.lat,
      lng: latlng.lng,
      name: 'Bollard',
      locked: false
    };
    state.elements.push(item);
    state.selectedId = item.id;
    pushHistory();
    renderAll();
    persist();
  }

  function getShipPoints(e) {
    const l = e.length;
    const w = e.width;
    const local = [
      { x: -l * 0.5, y: -w * 0.5 },
      { x: -l * 0.5, y: w * 0.5 },
      { x: l * 0.32, y: w * 0.5 },
      { x: l * 0.5, y: 0 },
      { x: l * 0.32, y: -w * 0.5 }
    ];
    return local.map((p) => localToLatLng({ lat: e.lat, lng: e.lng }, p.x, p.y, e.rotation));
  }

  function getRectPoints(e) {
    const hw = e.width * 0.5;
    const hh = e.height * 0.5;
    const local = [
      { x: -hw, y: -hh },
      { x: -hw, y: hh },
      { x: hw, y: hh },
      { x: hw, y: -hh }
    ];
    return local.map((p) => localToLatLng({ lat: e.lat, lng: e.lng }, p.x, p.y, e.rotation || 0));
  }

  function renderElement(e) {
    let layer = null;

    if (e.type === 'ship') {
      layer = L.polygon(getShipPoints(e), {
        color: e.id === state.selectedId ? '#fff25c' : '#e5f4ff',
        weight: e.id === state.selectedId ? 3 : 1,
        fillColor: e.color || '#4da6ff',
        fillOpacity: 0.85
      });

      const display = (e.name || '').trim();
      if (display) {
        const angle = ((e.rotation % 360) + 360) % 360;
        const flip = angle > 90 && angle < 270;
        const labelAngle = flip ? -(e.rotation + 180) : -e.rotation;
        const t = L.tooltip({ permanent: true, direction: 'center', className: 'ship-label' })
          .setLatLng([e.lat, e.lng])
          .setContent(`<span style="display:inline-block;transform:rotate(${labelAngle}deg);white-space:nowrap">${escapeHtml(display)}</span>`);
        state.mapLayer.addLayer(t);
      }
    } else if (e.type === 'warehouse') {
      layer = L.polygon(getRectPoints(e), {
        color: e.id === state.selectedId ? '#fff25c' : '#d5f3dd',
        weight: e.id === state.selectedId ? 3 : 1,
        fillColor: e.color || '#2b8b57',
        fillOpacity: 0.7
      });

      const boxName = (e.name || e.text || '').trim();
      if (boxName) {
        const t = L.tooltip({ permanent: true, direction: 'center', className: 'box-label' })
          .setLatLng([e.lat, e.lng])
          .setContent(`<span>${escapeHtml(boxName)}</span>`);
        state.mapLayer.addLayer(t);
      }
    } else if (e.type === 'label') {
      layer = L.marker([e.lat, e.lng], {
        icon: L.divIcon({
          className: 'label-icon',
          html: `<div style="color:${e.color || '#fff'};font-weight:700;text-shadow:0 0 3px #001">${escapeHtml(e.text || 'Label')}</div>`
        })
      });
    } else if (e.type === 'measurement') {
      layer = L.polyline([[e.a.lat, e.a.lng], [e.b.lat, e.b.lng]], {
        color: e.id === state.selectedId ? '#fff25c' : (e.color || '#f2bd4b'),
        weight: 3,
        dashArray: '8 6'
      });

      const mid = { lat: (e.a.lat + e.b.lat) * 0.5, lng: (e.a.lng + e.b.lng) * 0.5 };
      const meters = e.customMeters != null ? e.customMeters : distMeters(e.a, e.b);
      const t = L.tooltip({ permanent: true, direction: 'center', className: 'measure-label' })
        .setLatLng([mid.lat, mid.lng])
        .setContent(`${Number(meters).toFixed(1).replace(/\\.0$/, '')}m`);
      state.mapLayer.addLayer(t);
    } else if (e.type === 'bollard') {
      layer = L.circleMarker([e.lat, e.lng], {
        radius: state.touchMode ? 9 : 6,
        color: e.id === state.selectedId ? '#fff25c' : '#d8dee4',
        weight: e.id === state.selectedId ? 3 : 2,
        fillColor: '#000000',
        fillOpacity: 1
      });
    }

    if (!layer) return;
    layer.on('click', (ev) => {
      L.DomEvent.stop(ev);
      state.selectedId = e.id;
      renderAll();
    });
    layer.on('contextmenu', (ev) => {
      L.DomEvent.stop(ev);
      state.selectedId = e.id;
      openElementMenu(ev.originalEvent.clientX, ev.originalEvent.clientY, e);
      renderAll();
    });

    state.mapLayer.addLayer(layer);
    if (state.mode === 'move' && !e.locked) enableMove(e, layer);
    if (state.selectedId === e.id) renderGizmo(e);
  }

  function renderGizmo(e) {
    if (state.mode === 'rotate' && (e.type === 'ship' || e.type === 'warehouse')) {
      const len = e.type === 'ship' ? e.length * 0.58 : Math.max(e.width, e.height) * 0.6;
      const p = localToLatLng({ lat: e.lat, lng: e.lng }, len, 0, e.rotation || 0);
      const h = L.circleMarker([p.lat, p.lng], {
        radius: state.touchMode ? 12 : 8,
        color: '#ff5151',
        fillColor: '#ff5151',
        fillOpacity: 0.9
      }).addTo(state.mapLayer);
      h.on('mousedown touchstart', (ev) => {
        L.DomEvent.stop(ev);
        beginRotate(e);
      });
      return;
    }

    if (state.mode === 'resize') {
      if (e.type === 'measurement') {
        ['a', 'b'].forEach((k) => {
          const c = L.circleMarker([e[k].lat, e[k].lng], {
            radius: state.touchMode ? 12 : 8,
            color: '#35d46c',
            fillColor: '#35d46c',
            fillOpacity: 0.9
          }).addTo(state.mapLayer);
          c.on('mousedown touchstart', (ev) => {
            L.DomEvent.stop(ev);
            beginMeasureResize(e, k);
          });
        });
      } else if (e.type === 'ship' || e.type === 'warehouse') {
        const corners = e.type === 'ship'
          ? [
            { x: -e.length * 0.5, y: -e.width * 0.5 },
            { x: -e.length * 0.5, y: e.width * 0.5 },
            { x: e.length * 0.5, y: e.width * 0.5 },
            { x: e.length * 0.5, y: -e.width * 0.5 }
          ]
          : [
            { x: -e.width * 0.5, y: -e.height * 0.5 },
            { x: -e.width * 0.5, y: e.height * 0.5 },
            { x: e.width * 0.5, y: e.height * 0.5 },
            { x: e.width * 0.5, y: -e.height * 0.5 }
          ];
        corners.forEach((p) => {
          const pt = localToLatLng({ lat: e.lat, lng: e.lng }, p.x, p.y, e.rotation || 0);
          const c = L.circleMarker([pt.lat, pt.lng], {
            radius: state.touchMode ? 11 : 7,
            color: '#35d46c',
            fillColor: '#35d46c',
            fillOpacity: 0.95
          }).addTo(state.mapLayer);
          c.on('mousedown touchstart', (ev) => {
            L.DomEvent.stop(ev);
            beginResize(e);
          });
        });
      }
    }

    if (state.mode === 'move') {
      const o = e.type === 'measurement' ? e.a : e;
      const c = L.circleMarker([o.lat, o.lng], {
        radius: state.touchMode ? 10 : 6,
        color: '#4da6ff',
        fillColor: '#4da6ff',
        fillOpacity: 0.9
      }).addTo(state.mapLayer);
      c.on('mousedown touchstart', (ev) => {
        L.DomEvent.stop(ev);
        beginMove(e);
      });
    }
  }

  function pointerLatLng(ev) {
    if (ev.latlng) return ev.latlng;
    const oe = ev.originalEvent || ev;
    const touch = oe.touches ? oe.touches[0] : oe.changedTouches ? oe.changedTouches[0] : oe;
    const rect = state.map.getContainer().getBoundingClientRect();
    return state.map.containerPointToLatLng([touch.clientX - rect.left, touch.clientY - rect.top]);
  }

  function attachDrag(moveFn, upFn) {
    state.map.dragging.disable();
    const mm = (e) => moveFn(pointerLatLng(e));
    const mu = () => {
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup', mu);
      document.removeEventListener('touchmove', mm);
      document.removeEventListener('touchend', mu);
      state.map.dragging.enable();
      state.ignoreMapClicksUntil = Date.now() + 300;
      upFn();
    };
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup', mu);
    document.addEventListener('touchmove', mm, { passive: true });
    document.addEventListener('touchend', mu);
  }

  function beginMove(e) {
    if (e.locked) return;
    const start = deepCopy(e);
    attachDrag((latlng) => {
      if (e.type === 'measurement') {
        const dLat = latlng.lat - start.a.lat;
        const dLng = latlng.lng - start.a.lng;
        e.a.lat = start.a.lat + dLat;
        e.a.lng = start.a.lng + dLng;
        e.b.lat = start.b.lat + dLat;
        e.b.lng = start.b.lng + dLng;
      } else {
        e.lat = latlng.lat;
        e.lng = latlng.lng;
      }
      renderAll();
    }, () => {
      pushHistory();
      persist();
    });
  }

  function enableMove(e, layer) {
    layer.on('mousedown touchstart', (ev) => {
      L.DomEvent.stop(ev);
      beginMove(e);
    });
  }

  function beginRotate(e) {
    if (e.locked) return;
    const start = deepCopy(e);
    attachDrag((latlng) => {
      const dx = state.map.distance([start.lat, start.lng], [start.lat, latlng.lng]) * (latlng.lng >= start.lng ? 1 : -1);
      const dy = state.map.distance([start.lat, start.lng], [latlng.lat, start.lng]) * (latlng.lat >= start.lat ? 1 : -1);
      e.rotation = (Math.atan2(dy, dx) * 180) / Math.PI;
      renderAll();
    }, () => {
      pushHistory();
      persist();
    });
  }

  function beginResize(e) {
    if (e.locked) return;
    const start = deepCopy(e);
    attachDrag((latlng) => {
      const dx = state.map.distance([start.lat, start.lng], [start.lat, latlng.lng]) * (latlng.lng >= start.lng ? 1 : -1);
      const dy = state.map.distance([start.lat, start.lng], [latlng.lat, start.lng]) * (latlng.lat >= start.lat ? 1 : -1);
      const r = ((start.rotation || 0) * Math.PI) / 180;
      const lx = dx * Math.cos(-r) - dy * Math.sin(-r);
      const ly = dx * Math.sin(-r) + dy * Math.cos(-r);
      if (e.type === 'ship') {
        e.length = round2(Math.max(10, Math.abs(lx) * 2));
        e.width = round2(Math.max(4, Math.abs(ly) * 2));
      } else if (e.type === 'warehouse') {
        e.width = round2(Math.max(5, Math.abs(lx) * 2));
        e.height = round2(Math.max(5, Math.abs(ly) * 2));
      }
      renderAll();
    }, () => {
      pushHistory();
      persist();
    });
  }

  function beginMeasureResize(e, key) {
    if (e.locked) return;
    attachDrag((latlng) => {
      e[key].lat = latlng.lat;
      e[key].lng = latlng.lng;
      renderAll();
    }, () => {
      pushHistory();
      persist();
    });
  }

  function openMapMenu(x, y, latlng) {
    ui.mapMenu.innerHTML = '';
    const items = [
      ['🛳 Add Ship', () => addShip(latlng)],
      ['🏬 Add Warehouse', () => addWarehouse(latlng)],
      ['🏷 Add Label', () => addLabel(latlng)],
      ['📏 Add Measure', () => addMeasurement(latlng)],
      ['⚫ Add Bollard', () => addBollard(latlng)]
    ];
    items.forEach(([label, fn]) => {
      const b = document.createElement('button');
      b.textContent = label;
      b.onclick = () => { hideMenus(); fn(); };
      ui.mapMenu.appendChild(b);
    });
    showMenu(ui.mapMenu, x, y);
  }

  function openElementMenu(x, y, e) {
    ui.elementMenu.innerHTML = '';

    const lock = document.createElement('button');
    lock.textContent = e.locked ? '🔓 Unlock' : '🔒 Lock';
    lock.onclick = () => {
      e.locked = !e.locked;
      hideMenus();
      pushHistory();
      renderAll();
      persist();
    };

    const del = document.createElement('button');
    del.textContent = '🗑 Delete';
    del.onclick = () => {
      state.elements = state.elements.filter((x2) => x2.id !== e.id);
      state.selectedId = null;
      hideMenus();
      pushHistory();
      renderAll();
      persist();
    };

    const cancel = document.createElement('button');
    cancel.textContent = '✖ Cancel';
    cancel.onclick = hideMenus;

    ui.elementMenu.append(lock, del, cancel);
    showMenu(ui.elementMenu, x, y);
  }

  function showMenu(menu, x, y) {
    hideMenus();
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.classList.remove('hidden');
  }

  function hideMenus() {
    ui.mapMenu.classList.add('hidden');
    ui.elementMenu.classList.add('hidden');
    ui.touchActionsMenu.classList.add('hidden');
  }

  function escapeHtml(s) {
    return String(s)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  function bindToolbar() {
    ui.topBar.querySelectorAll('[data-mode]').forEach((b) => {
      b.onmousedown = (e) => {
        e.preventDefault();
        e.stopPropagation();
      };
      b.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        state.mode = b.dataset.mode;
        renderAll();
      };
    });

    ui.touchToggle.onclick = () => {
      state.touchMode = !state.touchMode;
      document.body.classList.toggle('touch-mode', state.touchMode);
      ui.touchBar.classList.toggle('hidden', !state.touchMode);
      renderAll();
    };

    ui.touchActionsBtn.onclick = () => ui.touchActionsMenu.classList.toggle('hidden');
    ui.touchActionsMenu.querySelectorAll('[data-add]').forEach((b) => {
      b.onclick = () => {
        const c = state.map.getCenter();
        if (b.dataset.add === 'ship') addShip(c);
        if (b.dataset.add === 'warehouse') addWarehouse(c);
        if (b.dataset.add === 'label') addLabel(c);
        if (b.dataset.add === 'measurement') addMeasurement(c);
        if (b.dataset.add === 'bollard') addBollard(c);
        ui.touchActionsMenu.classList.add('hidden');
      };
    });

    if (ui.filterBtn) {
      ui.filterBtn.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        const isHidden = ui.filterMenu.classList.contains('hidden');
        ui.filterMenu.classList.toggle('hidden', !isHidden);
        ui.filterBtn.textContent = isHidden ? 'Filter Overlays ▴' : 'Filter Overlays ▾';
      };
    }

    const filterEntries = [
      ['ship', ui.fltShip],
      ['warehouse', ui.fltWarehouse],
      ['label', ui.fltLabel],
      ['measurement', ui.fltMeasurement],
      ['bollard', ui.fltBollard]
    ];
    filterEntries.forEach(([type, input]) => {
      if (!input) return;
      input.checked = isTypeVisible(type);
      input.onchange = () => {
        state.filters[type] = !!input.checked;
        const s = selected();
        if (s && !isTypeVisible(s.type)) state.selectedId = null;
        renderAll();
      };
    });

    ui.zoomInBtn.onclick = () => state.map.zoomIn();
    ui.zoomOutBtn.onclick = () => state.map.zoomOut();
    ui.touchDeleteBtn.onclick = () => {
      deleteSelectedElement();
    };
    ui.touchLockBtn.onclick = () => {
      const s = selected();
      if (!s) return;
      s.locked = !s.locked;
      pushHistory();
      renderAll();
      persist();
    };
  }

  function bindPorts() {
    ui.togglePortsBtn.onclick = () => {
      ui.portsBody.classList.toggle('hidden');
      ui.togglePortsBtn.textContent = ui.portsBody.classList.contains('hidden') ? '▸' : '▾';
    };
    ui.newPortBtn.onclick = () => ui.newPortForm.classList.toggle('hidden');
    ui.pickPortPointBtn.onclick = () => {
      state.pickPortPoint = true;
      alert('Click on map to choose the new port coordinates.');
    };
    ui.createPortBtn.onclick = () => {
      const name = (ui.newPortName.value || '').trim();
      if (!name) return;
      let lat;
      let lng;
      const coords = (ui.newPortCoords.value || '').trim();
      if (coords.includes(',')) {
        const [a, b] = coords.split(',').map((x) => Number(x.trim()));
        if (Number.isFinite(a) && Number.isFinite(b)) {
          lat = a;
          lng = b;
        }
      }
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
        const c = state.map.getCenter();
        lat = c.lat;
        lng = c.lng;
      }
      const key = slug(name);
      if (state.ports.some((p) => p.key === key)) {
        alert('Port already exists.');
        return;
      }
      state.ports.push({ key, name, center: [lat, lng], zoom: state.map.getZoom() });
      savePorts();
      renderPorts();
      ui.portSelect.value = key;
      switchPort(key);
      ui.newPortForm.classList.add('hidden');
      ui.newPortName.value = '';
      ui.newPortCoords.value = '';
    };

    ui.portSelect.onchange = () => switchPort(ui.portSelect.value);
  }

  function renderPorts() {
    ui.portSelect.innerHTML = '';
    state.ports.forEach((p) => {
      const opt = document.createElement('option');
      opt.value = p.key;
      opt.textContent = p.name;
      ui.portSelect.appendChild(opt);
    });
    ui.portSelect.value = state.portKey;
  }

  async function switchPort(portKey) {
    await persistImmediate();
    const p = state.ports.find((x) => x.key === portKey);
    if (!p) return;
    state.portKey = p.key;
    state.map.setView(p.center, p.zoom || state.map.getZoom());
    await loadPort();
    renderPorts();
  }

  async function loadPort() {
    try {
      const data = await hostRequest('loadPortData', { portKey: state.portKey });
      state.elements = Array.isArray(data.elements) ? data.elements : [];
    } catch {
      state.elements = [];
    }
    state.selectedId = null;
    state.history = [];
    state.historyIndex = -1;
    pushHistory();
    renderAll();
  }

  function persist() {
    clearTimeout(state.saveTimer);
    state.saveTimer = setTimeout(() => { persistImmediate(); }, 250);
  }

  async function persistImmediate() {
    clearTimeout(state.saveTimer);
    const payload = { elements: state.elements, meta: { savedAt: new Date().toISOString() } };
    try {
      await hostRequest('savePortData', { portKey: state.portKey, data: payload });
    } catch {
      // ignore
    }
  }

  function setVisible(input, show) {
    const lab = input.previousElementSibling;
    if (lab) lab.classList.toggle('hidden', !show);
    input.classList.toggle('hidden', !show);
  }

  function bindDetails() {
    ui.detailsDeleteBtn.onclick = () => {
      deleteSelectedElement();
    };

    ui.detailsLockBtn.onclick = () => {
      const s = selected();
      if (!s) return;
      s.locked = !s.locked;
      pushHistory();
      renderAll();
      persist();
    };

    ui.fldName.oninput = () => {
      const s = selected();
      if (!s || (s.type !== 'ship' && s.type !== 'warehouse' && s.type !== 'bollard')) return;
      s.name = ui.fldName.value;
      renderAll();
      persist();
    };
    ui.fldLength.oninput = () => {
      const s = selected();
      if (!s || (s.type !== 'ship' && s.type !== 'warehouse')) return;
      const n = Number(ui.fldLength.value);
      if (!Number.isFinite(n)) return;
      if (s.type === 'ship') s.length = round2(Math.max(1, n));
      else s.height = round2(Math.max(1, n));
      renderAll();
      persist();
    };
    ui.fldWidth.oninput = () => {
      const s = selected();
      if (!s || (s.type !== 'ship' && s.type !== 'warehouse')) return;
      const n = Number(ui.fldWidth.value);
      if (!Number.isFinite(n)) return;
      if (s.type === 'ship') s.width = round2(Math.max(1, n));
      else s.width = round2(Math.max(1, n));
      renderAll();
      persist();
    };
    ui.fldColor.oninput = () => {
      const s = selected();
      if (!s) return;
      s.color = ui.fldColor.value;
      renderAll();
      persist();
    };
    ui.fldArrival.oninput = () => {
      const s = selected();
      if (!s || s.type !== 'ship') return;
      s.arrival = ui.fldArrival.value;
      renderVessels();
      persist();
    };
    ui.fldText.oninput = () => {
      const s = selected();
      if (!s || s.type !== 'label') return;
      s.text = ui.fldText.value;
      renderAll();
      persist();
    };
    ui.fldMeasure.oninput = () => {
      const s = selected();
      if (!s || s.type !== 'measurement') return;
      const clean = String(ui.fldMeasure.value).replace(/[^0-9.]/g, '');
      ui.fldMeasure.value = clean;
    };
    ui.fldMeasure.onblur = () => {
      const s = selected();
      if (!s || s.type !== 'measurement') return;
      const clean = String(ui.fldMeasure.value).replace(/[^0-9.]/g, '').trim();
      const n = Number(clean);
      if (Number.isFinite(n)) {
        s.customMeters = n;
        ui.fldMeasure.value = `${n}m`;
        renderAll();
        persist();
      }
    };
  }

  function refreshDetails() {
    const s = selected();
    ui.detailsPanel.classList.toggle('hidden', !s);
    if (!s) return;

    ui.detailsLockBtn.textContent = s.locked ? '🔓' : '🔒';
    ui.detailsLockBtn.title = s.locked ? 'Unlock element' : 'Lock element';

    const ship = s.type === 'ship';
    const wh = s.type === 'warehouse';
    const bollard = s.type === 'bollard';
    const lab = s.type === 'label';
    const ms = s.type === 'measurement';

    setVisible(ui.fldName, ship || wh || bollard);
    setVisible(ui.fldLength, ship || wh);
    setVisible(ui.fldArrival, ship);
    setVisible(ui.fldWidth, ship || wh);
    setVisible(ui.fldColor, ship || wh || lab || ms);
    setVisible(ui.fldText, lab);
    setVisible(ui.fldMapDistance, ms);
    setVisible(ui.fldMeasure, ms);

    if (ui.lblName) {
      if (ship) ui.lblName.textContent = 'Vessel Name';
      else if (wh) ui.lblName.textContent = 'Warehouse Name';
      else if (bollard) ui.lblName.textContent = 'Bollard Label';
    }
    if (ui.lblLength) ui.lblLength.textContent = 'Length (m)';

    if (ship) {
      ui.fldName.value = s.name || '';
      ui.fldLength.value = fmt2(s.length || 90);
      ui.fldWidth.value = fmt2(s.width || 15);
      ui.fldColor.value = s.color || '#4da6ff';
      ui.fldArrival.value = s.arrival || today();
    }
    if (wh) {
      const whName = s.name || s.text || 'Warehouse';
      ui.fldName.value = whName;
      s.name = whName;
      ui.fldLength.value = fmt2(s.height || 35);
      ui.fldWidth.value = fmt2(s.width || 70);
      ui.fldColor.value = s.color || '#2b8b57';
    }
    if (bollard) {
      const bollardName = s.name || 'Bollard';
      ui.fldName.value = bollardName;
      s.name = bollardName;
    }
    if (lab) {
      ui.fldColor.value = s.color || '#ffffff';
      ui.fldText.value = s.text || '';
    }
    if (ms) {
      const m = s.customMeters != null ? s.customMeters : distMeters(s.a, s.b);
      ui.fldColor.value = s.color || '#f2bd4b';
      ui.fldMapDistance.value = `${Number(distMeters(s.a, s.b)).toFixed(1).replace(/\\.0$/, '')} m`;
      ui.fldMeasure.value = `${Number(m).toFixed(1).replace(/\\.0$/, '')}m`;
    }
  }

  function renderVessels() {
    ui.vesselsBody.innerHTML = '';
    state.elements.filter((x) => x.type === 'ship').forEach((s) => {
      const d = document.createElement('div');
      d.className = `item ${s.id === state.selectedId ? 'active' : ''}`;
      d.innerHTML = `<div>${escapeHtml(s.name || 'Vessel')}</div><small>${escapeHtml(s.arrival || '')}</small>`;
      d.onclick = () => {
        state.selectedId = s.id;
        state.map.panTo([s.lat, s.lng]);
        renderAll();
      };
      ui.vesselsBody.appendChild(d);
    });
  }

  function renderAll() {
    state.mapLayer.clearLayers();
    state.elements.filter((e) => isTypeVisible(e.type)).forEach(renderElement);

    ui.topBar.querySelectorAll('[data-mode]').forEach((b) => {
      b.classList.toggle('active', b.dataset.mode === state.mode);
    });
    if (ui.transformHint) {
      ui.transformHint.textContent = 'Select an element, then use gizmo handles on the map.';
    }
    if (ui.transformStatus) {
      ui.transformStatus.textContent = `Transform mode: ${state.mode}.`;
    }
    refreshDetails();
    renderVessels();
  }

  function bindVessels() {
    ui.toggleVesselsBtn.onclick = () => {
      ui.vesselsBody.classList.toggle('hidden');
      ui.toggleVesselsBtn.textContent = ui.vesselsBody.classList.contains('hidden') ? '▸' : '▾';
    };
  }

  function bindKeyboard() {
    document.addEventListener('keydown', (e) => {
      const tag = document.activeElement && document.activeElement.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;

      if (e.key === 'w' || e.key === 'W') { state.mode = 'move'; renderAll(); }
      if (e.key === 's' || e.key === 'S') { state.mode = 'resize'; renderAll(); }
      if (e.key === 'r' || e.key === 'R') { state.mode = 'rotate'; renderAll(); }

      if (e.ctrlKey && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); undo(); }
      if (e.ctrlKey && (e.key === 'y' || e.key === 'Y')) { e.preventDefault(); redo(); }
    });
  }

  function setupMap() {
    state.map = L.map('map', { zoomControl: true, preferCanvas: true, maxZoom: 19 }).setView([53.69984, -0.872207], 16);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      maxNativeZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    }).addTo(state.map);

    state.mapLayer = L.layerGroup().addTo(state.map);

    state.map.on('click', (e) => {
      hideMenus();
      if (Date.now() < state.ignoreMapClicksUntil) return;
      if (state.pickPortPoint) {
        ui.newPortCoords.value = `${e.latlng.lat.toFixed(6)}, ${e.latlng.lng.toFixed(6)}`;
        state.pickPortPoint = false;
        return;
      }
      state.selectedId = null;
      renderAll();
    });

    state.map.on('contextmenu', (e) => {
      hideMenus();
      openMapMenu(e.originalEvent.clientX, e.originalEvent.clientY, e.latlng);
    });

    let longPressTimer = null;
    state.map.getContainer().addEventListener('touchstart', (ev) => {
      const touch = ev.touches[0];
      longPressTimer = setTimeout(() => {
        const rect = state.map.getContainer().getBoundingClientRect();
        const ll = state.map.containerPointToLatLng([touch.clientX - rect.left, touch.clientY - rect.top]);
        openMapMenu(touch.clientX, touch.clientY, ll);
      }, 500);
    }, { passive: true });
    state.map.getContainer().addEventListener('touchend', () => clearTimeout(longPressTimer), { passive: true });
  }

  async function init() {
    setupMap();
    bindToolbar();
    bindPorts();
    bindDetails();
    bindVessels();
    bindKeyboard();

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.menu') && !e.target.closest('.touch-actions')) hideMenus();
      if (!e.target.closest('.filters-wrap') && ui.filterMenu && !ui.filterMenu.classList.contains('hidden')) {
        ui.filterMenu.classList.add('hidden');
        if (ui.filterBtn) ui.filterBtn.textContent = 'Filter Overlays ▾';
      }
    });

    renderPorts();
    await loadPort();
  }

  window.addEventListener('beforeunload', () => {
    persistImmediate();
  });

  init();
})();


