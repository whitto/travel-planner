'use strict';

const $ = id => document.getElementById(id);
const tabs = ['Itinerary', 'Bookings', 'Places', 'Settings'];
const icons = { Flights: '✈', Hotels: '▤', Transport: '⇄', Activities: '☆', Other: '•' };
let trip;
let db;
let view = 'Itinerary';
let category = 'All';
let schedule = 'All dates';
let busy = false;
let ready = false;
let offlineReady = false;
let waitingWorker;
const fullDate = new Intl.DateTimeFormat('en-AU', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
const shortDate = new Intl.DateTimeFormat('en-AU', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' });
const asDate = date => new Date(date + 'T00:00:00Z');

function node(tag, value, className) {
  const element = document.createElement(tag);
  if (value != null) element.textContent = value;
  if (className) element.className = className;
  return element;
}
function button(label, handler, className) {
  const element = node('button', label, className);
  element.type = 'button';
  element.addEventListener('click', handler);
  return element;
}
function link(label, href) {
  const element = node('a', label, 'action');
  element.href = href;
  element.target = '_blank';
  element.rel = 'noopener noreferrer';
  return element;
}
function tell(message) {
  $('notice').textContent = message;
  $('notice').hidden = !message;
}
function today() {
  const date = new Date();
  return Itinerary.dateKey(date.getFullYear(), date.getMonth() + 1, date.getDate());
}
function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('travel-planner', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('planner');
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(Error('Close other Travel Planner tabs and reload.'));
  });
}
function storage(mode, operation) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('planner', mode);
    const request = operation(transaction.objectStore('planner'));
    transaction.oncomplete = () => resolve(request.result);
    transaction.onabort = () => reject(transaction.error || Error('Local storage transaction was interrupted.'));
    transaction.onerror = () => reject(transaction.error);
  });
}
function startFile(id) {
  if (!ready || busy) return;
  $(id).value = '';
  $(id).click();
}
function importReport(data) {
  const box = node('div', null, 'card');
  const count = data.records.length;
  box.append(node('p', `${count} ${count === 1 ? 'item' : 'items'} imported from ${data.sheetName}.`));
  if (data.extraSheets) box.append(node('p', `${data.extraSheets} additional worksheet(s) skipped.`));
  if (data.invalid.length) {
    const details = node('details');
    details.append(node('summary', `${data.invalid.length} row(s) skipped because the date in column A was missing or invalid`));
    const list = node('ul');
    for (const row of data.invalid) list.append(node('li', `Row ${row.row} · ${row.value || 'Blank date'}`));
    details.append(list);
    box.append(details);
  }
  return box;
}
function restoreShape(data) {
  const candidate = data?.trip;
  if (data?.version !== 1 || !candidate || !candidate.name || !Array.isArray(candidate.records) ||
      !candidate.records.length || !Array.isArray(candidate.invalid) ||
      ![candidate.name, candidate.filename, candidate.sheetName].every(value => typeof value === 'string') ||
      !Number.isInteger(candidate.extraSheets) || candidate.extraSheets < 0) throw Error('Invalid Travel Planner backup.');
  const textKeys = ['name', 'type', 'place', 'address', 'country', 'ref', 'url'];
  for (const item of candidate.records) {
    const date = item?.date;
    if (typeof date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
        Itinerary.dateKey(...date.split('-').map(Number)) !== date ||
        !Number.isInteger(item.row) || item.row < 2 || typeof item.name !== 'string' ||
        ![...Itinerary.categories.slice(1), 'Other'].includes(item.category) ||
        !(item.time === null || (Number.isInteger(item.time) && item.time >= 0 && item.time < 1440)) ||
        !textKeys.every(key => item[key] === undefined || typeof item[key] === 'string') ||
        !Array.isArray(item.fields) || item.fields.some(field => typeof field?.label !== 'string' ||
          typeof field?.value !== 'string')) throw Error('Invalid itinerary item in backup.');
  }
  if (candidate.invalid.some(row => !Number.isInteger(row?.row) || row.row < 2 || typeof row.value !== 'string')) {
    throw Error('Invalid import report in backup.');
  }
  Itinerary.sort(candidate.records);
  return candidate;
}
async function receive(file, backup) {
  if (!ready || !file || busy) return;
  busy = true;
  document.body.setAttribute('aria-busy', 'true');
  tell('Reading ' + file.name + '…');
  try {
    if (file.size > 25 * 1024 * 1024) throw Error('Choose a file smaller than 25 MB.');
    let next;
    if (backup) {
      next = restoreShape(JSON.parse(await file.text()));
    } else {
      if (!/\.xlsx?$/i.test(file.name)) throw Error('Choose an .xls or .xlsx file.');
      next = Itinerary.workbook(await file.arrayBuffer(), file.name);
      if (!next.records.length) {
        tell('No valid itinerary items were found.');
        $('notice').append(importReport(next));
        return;
      }
      next.filename = file.name;
      next.name = file.name.replace(/\.xlsx?$/i, '').replaceAll('_', ' ').trim() || 'Travel Planner';
    }
    if (trip && !confirm(`Replace “${trip.name}” with “${next.name}” on this device? ${next.records.length} valid item(s), ${next.invalid.length} skipped row(s).`)) {
      tell('Replacement cancelled.');
      return;
    }
    await storage('readwrite', store => store.put(next, 'trip'));
    trip = next;
    setView('Itinerary');
    tell(backup ? 'Backup restored.' : 'Workbook saved on this device.');
    $('notice').append(importReport(next));
  } catch (error) {
    tell(`Could not ${backup ? 'restore the backup' : 'import the workbook'}. ${error.message}`);
  } finally {
    busy = false;
    document.body.removeAttribute('aria-busy');
  }
}
function exportBackup() {
  const url = URL.createObjectURL(new Blob([JSON.stringify({ version: 1, trip }, null, 2)], { type: 'application/json' }));
  const anchor = node('a');
  anchor.href = url;
  anchor.download = 'travel-planner-backup.json';
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
async function removeTrip() {
  if (busy || !confirm(`Delete “${trip.name}” from this browser?`)) return;
  busy = true;
  try {
    await storage('readwrite', store => store.delete('trip'));
    trip = undefined;
    setView('Itinerary');
    tell('Itinerary deleted from this browser.');
  } catch (error) {
    tell('Could not delete the itinerary. ' + error.message);
  } finally { busy = false; }
}
function maps(target) {
  const actions = node('div', null, 'actions');
  if (target) {
    const query = encodeURIComponent(target);
    actions.append(link('Apple Maps', 'https://maps.apple.com/?q=' + query),
      link('Google Maps', 'https://www.google.com/maps/search/?api=1&query=' + query));
  }
  return actions;
}
const destination = item => [...new Set([item.place, item.address, item.country].filter(Boolean))].join(', ');
function itemCard(item, includeDate) {
  const box = node('details', null, 'item');
  const summary = node('summary');
  summary.append(node('strong', `${icons[item.category]}  ${item.name}`));
  const time = item.time == null ? '' : `${String(Math.floor(item.time / 60)).padStart(2, '0')}:${String(item.time % 60).padStart(2, '0')}`;
  const description = [includeDate && shortDate.format(asDate(item.date)), time, destination(item), includeDate && item.ref].filter(Boolean).join(' · ');
  if (description) summary.append(node('span', description, 'hint'));
  box.append(summary);
  const dl = node('dl');
  for (const field of item.fields) dl.append(node('dt', field.label), node('dd', field.value));
  box.append(dl);
  const actions = maps(destination(item));
  if (item.ref) actions.prepend(button('Copy reference', async () => {
    try { await navigator.clipboard.writeText(item.ref); tell('Booking reference copied.'); }
    catch { prompt('Copy this booking reference', item.ref); }
  }));
  const url = Itinerary.safeURL(item.url);
  if (url) actions.append(link('Booking details', url));
  if (actions.childElementCount) box.append(actions);
  return box;
}
function matches(item, query) {
  return [item.date, item.name, item.category, ...item.fields.flatMap(field => [field.label, field.value])]
    .some(value => value.toLowerCase().includes(query));
}
function settings(content) {
  const box = node('section', null, 'card');
  box.append(node('h3', 'Workbook'), node('p', trip.filename), importReport(trip));
  box.append(button('Refresh from Excel / Change spreadsheet', () => startFile('excel'), 'primary'));
  content.append(box);
  const local = node('section', null, 'card');
  local.append(node('h3', 'On this device'),
    node('p', 'Your itinerary is saved in this browser. The browser or operating system can clear local storage. Keep the original workbook as your master copy.'),
    node('p', 'Backups contain the full itinerary. Keep them somewhere private. Maps and booking sites may require internet access.'));
  const actions = node('div', null, 'actions');
  actions.append(button('Export backup', exportBackup), button('Restore backup', () => startFile('backup')),
    button('Delete itinerary', removeTrip));
  local.append(actions);
  content.append(local);
}
function renderContent() {
  if (!trip) return;
  const content = $('content');
  content.replaceChildren();
  if (view === 'Settings') { settings(content); return; }
  const query = $('search').value.trim().toLowerCase();
  const rows = trip.records.filter(item => matches(item, query));
  if (view === 'Places') {
    const places = new Map();
    for (const item of rows) {
      const target = destination(item);
      if (!target) continue;
      const key = target.toLocaleLowerCase('en-AU');
      const group = places.get(key) ?? { target, count: 0 };
      group.count++;
      places.set(key, group);
    }
    for (const { target, count } of places.values()) {
      const box = node('section', null, 'card');
      box.append(node('strong', target), node('p', `${count} itinerary item(s)`, 'hint'), maps(target));
      content.append(box);
    }
  } else {
    let lastDate;
    const current = today();
    for (const item of rows) {
      if (category !== 'All' && item.category !== category) continue;
      if (view === 'Bookings' && !item.ref && !Itinerary.safeURL(item.url)) continue;
      if (view === 'Itinerary') {
        if (schedule === 'Today' && item.date !== current) continue;
        if (schedule === 'Upcoming' && item.date <= current) continue;
        if (item.date !== lastDate) {
          lastDate = item.date;
          const active = lastDate === current;
          content.append(node('h3', (active ? 'Today · ' : '') + fullDate.format(asDate(lastDate)), active ? 'today' : ''));
        }
      }
      content.append(itemCard(item, view === 'Bookings'));
    }
  }
  if (!content.childElementCount) {
    const message = { Places: 'No destinations match. Add a city, address or country in Excel.',
      Bookings: 'No matching items with a booking reference or link.',
      Itinerary: 'No itinerary items match this view.' };
    content.append(node('p', message[view], 'card'));
  }
}
function renderFilters() {
  const filters = $('filters');
  filters.replaceChildren();
  if (!['Itinerary', 'Bookings'].includes(view)) return;
  const groups = [Itinerary.categories];
  if (view === 'Itinerary') groups.push(['All dates', 'Today', 'Upcoming']);
  for (const [index, labels] of groups.entries()) {
    for (const label of labels) {
      const element = button(label, () => {
        if (index) schedule = label;
        else category = label;
        renderFilters();
        renderContent();
        [...filters.children].find(child => child.textContent === label)?.focus();
      });
      element.setAttribute('aria-pressed', String(label === (index ? schedule : category)));
      filters.append(element);
    }
  }
}
function updateStats() {
  const stats = $('stats');
  stats.replaceChildren();
  if (trip) {
    const days = new Set(trip.records.map(item => item.date)).size;
    const countries = new Set(trip.records.map(item => item.country?.trim().toLowerCase()).filter(Boolean)).size;
    stats.append(node('span', `${days} scheduled ${days === 1 ? 'day' : 'days'}`, 'stat'));
    if (countries) stats.append(node('span', `${countries} ${countries === 1 ? 'country' : 'countries'}`, 'stat'));
  }
  const status = navigator.onLine ? (offlineReady ? 'Offline ready' : 'Online') : 'Offline';
  stats.append(node('span', status, 'stat'));
  if (waitingWorker) stats.append(button('Update app', () => waitingWorker.postMessage('activate-update')));
}
function render() {
  $('welcome').hidden = !!trip;
  $('planner').hidden = !trip;
  document.querySelector('nav').hidden = !trip;
  $('title').textContent = trip?.name ?? 'Travel Planner';
  $('subtitle').textContent = 'Your entire journey in one place';
  if (trip) {
    const first = trip.records[0].date;
    const last = trip.records.at(-1).date;
    const start = shortDate.format(asDate(first));
    $('subtitle').textContent = first === last ? start : `${start} – ${shortDate.format(asDate(last))}`;
  }
  updateStats();
  $('view-title').textContent = view;
  $('controls').hidden = view === 'Settings';
  $('search-view').textContent = view.toLowerCase();
  for (const tab of $('tabs').children) {
    if (tab.textContent === view) tab.setAttribute('aria-current', 'page');
    else tab.removeAttribute('aria-current');
  }
  renderFilters();
  renderContent();
}
function setView(next) {
  view = next;
  category = 'All';
  schedule = 'All dates';
  $('search').value = '';
  render();
  window.scrollTo(0, 0);
}
for (const tab of tabs) $('tabs').append(button(tab, () => { tell(''); setView(tab); }));
document.querySelector('[data-action=choose]').onclick = () => startFile('excel');
document.querySelector('[data-action=restore]').onclick = () => startFile('backup');
$('excel').onchange = event => receive(event.target.files[0], false);
$('backup').onchange = event => receive(event.target.files[0], true);
$('search').oninput = renderContent;
window.addEventListener('online', updateStats);
window.addEventListener('offline', updateStats);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && view === 'Itinerary') renderContent();
});
(async () => {
  try {
    db = await openDB();
    trip = await storage('readonly', store => store.get('trip'));
    ready = true;
  } catch (error) {
    tell('Could not open local storage. ' + error.message);
  }
  render();
})();
if ('serviceWorker' in navigator) {
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (hadController) location.reload();
  });
  navigator.serviceWorker.register('./sw.js').then(registration => {
    const update = () => { waitingWorker = registration.waiting; updateStats(); };
    update();
    registration.addEventListener('updatefound', () => {
      registration.installing?.addEventListener('statechange', update);
    });
    navigator.serviceWorker.ready.then(() => { offlineReady = true; updateStats(); });
  }).catch(() => tell('Offline setup failed. Reload while connected to try again.'));
}
