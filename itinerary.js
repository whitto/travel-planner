/* Workbook normalization shared by the app and its tests. */
globalThis.Itinerary = (() => {
  const categories = ['All', 'Flights', 'Hotels', 'Transport', 'Activities'];
  const aliases = {
    type: /^(type|item type|booking type|category|travel type)$/,
    name: /^(name|title|description|item|hotel name|hotel|accommodation|accomodation|property|flight|activity|train|service)$/,
    place: /^(city|town|destination|location|place)$/,
    address: /^(address|street address|hotel address|property address|address area)$/,
    country: /^(country|destination country)$/,
    ref: /^(ref|reference|booking ref|booking reference|booking reference number|confirmation|confirmation number|confirmation no|confirmation code|reservation number|pnr)$/,
    url: /^(booking url|booking link|url|website|booking website)$/,
    time: /^(time|departure time|start time|check in time|check in|checkin)$/
  };
  const normalize = key => key.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const pad = value => String(value).padStart(2, '0');
  function dateKey(year, month, day) {
    if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return '';
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCMonth() === month - 1 && date.getUTCDate() === day
      ? `${year}-${pad(month)}-${pad(day)}` : '';
  }
  function readDate(value, date1904) {
    if (typeof value === 'number') {
      const parsed = XLSX.SSF.parse_date_code(value, { date1904 });
      return parsed ? dateKey(parsed.y, parsed.m, parsed.d) : '';
    }
    const input = String(value ?? '').trim();
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(input);
    if (iso) return dateKey(+iso[1], +iso[2], +iso[3]);
    const au = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/.exec(input);
    if (!au) return '';
    const year = +au[3];
    return dateKey(au[3].length === 2 ? year + (year < 70 ? 2000 : 1900) : year, +au[2], +au[1]);
  }
  function readTime(value) {
    if (typeof value === 'number') return value >= 0 && value < 1 ? Math.floor(value * 1440 + 1e-7) : null;
    const match = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm)?$/i.exec(String(value ?? '').trim());
    if (!match || +match[2] > 59) return null;
    const hour = +match[1];
    if (match[3]) {
      if (hour < 1 || hour > 12) return null;
      return (hour % 12 + (match[3].toLowerCase() === 'pm' ? 12 : 0)) * 60 + +match[2];
    }
    return hour < 24 ? hour * 60 + +match[2] : null;
  }
  function category(type = "", keys) {
    const input = normalize(type) || keys.join(' ');
    if (/\b(flights?|airline|airplane|plane)\b/.test(input)) return 'Flights';
    if (/\b(hotels?|accommodation|hostel|lodging|resort|apartment)\b/.test(input)) return 'Hotels';
    if (/\b(transport|trains?|rail|buses|bus|taxi|ferry|transfer|car)\b/.test(input)) return 'Transport';
    if (/\b(activit(?:y|ies)|tour|tickets?|event|museum|sightseeing)\b/.test(input)) return 'Activities';
    return 'Other';
  }
  function workbook(buffer, filename = '') {
    const bytes = new Uint8Array(buffer);
    const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    const ole = bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
    const biff = bytes[0] === 0x09 && [0, 2, 4, 8].includes(bytes[1]);
    if (!zip && !ole && !biff) throw Error('This file is not a supported Excel workbook.');
    const sheetNames = XLSX.read(bytes, { type: 'array', bookSheets: true }).SheetNames;
    const dailyPlan = sheetNames.find(name => normalize(name) === 'daily plan');
    const sheetName = dailyPlan || sheetNames[0];
    if (!sheetName) throw Error('The workbook contains no worksheets.');
    const book = XLSX.read(bytes, {
      type: 'array', dense: true, sheets: sheetName, cellFormula: false, cellHTML: false, bookVBA: false
    });
    const sheet = book.Sheets[sheetName];
    if (!sheet?.['!ref']) throw Error('The selected worksheet is empty.');
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (range.e.r > 100000 || range.e.c > 1023) throw Error('The selected worksheet exceeds 100,001 rows or 1,024 columns.');
    const rows = sheet['!data'] ?? [];
    const records = [];
    const invalid = [];
    const dateHeader = dailyPlan ? rows.slice(0, 25).findIndex(row => normalize(String(row?.[0]?.v ?? '')) === 'date') : -1;
    const headerIndex = Math.max(0, dateHeader);
    const headers = rows[headerIndex];
    const hasSummary = dailyPlan && normalize(String(headers?.[15]?.v ?? '')) === 'summary' &&
      normalize(String(headers?.[16]?.v ?? '')) === 'value';
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const filenameYear = /(?:^|[^0-9])((?:19|20)[0-9]{2})(?![0-9])/.exec(filename);
    let year = filenameYear ? Number(filenameYear[1]) : null;
    let lastMonth = null;
    for (let index = headerIndex + 1; index <= range.e.r; index++) {
      const cells = rows[index];
      if (!cells?.some((cell, column) => (!hasSummary || (column !== 15 && column !== 16)) &&
        cell?.v != null && String(cell.v).trim())) continue;
      const dateValue = cells[0]?.v;
      let date = readDate(dateValue, !!book.Workbook?.WBProps?.date1904);
      if (dailyPlan && date) {
        year = Number(date.slice(0, 4));
        lastMonth = Number(date.slice(5, 7));
      } else if (dailyPlan && year) {
        const abbreviated = /^(\d{1,2})\s+([a-z]{3})$/i.exec(String(dateValue ?? '').trim());
        if (abbreviated) {
          const month = months.indexOf(abbreviated[2].toLowerCase()) + 1;
          const nextYear = lastMonth !== null && lastMonth - month > 6 ? year + 1 : year;
          date = dateKey(nextYear, month, Number(abbreviated[1]));
          if (date) {
            year = nextYear;
            lastMonth = month;
          }
        }
      }
      if (!date) {
        invalid.push({ row: index + 1, value: String(cells[0]?.w ?? dateValue ?? '') });
        continue;
      }
      const fields = [];
      const found = {};
      let activity = '';
      let accommodation = '';
      const keys = [];
      cells.forEach((cell, column) => {
        if (!column || cell?.v == null || !String(cell.v).trim()) return;
        const label = String(headers?.[column]?.v ?? '').trim() || `Column ${XLSX.utils.encode_col(column)}`;
        const value = String(cell.w ?? cell.v);
        fields.push({ label, value });
        const key = normalize(label);
        keys.push(key);
        if (key === 'activity') activity = value;
        if (/^(accommodation|accomodation|hotel|hotel name)$/.test(key)) accommodation = value;
        if (!found.url && aliases.url.test(key)) found.url = safeURL(cell.l?.Target) || value;
        for (const [field, pattern] of Object.entries(aliases)) {
          if (found[field] == null && pattern.test(key)) {
            found[field] = field === 'time' ? readTime(cell.v) : value;
          }
        }
      });
      const time = found.time ?? (typeof dateValue === 'number' && dateValue % 1 ? readTime(dateValue % 1) : null);
      const inferred = activity ? category(activity, []) : 'Other';
      const kind = dailyPlan && !found.type
        ? (inferred !== 'Other' ? inferred : accommodation ? 'Hotels' : activity ? 'Activities' : 'Other')
        : category(found.type, keys);
      records.push({ ...found, row: index + 1, date, time, category: kind,
        name: (dailyPlan && (activity || accommodation || found.place)) || found.name || found.type || fields[0]?.value || 'Itinerary item', fields });
    }
    sort(records);
    return { records, invalid, sheetName, extraSheets: sheetNames.length - 1 };
  }
  function sort(records) {
    return records.sort((a, b) => a.date.localeCompare(b.date) ||
      (a.time ?? 1440) - (b.time ?? 1440) || a.row - b.row);
  }
  function safeURL(value) {
    try {
      const url = new URL(value);
      return ['https:', 'http:'].includes(url.protocol) && !url.username && !url.password ? url.href : '';
    } catch { return ''; }
  }
  return { categories, workbook, dateKey, sort, safeURL };
})();
