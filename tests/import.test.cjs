const assert = require('node:assert/strict');
const { test } = require('node:test');
globalThis.XLSX = require('../vendor/xlsx.min.js');
require('../itinerary.js');

function fixture(type) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Date'], ['01/01/2000']]), 'Ignore me');
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Example daily itinerary'], [],
    ['Date', 'Type', 'Activity', 'Booking Ref', 'Start Time', 'Address / Area', 'Booking URL', 'Extra detail'],
    ['5 Jan', 'Train', 'Example train', 'DEMO03', 0.5, 'Demo Station', 'https://example.org'],
    ['31 Dec', 'Flight', 'Example flight', 'DEMO01', 0.25, 'Demo Airport'],
    ['5 Jan', 'Hotel', 'Example hotel', 'DEMO02', undefined, 'Demo Street', 'javascript:alert(1)', '<b>Plain text</b>'],
    [],
    ['31/02/2027', 'Other', 'Invalid'],
    [46388, 'Other', 'Numeric date'],
    ['05/01/2027', 'Other', 'Second untimed item']
  ]);
  // Date-only rows in Daily Plan are assumed to be in calendar order for year rollover.
  XLSX.utils.book_append_sheet(book, sheet, 'Daily Plan');
  return XLSX.write(book, { type: 'buffer', bookType: type });
}
for (const type of ['xls', 'xlsx']) {
  test(type + ' named sheet, dates, chronology, invalid rows and arbitrary fields', () => {
    const result = Itinerary.workbook(fixture(type), 'Example 2027.' + type);
    assert.equal(result.sheetName, 'Daily Plan');
    assert.equal(result.extraSheets, 1);
    assert.deepEqual(result.records.map(item => item.name),
      ['Numeric date', 'Example train', 'Second untimed item', 'Example flight', 'Example hotel']);
    assert.equal(result.records[1].time, 720);
    assert.equal(result.records[1].address, 'Demo Station');
    assert.equal(result.records[1].ref, 'DEMO03');
    const hotel = result.records.at(-1);
    assert.equal(hotel.date, '2028-01-05');
    assert.equal(hotel.fields.at(-1).value, '<b>Plain text</b>');
    assert.deepEqual(result.invalid, [{ row: 8, value: '31/02/2027' }]);
  });
}
test('first-sheet fallback and source order for untimed dates', () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['Date', 'Hotel Name'], ['03/12/2026', 'A'], ['2026-12-03', 'B']
  ]), 'Itinerary');
  const result = Itinerary.workbook(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  assert.deepEqual(result.records.map(item => item.name), ['A', 'B']);
  assert.equal(result.records[0].category, 'Hotels');
});
test('rejects corrupt workbooks and unsafe booking URLs', () => {
  assert.throws(() => Itinerary.workbook(Buffer.from('invalid')), /supported Excel workbook/);
  assert.equal(Itinerary.safeURL('javascript:alert(1)'), '');
  assert.equal(Itinerary.safeURL('https://user:password@example.org'), '');
  assert.equal(Itinerary.safeURL('https://example.org'), 'https://example.org/');
});
