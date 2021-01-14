type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet

type Decoder<T> = (keys: string[], values: any[]) => T

function _rawDecoder<T>(keys: string[], values: any[]): T {
  return keys.reduce((data, key, n) => {
    data[key] = values[n]
    return data
  }, ({} as T))
}

function decodeRangeWith<T>(range: Range, decoder: Decoder<T>): T[] {
  const data: T[] = []
  const values = range.getValues()
  const keys = values[0]

  for (let i = 1; i < values.length; i++) {
    data.push(decoder(keys, values[i]))
  }

  return data
}

function decodeRange<T>(range: Range): T[] {
  return decodeRangeWith(range, _rawDecoder) as T[]
}

function decodeSheet<T>(sheet: Sheet, decoder: Decoder<T> = _rawDecoder as Decoder<T>): T[] {
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
  return decodeRangeWith(range, decoder)
}

function append<T>(sheet: Sheet, keys: string[], data: T[]): void {
  const start = sheet.getLastRow() + 1
  const range = sheet.getRange(start, 1, data.length, keys.length)
  const values = data.map(entry => keys.map(key => entry[key]))

  range.setValues(values)
}

function overwrite<T>(sheet: Sheet, header: string[], data: T[]) {
  const lastRow = sheet.getLastRow()
  if (lastRow > 0) {
    sheet.insertRowsAfter(lastRow, data.length + 1)
    sheet.deleteRows(1, lastRow)
  }

  const range = sheet.getRange(1, 1, data.length + 1, header.length)
  const values = [header]
  for (const item of data) {
    values.push(header.map(key => item[key]))
  }
  range.setValues(values)
}

export {}