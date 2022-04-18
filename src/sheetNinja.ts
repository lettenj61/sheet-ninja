type Range = GoogleAppsScript.Spreadsheet.Range
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
type Sheet = GoogleAppsScript.Spreadsheet.Sheet

type Decoder<T> = (keys: string[], values: any[]) => T

function _rawDecoder<T>(keys: string[], values: any[]): T {
  return keys.reduce((data, key, n) => {
    data[key] = values[n]
    return data
  }, {} as T)
}


// CORE FUNCTIONS

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

function decodeSheet<T = Record<string, any>>(sheet: Sheet): T[] {
  return decodeSheetWith(sheet, _rawDecoder as Decoder<T>)
}

function decodeSheetWith<T>(sheet: Sheet, decoder: Decoder<T>): T[] {
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
  return decodeRangeWith(range, decoder)
}

function decodeSheetMetadata<T>(sheet: Sheet): T {
  return sheet.getDeveloperMetadata().reduce((bag, metadata) => {
    const key = metadata.getKey()
    bag[key] = metadata.getValue()
    return bag
  }, {} as T)
}

function append<T>(sheet: Sheet, keys: string[], data: T[]): void {
  if (!data.length) return

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

function updateOrInsertBy<T, Id>(
  sheet: Sheet,
  header: string[],
  data: T[],
  toKey: (value: T) => Id,
  duplicate: boolean = false
): void {
  if (!data.length) {
    return
  }
  const merged = data.reduce<{ seen: Set<Id>; bag: T[] }>(
    (state, item) => {
      const key = toKey(item)
      if (!state.seen.has(key)) {
        state.seen.add(key)
        state.bag.push(item)
      }
      return state
    },
    { seen: new Set(), bag: [] }
  ).bag

  const prevData: T[] = decodeSheet(sheet)
  const newRecords: T[] = []
  for (const upd of merged) {
    let found = false
    for (let i = 0; i < prevData.length; i++) {
      const prev = prevData[i]
      const currentKey = toKey(upd)
      const oldKey = toKey(prev)
      if (currentKey === oldKey) {
        found = true
        const updated: T = Object.assign({}, prev, upd)
        const newValues = header.map(k => updated[k])
        const range = sheet.getRange(i + 2, 1, 1, header.length)
        range.setValues([newValues])
        if (!duplicate) {
          break
        }
      }
    }

    if (!found) {
      newRecords.push(upd)
    }
  }

  if (newRecords.length) {
    append(sheet, header, newRecords)
  }
}

function copySheet(src: Sheet, dest: Spreadsheet, newName: string): Sheet {
  const copied = src.copyTo(dest)
  copied.setName(newName)
  return copied
}

function clearContents(sheet: Sheet, startRow: number, numColumns: number): void {
  const lastRow = sheet.getLastRow()
  const numRows = lastRow - (startRow - 1)
  if (numColumns < 1) return
  const range = sheet.getRange(startRow, 1, numRows, numColumns)
  range.clearContent()
}


// O/R MAPPER FEATURES

function createMapper<T, Id extends string | number>(init: MapperInit<T, Id>): Mapper<T, Id> {
  return new Mapper(init)
}

export type MapperInit<T, Id extends string | number> = {
  sheetId: string
  sheetName: string
  keys: (keyof T)[]
  toId: (value: T) => Id
}

export class Mapper<T, Id extends string | number> {
  private readonly init: MapperInit<T, Id>
  private readonly sheet: Sheet

  constructor(init: MapperInit<T, Id>) {
    this.init = init
    this.sheet = Mapper.openSheet(init.sheetId, init.sheetName)
  }

  private get keysAsString(): string[] {
    return this.init.keys as string[]
  }

  readAll(): T[] {
    return decodeSheet(this.sheet)
  }

  findById(id: Id): T | undefined {
    const records = this.readAll()

    return records.find(item => this.init.toId(item) === id)
  }

  upsert(data: T[]): void {
    updateOrInsertBy(this.sheet, this.keysAsString, data, val => this.init.toId(val))
  }

  deleteBy(pred: (item?: T) => boolean): void {
    const data = this.readAll().filter(item => !pred(item))
    overwrite(this.sheet, this.keysAsString, data)
  }

  static openSheet(sheetId: string, sheetName: string): Sheet {
    try {
      const workbook = SpreadsheetApp.openById(sheetId)
      const sheet = workbook.getSheetByName(sheetName)

      return sheet
    } catch (ex) {
      throw ex
    }
  }
}

export {
  decodeRange,
  decodeRangeWith,
  decodeSheet,
  decodeSheetWith,
  decodeSheetMetadata,
  append,
  overwrite,
  updateOrInsertBy,
  copySheet,
  clearContents,
  createMapper,
}
