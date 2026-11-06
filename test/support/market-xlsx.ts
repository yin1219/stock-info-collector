// Minimal synthetic OOXML archive; no licensed CMoney rows are copied.
export function workbook(rows: string[][]): Buffer {
  const xml =
    "<worksheet><sheetData>" +
    rows
      .map(
        (row, i) =>
          "<row>" +
          row
            .map(
              (v, j) =>
                `<c r="${String.fromCharCode(65 + j)}${i + 1}" t="inlineStr"><is><t>${v}</t></is></c>`,
            )
            .join("") +
          "</row>",
      )
      .join("") +
    "</sheetData></worksheet>";
  const name = Buffer.from("xl/worksheets/sheet1.xml"),
    data = Buffer.from(xml);
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50);
  local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(name.length, 26);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50);
  central.writeUInt32LE(data.length, 20);
  central.writeUInt32LE(data.length, 24);
  central.writeUInt16LE(name.length, 28);
  const directory = Buffer.concat([central, name]);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(1, 8);
  end.writeUInt16LE(1, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(local.length + name.length + data.length, 16);
  return Buffer.concat([local, name, data, directory, end]);
}
