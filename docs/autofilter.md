# The AutoFilter

Excel's AutoFilter (Data > Filter) is a range of a sheet, outside a table, with filter buttons on its header row. A source tells the sheet where a sheet's AutoFilter is with `SheetInfo.autoFilter` (`{ range: 'A1:D20' }`, the header row first), and turns it on or off with `setAutoFilter(sheet, range | null)`, a single undo step. `./local` gives xlcalc's `SheetInput.autoFilter` and sets it through `setLayout`.
