# Hyperlinks, notes, Clear and the default font

## The workbook's default font

A sheet whose `SheetInfo` has `font` (`{ name, size }`, the workbook's Normal style) draws every cell without a font of its own in that font, as Excel draws it: the grid's table takes the font's name and its size in the grid's pixels (11 pt is 12.5 px, the rest in proportion), and the cells inherit it. A cell's own font name or size takes precedence. The row numbers and column letters keep the interface font, `--xs-font`. Without `font` nothing changes: cells are drawn in `--xs-font` and `--xs-font-size`. The row heights are the source's either way (`rowHeights`, `defaultRowHeight`), and overflow into empty neighbors goes by the column widths, so neither depends on the font.
