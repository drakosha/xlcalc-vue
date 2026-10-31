// The texts of sheet protection, English and Russian, as Excel for Windows 16.0 words them in its alert and its Protect Sheet
// and Format Cells dialogs; Unprotect Sheet and the note under Protect Sheet's boxes are Excel for Mac 16.114's (Windows has no
// such note). src/messages.ts includes them in its sets.
import type { ProtectionMessages } from './contract/messages'

export const enProtection: ProtectionMessages = {
  protectedAlert: "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.",
  menuProtectSheet: 'Protect Sheet...',
  menuUnprotectSheet: 'Unprotect Sheet',
  protectTitle: 'Protect Sheet',
  protectContents: 'Protect worksheet and contents of locked cells',
  protectAllow: 'Allow all users of this worksheet to:',
  protectBox: (box) =>
    ({
      selectLockedCells: 'Select locked cells',
      selectUnlockedCells: 'Select unlocked cells',
      formatCells: 'Format cells',
      formatColumns: 'Format columns',
      formatRows: 'Format rows',
      insertColumns: 'Insert columns',
      insertRows: 'Insert rows',
      insertHyperlinks: 'Insert hyperlinks',
      deleteColumns: 'Delete columns',
      deleteRows: 'Delete rows',
      sort: 'Sort',
      autoFilter: 'Use AutoFilter',
      pivotTables: 'Use PivotTable and PivotChart',
      objects: 'Edit objects',
      scenarios: 'Edit scenarios',
    })[box],
  fmtProtection: 'Protection',
  fmtLocked: 'Locked',
  fmtHidden: 'Hidden',
  fmtProtectionNote: 'Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect group, Protect Sheet button).',
}

export const ruProtection: ProtectionMessages = {
  protectedAlert: 'Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль.',
  menuProtectSheet: 'Защитить лист...',
  menuUnprotectSheet: 'Снять защиту листа',
  protectTitle: 'Защитить лист',
  protectContents: 'Защитить лист и содержимое защищаемых ячеек',
  protectAllow: 'Разрешить всем пользователям этого листа:',
  protectBox: (box) =>
    ({
      selectLockedCells: 'Выделение заблокированных ячеек',
      selectUnlockedCells: 'Выделение незаблокированных ячеек',
      formatCells: 'Форматирование ячеек',
      formatColumns: 'Форматирование столбцов',
      formatRows: 'Форматирование строк',
      insertColumns: 'Вставку столбцов',
      insertRows: 'Вставку строк',
      insertHyperlinks: 'Вставку гиперссылок',
      deleteColumns: 'Удаление столбцов',
      deleteRows: 'Удаление строк',
      sort: 'Сортировку',
      autoFilter: 'Использование автофильтра',
      pivotTables: 'Использование сводной таблицы и сводной диаграммы',
      objects: 'Изменение объектов',
      scenarios: 'Изменение сценариев',
    })[box],
  fmtProtection: 'Защита',
  fmtLocked: 'Защищаемая ячейка',
  fmtHidden: 'Скрыть формулы',
  fmtProtectionNote: 'Блокирование ячеек и скрытие формул действуют только после защиты листа (вкладка "Рецензирование", группа "Изменения", кнопка "Защитить лист").',
}
