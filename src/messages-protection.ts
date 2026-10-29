// The texts of sheet protection, English and Russian, as Excel for Mac 16.114 words them (its strman tables and its Protect
// Sheet and Format Cells dialogs); src/messages.ts includes them in its sets.
import type { ProtectionMessages } from './contract/messages'

export const enProtection: ProtectionMessages = {
  protectedAlert: "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.",
  menuProtectSheet: 'Protect Sheet…',
  menuUnprotectSheet: 'Unprotect Sheet',
  protectTitle: 'Protect Sheet',
  protectContents: 'Protect the sheet and contents of locked cells.',
  protectAllow: 'Allow users of this sheet to:',
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
  protectNote: 'All cells are locked by default, but can be formatted as unlocked.',
  fmtProtection: 'Protection',
  fmtLocked: 'Locked',
  fmtHidden: 'Hidden',
  fmtProtectionNote: 'To lock cells or hide formulas, protect the worksheet. On the Review tab, click Protect Sheet.',
}

export const ruProtection: ProtectionMessages = {
  protectedAlert: 'Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль.',
  menuProtectSheet: 'Защитить лист…',
  menuUnprotectSheet: 'Снять защиту листа',
  protectTitle: 'Защитить лист',
  protectContents: 'Защита листа и содержимого заблокированных ячеек.',
  protectAllow: 'Разрешить всем пользователям этого листа:',
  protectBox: (box) =>
    ({
      selectLockedCells: 'Выделить заблокированные ячейки',
      selectUnlockedCells: 'Выделить незаблокированные ячейки',
      formatCells: 'Форматировать ячейки',
      formatColumns: 'Форматировать столбцы',
      formatRows: 'Форматировать строки',
      insertColumns: 'Вставить столбцы',
      insertRows: 'Вставить строки',
      insertHyperlinks: 'Вставить гиперссылки',
      deleteColumns: 'Удалить столбцы',
      deleteRows: 'Удалить строки',
      sort: 'Сортировать',
      autoFilter: 'Использование автофильтра',
      pivotTables: 'Использовать сводную таблицу и сводную диаграмму',
      objects: 'Изменить объекты',
      scenarios: 'Изменить сценарии',
    })[box],
  protectNote: 'Все ячейки по умолчанию заблокированы, однако их можно пометить как незаблокированные.',
  fmtProtection: 'Защита',
  fmtLocked: 'Заблокировать ячейку',
  fmtHidden: 'Скрыть формулу',
  fmtProtectionNote: 'Чтобы заблокировать ячейки или скрыть формулы, защитите лист. Для этого на вкладке "Рецензирование" выберите команду "Защитить лист".',
}
