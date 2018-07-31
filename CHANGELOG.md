# Changelog

## Unreleased

### Changed 

- `skip_rows` > `rows_shift`; `skip_columns` > `columns_shift`

### Added

- `columns_shift`, `row_shift` for cells
- `ignore_height` to ignore particular cell rendered height while computing next row

## 0.3.0

### Added

- Column and row sizing via `column_widths` and `row_height`
- Number formatting via `number_format`

## 0.2.0

### Added

- New row parameters: `skip_rows`, `skip_columns` 

### Fixed

- Empty row rendering

## 0.1.2

### Fixed

- Fix cell range size is greater that given size by 1 (e.g. 2x2 range equals to A1:B2 not A1:C3)

## 0.1.1

### Fixed

- Render multiple cells with different width on same row 

## 0.1.0

### Added 

- Create xlsx from json (see README)