package esb

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func mergeCell(f *excelize.File, sheet string, c1, r1, c2, r2 int) error {
	topLeftCell, err := excelize.CoordinatesToCellName(c1, r1)
	if err != nil {
		return err
	}

	bottomRightCell, err := excelize.CoordinatesToCellName(c2, r2)
	if err != nil {
		return err
	}

	return f.MergeCell(sheet, topLeftCell, bottomRightCell)
}

func renderHeaders(f *excelize.File, sheet string, cols Columns, maxDepth int) error {
	currentCol := 1

	var render func(Columns, int) error
	render = func(cols Columns, row int) error {
		for _, col := range cols {
			startCol := currentCol

			cell, _ := excelize.CoordinatesToCellName(currentCol, row)
			if err := f.SetCellValue(sheet, cell, col.Title); err != nil {
				return err
			}

			if len(col.NestedCols) == 0 {
				if row < maxDepth {
					if err := mergeCell(f, sheet, startCol, row, startCol, maxDepth); err != nil {
						return err
					}
				}
				currentCol++
			} else {
				span := col.CountLeaves()

				if span > 1 {
					if err := mergeCell(f, sheet, startCol, row, startCol+span-1, row); err != nil {
						return err
					}
				}

				if err := render(col.NestedCols, row+1); err != nil {
					return err
				}

				currentCol = startCol + span
			}
		}

		return nil
	}

	return render(cols, 1)
}

func renderData(f *excelize.File, sheet string, cols []Column, row any, rowIdx int) error {
	currentCol := 1

	var render func([]Column, any, int) error
	render = func(cols []Column, row any, rowIdx int) error {
		for _, c := range cols {
			val := c.Value(row)

			if len(c.NestedCols) == 0 {
				cell, _ := excelize.CoordinatesToCellName(currentCol, rowIdx)

				if err := f.SetCellValue(sheet, cell, val); err != nil {
					return fmt.Errorf("set data cell: %w", err)
				}

				currentCol++
			} else {
				if err := render(c.NestedCols, val, rowIdx); err != nil {
					return err
				}
			}
		}

		return nil
	}

	return render(cols, row, rowIdx)
}

func (r *Report[T]) applyStyle(f *excelize.File, sheet string, fn StyleFunc, startRow, endRow int) error {
	leafCount := 0
	for _, c := range r.columns {
		leafCount += c.CountLeaves()
	}

	if leafCount == 0 {
		return nil
	}

	style := excelize.Style{}

	if fn != nil {
		fn(&style)
	}

	styleID, err := f.NewStyle(&style)
	if err != nil {
		return err
	}

	start, _ := excelize.CoordinatesToCellName(1, startRow)
	end, _ := excelize.CoordinatesToCellName(leafCount, endRow)

	return f.SetCellStyle(sheet, start, end, styleID)
}

func (r *Report[T]) WriteTo(f *excelize.File, sheetName string, rows []T) (*excelize.File, error) {
	sheetIndex, err := f.NewSheet(sheetName)
	if err != nil {
		return f, err
	}

	f.SetActiveSheet(sheetIndex)

	maxDepth := Columns(r.columns).Depth()
	if err = renderHeaders(f, sheetName, r.columns, maxDepth); err != nil {
		return f, err
	}

	if err = r.applyStyle(f, sheetName, r.styles.headerStyle, 1, maxDepth); err != nil {
		return f, err
	}

	dataStart := maxDepth + 1
	for idx, row := range rows {
		if err = renderData(f, sheetName, r.columns, any(row), dataStart+idx); err != nil {
			return f, err
		}
	}

	if len(rows) > 0 {
		if err = r.applyStyle(f, sheetName, r.styles.bodyStyle, dataStart, dataStart+len(rows)-1); err != nil {
			return f, err
		}
	}

	return f, nil
}
