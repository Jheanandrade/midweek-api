package assigner

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"sort"
	"strings"
	"time"
)

const (
	sheetName = "Sheet1"
)

type Designated struct {
	Name            string
	LastDesignation string
}

func LoadAvailableDesignatesFromFile(f *excelize.File) (map[string][]Designated, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("error reading tab %s: %v", sheetName, err)
	}
	if len(rows) < 2 {
		return nil, fmt.Errorf("sheet without enough data")
	}

	headers := rows[0]
	designates := parseDesignatedRows(f, rows[1:], headers)

	for function, list := range designates {
		sort.SliceStable(list, func(i, j int) bool {
			return compareByDatePriority(list[i], list[j])
		})
		designates[function] = list
	}

	return designates, nil
}

func parseDesignatedRows(f *excelize.File, dataRows [][]string, headers []string) map[string][]Designated {
	result := make(map[string][]Designated)

	for rowIdx, row := range dataRows {
		if len(row) == 0 {
			continue
		}
		name := row[0]

		for colIdx := 1; colIdx < len(headers)-1; colIdx += 2 {
			function := strings.TrimSpace(headers[colIdx])
			aptCell := cellPosition(colIdx+1, rowIdx+2)
			dateCell := cellPosition(colIdx+2, rowIdx+2)

			aptVal, _ := f.GetCellValue(sheetName, aptCell)
			dateVal, _ := f.GetCellValue(sheetName, dateCell)

			if isDesignated(aptVal) {
				designated := Designated{
					Name:            name,
					LastDesignation: strings.TrimSpace(dateVal),
				}
				result[function] = append(result[function], designated)
			}
		}
	}
	return result
}

// ✅ Ordena por: vazios > datas antigas > datas inválidas
func compareByDatePriority(a, b Designated) bool {
	// Prioridade 1: sem data
	if a.LastDesignation == "" && b.LastDesignation != "" {
		return true
	}
	if b.LastDesignation == "" && a.LastDesignation != "" {
		return false
	}
	if a.LastDesignation == "" && b.LastDesignation == "" {
		return a.Name < b.Name // desempatador por nome
	}

	// Ambas têm data — compara
	dateA, errA := time.Parse("02/01/2006", a.LastDesignation)
	dateB, errB := time.Parse("02/01/2006", b.LastDesignation)

	if errA != nil && errB == nil {
		return false
	}
	if errB != nil && errA == nil {
		return true
	}
	if errA != nil && errB != nil {
		return a.Name < b.Name
	}

	return dateA.Before(dateB)
}

func isDesignated(value string) bool {
	v := strings.ToLower(strings.TrimSpace(value))
	return v == "1" || v == "true"
}

func cellPosition(colIndex int, row int) string {
	colName, _ := excelize.ColumnNumberToName(colIndex)
	return fmt.Sprintf("%s%d", colName, row)
}
