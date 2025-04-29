package writer

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"midweek-project/internal/parser"
	"sort"
	"strings"
)

const (
	sectionTreasures  = "TESOUROS DA PALAVRA DE DEUS"
	sectionMinistry   = "FAÇA SEU MELHOR NO MINISTÉRIO"
	sectionChristians = "NOSSA VIDA CRISTÃ"
)

func WriteToBuffer(meetings []parser.MeetingData, out io.Writer) error {
	f := excelize.NewFile()
	styles := createStyles(f)

	first := true
	for _, meeting := range meetings {
		if meeting.MeetingDate == "" {
			continue
		}
		sheet := meeting.MeetingDate
		if first {
			_ = f.SetSheetName("Sheet1", sheet)
			first = false
		} else {
			_, _ = f.NewSheet(sheet)
		}

		prepareSheetLayout(f, sheet)
		row := writeHeader(f, sheet, meeting, styles)

		sections := []struct {
			Name  string
			Items map[string]string
		}{
			{sectionTreasures, meeting.TreasuresFromGodsWord},
			{sectionMinistry, meeting.ApplyYourselfToTheFieldMinistry},
			{sectionChristians, meeting.LivingAsChristians},
		}

		for _, sec := range sections {
			row = writeSection(f, sheet, row, sec.Name, sec.Items, meeting, styles)
		}

		writeFooter(f, sheet, row, meeting, styles)
	}

	return f.Write(out)
}

func createStyles(f *excelize.File) map[string]int {
	gray, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#FFFFFF", Bold: true, Family: "Calibri", Size: 12},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#575a5d"}, Pattern: 1},
	})
	orange, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#FFFFFF", Bold: true, Family: "Calibri", Size: 12},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#be8900"}, Pattern: 1},
	})
	wine, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#FFFFFF", Bold: true, Family: "Calibri", Size: 12},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#7e0024"}, Pattern: 1},
	})
	content, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Family: "Calibri", Size: 12},
	})
	bold, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Family: "Calibri", Size: 12, Bold: true},
	})
	small, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Family: "Calibri", Size: 8},
	})

	return map[string]int{
		"gray":    gray,
		"orange":  orange,
		"wine":    wine,
		"content": content,
		"bold":    bold,
		"small":   small,
	}
}

func prepareSheetLayout(f *excelize.File, sheet string) {
	_ = f.SetColWidth(sheet, "A", "A", 70)
	_ = f.SetColWidth(sheet, "B", "B", 15)
	_ = f.SetColWidth(sheet, "C", "C", 15)
}

func writeHeader(f *excelize.File, sheet string, m parser.MeetingData, s map[string]int) int {
	row := 1
	setStyledCell(f, sheet, row, "A", "CONGREGAÇÃO VILA CABRAL", s["bold"], true, true)
	row += 2
	setStyledCell(f, sheet, row, "A", "Semana: "+m.MeetingDate, s["bold"], false, false)
	setStyledCell(f, sheet, row, "C", "Presidente: "+getDesignated(m, "Presidente"), s["small"], false, false)
	row++
	setStyledCell(f, sheet, row, "C", "Conselheiro sala B: "+getDesignated(m, "Conselheiro Sala B"), s["small"], false, false)
	row++
	setStyledCell(f, sheet, row, "A", "Cântico Inicial: "+m.InitSong, s["content"], false, false)
	setStyledCell(f, sheet, row, "C", "Oração: "+getDesignated(m, "Oração"), s["small"], false, false)
	row++
	setStyledCell(f, sheet, row, "A", "Comentários iniciais (1 min)", s["content"], false, false)
	return row + 1
}

func writeSection(f *excelize.File, sheet string, row int, name string, items map[string]string, m parser.MeetingData, s map[string]int) int {
	sectionStyle := map[string]int{
		sectionTreasures:  s["gray"],
		sectionMinistry:   s["orange"],
		sectionChristians: s["wine"],
	}

	setStyledCell(f, sheet, row, "A", name, sectionStyle[name], false, false)

	if name == sectionMinistry {
		setStyledCell(f, sheet, row, "B", "Sala B", s["small"], false, false)
		setStyledCell(f, sheet, row, "C", "Salão principal", s["small"], false, false)
	}
	row++

	if name == sectionChristians {
		setStyledCell(f, sheet, row, "A", "Cântico: "+m.MidSong, s["content"], false, false)
		row++
	}

	keys := getSortedKeys(items)
	for _, k := range keys {
		text := items[k]
		setStyledCell(f, sheet, row, "A", text, s["content"], false, false)

		lowerText := strings.ToLower(text)

		if name == sectionMinistry {
			valA := getDesignated(m, k+".A")
			valB := getDesignated(m, k+".B")
			if valA != "" {
				setStyledCell(f, sheet, row, "C", valA, s["small"], false, false)
			}
			if valB != "" {
				setStyledCell(f, sheet, row, "B", valB, s["small"], false, false)
			}
		} else if name == sectionTreasures && strings.Contains(lowerText, "leitura da bíblia") {
			// Caso especial: Leitura da Bíblia → dois leitores
			valA := getDesignated(m, k+".A")
			valB := getDesignated(m, k+".B")
			if valA != "" {
				setStyledCell(f, sheet, row, "C", valA, s["small"], false, false)
			}
			if valB != "" {
				setStyledCell(f, sheet, row, "B", valB, s["small"], false, false)
			}
		} else {
			// Caso padrão
			if strings.Contains(lowerText, "estudo bíblico de congregação") {
				setStyledCell(f, sheet, row, "B", "Dirigente/Leitor", s["small"], false, false)
			}
			if d := getDesignated(m, k); d != "" {
				setStyledCell(f, sheet, row, "C", d, s["small"], false, false)
			}
		}

		row++
	}
	return row + 1
}

func writeFooter(f *excelize.File, sheet string, row int, m parser.MeetingData, s map[string]int) {
	setStyledCell(f, sheet, row, "A", "Comentários finais (3 min)", s["content"], false, false)
	row++
	setStyledCell(f, sheet, row, "A", "Cântico Final: "+m.FinalSong, s["content"], false, false)
	setStyledCell(f, sheet, row, "C", "Oração: "+getDesignated(m, "OraçãoFinal"), s["small"], false, false)
}

func setStyledCell(f *excelize.File, sheet string, row int, col string, value string, style int, center bool, upper bool) {
	if upper {
		value = strings.ToUpper(value)
	}
	cell := fmt.Sprintf("%s%d", col, row)
	_ = f.SetCellValue(sheet, cell, value)
	_ = f.SetRowHeight(sheet, row, 24)
	if center {
		centerStyle, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
			Font: &excelize.Font{
				Bold:   true,
				Size:   12,
				Color:  "#000000",
				Family: "Calibri",
			},
		})
		_ = f.SetCellStyle(sheet, cell, cell, centerStyle)
	} else {
		_ = f.SetCellStyle(sheet, cell, cell, style)
	}
}

func getDesignated(m parser.MeetingData, key string) string {
	if name, ok := m.Designated[key]; ok {
		return name
	}
	return ""
}

func getSortedKeys(m map[string]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	return keys
}

func GenerateDesignationsDoc(meetings []parser.MeetingData, period string) ([]byte, error) {
	var builder strings.Builder

	for _, meeting := range meetings {
		date := meeting.MeetingDate

		keysTreasures := getSortedKeysDoc(meeting.TreasuresFromGodsWord)
		for _, key := range keysTreasures {
			partText := strings.ToLower(meeting.TreasuresFromGodsWord[key])

			if strings.Contains(partText, "leitura da bíblia") {
				designationA := meeting.Designated[key+".A"]
				if designationA != "" {
					builder.WriteString(buildDesignationBlock(designationA, date, key, "Salão principal"))
				}
				designationB := meeting.Designated[key+".B"]
				if designationB != "" {
					builder.WriteString(buildDesignationBlock(designationB, date, key, "Sala B"))
				}
			}
		}

		keysMinistry := getSortedKeys(meeting.ApplyYourselfToTheFieldMinistry)
		for _, key := range keysMinistry {
			designationA := meeting.Designated[key+".A"]
			if designationA != "" {
				builder.WriteString(buildDesignationBlock(designationA, date, key, "Salão principal"))
			}
			designationB := meeting.Designated[key+".B"]
			if designationB != "" {
				builder.WriteString(buildDesignationBlock(designationB, date, key, "Sala B"))
			}
		}
	}

	return []byte(builder.String()), nil
}

func buildDesignationBlock(designation, date string, partNumber string, location string) string {
	studentName := designation
	helperName := ""
	if strings.Contains(designation, "/") {
		parts := strings.SplitN(designation, "/", 2)
		studentName = strings.TrimSpace(parts[0])
		helperName = strings.TrimSpace(parts[1])
	}

	var sb strings.Builder
	sb.WriteString("DESIGNAÇÃO PARA A REUNIÃO\n")
	sb.WriteString("NOSSA VIDA E MINISTÉRIO CRISTÃO\n\n")
	sb.WriteString(fmt.Sprintf("Nome: %s\n", studentName))
	if helperName != "" {
		sb.WriteString(fmt.Sprintf("Ajudante: %s\n", helperName))
	}
	sb.WriteString(fmt.Sprintf("Data: %s\n", date))
	sb.WriteString(fmt.Sprintf("Número da parte: %s\n", partNumber))
	sb.WriteString(fmt.Sprintf("Local: %s\n", location))
	sb.WriteString("\n")
	sb.WriteString("Observação para o estudante: A lição e a fonte\n")
	sb.WriteString("de matéria para a sua designação estão na Apostila\n")
	sb.WriteString("da Reunião Vida e Ministério. Veja as instruções para a\n")
	sb.WriteString("parte que estão nas Instruções para a Reunião Nossa\n")
	sb.WriteString("Vida e Ministério Cristão (S-38).\n")
	sb.WriteString("S-89-T 11/23\n")
	sb.WriteString("\n\n")

	return sb.String()
}

func getSortedKeysDoc(m map[string]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	return keys
}
