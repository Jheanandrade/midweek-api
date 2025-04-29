package assigner

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"midweek-project/internal/parser"
	"regexp"
	"sort"
	"strings"
)

const (
	FUNC_PRESIDENTE          = "Presidente"
	FUNC_CONSELHEIRO         = "Conselheiro Sala B"
	FUNC_ORACAO              = "Oração"
	FUNC_ORACAO_FINAL        = "OraçãoFinal"
	FUNC_LEITOR_BIBLIA_A     = "Leitor - Leitura da Bíblia - A"
	FUNC_LEITOR_BIBLIA_B     = "Leitor - Leitura da Bíblia - B"
	FUNC_DISCURSO_TESOUROS   = "Discurso - Tesouros da Palavra de Deus"
	FUNC_JOIAS               = "Joías Espirituais - Tesouros da Palavra de Deus"
	FUNC_DISCURSO_MINISTERIO = "Discursos -  Faça Seu Melhor no Ministério"
	FUNC_DISCURSO_CRISTA     = "Discursos - Nossa Vida Cristã"
	FUNC_ESTUDO_BIBLICO      = "Estudo Bíblico - Nossa Vida Cristã"
	FUNC_LEITOR_ESTUDO       = "Leitor - Estudo Biblíco"
	FUNC_TITULAR_A_HOMEM     = "Titular - A (Homem)"
	FUNC_AJUDANTE_A_HOMEM    = "Ajudante - A (Homem)"
	FUNC_TITULAR_A_MULHER    = "Titular - A (Mulher)"
	FUNC_AJUDANTE_A_MULHER   = "Ajudante - A (Mulher)"
	FUNC_TITULAR_B_HOMEM     = "Titular - B (Homem)"
	FUNC_AJUDANTE_B_HOMEM    = "Ajudante - B (Homem)"
	FUNC_TITULAR_B_MULHER    = "Titular - B (Mulher)"
	FUNC_AJUDANTE_B_MULHER   = "Ajudante - B (Mulher)"
)

func AssignToMeetings(meetings []parser.MeetingData, pool map[string][]Designated, f *excelize.File) ([]parser.MeetingData, error) {
	if len(meetings) == 0 {
		return nil, fmt.Errorf("meeting list is empty")
	}
	if len(pool) == 0 {
		return nil, fmt.Errorf("designation pool is empty")
	}

	for i, meeting := range meetings {
		used := map[string]bool{}
		designated := make(map[string]string)
		date := meeting.MeetingDate

		assignTreasures(meeting, designated, pool, used, f, date, true)
		assignMinistry(meeting, designated, pool, used, f, date, true)
		assignChristians(meeting, designated, pool, used, f, date, true)

		assignFunction(FUNC_PRESIDENTE, designated, pool, used, f, date, true)
		assignFunction(FUNC_CONSELHEIRO, designated, pool, used, f, date, true)

		initPrayer := pickUniqueExcluding(FUNC_ORACAO, pool, f, date, used, "", false)
		designated[FUNC_ORACAO] = initPrayer
		_ = updateDesignationDate(f, FUNC_ORACAO, initPrayer, date)

		finalPrayer := pickUniqueExcluding(FUNC_ORACAO, pool, f, date, used, initPrayer, false)
		designated[FUNC_ORACAO_FINAL] = finalPrayer
		_ = updateDesignationDate(f, FUNC_ORACAO_FINAL, finalPrayer, date)

		meetings[i].Designated = designated
	}
	return meetings, nil
}

func assignFunction(function string, dest map[string]string, pool map[string][]Designated, used map[string]bool, f *excelize.File, date string, exclusive bool) {
	dest[function] = pickUniqueAndRotate(function, pool, f, date, used, exclusive)
}

func assignTreasures(m parser.MeetingData, dest map[string]string, pool map[string][]Designated, used map[string]bool, f *excelize.File, date string, exclusive bool) {
	for _, key := range getSortedKeys(m.TreasuresFromGodsWord) {
		text := strings.ToLower(m.TreasuresFromGodsWord[key])

		switch {
		case strings.Contains(text, "leitura da bíblia"):
			assignedA := pickUniqueExcluding(FUNC_LEITOR_BIBLIA_A, pool, f, date, used, "", exclusive)
			assignedB := pickUniqueExcluding(FUNC_LEITOR_BIBLIA_B, pool, f, date, used, assignedA, exclusive)
			dest[key+".A"] = assignedA
			dest[key+".B"] = assignedB

		case strings.Contains(text, "joias espirituais"):
			dest[key] = pickUniqueAndRotate(FUNC_JOIAS, pool, f, date, used, exclusive)

		default:
			dest[key] = pickUniqueAndRotate(FUNC_DISCURSO_TESOUROS, pool, f, date, used, exclusive)
		}
	}
}

func assignMinistry(meeting parser.MeetingData, dest map[string]string, pool map[string][]Designated, used map[string]bool, f *excelize.File, date string, exclusive bool) {
	keys := getSortedKeys(meeting.ApplyYourselfToTheFieldMinistry)
	total := len(keys)
	maleSlots := 1
	femaleSlots := total - maleSlots

	discourseKeys := []string{}
	nonDiscourseKeys := []string{}

	for _, key := range keys {
		text := strings.ToLower(meeting.ApplyYourselfToTheFieldMinistry[key])
		switch {
		case strings.Contains(text, "discurso"):
			discourseKeys = append(discourseKeys, key)

		default:
			nonDiscourseKeys = append(nonDiscourseKeys, key)
		}
	}

	for _, key := range discourseKeys {
		assignedA := pickUniqueExcluding(FUNC_DISCURSO_MINISTERIO, pool, f, date, used, "", exclusive)
		assignedB := pickUniqueExcluding(FUNC_DISCURSO_MINISTERIO, pool, f, date, used, assignedA, exclusive)
		dest[key+".A"] = assignedA
		dest[key+".B"] = assignedB
	}

	maleSlots -= len(discourseKeys)
	if maleSlots < 0 {
		maleSlots = 0
	}
	femaleSlots = total - len(discourseKeys) - maleSlots

	for i, key := range nonDiscourseKeys {
		if i < femaleSlots {
			holderA := pickUniqueExcluding(FUNC_TITULAR_A_MULHER, pool, f, date, used, "", exclusive)
			helperA := pickUniqueExcluding(FUNC_AJUDANTE_A_MULHER, pool, f, date, used, holderA, exclusive)
			holderB := pickUniqueExcluding(FUNC_TITULAR_B_MULHER, pool, f, date, used, "", exclusive)
			helperB := pickUniqueExcluding(FUNC_AJUDANTE_B_MULHER, pool, f, date, used, holderB, exclusive)
			dest[key+".A"] = fmt.Sprintf("%s/%s", holderA, helperA)
			dest[key+".B"] = fmt.Sprintf("%s/%s", holderB, helperB)
		} else {
			holderA := pickUniqueExcluding(FUNC_TITULAR_A_HOMEM, pool, f, date, used, "", exclusive)
			helperA := pickUniqueExcluding(FUNC_AJUDANTE_A_HOMEM, pool, f, date, used, holderA, exclusive)
			holderB := pickUniqueExcluding(FUNC_TITULAR_B_HOMEM, pool, f, date, used, "", exclusive)
			helperB := pickUniqueExcluding(FUNC_AJUDANTE_B_HOMEM, pool, f, date, used, holderB, exclusive)
			dest[key+".A"] = fmt.Sprintf("%s/%s", holderA, helperA)
			dest[key+".B"] = fmt.Sprintf("%s/%s", holderB, helperB)
		}
	}
}

func assignChristians(m parser.MeetingData, dest map[string]string, pool map[string][]Designated, used map[string]bool, f *excelize.File, date string, exclusive bool) {
	for _, key := range getSortedKeys(m.LivingAsChristians) {
		text := strings.ToLower(m.LivingAsChristians[key])

		switch {
		case strings.Contains(text, "estudo bíblico de congregação"):
			leader := pickUniqueAndRotate(FUNC_ESTUDO_BIBLICO, pool, f, date, used, exclusive)
			reader := pickUniqueExcluding(FUNC_LEITOR_ESTUDO, pool, f, date, used, leader, exclusive)
			dest[key] = fmt.Sprintf("%s/%s", leader, reader)

		default:
			dest[key] = pickUniqueAndRotate(FUNC_DISCURSO_CRISTA, pool, f, date, used, exclusive)

		}
	}
}

func pickUniqueAndRotate(role string, pool map[string][]Designated, f *excelize.File, meeting string, used map[string]bool, exclusive bool) string {
	return pickUniqueExcluding(role, pool, f, meeting, used, "", exclusive)
}

func pickUniqueExcluding(role string, pool map[string][]Designated, f *excelize.File, meeting string, used map[string]bool, exclude string, exclusive bool) string {
	list := pool[role]
	for i := 0; i < len(list); i++ {
		name := list[i].Name
		if (!exclusive || !used[name]) && name != exclude {
			pool[role] = append(list[i+1:], list[:i+1]...)
			if exclusive {
				used[name] = true
			}
			_ = updateDesignationDate(f, role, name, meeting)
			return name
		}
	}
	if len(list) > 0 {
		name := list[0].Name
		pool[role] = append(list[1:], list[0])
		if exclusive {
			used[name] = true
		}
		_ = updateDesignationDate(f, role, name, meeting)
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

func updateDesignationDate(f *excelize.File, role string, name string, meetingDate string) error {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}

	headers := rows[0]
	publicadoresIdx := -1
	roleColIdx := -1

	for i, h := range headers {
		h = strings.TrimSpace(h)
		if strings.EqualFold(h, "Publicadores") {
			publicadoresIdx = i
		}
		if h == role {
			roleColIdx = i
		}
	}
	if publicadoresIdx == -1 {
		return fmt.Errorf("'Publishers' column not found")
	}
	if roleColIdx == -1 || roleColIdx+1 >= len(headers) {
		return fmt.Errorf("column for %s not found", role)
	}

	date := extractLastDateFromMeeting(meetingDate)
	if date == "" {
		return fmt.Errorf("invalid date for meeting: %s", meetingDate)
	}

	for i, row := range rows {
		if i == 0 || len(row) == 0 {
			continue
		}
		if publicadoresIdx >= len(row) {
			continue
		}
		if strings.TrimSpace(row[publicadoresIdx]) == name {
			colName, _ := excelize.ColumnNumberToName(roleColIdx + 2)
			cell := fmt.Sprintf("%s%d", colName, i+1)
			return f.SetCellValue(sheetName, cell, date)
		}
	}

	return fmt.Errorf("designated %s not found", name)
}

func extractLastDateFromMeeting(meeting string) string {
	re := regexp.MustCompile(`(?i)a\s+(\d{1,2})\s+de\s+([a-zç]+)`)
	match := re.FindStringSubmatch(meeting)
	if len(match) < 3 {
		return ""
	}
	day := match[1]
	month := strings.ToLower(match[2])
	months := map[string]string{
		"janeiro": "01", "fevereiro": "02", "março": "03", "marco": "03", "abril": "04",
		"maio": "05", "junho": "06", "julho": "07", "agosto": "08", "setembro": "09",
		"outubro": "10", "novembro": "11", "dezembro": "12",
	}
	monthNum, ok := months[month]
	if !ok {
		return ""
	}
	return fmt.Sprintf("%02s/%s/2025", day, monthNum)
}
