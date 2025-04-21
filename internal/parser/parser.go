package parser

import (
	"bufio"
	"fmt"
	"midweek-project/internal/util"
	"regexp"
	"strings"
)

type Section map[string]string

type MeetingData struct {
	MeetingDate                     string
	InitSong                        string
	MidSong                         string
	FinalSong                       string
	TreasuresFromGodsWord           Section
	ApplyYourselfToTheFieldMinistry Section
	LivingAsChristians              Section
	Designated                      map[string]string
}

const (
	SectionTreasures = "TESOUROS DA PALAVRA DE DEUS"
	SectionMinistry  = "FAÇA SEU MELHOR NO MINISTÉRIO"
	SectionChristian = "NOSSA VIDA CRISTÃ"
)

var (
	reTopic = regexp.MustCompile(`(?m)^\s*(\d{1,3})[.\xEF\xBC\x8E]?\s+(.+?)\s*\(\s*(\d{1,3})\s*(?:minutos|min)\s*\)`)
	reSong  = regexp.MustCompile(`(?i)C[âa]ntico[\s\xA0]+(\d+)`)
	reDate  = regexp.MustCompile(`(?i)(\d{1,2})(?:\s+de\s+([a-zç]+))?\s+a\s+(\d{1,2})(?:º?\.?|\.º)?\s+de\s+([a-zç]+)`)
)

func ParseAllMeetings(contents []string) ([]MeetingData, error) {
	var meetings []MeetingData
	for _, content := range contents {
		meeting := parseTxtMeeting(content)
		if meeting.MeetingDate != "" {
			meetings = append(meetings, meeting)
		}
	}
	return meetings, nil
}

func parseTxtMeeting(content string) MeetingData {
	meeting := MeetingData{
		TreasuresFromGodsWord:           make(Section),
		ApplyYourselfToTheFieldMinistry: make(Section),
		LivingAsChristians:              make(Section),
		Designated:                      make(map[string]string),
	}

	var currentSection string
	scanner := bufio.NewScanner(strings.NewReader(content))
	for scanner.Scan() {
		line := util.NormalizeLine(scanner.Text())

		if meeting.MeetingDate == "" {
			meeting.MeetingDate = extractDateFromLine(line)
			if meeting.MeetingDate != "" {
				continue
			}
		}

		if sec := detectSection(line); sec != "" {
			currentSection = sec
			continue
		}

		assignSongIfApplicable(&meeting, line)
		assignTopicIfApplicable(&meeting, currentSection, line)
	}

	return meeting
}

func extractDateFromLine(line string) string {
	match := reDate.FindStringSubmatch(line)
	if len(match) > 0 {
		if match[2] != "" {
			return fmt.Sprintf("%s de %s a %s de %s", match[1], match[2], match[3], match[4])
		}
		return fmt.Sprintf("%s a %s de %s", match[1], match[3], match[4])
	}
	return ""
}

func detectSection(line string) string {
	switch {
	case strings.EqualFold(line, "Tesouros da Palavra de Deus"):
		return SectionTreasures
	case strings.EqualFold(line, "FAÇA SEU MELHOR NO MINISTÉRIO"):
		return SectionMinistry
	case strings.EqualFold(line, "Nossa vida cristã"):
		return SectionChristian
	default:
		return ""
	}
}

func assignSongIfApplicable(meeting *MeetingData, line string) {
	if !reSong.MatchString(line) {
		return
	}
	switch {
	case meeting.InitSong == "":
		meeting.InitSong = line
	case meeting.MidSong == "":
		meeting.MidSong = line
	case meeting.FinalSong == "":
		meeting.FinalSong = line
	}
}

func assignTopicIfApplicable(meeting *MeetingData, section, line string) {
	if section == "" || !reTopic.MatchString(line) {
		return
	}
	match := reTopic.FindStringSubmatch(line)
	if len(match) == 0 {
		return
	}
	topicNum := match[1]
	topicText := match[0]

	switch section {
	case SectionTreasures:
		meeting.TreasuresFromGodsWord[topicNum] = topicText
	case SectionMinistry:
		meeting.ApplyYourselfToTheFieldMinistry[topicNum] = topicText
	case SectionChristian:
		meeting.LivingAsChristians[topicNum] = topicText
	}
}
