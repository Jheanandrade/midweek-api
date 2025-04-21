package util

import (
	"archive/zip"
	"bytes"
	"fmt"
	"github.com/saintfish/chardet"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/transform"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

const (
	rtfExtension          = ".rtf"
	txtExtension          = ".txt"
	libreOfficeEnvVar     = "LIBREOFFICE_PATH"
	defaultLibreOfficeCmd = "libreoffice"
)

func SearchZipFileByKeyword(directory, keyword string) (string, error) {
	var foundPath string

	err := filepath.Walk(directory, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info == nil || info.IsDir() {
			return nil
		}

		if isMatchingZipFile(path, keyword) {
			foundPath = path
		}
		return nil
	})

	if err != nil {
		return "", fmt.Errorf("error while traversing directory: %w", err)
	}

	return foundPath, nil
}

func isMatchingZipFile(path, keyword string) bool {
	pathLower := strings.ToLower(path)
	keywordLower := strings.ToLower(keyword)
	return strings.HasSuffix(pathLower, ".zip") && strings.Contains(pathLower, keywordLower)
}

func UnzipRTFFiles(pathUnzipped, tmp string) ([]string, error) {
	var rtfPaths []string
	reader, err := zip.OpenReader(pathUnzipped)
	if err != nil {
		return nil, fmt.Errorf("unable to unzip: %w", err)
	}

	defer func(reader *zip.ReadCloser) {
		_ = reader.Close()
	}(reader)

	for _, file := range reader.File {
		if filepath.Ext(file.Name) != rtfExtension {
			continue
		}
		fullPath := filepath.Join(tmp, file.Name)
		outFile, err := os.Create(fullPath)
		if err != nil {
			return nil, err
		}
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		_, _ = io.Copy(outFile, rc)
		_ = outFile.Close()
		_ = rc.Close()
		rtfPaths = append(rtfPaths, fullPath)
	}
	return rtfPaths, nil
}

func NormalizeLine(line string) string {
	line = strings.TrimSpace(line)
	return strings.ReplaceAll(line, "\u00A0", " ")
}

func ConvertRTFToTxtInMemory(rtfPaths []string) ([]string, error) {
	var txtContents []string
	libreOfficeCmd := getLibreOfficeCommand()

	for _, path := range rtfPaths {
		outputPath := buildOutputPath(path)

		if err := runLibreOfficeConversion(libreOfficeCmd, path, outputPath); err != nil {
			return nil, err
		}

		content, err := readAndDecodeFile(outputPath)
		if err != nil {
			return nil, err
		}

		txtContents = append(txtContents, content)

		_ = os.Remove(outputPath)
	}

	return txtContents, nil
}

func getLibreOfficeCommand() string {
	if cmd := os.Getenv(libreOfficeEnvVar); cmd != "" {
		return cmd
	}
	return defaultLibreOfficeCmd
}

func buildOutputPath(inputPath string) string {
	baseName := filepath.Base(inputPath[:len(inputPath)-len(filepath.Ext(inputPath))]) + txtExtension
	return filepath.Join(os.TempDir(), baseName)
}

func runLibreOfficeConversion(cmdPath, inputPath, outputPath string) error {
	cmd := exec.Command(cmdPath, "--headless", "--convert-to", "txt:Text", "--outdir", os.TempDir(), inputPath)

	var stderr bytes.Buffer
	cmd.Stderr = &stderr

	if err := cmd.Run(); err != nil {
		return fmt.Errorf("failed to convert %s using libreoffice: %v\n%s", inputPath, err, stderr.String())
	}
	return nil
}

func readAndDecodeFile(filePath string) (string, error) {
	rawContent, err := os.ReadFile(filePath)
	if err != nil {
		return "", fmt.Errorf("failed to read converted file %s: %w", filePath, err)
	}

	encoding, err := detectEncoding(rawContent)
	if err != nil {
		return "", err
	}

	switch encoding {
	case "windows-1252":
		utf8Reader := transform.NewReader(bytes.NewReader(rawContent), charmap.Windows1252.NewDecoder())
		decoded, err := io.ReadAll(utf8Reader)
		if err != nil {
			return "", fmt.Errorf("failed to decode windows-1252: %w", err)
		}
		return string(decoded), nil
	default:
		return string(rawContent), nil
	}
}

func detectEncoding(data []byte) (string, error) {
	detector := chardet.NewTextDetector()
	result, err := detector.DetectBest(data)
	if err != nil {
		return "", fmt.Errorf("failed to detect encoding: %w", err)
	}
	return result.Charset, nil
}
