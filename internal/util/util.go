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
	libreOfficeEnvVar     = "LIBREOFFICE_PATH"
	defaultLibreOfficeCmd = "libreoffice"
)

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

func getLibreOfficeCommand() string {
	if cmd := os.Getenv(libreOfficeEnvVar); cmd != "" {
		return cmd
	}
	return defaultLibreOfficeCmd
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

func ConvertSingleRTFToTXT(inputPath, outputPath string) error {
	libreOfficeCmd := getLibreOfficeCommand()

	cmd := exec.Command(libreOfficeCmd, "--headless", "--convert-to", "txt:Text", "--outdir", filepath.Dir(outputPath), inputPath)

	var stderr bytes.Buffer
	cmd.Stderr = &stderr

	if err := cmd.Run(); err != nil {
		return fmt.Errorf("failed to convert %s using libreoffice: %v\n%s", inputPath, err, stderr.String())
	}
	return nil
}

func ListTxtFilesForPeriod(period string) ([]string, error) {
	dir := filepath.Join("data/unzipped", period)
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, fmt.Errorf("failed to read txt directory: %w", err)
	}

	var txtPaths []string
	for _, entry := range entries {
		if !entry.IsDir() && filepath.Ext(entry.Name()) == ".txt" {
			txtPaths = append(txtPaths, filepath.Join(dir, entry.Name()))
		}
	}
	return txtPaths, nil
}

func ReadTxtFiles(txtPaths []string) ([]string, error) {
	var contents []string
	for _, path := range txtPaths {
		content, err := readAndDecodeFile(path)
		if err != nil {
			return nil, fmt.Errorf("failed to read txt file %s: %w", path, err)
		}
		contents = append(contents, content)
	}
	return contents, nil
}
