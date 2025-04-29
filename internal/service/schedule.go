package service

import (
	"archive/zip"
	"bytes"
	"context"
	"fmt"
	"io"
	"midweek-project/internal/assigner"
	"midweek-project/internal/parser"
	"midweek-project/internal/util"
	"midweek-project/internal/writer"
	"mime/multipart"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	zipInputPath = "data/unzipped"
)

func ListZipFiles(ctx context.Context) ([]string, error) {
	var files []string

	entries, err := os.ReadDir(zipInputPath)
	if err != nil {
		return nil, err
	}

	for _, entry := range entries {
		if !entry.IsDir() {
			files = append(files, entry.Name())
		}
	}

	return files, nil
}

func DeleteZipFile(filename string) error {
	path := filepath.Join(zipInputPath, filename)
	return os.Remove(path)
}

func StoreZipFile(file multipart.File, filename string) error {
	period := strings.TrimSuffix(filename, filepath.Ext(filename))

	tempZipPath := filepath.Join(os.TempDir(), filename)

	out, err := os.Create(tempZipPath)
	if err != nil {
		return fmt.Errorf("failed to create temp zip file: %w", err)
	}

	defer func(out *os.File) {
		_ = out.Close()
	}(out)

	if _, err := io.Copy(out, file); err != nil {
		return fmt.Errorf("failed to copy zip contents: %w", err)
	}

	destDir := filepath.Join(zipInputPath, period)
	if err := os.MkdirAll(destDir, os.ModePerm); err != nil {
		return fmt.Errorf("failed to create dest dir: %w", err)
	}

	rtfPaths, err := util.UnzipRTFFiles(tempZipPath, destDir)
	if err != nil {
		return fmt.Errorf("failed to unzip: %w", err)
	}

	for _, rtfPath := range rtfPaths {
		outputPath := strings.TrimSuffix(rtfPath, filepath.Ext(rtfPath)) + ".txt"

		if err := util.ConvertSingleRTFToTXT(rtfPath, outputPath); err != nil {
			return fmt.Errorf("failed to convert rtf to txt: %w", err)
		}

		_ = os.Remove(rtfPath)
	}

	_ = os.Remove(tempZipPath)

	return nil
}

func ProcessSchedule(designates io.Reader, period string) ([]byte, error) {
	var buf bytes.Buffer
	if _, err := io.Copy(&buf, designates); err != nil {
		return nil, err
	}

	excelFile, err := excelize.OpenReader(&buf)
	if err != nil {
		return nil, err
	}

	designatesPool, err := assigner.LoadAvailableDesignatesFromFile(excelFile)
	if err != nil {
		return nil, err
	}

	txtPaths, err := util.ListTxtFilesForPeriod(period)
	if err != nil {
		return nil, err
	}

	txtContents, err := util.ReadTxtFiles(txtPaths)
	if err != nil {
		return nil, err
	}

	meetings, err := parser.ParseAllMeetings(txtContents)
	if err != nil {
		return nil, err
	}

	meetingsWithDesignates, err := assigner.AssignToMeetings(meetings, designatesPool, excelFile)
	if err != nil {
		return nil, err
	}

	docContent, err := writer.GenerateDesignationsDoc(meetingsWithDesignates, period)
	if err != nil {
		return nil, err
	}

	var midweekBuffer bytes.Buffer
	if err := writer.WriteToBuffer(meetingsWithDesignates, &midweekBuffer); err != nil {
		return nil, err
	}

	var designatesBuffer bytes.Buffer
	if err := excelFile.Write(&designatesBuffer); err != nil {
		return nil, err
	}

	var zipBuffer bytes.Buffer
	zipWriter := zip.NewWriter(&zipBuffer)

	writeToZip(zipWriter, fmt.Sprintf("%s.xlsx", period), midweekBuffer.Bytes())
	writeToZip(zipWriter, "designates.xlsx", designatesBuffer.Bytes())
	writeToZip(zipWriter, fmt.Sprintf("%s.docx", period), docContent)

	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return zipBuffer.Bytes(), nil
}

func writeToZip(zipWriter *zip.Writer, filename string, data []byte) {
	f, _ := zipWriter.Create(filename)
	_, _ = f.Write(data)
}
