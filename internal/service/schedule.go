package service

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"
	"midweek-project/internal/assigner"
	"midweek-project/internal/parser"
	"midweek-project/internal/util"
	"midweek-project/internal/writer"
	"mime/multipart"
	"os"
	"path/filepath"

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

func DeleteZipFile(ctx context.Context, filename string) error {
	path := filepath.Join(zipInputPath, filename)
	return os.Remove(path)
}

func StoreZipFile(ctx context.Context, file multipart.File, filename string) error {
	if err := os.MkdirAll(zipInputPath, os.ModePerm); err != nil {
		return err
	}

	destPath := filepath.Join(zipInputPath, filename)
	destFile, err := os.Create(destPath)
	if err != nil {
		return err
	}

	defer func(destFile *os.File) {
		_ = destFile.Close()
	}(destFile)

	_, err = io.Copy(destFile, file)
	return err
}

func ProcessSchedule(designates io.Reader, period string) ([]byte, error) {
	zipFilePath, err := util.SearchZipFileByKeyword(zipInputPath, period)
	if err != nil {
		return nil, errors.New("no zip file found")
	}

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

	rtfPaths, err := util.UnzipRTFFiles(zipFilePath, zipInputPath)
	if err != nil {
		return nil, err
	}

	txtContents, err := util.ConvertRTFToTxtInMemory(rtfPaths)
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

	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return zipBuffer.Bytes(), nil
}

func writeToZip(zipWriter *zip.Writer, filename string, data []byte) {
	f, _ := zipWriter.Create(filename)
	_, _ = f.Write(data)
}
