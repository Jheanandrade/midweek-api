package handler

import (
	"github.com/labstack/echo/v4"
	"midweek-project/internal/service"
	"mime/multipart"
	"net/http"
)

func ListZipFiles(c echo.Context) error {
	files, err := service.ListZipFiles(c.Request().Context())
	if err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{
			"error": err.Error(),
		})
	}

	return c.JSON(http.StatusOK, files)
}

func DeleteZipFile(c echo.Context) error {
	filename := c.QueryParam("filename")
	if filename == "" {
		return c.JSON(http.StatusBadRequest, map[string]string{
			"error": "Missing filename",
		})
	}

	err := service.DeleteZipFile(c.Request().Context(), filename)
	if err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{
			"error": err.Error(),
		})
	}

	return c.NoContent(http.StatusOK)
}

func GenerateSchedule(c echo.Context) error {
	fileHeader, err := c.FormFile("designates")
	if err != nil {
		return c.JSON(http.StatusBadRequest, map[string]string{
			"error": "Missing designates file",
		})
	}

	src, err := fileHeader.Open()
	if err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{
			"error": "Unable to open uploaded file",
		})
	}

	defer func(src multipart.File) {
		_ = src.Close()
	}(src)

	period := c.FormValue("period")
	if period == "" {
		return c.JSON(http.StatusBadRequest, map[string]string{
			"error": "Missing period {{period}}",
		})
	}

	zipBytes, err := service.ProcessSchedule(src, period)
	if err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{
			"error": err.Error(),
		})
	}

	return c.Blob(http.StatusOK, "application/zip", zipBytes)
}

func HandleUploadZip(c echo.Context) error {
	fileHeader, err := c.FormFile("zip")
	if err != nil {
		return c.JSON(http.StatusBadRequest, map[string]string{"error": "ZIP file is required"})
	}

	file, err := fileHeader.Open()
	if err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{"error": "Failed to open file"})
	}

	defer func(file multipart.File) {
		_ = file.Close()
	}(file)

	zipFilename := fileHeader.Filename
	if err := service.StoreZipFile(c.Request().Context(), file, zipFilename); err != nil {
		return c.JSON(http.StatusInternalServerError, map[string]string{"error": err.Error()})
	}

	return c.JSON(http.StatusOK, map[string]string{"message": "ZIP file uploaded successfully"})
}
