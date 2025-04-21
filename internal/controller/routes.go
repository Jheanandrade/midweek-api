package controller

import (
	"github.com/labstack/echo/v4"
	"midweek-project/internal/handler"
)

func RegisterRoutes(e *echo.Echo) {
	e.POST("/generate-schedule", handler.GenerateSchedule)

	e.POST("/upload-zip", handler.HandleUploadZip)
	e.GET("/list-zip-files", handler.ListZipFiles)
	e.DELETE("/delete-zip-file", handler.DeleteZipFile)
}
