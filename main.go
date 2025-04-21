package main

import (
	"github.com/labstack/echo/v4"
	"github.com/labstack/echo/v4/middleware"
	"midweek-project/internal/controller"
)

func main() {
	e := echo.New()

	e.Use(middleware.Logger())
	e.Use(middleware.Recover())

	controller.RegisterRoutes(e)

	e.Logger.Fatal(e.Start(":8080"))
}
