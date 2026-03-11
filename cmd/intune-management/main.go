package main

import (
	"fmt"
	"os"

	"intune-management/internal/app"
)

func main() {
	if err := app.Run(); err != nil {
		fmt.Println("Program error:", err)
		os.Exit(1)
	}
}
