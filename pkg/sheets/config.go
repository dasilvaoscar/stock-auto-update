package sheets

import (
	"fmt"
	"os"
)

type Config struct {
	CredentialsFile string
	SpreadsheetID   string
	SheetName       string
	
	PVPColumn       int
	DYColumn        int
	FIIsStartIndex  int
	
	RequestTimeout  int
	RequestDelay    int
}

func GetConfigs() *Config {
	spreadsheetID := os.Getenv("SPREADSHEET_ID")

	if spreadsheetID == "" {
		fmt.Println("ERRO: ID da planilha não definido!")
		fmt.Println("Defina a variável de ambiente SPREADSHEET_ID ou edite o arquivo config.go")
		fmt.Println("Exemplo: export SPREADSHEET_ID=\"1A2B3C4D5E6F7G8H9I0J\"")
		os.Exit(1)
	}

	config := &Config{
		CredentialsFile: "credentials.json",
		SpreadsheetID:  	spreadsheetID,
		SheetName:       "Testes",
		PVPColumn:       12,
		DYColumn:        3,
		FIIsStartIndex:  2,
		RequestTimeout:  30,
		RequestDelay:    1,
	}

	return config
}
