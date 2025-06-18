package sheets

import "os"

// Config contém todas as configurações do aplicativo
type Config struct {
	// Google Sheets
	CredentialsFile string
	SpreadsheetID   string
	SheetName       string
	
	// Colunas
	PVPColumn       int
	DYColumn        int
	FIIsStartIndex  int
	
	// Yahoo Finance
	RequestTimeout  int // em segundos
	RequestDelay    int // em segundos, delay entre requests
}

// DefaultConfig retorna a configuração padrão
func DefaultConfig() *Config {
	spreadsheetID := os.Getenv("SPREADSHEET_ID")

	return &Config{
		CredentialsFile: "credentials.json",
		SpreadsheetID:  	spreadsheetID,
		SheetName:       "Testes",
		PVPColumn:       12,
		DYColumn:        3,
		FIIsStartIndex:  2,
		RequestTimeout:  30,
		RequestDelay:    1,
	}
}
