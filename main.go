package main

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"
	"github.com/joho/godotenv"

	"stock-auto-update/pkg/sheets"
)

func main() {
	godotenv.Load()

	config := sheets.DefaultConfig()
	
	if spreadsheetID := os.Getenv("SPREADSHEET_ID"); spreadsheetID != "" {
		config.SpreadsheetID = spreadsheetID
	}
	
	if config.SpreadsheetID == "" {
		fmt.Println("ERRO: ID da planilha não definido!")
		fmt.Println("Defina a variável de ambiente SPREADSHEET_ID ou edite o arquivo config.go")
		fmt.Println("Exemplo: export SPREADSHEET_ID=\"1A2B3C4D5E6F7G8H9I0J\"")
		os.Exit(1)
	}

	fmt.Println("🚀 Iniciando Stock Auto Update em Go...")
	fmt.Printf("📊 Planilha ID: %s\n", config.SpreadsheetID)
	fmt.Printf("📋 Aba: %s\n", config.SheetName)

	sheetsClient, err := sheets.NewSheetsClient(config)
	if err != nil {
		log.Fatalf("❌ Erro ao criar cliente do Sheets: %v", err)
	}

	fmt.Println("📈 Buscando lista de FIIs...")
	fiis, err := sheetsClient.GetFIIs(config.SheetName)
	if err != nil {
		log.Fatalf("❌ Erro ao obter FIIs: %v", err)
	}

	if len(fiis) == 0 {
		fmt.Println("⚠️ Nenhum FII encontrado na planilha")
		return
	}

	fmt.Printf("✅ Encontrados %d FIIs para processar\n", len(fiis))
	fmt.Println("---")

	var processedCount, errorCount, updatedPVP, updatedDY int
	startTime := time.Now()

	for i, ticker := range fiis {
		if strings.TrimSpace(ticker) == "" {
			continue
		}

		fmt.Printf("📊 [%d/%d] Processando %s...\n", i+1, len(fiis), ticker)
		processedCount++
	}
	
	fmt.Println("---")
	duration := time.Since(startTime)
	fmt.Println("🎉 Processamento concluído!")
	fmt.Printf("📊 Estatísticas:\n")
	fmt.Printf("   • FIIs processados: %d\n", processedCount)
	fmt.Printf("   • P/VP atualizados: %d\n", updatedPVP)
	fmt.Printf("   • DY atualizados: %d\n", updatedDY)
	fmt.Printf("   • Erros: %d\n", errorCount)
	fmt.Printf("   • Tempo total: %v\n", duration.Round(time.Second))
	
	if errorCount > 0 {
		fmt.Printf("⚠️ %d erros ocorreram durante o processamento\n", errorCount)
	}
}
