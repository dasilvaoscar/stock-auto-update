package sheets

import (
	"context"
	"fmt"
	"io/ioutil"
	"strings"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// SheetsClient encapsula as operações com Google Sheets
type SheetsClient struct {
	service       *sheets.Service
	spreadsheetID string
	config        *Config
}

// NewSheetsClient cria um novo cliente para Google Sheets
func NewSheetsClient(config *Config) (*SheetsClient, error) {
	ctx := context.Background()
	
	// Ler o arquivo de credenciais
	credentialsJSON, err := ioutil.ReadFile(config.CredentialsFile)
	if err != nil {
		return nil, fmt.Errorf("erro ao ler arquivo de credenciais: %v", err)
	}

	// Configurar as credenciais
	credentials, err := google.JWTConfigFromJSON(credentialsJSON, 
		"https://www.googleapis.com/auth/spreadsheets",
		"https://www.googleapis.com/auth/drive")
	if err != nil {
		return nil, fmt.Errorf("erro ao configurar credenciais: %v", err)
	}

	client := credentials.Client(ctx)
	
	// Criar o serviço do Sheets
	service, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("erro ao criar serviço do Sheets: %v", err)
	}

	return &SheetsClient{
		service:       service,
		spreadsheetID: config.SpreadsheetID,
		config:        config,
	}, nil
}

// SetSpreadsheetID define o ID da planilha
func (sc *SheetsClient) SetSpreadsheetID(spreadsheetID string) {
	sc.spreadsheetID = spreadsheetID
}

// GetColumnValues retorna todos os valores de uma coluna
func (sc *SheetsClient) GetColumnValues(sheetName string, column int) ([]string, error) {
	// Converter número da coluna para letra (A=1, B=2, etc.)
	columnLetter := columnNumberToLetter(column)
	
	readRange := fmt.Sprintf("%s!%s:%s", sheetName, columnLetter, columnLetter)
	
	resp, err := sc.service.Spreadsheets.Values.Get(sc.spreadsheetID, readRange).Do()
	if err != nil {
		return nil, fmt.Errorf("erro ao ler coluna %s: %v", columnLetter, err)
	}

	var values []string
	for _, row := range resp.Values {
		if len(row) > 0 && row[0] != nil {
			values = append(values, fmt.Sprintf("%v", row[0]))
		} else {
			values = append(values, "")
		}
	}

	return values, nil
}

// UpdateCell atualiza uma célula específica
func (sc *SheetsClient) UpdateCell(sheetName string, row, col int, value interface{}) error {
	// Converter número da coluna para letra
	columnLetter := columnNumberToLetter(col)
	
	cellRange := fmt.Sprintf("%s!%s%d", sheetName, columnLetter, row)
	
	var valueToUpdate interface{}
	switch v := value.(type) {
	case float64:
		valueToUpdate = v
	case string:
		valueToUpdate = v
	case int:
		valueToUpdate = v
	default:
		valueToUpdate = fmt.Sprintf("%v", v)
	}

	valueRange := &sheets.ValueRange{
		Values: [][]interface{}{{valueToUpdate}},
	}

	_, err := sc.service.Spreadsheets.Values.Update(sc.spreadsheetID, cellRange, valueRange).
		ValueInputOption("RAW").Do()
	
	if err != nil {
		return fmt.Errorf("erro ao atualizar célula %s: %v", cellRange, err)
	}

	return nil
}

// GetFIIs retorna a lista de FIIs da planilha
func (sc *SheetsClient) GetFIIs(sheetName string) ([]string, error) {
	values, err := sc.GetColumnValues(sheetName, 1)

	
	if err != nil {
		return nil, err
	}

	if len(values) <= sc.config.FIIsStartIndex {
		return []string{}, nil
	}

	// Filtrar valores vazios
	var fiis []string
	for _, fii := range values[sc.config.FIIsStartIndex:] {
		trimmed := strings.TrimSpace(fii)
		if trimmed != "" {
			fiis = append(fiis, trimmed)
		}
	}

	return fiis, nil
}

// GetIndexByValue encontra o índice de um valor em uma coluna específica
func (sc *SheetsClient) GetIndexByValue(sheetName string, column int, searchValue string) (int, error) {
	values, err := sc.GetColumnValues(sheetName, column)
	if err != nil {
		return 0, err
	}

	for i, value := range values {
		if strings.TrimSpace(value) == strings.TrimSpace(searchValue) {
			return i + 1, nil // +1 porque as planilhas são 1-indexed
		}
	}

	return 0, fmt.Errorf("valor '%s' não encontrado na coluna %d", searchValue, column)
}

// BatchUpdate atualiza múltiplas células de uma vez
func (sc *SheetsClient) BatchUpdate(updates []BatchUpdateRequest) error {
	if len(updates) == 0 {
		return nil
	}

	var valueRanges []*sheets.ValueRange
	
	for _, update := range updates {
		columnLetter := columnNumberToLetter(update.Column)
		cellRange := fmt.Sprintf("%s!%s%d", update.SheetName, columnLetter, update.Row)
		
		valueRange := &sheets.ValueRange{
			Range:  cellRange,
			Values: [][]interface{}{{update.Value}},
		}
		valueRanges = append(valueRanges, valueRange)
	}

	batchUpdateRequest := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "RAW",
		Data:            valueRanges,
	}

	_, err := sc.service.Spreadsheets.Values.BatchUpdate(sc.spreadsheetID, batchUpdateRequest).Do()
	if err != nil {
		return fmt.Errorf("erro ao fazer batch update: %v", err)
	}

	return nil
}

// BatchUpdateRequest representa uma solicitação de atualização em lote
type BatchUpdateRequest struct {
	SheetName string
	Row       int
	Column    int
	Value     interface{}
}

// columnNumberToLetter converte um número de coluna para letra (1=A, 2=B, etc.)
func columnNumberToLetter(column int) string {
	result := ""
	for column > 0 {
		column-- // Ajustar para base 0
		result = string(rune('A'+column%26)) + result
		column /= 26
	}
	return result
}
