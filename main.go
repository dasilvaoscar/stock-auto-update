package main

import (
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"sort"
	"strconv"

	"github.com/joho/godotenv"
)

func main() {
	godotenv.Load()

	stock := getStockInformation("BBAS3")
	fmt.Println("Test", stock)
}

func getStockInformation(ticker string) *Stock {
	uri := "www.alphavantage.co"
	query := fmt.Sprintf("symbol=%s.SA&apikey=%s", ticker, os.Getenv("ALPHAVANTAGE_API_KEY_TEST"))
	url := fmt.Sprintf("https://%s/query?function=TIME_SERIES_MONTHLY_ADJUSTED&%s", uri, query)

	response, err := http.Get(url)

	if err != nil || response.StatusCode != 200 {
		panic(response.StatusCode)
	}

	defer response.Body.Close()

	var apiResponse AlphaVantageResponse
	if err := json.NewDecoder(response.Body).Decode(&apiResponse); err != nil {
		fmt.Printf("Erro ao decodificar o JSON: %v\n", err)
		panic(err)
	}

	var D5YAmount float64 = 5

	allDates := make([]string, 0, len(apiResponse.MonthlyAdjusted))

	for date := range apiResponse.MonthlyAdjusted {
		allDates = append(allDates, date)
	}
	
	sort.Sort(sort.Reverse(sort.StringSlice(allDates)))
	
	for i, date := range allDates {
		if i >= 12 {
			break
		}
		
		manthlyData := apiResponse.MonthlyAdjusted[date]
		dividend := manthlyData.Divident
		
		dividendValue, _ := strconv.ParseFloat(dividend, 64)
		
		D5YAmount = D5YAmount + dividendValue
	}

	var D5YPercentual float64 = (D5YAmount / 21.79) * 100

	return &Stock{
		Ticker: ticker,
		D5Y:    D5YPercentual,
	}
}

type Stock struct {
	Ticker string
	D5Y    float64
}

type AlphaVantageResponse struct {
	MetaData        MetaData             `json:"Meta Data"`
	MonthlyAdjusted map[string]DailyData `json:"Monthly Adjusted Time Series"`
}

type DailyData struct {
	Open          string `json:"1. open"`
	High          string `json:"2. high"`
	Low           string `json:"3. low"`
	Close         string `json:"4. close"`
	AdjustedClose string `json:"5. adjusted close"`
	Volume        string `json:"6. volume"`
	Divident      string `json: "7. dividend amount"`
}

type MetaData struct {
	Information   string `json:"1. Information"`
	Symbol        string `json:"2. Symbol"`
	LastRefreshed string `json:"3. Last Refreshed"`
	OutputSize    string `json:"4. Output Size"`
	TimeZone      string `json:"5. Time Zone"`
}
