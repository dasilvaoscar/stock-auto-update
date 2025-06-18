package crawler

import (
	"fmt"
	"io"
	"net/http"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/PuerkitoBio/goquery"
)

type FiiData struct {
	Codigo              string  `json:"codigo"`
	Nome                string  `json:"nome"`
	Preco               float64 `json:"preco"`
	DividendYield       float64 `json:"dividend_yield"`
	PVP                 float64 `json:"p_vp"`
	LiquidezMediaDiaria float64 `json:"liquidez_media_diaria"`
	Patrimonio          float64 `json:"patrimonio"`
	ValorPatrimonial    float64 `json:"valor_patrimonial"`
	Rentabilidade12m    float64 `json:"rentabilidade_12m"`
}

type InvestidorCrawler struct {
	client *http.Client
}

func NewInvestidorCrawler() *InvestidorCrawler {
	return &InvestidorCrawler{
		client: &http.Client{
			Timeout: 30 * time.Second,
		},
	}
}

func (ic *InvestidorCrawler) CrawlFii(codigo string) (*FiiData, error) {
	url := fmt.Sprintf("https://statusinvest.com.br/fundos-imobiliarios/%s", strings.ToLower(codigo))

	fmt.Printf("Fazendo requisição para: %s\n", url)

	resp, err := ic.makeRequest(url)
	if err != nil {
		return nil, fmt.Errorf("erro na requisição: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return nil, fmt.Errorf("status code: %d", resp.StatusCode)
	}

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("erro ao ler resposta: %v", err)
	}

	doc, err := goquery.NewDocumentFromReader(strings.NewReader(string(body)))
	if err != nil {
		return nil, fmt.Errorf("erro ao parsear HTML: %v", err)
	}

	fmt.Println("doc", doc)

	fii := &FiiData{
		Codigo: strings.ToUpper(codigo),
	}

	// Extrai nome do FII
	doc.Find("h1, .ticker, .company-name").Each(func(i int, s *goquery.Selection) {
		if fii.Nome == "" {
			text := strings.TrimSpace(s.Text())
			if text != "" && !strings.Contains(text, codigo) {
				fii.Nome = text
			}
		}
	})

	// Extrai preço atual
	doc.Find(".price, [data-element='price']").Each(func(i int, s *goquery.Selection) {
		if fii.Preco == 0 {
			priceText := strings.TrimSpace(s.Text())
			fii.Preco = extractNumber(priceText)
		}
	})

	// Extrai indicadores dos cards
	doc.Find(".card").Each(func(i int, card *goquery.Selection) {
		title := strings.TrimSpace(card.Find(".card-title, .title, h3, h4").Text())
		value := strings.TrimSpace(card.Find(".card-value, .value, .number").Text())

		switch {
		case strings.Contains(strings.ToLower(title), "dividend yield"):
			fii.DividendYield = extractNumber(value)
		case strings.Contains(strings.ToLower(title), "p/vp"):
			fii.PVP = extractNumber(value)
		case strings.Contains(strings.ToLower(title), "liquidez"):
			fii.LiquidezMediaDiaria = extractNumber(value)
		case strings.Contains(strings.ToLower(title), "patrimônio"):
			fii.Patrimonio = extractNumber(value)
		case strings.Contains(strings.ToLower(title), "valor patrimonial"):
			fii.ValorPatrimonial = extractNumber(value)
		case strings.Contains(strings.ToLower(title), "rentabilidade"):
			fii.Rentabilidade12m = extractNumber(value)
		}
	})

	// Fallback: busca por atributos de dados específicos
	doc.Find("[title], [data-title]").Each(func(i int, s *goquery.Selection) {
		title, exists := s.Attr("title")
		if !exists {
			title, _ = s.Attr("data-title")
		}

		value := strings.TrimSpace(s.Text())

		switch strings.ToLower(title) {
		case "dividend yield":
			if fii.DividendYield == 0 {
				fii.DividendYield = extractNumber(value)
			}
		case "p/vp":
			if fii.PVP == 0 {
				fii.PVP = extractNumber(value)
			}
		}
	})

	return fii, nil
}

func (ic *InvestidorCrawler) makeRequest(url string) (*http.Response, error) {
	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return nil, err
	}

	req.Header.Set("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36")
	req.Header.Set("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")
	req.Header.Set("Accept-Encoding", "gzip, deflate, br, zstd")
	req.Header.Set("Accept-Language", "pt-BR,pt;q=0.9")
	req.Header.Set("Cache-Control", "max-age=0")
	req.Header.Set("Cookie", "hash-ads_popup=eyJpdiI6ImdoRVBaVDRaTHV3c1ZrL2d0ZHJSemc9PSIsInZhbHVlIjoiWURPU1FrYmlxYVRibDJWVktFMW4xOTNMWFdnMVV1ekExSC80dmpSKy9UUnpxQmtubFl1MFZ4Vm5JUnF5SGtrWnFySEp1YnZhd1lZeUhoR29sNDM5ajE5c1kyTnlYenFrbDg5SjZnMDJsSnc9IiwibWFjIjoiYmIyODcxN2E4NmQ0NTY4ZTMzNTg1NTNkZGJmODJjZThlZjQ0Nzk2YzY0NmJiODY0OTNlZWFhODI3MTUyMWJjMiIsInRhZyI6IiJ9; buy_and_hold=1; sib_cuid=2ddcab91-caab-424d-b372-1e4eb965df05; __gads=ID=a6b880a2caaff2ff:T=1747060692:RT=1747060692:S=ALNI_Mag9Z-qhfUxQteSbZH7So3thTT-Mw; __gpi=UID=0000101f0ba0495f:T=1747060692:RT=1747060692:S=ALNI_MZjYHR_-RGkRCN-doQT3F6d1zUzOQ; __eoi=ID=96c2c674c336ce9c:T=1747060692:RT=1747060692:S=AA-Afjbl6q1m-xZLfjJKe0eEQslx; _fbp=fb.2.1747060693082.401720723279702238; remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d=eyJpdiI6IklzdFo0dGg3VkQrOHFMdlVIb0RES1E9PSIsInZhbHVlIjoicUU3M21iKzNKY0pxVEpqNmJZMEtpeFVKL0ZlS3hubHJRQnBDS2JXSUhIRWRKRHEzcitvTmNKOWJqeEJmS1B2M1MxNnNNdU5HTnRGdTdDRjUwUmJxZCtyM1hiaUxEaGZEeWw5bWpocTRDcDlQVys1cGU3Y1VBdVZmQUlqcWVmblBrMkRyNFdxakVMa3YrTjRHdHgvRVN3PT0iLCJtYWMiOiJkZmQ3OWQyMjMwOWI2NDEzNjExNWYyNDVkMWRhNDc3NGI0OWRmNGE0MGRmNGViY2Y4MDg5NTk4NjUxYjNmMjJkIiwidGFnIjoiIn0%3D; _hjSessionUser_2058627=eyJpZCI6ImE3MTk5YjAwLTRmOTYtNTViYi04ZDEyLTZjMmEzMTBhNjRhOSIsImNyZWF0ZWQiOjE3NDcwNjE0MDE4MzgsImV4aXN0aW5nIjp0cnVlfQ==; _CEFT=EgNwlgpg7hAmBcAhA7IgZgDQNYCYCSADgBwCCA0AIwBWAionlAGoCiANlQCxA%3D%3D%3D; _ga_5EQ309XQDG=deleted; _ga_267511960=deleted; _gcl_au=1.1.1206915167.1747101344.630377774.1749507179.1749507179; pro_ref_v2=eyJpdiI6InNOY1V1ZDVSTmNPME5hT2Y5dDF6T2c9PSIsInZhbHVlIjoibmV1NU1TOTBWMXVveTZJbWcraDZRbFJDb01mL1pDeHoxeXN1d3puWkIxdzA5YkdtOHpJOXo3R1ZtREtQMUtQcyIsIm1hYyI6ImVmOTZkZDdhOTg0MWI2NDI2NGEyY2UyNmMwYjBmZDYwNTBlYThlYzhhNmUwM2Y2ODJiY2U5MTUzNDEyNzRhMjkiLCJ0YWciOiIifQ%3D%3D; cebs=1; cookie_contador_whole-site=eyJpdiI6Ino4WFBndVpTbUhDT2xiS0JVUnVlb3c9PSIsInZhbHVlIjoibFBIdU5VVXh0d1JZVURVK0JwNjh3L205TVJta3BZYm9BaTdBYjl3YVpUcm1qZ0ZyMTJJYkVHSmhvYWl6RUlZcDRWQUpEZHdpRDlNZVFNRjE4bHA2Y0E9PSIsIm1hYyI6IjU4MjE0Y2U0Mjc0OWEyMWExOGQ4NGIxZWZjNjZiZDg3ZmIyMTEwYmJlNGY5YTBmZjA5NmI3YTA0OWE3ZWIzYTgiLCJ0YWciOiIifQ%3D%3D; _gid=GA1.3.705099812.1750188134; _ce.clock_data=6%2C187.85.159.154%2C1%2Cd6dcc0a6def5582f8d3a9f7f2addb88b%2CChrome%2CBR; _clck=1o8unhz%7C2%7Cfwv%7C0%7C1958; _ga_267511960=GS2.1.s1750211543$o82$g1$t1750212864$j51$l0$h0; _ga=GA1.3.1459790563.1747060691; _ga_5EQ309XQDG=GS2.1.s1750211543$o63$g1$t1750212864$j52$l0$h1486361822; cebsp_=17; XSRF-TOKEN=eyJpdiI6InJsbUp6Ujk2NGJ4TVAxV2pWK2ZCenc9PSIsInZhbHVlIjoiRkRoc1o5MEljMDdnVGZnRHArRTM1QkF2UmNhMnRpRi9WNHBmUzJTeVhKUnRabW9ockxMR2tPVzJNMmVTVVJMbGc3UlpENGdtSXBJeFVXTmJZc1dUQk9vYkM1QjdURmlNUE85T05KbHpCRDJoejVraGdoRVp2c3hzTGtaTnRMcUoiLCJtYWMiOiI5NzE2ZDY2NGYxNzVjMGM0ZjliZTJkMTZhNGU4MTVkNzgyYTczZmI4ZDk0NWQzNWUzODVmZDZlZWRhZGU1NmI2IiwidGFnIjoiIn0%3D; laravel_session=eyJpdiI6ImhFbWRyQjBpNWhoeVZ1dy9NWndQb2c9PSIsInZhbHVlIjoieW5CYzAvdFB1YnRMNGlhSHFsTVRDSVJBL2YvaUs0bTZHYmhVOVFsUGtVOGoxYkpSNWRHbUY4dytPMG40a2V5U0RCVWtuaDJOMDkxcEpuWWxXbEo3RWJHY1lOcEEyb3JZYWJLODVnZUxTdHNRZExvcmRBYnlKQ2tXTkpCMkJZcHQiLCJtYWMiOiI0ZmEzMTRjY2UwNjk3ZDBmN2YxMjFiYjJlOTk2YzcyMzRlOWJlOTNmYTRhZDQwMThiYWJlMzFlZDZhMjY3N2MxIiwidGFnIjoiIn0%3D; sc_is_visitor_unique=rx12923067.1750213750.D89C696FFBFC41D99300312826066F36.58.50.41.39.38.33.29.26.24; _clsk=unch8z%7C1750213750733%7C4%7C0%7Cb.clarity.ms%2Fcollect; _ce.s=v~74d9a0e5f2ecdb967d44835c393fe793c31e0974~lcw~1750213777558~vir~returning~lva~1750212865125~vpv~21~flvl~%2CB7BfXk2Ip8A%3A1jQBIwVElj4~v11.cs~442075~v11.s~435f3680-385e-11f0-aaf1-e7daf9a6bbeb~v11.vs~74d9a0e5f2ecdb967d44835c393fe793c31e0974~v11ls~435f3680-385e-11f0-aaf1-e7daf9a6bbeb~v11.fsvd~e30%3D~gtrk.la~mc1c3e2u~lcw~1750213777926")
	req.Header.Set("Priority", "u=0, i")
	req.Header.Set("Sec-Ch-Ua", `"Google Chrome";v="135", "Not-A.Brand";v="8", "Chromium";v="135"`)
	req.Header.Set("Sec-Ch-Ua-Mobile", "?0")
	req.Header.Set("Sec-Ch-Ua-Platform", `"macOS"`)
	req.Header.Set("Sec-Fetch-Dest", "document")
	req.Header.Set("Sec-Fetch-Mode", "navigate")
	req.Header.Set("Sec-Fetch-Site", "same-origin")
	req.Header.Set("Sec-Fetch-User", "?1")
	req.Header.Set("Upgrade-Insecure-Requests", "1")

	return ic.client.Do(req)
}

func extractNumber(text string) float64 {
	// Remove caracteres não numéricos, mantendo vírgulas, pontos e sinais
	re := regexp.MustCompile(`[^\d,.-]`)
	cleaned := re.ReplaceAllString(text, "")

	// Substitui vírgula por ponto para conversão
	cleaned = strings.Replace(cleaned, ",", ".", -1)

	// Remove pontos extras (exceto o último, que é decimal)
	parts := strings.Split(cleaned, ".")
	if len(parts) > 2 {
		// Reconstrói o número: partes[0] + partes[1] + "." + última parte
		cleaned = strings.Join(parts[:len(parts)-1], "") + "." + parts[len(parts)-1]
	}

	num, err := strconv.ParseFloat(cleaned, 64)
	if err != nil {
		return 0
	}

	return num
}
