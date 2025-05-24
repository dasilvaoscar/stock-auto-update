import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
import warnings

PVP_COLUMN = 8
DY_COLUMN = 5

def get_sheets():
  SERVICE_ACCOUNT_FILE = 'credentials.json'
  SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

  creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
  client = gspread.authorize(creds)

  spreadsheet = client.open("Orçamento mensal - 2025")

  return {
    "fiis": spreadsheet.worksheet("FIIs")
  }


def get_fiis(sheet):
  FIIS_START_INDEX = 19
  cols = sheet.col_values(1)
  return cols[FIIS_START_INDEX:]

def get_index_by_value(sheet, numero_coluna: int, valor_procurado: str) -> int | None:
    try:
      coluna = sheet.col_values(numero_coluna)
      indice = coluna.index(valor_procurado)
      return indice + 1
    except ValueError:
      return None

sheets = get_sheets()
fii_sheet = sheets['fiis']

all_infos = []

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    for ticker in get_fiis(fii_sheet):
      fii = yf.Ticker(ticker + ".SA")

      info = fii.info
      vp = info.get("bookValue")
      dy = info.get("dividendYield")

      ROW = get_index_by_value(fii_sheet, 1, ticker)

      if vp:
        preco = fii.history(period="1d")["Close"][-1]
        pvp = preco / vp
        fii_sheet.update_cell(ROW, PVP_COLUMN, pvp + 0.01)
      else:
        fii_sheet.update_cell(ROW, PVP_COLUMN, 'N/A')
      print('P/VP updated!')

      if dy:
        fii_sheet.update_cell(ROW, DY_COLUMN, dy / 100)
      else:
        fii_sheet.update_cell(ROW, DY_COLUMN, 'N/A')
      print('Dy updated!')