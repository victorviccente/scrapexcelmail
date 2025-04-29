import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime
import random
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

load_dotenv()

class StockScraper:
    def __init__(self):
        self.url = "https://finance.yahoo.com/most-active"
        self.user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Safari/605.1.15",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:90.0) Gecko/20100101 Firefox/90.0"
        ]
        self.headers = {
            "User-Agent": random.choice(self.user_agents),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Cache-Control": "max-age=0"
        }

    def scrape_most_active_stocks(self):
        try:
            response = requests.get(self.url, headers=self.headers)
            response.raise_for_status()

            if "consent.yahoo.com" in response.url:
                print("Detected consent page. Yahoo Finance is now requiring consent management.")
                print("Consider using the Yahoo Finance API or an alternative data source.")
                return None

            soup = BeautifulSoup(response.text, 'html.parser')
            table = self._find_table(soup)

            if not table:
                print("Could not locate the table with stock data. Page structure may have changed.")
                with open("yahoo_finance_debug.html", "w", encoding="utf-8") as f:
                    f.write(response.text)
                print("Saved HTML to yahoo_finance_debug.html for debugging")
                return None

            return self._parse_table(table)

        except requests.exceptions.RequestException as e:
            print(f"Error fetching data: {e}")
            return None
        except Exception as e:
            print(f"Error processing data: {e}")
            return None

    def _find_table(self, soup):
        table = soup.find('table', {'data-test': 'table'})
        if not table:
            tables = soup.find_all('table')
            return tables[0] if tables else None
        return table

    def _parse_table(self, table):
        stock_data = {
            'Symbol': [], 'Name': [], 'Price': [], 'Change': [],
            'Change %': [], 'Volume': [], 'Market Cap': []
        }

        rows = table.find('tbody').find_all('tr') if table.find('tbody') else table.find_all('tr')[1:]

        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 6:
                try:
                    stock_data['Symbol'].append(cells[0].get_text(strip=True))
                    stock_data['Name'].append(self._get_name(cells[1]))
                    stock_data['Price'].append(cells[2].get_text(strip=True))
                    stock_data['Change'].append(cells[3].get_text(strip=True))
                    stock_data['Change %'].append(cells[4].get_text(strip=True))
                    stock_data['Volume'].append(cells[5].get_text(strip=True))
                    stock_data['Market Cap'].append(cells[6].get_text(strip=True) if len(cells) > 6 else "N/A")
                except Exception:
                    continue

        return pd.DataFrame(stock_data) if stock_data['Symbol'] else None

    def _get_name(self, cell):
        name_element = cell.find('a')
        return name_element.get_text(strip=True) if name_element else cell.get_text(strip=True)

class ExcelFormatter:
    def __init__(self):
        self.header_format = {'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1}
        self.price_format = {'num_format': '$#,##0.00'}
        self.percent_format = {'num_format': '0.00%'}
        self.volume_format = {'num_format': '#,##0'}
        self.positive_format = {'font_color': 'green'}
        self.negative_format = {'font_color': 'red'}

    def create_formatted_excel(self, df, file_path):
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Most Active Stocks', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Most Active Stocks']

                self._format_headers(df, workbook, worksheet)
                self._format_columns(df, workbook, worksheet)
                self._apply_conditional_formatting(df, workbook, worksheet)
                self._set_column_widths(df, worksheet)
                self._add_metadata(workbook, worksheet)

            print(f"Excel file created successfully: {file_path}")
            return True

        except Exception as e:
            print(f"Error creating Excel file: {e}")
            return False

    def _format_headers(self, df, workbook, worksheet):
        header_format = workbook.add_format(self.header_format)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    def _format_columns(self, df, workbook, worksheet):
        worksheet.set_column('C:C', 12, workbook.add_format(self.price_format))
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:E', 12)
        worksheet.set_column('F:F', 15, workbook.add_format(self.volume_format))
        worksheet.set_column('G:G', 15)

    def _apply_conditional_formatting(self, df, workbook, worksheet):
        positive_format = workbook.add_format(self.positive_format)
        negative_format = workbook.add_format(self.negative_format)
        change_col = df.columns.get_loc('Change')

        for row in range(1, len(df) + 1):
            try:
                change_val = self._parse_change_value(df.iloc[row - 1, change_col])
                format_type = positive_format if change_val > 0 else negative_format if change_val < 0 else None
                if format_type:
                    worksheet.write(row, change_col, change_val, format_type)
            except Exception:
                continue

    def _parse_change_value(self, change_val):
        if isinstance(change_val, str):
            change_val = change_val.replace(',', '').replace('$', '').strip()
            if change_val.startswith('+'):
                change_val = change_val[1:]
            try:
                return float(change_val)
            except ValueError:
                return 0
        return change_val

    def _set_column_widths(self, df, worksheet):
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col) + 2)
            worksheet.set_column(i, i, column_len)

    def _add_metadata(self, workbook, worksheet):
        title_format = workbook.add_format({'bold': True, 'font_size': 14})
        worksheet.write('A1', 'Yahoo Finance Most Active Stocks', title_format)
        worksheet.write('A2', f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

class EmailSender:
    def __init__(self):
        self.sender_email = os.getenv("SENDER_EMAIL")
        self.app_password = os.getenv("APP_PASSWORD")
        self.recipient_email = os.getenv("RECIPIENT_EMAIL")

    def send_email(self, subject, body, file_paths):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.recipient_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            self._attach_files(msg, file_paths)
            self._send_smtp(msg)

            print(f"Email enviado com sucesso para {self.recipient_email}")
            return True

        except Exception as e:
            print(f"Erro ao enviar email: {e}")
            return False

    def _attach_files(self, msg, file_paths):
        for file_path in file_paths:
            if os.path.exists(file_path):
                with open(file_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    file_name = os.path.basename(file_path)
                    part.add_header('Content-Disposition', f'attachment; filename= {file_name}')
                    msg.attach(part)
                    print(f"Arquivo anexado: {file_name}")
            else:
                print(f"Arquivo não encontrado: {file_path}")

    def _send_smtp(self, msg):
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(self.sender_email, self.app_password)
            server.sendmail(self.sender_email, self.recipient_email, msg.as_string())

def main():
    print(f"Buscando ações mais ativas do Yahoo Finance em {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    time.sleep(1)

    scraper = StockScraper()
    df = scraper.scrape_most_active_stocks()

    if df is not None and not df.empty:
        print("\nAções Mais Ativas:")
        print("=" * 80)
        print(df)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        csv_filename = f"yahoo_most_active_{timestamp}.csv"
        excel_filename = f"yahoo_most_active_{timestamp}.xlsx"

        df.to_csv(csv_filename, index=False)
        print(f"\nDados salvos em {csv_filename}")

        formatter = ExcelFormatter()
        if formatter.create_formatted_excel(df, excel_filename):
            print(f"Dados formatados salvos em {excel_filename}")
            excel_full_path = os.path.abspath(excel_filename)
            print(f"Caminho completo para o arquivo Excel: {excel_full_path}")

            email_sender = EmailSender()
            subject = f"Ações Mais Ativas do Yahoo Finance - {datetime.now().strftime('%d/%m/%Y')}"
            body = f"""
Olá,

Segue em anexo os dados das ações mais ativas do Yahoo Finance coletados em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}.

Este email foi enviado automaticamente pelo sistema de raspagem de dados do Yahoo Finance.

Atenciosamente, Equipe 2V!
"""
            files_to_attach = [excel_full_path, os.path.abspath(csv_filename)]
            if email_sender.send_email(subject, body, files_to_attach):
                print("Email com dados das ações enviado com sucesso!")
            else:
                print("Não foi possível enviar o email com os dados das ações.")
    else:
        print("Falha ao recuperar dados.")

if __name__ == "__main__":
    main()