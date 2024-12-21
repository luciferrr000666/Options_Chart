import requests
import pandas as pd
import datetime
import time
import streamlit as st
from ta import volatility, momentum, trend
import plotly.graph_objects as go
from io import BytesIO


class ContractAnalyzer:
    def __init__(self, company_name, start_epoch, end_epoch, interval):
        self.company_name = company_name
        self.start_epoch = start_epoch
        self.end_epoch = end_epoch
        self.interval = interval
        self.search_id = None
        self.nse_scrip_code = None
        self.call_contract_id = None
        self.put_contract_id = None
        self.headers = {"accept": "application/json"}
        self.base_url = "https://groww.in/v1/api"

    def fetch_ticker_details(self):
        try:
            url = f"{self.base_url}/search/v3/query/global/st_query?from=0&query={self.company_name}&size=10&web=true"
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            data = response.json()

            items = data.get('data', {}).get('content', [])
            if not items:
                st.warning(f"No results found for company: {self.company_name}")
                return False

            self.search_id = items[0].get('search_id')
            self.nse_scrip_code = items[0].get('nse_scrip_code')
            if not self.search_id or not self.nse_scrip_code:
                st.warning(f"Search ID or NSE scrip code not found for company: {self.company_name}")
                return False

            return True
        except Exception as e:
            st.error(f"Error fetching ticker details: {e}")
            return False

    def fetch_closest_contract_ids(self):
        try:
            derivatives_url = f"{self.base_url}/option_chain_service/v1/option_chain/derivatives/{self.search_id}"
            response = requests.get(derivatives_url, headers=self.headers)
            response.raise_for_status()
            derivatives_data = response.json()

            option_chains = derivatives_data.get('optionChain', {}).get('optionChains', [])
            if not option_chains:
                st.warning(f"No option chains found for {self.company_name}")
                return False

            live_prices_url = f"{self.base_url}/stocks_data/v1/tr_live_prices/exchange/NSE/segment/CASH/{self.nse_scrip_code}/latest"
            response = requests.get(live_prices_url, headers=self.headers)
            response.raise_for_status()
            live_price = response.json().get('ltp')

            closest_chain = min(option_chains, key=lambda x: abs(x['strikePrice'] / 100 - live_price))
            self.call_contract_id = closest_chain.get("callOption", {}).get("growwContractId")
            self.put_contract_id = closest_chain.get("putOption", {}).get("growwContractId")

            if not self.call_contract_id or not self.put_contract_id:
                st.warning(f"Unable to find valid contract IDs for {self.company_name}")
                return False

            return True
        except Exception as e:
            st.error(f"Error fetching contract IDs: {e}")
            return False

    def fetch_contract_price_details(self, contract_id):
        try:
            if not contract_id:
                return None

            url = f"{self.base_url}/stocks_fo_data/v4/charting_service/chart/exchange/NSE/segment/FNO/{contract_id}?endTimeInMillis={self.end_epoch}&intervalInMinutes={self.interval}&startTimeInMillis={self.start_epoch}"
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            data = response.json()

            candles = data.get("candles", [])
            if not candles:
                return None

            formatted_data = [
                {
                    "time": datetime.datetime.fromtimestamp(candle[0]).strftime('%Y-%m-%d %H:%M:%S'),
                    "open": candle[1],
                    "high": candle[2],
                    "low": candle[3],
                    "close": candle[4],
                    "volume": candle[5],
                }
                for candle in candles
            ]

            return pd.DataFrame(formatted_data)
        except Exception as e:
            st.error(f"Error fetching price details for contract ID {contract_id}: {e}")
            return None

    def calculate_technical_indicators(self, data):
        try:
            if data is None or len(data) < 20:
                st.warning("Not enough data to calculate indicators.")
                return None

            rsi = momentum.RSIIndicator(data['close'], window=14)
            data['RSI'] = rsi.rsi()
            data['RSI_SMA_14'] = data['RSI'].rolling(window=14).mean()
            data['RSI_SMA-14_Difference'] = data['RSI'] - data['RSI_SMA_14']
            data['SMA_20'] = data['close'].rolling(window=20).mean()
            bb = volatility.BollingerBands(data['close'], window=20)
            data['BB_Upper'] = bb.bollinger_hband()
            data['BB_Lower'] = bb.bollinger_lband()
            data['BollingerBandWidth'] = bb.bollinger_wband()

            macd = trend.MACD(data['close'], window_slow=26, window_fast=12, window_sign=9)
            data['MACD_Histogram'] = macd.macd_diff()

            adx = trend.ADXIndicator(data['high'], data['low'], data['close'], window=14)
            data['ADX'] = adx.adx()

            tsi = momentum.TSIIndicator(data['close'], window_slow=25, window_fast=13)
            data['TSI'] = tsi.tsi()

            return data.dropna()
        except Exception as e:
            st.error(f"Error calculating technical indicators: {e}")
            return None

    # For excel data    
    def analysis_summary(self):
        if not self.fetch_ticker_details():
            return None

        if not self.fetch_closest_contract_ids():
            return None

        call_data = self.fetch_contract_price_details(self.call_contract_id)
        put_data = self.fetch_contract_price_details(self.put_contract_id)

        call_technical = self.calculate_technical_indicators(call_data)
        put_technical = self.calculate_technical_indicators(put_data)

        call_summary = call_technical.iloc[-1].to_dict() if call_technical is not None and not call_technical.empty else None
        put_summary = put_technical.iloc[-1].to_dict() if put_technical is not None and not put_technical.empty else None

        if call_summary:
            call_summary["contractId"] = self.call_contract_id
        if put_summary:
            put_summary["contractId"] = self.put_contract_id

        return {
            "Call": call_summary,
            "Put": put_summary,
        }

    # For plotting charts
    def analyze(self):
        if not self.fetch_ticker_details():
            return None

        if not self.fetch_closest_contract_ids():
            return None

        call_data = self.fetch_contract_price_details(self.call_contract_id)
        put_data = self.fetch_contract_price_details(self.put_contract_id)

        call_technical = self.calculate_technical_indicators(call_data)
        put_technical = self.calculate_technical_indicators(put_data)

        return {
            "Call": {"data": call_technical, "contractId": self.call_contract_id},
            "Put": {"data": put_technical, "contractId": self.put_contract_id},
        }


def plot_candlestick_chart(data, title, key):
    fig = go.Figure(data=[go.Candlestick(
        x=data['time'],
        open=data['open'],
        high=data['high'],
        low=data['low'],
        close=data['close'],
        name=title
    )])

    if 'BB_Upper' in data.columns and 'BB_Lower' in data.columns:
        fig.add_trace(go.Scatter(
            x=data['time'],
            y=data['BB_Upper'],
            mode='lines',
            name='BB Upper',
            line=dict(color='green', width=1, dash='dot')
        ))
        fig.add_trace(go.Scatter(
            x=data['time'],
            y=data['BB_Lower'],
            mode='lines',
            name='BB Lower',
            line=dict(color='red', width=1, dash='dot')
        ))

    if 'SMA_20' in data.columns:
        fig.add_trace(go.Scatter(
            x=data['time'],
            y=data['SMA_20'],
            mode='lines',
            name='SMA 20',
            line=dict(color='blue', width=1.5)
        ))

    fig.update_layout(
        title=title,
        xaxis_title="Date",
        yaxis_title="Price",
        yaxis2=dict(
            title="Volume",
            overlaying="y",
            side="right",
            showgrid=False
        ),
        xaxis_rangeslider_visible=False,
        template="plotly_dark"
    )
    st.plotly_chart(fig, key=key)


def main():
    st.title("Options Contract Analyzer")

    uploaded_file = st.file_uploader("Upload Ticker CSV File", type="csv")
    start_date = st.date_input("Start Date")
    start_time = st.time_input("Start Time")
    end_date = st.date_input("End Date")
    end_time = st.time_input("End Time")
    interval = st.number_input("Interval in Minutes", min_value=1, value=15)
    output_path = st.text_input("Enter output path & filename (e.g., output.xlsx)", value="analysis.xlsx")

    if st.button("Analyze"):
        if uploaded_file:
            start_epoch = int(time.mktime(datetime.datetime.combine(start_date, start_time).timetuple()) * 1000)
            end_epoch = int(time.mktime(datetime.datetime.combine(end_date, end_time).timetuple()) * 1000)

            tickers_df = pd.read_csv(uploaded_file)
            call_data_list = []
            put_data_list = []

            for ticker in tickers_df['Ticker']:
                st.write(f"Processing {ticker}...")
                analyzer = ContractAnalyzer(ticker.strip(), start_epoch, end_epoch, interval)

                # Fetch detailed results for plotting
                results = analyzer.analyze()

                # Fetch summary results for saving in Excel
                contracts_results = analyzer.analysis_summary()

                if results and results['Call'] and results['Call']['data'] is not None:
                    plot_candlestick_chart(
                        results['Call']['data'], 
                        f"CALL - {results['Call']['contractId']}", 
                        key=f"{ticker}_CALL"
                    )
                    if contracts_results and contracts_results['Call']:
                        call_data_list.append({"Ticker": ticker, **contracts_results['Call']})

                if results and results['Put'] and results['Put']['data'] is not None:
                    plot_candlestick_chart(
                        results['Put']['data'], 
                        f"PUT - {results['Put']['contractId']}", 
                        key=f"{ticker}_PUT"
                    )
                    if contracts_results and contracts_results['Put']:
                        put_data_list.append({"Ticker": ticker, **contracts_results['Put']})

            if call_data_list or put_data_list:
                output = BytesIO()
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    if call_data_list:
                        call_df = pd.DataFrame(call_data_list)
                        call_df.to_excel(writer, index=False, sheet_name="CALL")

                    if put_data_list:
                        put_df = pd.DataFrame(put_data_list)
                        put_df.to_excel(writer, index=False, sheet_name="PUT")

                st.success(f"Results saved to {output_path}")
            else:
                st.error("Please upload a CSV file and provide an output path.")

if __name__ == "__main__":
    main()
