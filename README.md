# Value At Risk

## This program can:
- Calculate VaR using three methods (Historic, Variance-Covariance, Monte-Carlo)
- Perform VaR backtesting

## Required Libraries
- urllib3
- numpy
- pandas
- scipy
- yfinance
- plotly
- python-docx

## Install Dependencies

```
pip install -r requirements.txt
```

## How to Run

1. The program expects a file named `TICKERS.txt`, where each line represents one yfinance ticker (an example file is included in the repository).

2. Execute: ```python main.py```

3. The program creates a folder `reports`, where it outputs the generated `.docx` reports. Each report is named after its yfinance ticker.
