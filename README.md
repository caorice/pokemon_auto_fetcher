# Pokemon Price Auto Fetcher


## Install Dependencies

```bash

pip3 install -r requirements.txt

```

## Get Template Excel File

```bash

python3 main.py template.xlsx -d

```

## Update Latest Price Data Into Excel File

```bash

python3 main.py example.xlsx

```

## Update With Max Price And Min Price 

```bash

python3 main.py example.xlsx --min 200 --max 20000

```

## Update Latest Price Data From Ebay

```bash

python3 main.py example.xlsx -p ebay

```

## Update Latest Price Data From 130point "Search Ebay Sales" Tab

```bash

python3 main.py example.xlsx -p 130point

```

## Update Latest Price Data From 130point "All Marketplace Sales" Tab

```bash

python3 main.py example.xlsx -p 130point-all

```

## Use "socks5://127.0.0.1:7890" As An Example Of Proxy To Fetch Data

```bash

PROXY=socks5://127.0.0.1:7890 python3 main.py example.xlsx

```