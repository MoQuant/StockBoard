'''
sheet = xw.Book("Sheet.xlsm").sheet[0]
sheet.range("").value
'''

import asyncio
import aiohttp
import json
import numpy as np
import threading
import xlwings as xw
import time 

class Stocks(threading.Thread):

    def __init__(self, tickers=['SPY'], limit=50):
        threading.Thread.__init__(self)
        self.tickers = tickers
        self.storage = []
        self.limit = limit

        self.url = 'wss://stream.data.alpaca.markets/v2/iex'

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self.stockData())

    async def stockData(self):
        async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(ssl=False)) as session:
            async with session.ws_connect(self.url) as client:
                auth = {"action": "auth", "key": "", "secret": ""}
                await client.send_str(json.dumps(auth))

                msg = {"action":"subscribe","trades":self.tickers}
                await client.send_str(json.dumps(msg))

                while True:
                    resp = await client.receive()
                    resp = json.loads(resp.data)

                    if resp[0]['T'] == 't':
                        price = float(resp[0]['p'])
                        volume = float(resp[0]['s'])

                        if len(self.storage) == 0:
                            self.storage.append([price, volume, 0])
                        else:
                            self.storage.append([price, volume, price/oldPrice - 1.0])
                        

                        if len(self.storage) > self.limit:
                            del self.storage[0]

                        oldPrice = price

book = xw.Book("Board.xlsx").sheets[0]

client = Stocks(limit=30)
client.start()

while True:
    storage = client.storage
    output_range = f'B4:D{4 + len(storage)}'
    book.range(output_range).value = storage
    print('Store: ', len(storage))
    time.sleep(0.5)

client.join()
    

