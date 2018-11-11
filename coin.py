import json
import requests
from datetime import datetime
import xlsxwriter
import time

def market_info():
    convert = 'USD'

    global_url = 'https://api.coinmarketcap.com/v2/global/' + '?convert=' + convert

    request = requests.get(global_url)
    results = request.json()

    #print(json.dumps(results, sort_keys=True, indent=4))

    data = results['data']
    active_currencies = data['active_cryptocurrencies']
    active_markets = data['active_markets']
    bitcoin_percentage = data['bitcoin_percentage_of_market_cap']
    last_updated_timestamp = data['last_updated']
    global_cap = int(data['quotes']['USD']['total_market_cap'])
    global_volume = int(data['quotes']['USD']['total_volume_24h'])

    active_currencies_string = '{:,}'.format(active_currencies)
    active_markets_string = '{:,}'.format(active_markets)
    global_cap_string = '{:,}'.format(global_cap)
    global_volume_string = '{:,}'.format(global_volume)

    last_updated_string = datetime.fromtimestamp(last_updated_timestamp).strftime('%B %d, %Y at %I:%M%p')

    global_cap_string = '$' + global_cap_string
    global_volume_string = '$' + global_volume_string
    bitcoin_percentage = '%' + str(bitcoin_percentage)

    print('\nActive Currencies: {}\nActive Markets: {}\nGlobal Market Cap: {}\n24h Volume: {}\nBitcoin Dominance: {}'.format(active_currencies_string,active_markets_string,global_cap_string,global_volume_string,bitcoin_percentage))

    print('\nThis information was updated on ' + last_updated_string + '.')





def coin():
    # open excel workbooks
    crypto_workbook = xlsxwriter.Workbook('coins.xlsx')

    # add a sheet
    crypto_sheet = crypto_workbook.add_worksheet()

    # add headers to the sheet
    crypto_sheet.write('A1',"Rank")
    crypto_sheet.write('B1',"Name")
    crypto_sheet.write('C1',"Symbol")
    crypto_sheet.write('D1',"Market Cap")
    crypto_sheet.write('E1',"Price")
    crypto_sheet.write('F1',"24h Volume")
    crypto_sheet.write('G1',"Circulating Supply")
    crypto_sheet.write('H1',"Total Supply")
    crypto_sheet.write('I1',"Hour Change")
    crypto_sheet.write('J1',"Day Change")
    crypto_sheet.write('K1',"Week Change")
    
    i = 0 #loop tracker
    start = 340 #first coin rank
    f = 1
    limit = 7 #page
    while i < limit:
        time.sleep(2)
        ticker_url = 'https://api.coinmarketcap.com/v2/ticker/?' + 'start=' + str(start) + '&limit=100' + '&sort=rank&structure=array'
        request = requests.get(ticker_url)
        results = request.json()

        #print(json.dumps(results, sort_keys=True, indent=4))

        data = results['data']

        print()
        for currency in data:
            rank = currency['rank']
            name = currency['name']
            symbol = currency['symbol']

            circulating_supply = int(currency['circulating_supply'])
            total_supply = int(currency['total_supply'])

            quotes = currency['quotes']['USD']
            market_cap = quotes['market_cap']
            hour_change = quotes['percent_change_1h']
            day_change = quotes['percent_change_24h']
            week_change = quotes['percent_change_7d']
            price = quotes['price']
            volume = quotes['volume_24h']

            volume_string = '{:,}'.format(volume)
            market_cap_string = '{:,}'.format(market_cap)
            circulating_supply_string = '{:,}'.format(circulating_supply)
            total_supply_string = '{:,}'.format(total_supply)

            if(market_cap <= 10000000.0 and market_cap >= 999999.9 and volume >=499999.9):
                crypto_sheet.write(f,0,rank)
                crypto_sheet.write(f,1,name)
                crypto_sheet.write(f,2,symbol)
                crypto_sheet.write(f,3,'$' + market_cap_string)
                crypto_sheet.write(f,4,'$' + str(price))
                crypto_sheet.write(f,5,'$' + volume_string)
                crypto_sheet.write(f,6,circulating_supply_string)
                crypto_sheet.write(f,7,total_supply_string)
                crypto_sheet.write(f,8,str(hour_change) + '%')
                crypto_sheet.write(f,9,str(day_change) + '%')
                crypto_sheet.write(f,10,str(week_change) + '%')
                f += 1
                print("Successful")
            else:
                print("Fail")    
            """
            print(str(rank) + ': ' + name + ' (' + symbol + ')')
            print('Market cap: \t\t$' + market_cap_string)
            print('Price: \t\t\t$' + str(price))
            print('24h Volume: \t\t$' + volume_string)
            print('Hour change: \t\t' + str(hour_change) + '%')
            print('Day change: \t\t' + str(day_change) + '%')
            print('Week change: \t\t' + str(week_change) + '%')
            print('Circulating supply: \t' + circulating_supply_string)
            print('Total supply: \t\t' + total_supply_string)
            print('Percentage circulating: ' + str(int(circulating_supply / total_supply * 100)) + '%')
            print()
            """
        i += 1
        start +=100
        if(i == limit):        
            crypto_workbook.close()
            break
coin()    