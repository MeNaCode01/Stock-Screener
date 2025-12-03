import yfinance as yf
from pandas_datareader import data as pdr
import yahooquery as yq
import datetime as dt
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
import pandas as pd
from pandas import ExcelWriter
# import fix_yahoo_finance as fyf
# import pandas_datareader as pdr
import yahoo_fin.stock_info as si

yf.pdr_override() 
start =dt.datetime(2010,12,1)
now = dt.datetime.now()

root = Tk()
ftypes = [(".xlsm","*.xlsx",".xls")]
ttl  = "Title"
dir1 = 'C:\\'
filePath = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)
# filePath=r"C:\Users\MANAV\Desktop\stock screener\National_Stock_Exchange_of_India_Ltd.xlsx"

stocklist = pd.read_excel(filePath)
# stocklist=stocklist.head(1)

exportList= pd.DataFrame(columns=['Stock', "50 Day MA", "150 Day Ma", "200 Day MA", "52 Week Low", "52 week High"])

for i in stocklist.index:
	stock=str(stocklist["Symbol"][i])
	print(stock)

	try:
		df = pdr.get_data_yahoo(f'{stock}.NS', start, now)

		smaUsed=[50,150,200]
		for x in smaUsed:
			sma=x
			df['SMA_' + str(sma)] = round(df.Close.rolling(window = sma).mean())
			# print(df)


		currentClose=df["Adj Close"][-1]
		moving_average_50=df["SMA_50"][-1]
		moving_average_150=df["SMA_150"][-1]
		moving_average_200=df["SMA_200"][-1]
		low_of_52week=min(df["Adj Close"][-260:])
		high_of_52week=max(df["Adj Close"][-260:])
		print(currentClose, moving_average_50, moving_average_150, moving_average_200, low_of_52week, high_of_52week)
		try:
			moving_average_200_20 = df["SMA_200"][-20]

		except Exception:
			moving_average_200_20=0

		#Condition 1: Current Price > 150 SMA and > 200 SMA
		if(currentClose>moving_average_150>moving_average_200):
			cond_1=True
		else:
			cond_1=False
		#Condition 2: 150 SMA and > 200 SMA
		if(moving_average_150>moving_average_200):
			cond_2=True
		else:
			cond_2=False
		#Condition 3: 200 SMA trending up for at least 1 month (ideally 4-5 months)
		if(moving_average_200>moving_average_200_20):
			cond_3=True
		else:
			cond_3=False
		#Condition 4: 50 SMA> 150 SMA and 50 SMA> 200 SMA
		if(moving_average_50>moving_average_150>moving_average_200):
			#print("Condition 4 met")
			cond_4=True
		else:
			#print("Condition 4 not met")
			cond_4=False
		#Condition 5: Current Price > 50 SMA
		if(currentClose>moving_average_50):
			cond_5=True
		else:
			cond_5=False
		#Condition 6: Current Price is at least 30% above 52 week low (Many of the best are up 100-300% before coming out of consolidation)
		if(currentClose>=(1.3*low_of_52week)):
			cond_6=True
		else:
			cond_6=False
		#Condition 7: Current Price is within 25% of 52 week high
		if(currentClose>=(.75*high_of_52week)):
			cond_7=True
		else:
			cond_7=False
		
		if(cond_1 and cond_2 and cond_3 and cond_4 and cond_5 and cond_6 and cond_7):
			exportList = exportList._append({'Stock': stock, "50 Day MA": moving_average_50, "150 Day Ma": moving_average_150, "200 Day MA": moving_average_200, "52 Week Low": low_of_52week, "52 week High": high_of_52week}, ignore_index=True)
	except Exception:
		print("No data on "+stock)

print(exportList)
currenttime = str(dt.datetime.now().date()) + '_' + str(dt.datetime.now().time()).replace('.', '-').replace(":", "-")
os.makedirs(os.path.dirname(filePath)+f"/{currenttime}")

for i in exportList.index:
	stock=str(exportList["Stock"][i])
#     print(stock)
	# print(i)

	try:
		df = pdr.get_data_yahoo(f'{stock}.NS', start, now)

		smaUsed=[50,150,200]
		for x in smaUsed:
			sma=x
			df['SMA_' + str(sma)] = round(df.Close.rolling(window = sma).mean())


		currentClose=df["Adj Close"][-1]
		moving_average_50=df["SMA_50"][-1]
		moving_average_150=df["SMA_150"][-1]
		moving_average_200=df["SMA_200"][-1]
		low_of_52week=min(df["Adj Close"][-260:])
		high_of_52week=max(df["Adj Close"][-260:])
		
		plt.figure(figsize=(12,6))
		plt.plot(df.Close, label="Closing Price", color="green")
		plt.plot(df.SMA_150, label="SMA_150", color="b")
		plt.plot(df.SMA_200, label="SMA_200", color="r")
		plt.legend()
		plt.xlabel("YEARS")
		plt.ylabel("PRICE")
		plt.title(stock)
		plt.savefig(f'{os.path.dirname(filePath)}/{currenttime}/{stock}_{currenttime}')
		
	except Exception as e:
		print(e)

# newFile=os.path.dirname(filePath)+"/ScreenOutput.xlsx"
newFile = f'{os.path.dirname(filePath)}/{currenttime}/ScreenOutput.xlsx'

# print(df.head(200))


writer= ExcelWriter(newFile)
exportList.to_excel(writer,"Sheet1")
writer.close()

currentPrice=[]
marketCap = []
PE_Ratio = []
stocks = []

# for i in exportList.index:
#     stock=str(exportList["Stock"][i])
#     stocks.append(stock)
for i in exportList.index:
	stock = str(exportList["Stock"][i])
	stocks.append(stock)

	try:
		df1  = pdr.get_data_yahoo(f'{stock}.NS', start, now)
		currentPrice.append(df.Close[-1])

		mc = yq.Ticker(f'{stock}.NS').price[f'{stock}.NS']['marketCap']
		marketCap.append(mc)

		pe = yq.Ticker(f'{stock}.NS').quotes[f'{stock}.NS']['trailingPE']
		# pe = dic['trailingPE']
		PE_Ratio.append(pe)

		# print(df1, mc, pe)

        # pe = yq.Ticker(f'{stock}.NS').quotes[f'{stock}.NS']


	except Exception as e:
		print("No data on "+stock)
		print(e)

	# try:        
	# 	df1 = pdr.get_data_yahoo(f'{stock}.NS', start, now)
	# 	currentPrice._append(df.Close[-1])
        
	# 	# df2 = pdr.get_quote_yahoo(f'{stock}.NS')['marketCap']
	# 	mc = yq.Ticker(f'{stock}.NS').price[f'{stock}.NS']['marketCap']
	# 	marketCap.append(mc)
	# 	# marketCap._append(df2.iloc[0])
        
    #     # stock_info = si.get_quote_table(f'{stock}.NS')
    #     # df3 = pd.DataFrame.from_dict(stock_info,orient='index')
    #     # PE_Ratio._append(float(df3.loc['PE Ratio (TTM)']))

	# except Exception:
	# 	print("No data on "+stock)
        
Market_Cap_in_mil = [i/100000000 for i in marketCap]
print('Current Price: ', currentPrice)
print('Market Cap in mill:', Market_Cap_in_mil)
print('PE RATIO: ',PE_Ratio)
