import openpyxl
import os.path
import datetime

from pandas_datareader import data
import pandas as pd
from openpyxl.utils import get_column_letter
import numpy as np

import urllib
from bs4 import BeautifulSoup as bs
import requests

import bs4
import re
import logging
import more_itertools
import tqdm

from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import win32clipboard
from openpyxl.styles import Alignment
import time

import matplotlib.pyplot as plt
from sklearn.decomposition import PCA






class Portfolio:


	def __init__(self, directory, file_name):


		if file_name[-4] + file_name[-3] + file_name[-2] + file_name[-1] != "xlsx":
			file_name = file_name + ".xlsx"

		file_location = os.path.join(directory, file_name)


		if not os.path.isfile(file_location):

			wb = openpyxl.Workbook()

			wb["Sheet"].title = "Prices"
			wb.create_sheet("Stocks")
			wb.create_sheet("Dates")
			wb.create_sheet("Accounts")



			wb.save(file_location)

			self.stocks = np.array([])
			self.shares = np.array([])
			self.equity = np.array([])

			self.borrowing_rate = 0.05

			self.assets = pd.DataFrame(np.zeros(shape = (1,3)))
			self.assets.columns = ["Equity", "Debt", "Net assets"]

			self.directory = directory
			self.file_name = file_name
			self.file_location = file_location

			self.xlsx = wb


			wb["Accounts"]["A1"] = 0.0
			wb["Accounts"]["B1"] = 0.0

			print("Empty portfolio created.")


		else:

			wb = openpyxl.load_workbook(file_location)

			s = wb["Stocks"]

			self.stocks = []
			self.shares = np.zeros(wb["Stocks"].max_row-1)

			j = 0
			

			for  i in range(wb["Stocks"].max_row-1):

				self.stocks.append(wb["Stocks"]["A" + str(i+2)].value) 
				self.shares[j] = wb["Stocks"]["B" + str(i+2)].value
				j += 1

			self.stocks = np.array(self.stocks)

			n = wb["Prices"].max_row

			self.equity = np.zeros(len(self.stocks))



			for j in range(len(self.stocks)):

				try:
					self.equity[j] = wb["Prices"][get_column_letter(j+1) + str(n)].value * self.shares[j]

				except:
					print(self.stocks[j] + " couldn't be priced...")


			self.borrowing_rate = 0.05

			self.borrowing_rate = 0.05

			cash = np.zeros(shape = (1,3))
			cash[0,0] = wb["Accounts"]["A1"].value
			cash[0,1] = wb["Accounts"]["B1"].value
			cash[0,2] = cash[0,0] + cash[0,1]

			self.assets = pd.DataFrame(cash)
			self.assets.columns = ["Equity", "Debt", "Net assets"]

			self.directory = directory
			self.file_name = file_name
			self.file_location = file_location

			self.xlsx = wb

			print("Existing portfolio loaded.")

			self.update()










	def add_stock(self, tickers):

		tickers = np.array(tickers)

		if len(self.stocks) == 0:

			for i in tickers:

				try:
					data.get_quote_yahoo(i)

				except:
					print(i + " couldn't be found... Deleting")

				else:
					self.stocks = np.append(self.stocks, i)
					self.equity = np.append(self.equity, 0.0)
					self.shares = np.append(self.shares, 0.0)	

					n = self.xlsx["Stocks"].max_row + 1
					self.xlsx["Stocks"]["A" + str(n)] = i
					self.xlsx["Stocks"]["B" + str(n)] = 0	

					self.xlsx.create_sheet(i)
					self.xlsx.create_sheet(i + "Calls")
					self.xlsx.create_sheet(i + "Puts")
					self.xlsx.create_sheet(i + "Historical_data")

					break	


		for i in tickers:

			if(len(np.where(self.stocks == i)[0]) == 0):

				try:
					data.get_quote_yahoo(i)

				except:
					print(i + " couldn't be found... Deleting")

				else:
					self.stocks = np.append(self.stocks, i)
					self.equity = np.append(self.equity, 0.0)
					self.shares = np.append(self.shares, 0.0)	

					n = self.xlsx["Stocks"].max_row + 1
					self.xlsx["Stocks"]["A" + str(n)] = i
					self.xlsx["Stocks"]["B" + str(n)] = 0

					self.xlsx.create_sheet(i)
					self.xlsx.create_sheet(i + "Calls")
					self.xlsx.create_sheet(i + "Puts")
					self.xlsx.create_sheet(i + "Historical_data")

			else:
				print(i + " is already in the portfolio")


		self.xlsx.save(self.file_location)




	def update(self):

		row = self.xlsx["Prices"].max_row + 1
		col = 0

		for i in self.stocks:

			col += 1

			try:
				new_price = data.get_quote_yahoo(i)

			except:
				print("Couldn't update " + i + "...")

			else:
				self.xlsx["Prices"][get_column_letter(col) + str(row)] = new_price["price"][0]
				self.equity[col-1] = self.shares[col-1] * new_price["price"][0]
		 

		now = datetime.datetime.now()
		date = str(now.year) + "-" + str(now.month) + "-" + str(now.day)

		n = self.xlsx["Dates"].max_row + 1

		self.xlsx["Dates"]["A" + str(n)] = date
		self.xlsx["Dates"]["B" + str(n)] = now.hour
		self.xlsx["Dates"]["C" + str(n)] = now.minute
		self.xlsx["Dates"]["D" + str(n)] = now.second

		self.assets["Equity"] = np.sum(self.equity)

		try:

			old_date = self.xlsx["Dates"]["A" + str(n-1)].value + " " + str(self.xlsx["Dates"]["B" + str(n-1)].value) + ":" + str(self.xlsx["Dates"]["C" + str(n-1)].value) + ":" + str(self.xlsx["Dates"]["D" + str(n-1)].value)
			old_date = datetime.datetime.strptime(old_date, "%Y-%m-%d %H:%M:%S")

		except:
			print("Couldn't update debt...")

		else:
			r = np.log(1 + self.borrowing_rate)
			t = (now - old_date).total_seconds() / (60*60*24*365)

			self.xlsx["Accounts"]["B1"].value *= np.exp(t*r)
			self.assets["Debt"] *= np.exp(t*r)

		self.assets["Net assets"] = self.assets["Equity"] + self.assets["Debt"]

		self.xlsx.save(self.file_location)

		print("Prices and Dates vectors successfully updated.")
		print("")

		print("Current assets:")
		print("")

		print(self.assets)





	def add_single_shares(self, ticker, amount):

		index = np.where(self.stocks == ticker)[0]

		if len(index) != 0:
			self.shares[index] += amount

			self.xlsx["Stocks"]["B" + str(index[0] + 2)].value += amount		

			price = data.get_quote_yahoo(ticker)["price"][0]

			self.xlsx["Accounts"]["B1"].value -= price*amount

			self.assets["Debt"] -= price*amount

			self.equity[index] = price*self.shares[index]
			self.assets["Equity"] = np.sum(self.equity)
			self.xlsx["Accounts"]["A1"].value = self.assets["Equity"][0]

			print(str(amount) + " " + ticker + " shares were added.")
			print("")

			print("Current assets:")
			print("")

			print(self.assets)

			self.xlsx.save(self.file_location)


		else:
			print(ticker + " not found within the portfolio")



	def add_vector_shares(self, amounts):


		if len(amounts) == len(self.shares):

			amounts = np.array(amounts)

			for index in range(len(amounts)):

				try:

					self.xlsx["Stocks"]["B" + str(index + 2)].value += amounts[index]
					self.shares[index] += amounts[index]
				
					price = data.get_quote_yahoo(self.stocks[index])["price"][0]

					self.xlsx["Accounts"]["B1"].value -= price*amounts[index]
					self.assets["Debt"] -= price*amounts[index]

					self.equity[index] = price*self.shares[index]

				except:

					print(self.stocks[index] + " couldn't be updated; canceling transaction...")


			self.xlsx["Accounts"]["A1"].value = np.sum(self.equity)
			self.assets["Equity"] = self.xlsx["Accounts"]["A1"].value
			self.xlsx.save(self.file_location)

			print("Added new shares...")
			print("")

			print("Current assets:")
			print("")

			print(self.assets)

		else:
			print("Couldn't update shares: incorrect numbers of dimentions")
			print("Provided vector has " + str(len(amounts)) + " elements; Portfolio object has " + str(len(self.shares)))




	def buy_call_put(self,ticker):

		if len(np.where(self.stocks == ticker)[0]) != 0:

			url = "https://www.barchart.com/stocks/quotes/" + ticker + "/options"

			browser = Chrome()
			browser.get(url)

			print("Webpage loaded.")
			print("")

			time.sleep(2)

			expiration_date_dropdown_menu = browser.find_element_by_class_name("expiration-name")
			date_options = expiration_date_dropdown_menu.find_elements_by_tag_name('option')

			expiration_dates = []

			for i in date_options:

				expiration_dates.append(i.text)

			print("Avaible expiration dates:")
			print("")

			print(expiration_dates)

			print("")
			choice = input("Enter the chosen date (copy-paste WITHOUT THE QUOTES):")

			url = url + "?expiration=" + choice

			browser.get(url)

			print("Webpage loaded.")
			print("")

			time.sleep(2)


			content = browser.find_element_by_css_selector("body")
			content.send_keys(Keys.CONTROL + "a")
			content.send_keys(Keys.CONTROL + "c")


			win32clipboard.OpenClipboard()
			clipboard_data = win32clipboard.GetClipboardData()
			win32clipboard.EmptyClipboard()
			win32clipboard.CloseClipboard()

			browser.close()

			string_list = np.array(clipboard_data.split("\r\n"))

			calls_index = np.where(string_list == "Calls")[0][0]
			puts_index = np.where(string_list == "Puts")[0][0]	
			expiration_index = calls_index - 1


			call_table_string = []
			n = int((puts_index - calls_index)/12)

			frm = calls_index + 2
			to = frm + 11

			for i in range(n):
				call_table_string.append(string_list[frm:to])
				frm = to + 1
				to = frm + 11


			put_table_string = []
			n = 0
			reached_end = False

			while not reached_end:
				reached_end = "Put Volume Tota" in string_list[puts_index + n]
				n += 1

			n = int(n / 12)

			frm = puts_index + 2
			to = frm + 11

			for i in range(n):
				put_table_string.append(string_list[frm:to])
				frm = to + 1
				to = frm + 11

			call_table_columns = string_list[calls_index + 1].split("\t")[0:11]
			put_table_columns = string_list[puts_index + 1].split("\t")[0:11]

			frames_dict = {"Call" : pd.DataFrame(call_table_string, columns = call_table_columns),
							"Put" : pd.DataFrame(put_table_string, columns = put_table_columns)}

			
			print("CALLS:")
			print("")
			print(frames_dict["Call"])

			print("")
			print("")

			print("PUTS:")
			print("")
			print(frames_dict["Put"])


			option_type = input("Purchase Call or Put? " )

			frame = frames_dict[option_type]

			print("")
			print("")
			print(option_type + ":")
			print("")
			print(frame)

			index = input("Enter the row index of the desired purchase: " )

			option = frame.iloc[int(index)]

			print("")
			print("")
			print("Chosen option:")
			print("")
			print("Expiration date: " + choice)
			print(option)

			n_purchase = input("# of options to purchase: ")


			n = self.xlsx[ticker].max_row + 1

			for i in range(len(option)):

				try:
					self.xlsx[ticker][get_column_letter(i+1) + str(n)].value = float(option[i])

				except:
					self.xlsx[ticker][get_column_letter(i+1) + str(n)].value = option[i]


			self.xlsx[ticker][get_column_letter(len(option)+1) + str(n)] = choice
			self.xlsx[ticker][get_column_letter(len(option)+2) + str(n)] = float(n_purchase)
			self.xlsx[ticker][get_column_letter(len(option)+3) + str(n)] = option_type

			self.assets["Debt"] -= float(option["Ask"]) * int(n_purchase)
			self.xlsx["Accounts"]["B1"].value -= float(option["Ask"]) * int(n_purchase)
			self.assets["Net assets"] -= float(option["Ask"]) * int(n_purchase)

			self.xlsx.save(self.file_location)

			print("")
			print("")
			print("")
			print("Purchase completed.")
			print("")
			print("Current Assets:")
			print("")
			print(self.assets)

		else:

			print(ticker + " wasn't found within the portfolio")



	def exercise_put_call(self):

		now = datetime.datetime.now()
		to_check = []
		to_rmv = []

		for i in self.stocks:

			n = self.xlsx[i].max_row

			if(n > 1):

				for j in range(n-1):

					then = self.xlsx[i]["L" + str(j+2)].value

					if type(then) == str:

						then = datetime.datetime.strptime(then, "%Y-%m-%d")

					if (then - now).total_seconds() < 0:

						to_check.append([i, self.xlsx[i]["M" + str(j+2)].value, self.xlsx[i]["N" + str(j+2)].value, self.xlsx[i]["L" + str(j+2)].value, self.xlsx[i]["A" + str(j+2)].value])
						to_rmv.append([i, j+2])



		if len(to_check) == 0:
			print("No call or put options to exercise")

		else:

			for i in to_check:
				to = i[3]
				info = data.DataReader(i[0], 'yahoo', to, to)
				info = info.iloc[np.shape(info)[0] - 1]

				if i[2] == "Call":

					if i[4] < info["Low"]:

						print("Buying " + str(i[1]) + " " + i[0] + " shares at " + str(i[4]) + " each.")

						index = np.where(self.stocks == i[0])[0][0]

						self.assets["Debt"] -= i[4] * i[1]
						self.xlsx["Accounts"]["B1"].value -= i[4] * i[1]
						self.shares[index] += i[1]
						self.xlsx["Stocks"]["B" + str(index+2)].value += i[1]


				else:

					if i[4] > info["High"]:

						print("Selling " + str(i[1]) + " " + i[0] + " shares at " + str(i[4]) + " each.")

						index = np.where(self.stocks == i[0])[0][0]

						self.assets["Debt"] += i[4] * i[1]
						self.xlsx["Accounts"]["B1"].value += i[4] * i[1]
						self.shares[index] -= i[1]
						self.xlsx["Stocks"]["B" + str(index+2)].value -= i[1]

						if(self.shares[index] < 0):
							self.assets["Debt"] -= self.shares[index] * info["High"]
							self.xlsx["Accounts"]["B1"].value -= self.shares[index] * info["High"]
							self.shares[index] = 0
							self.xlsx["Stocks"]["B" + str(index+2)].value = 0



			self.update()


			to_rmv = np.array([j for i in to_rmv for j in i])

			for i in self.stocks:

				to_del = np.where(to_rmv == i)[0]

				if len(to_del) != 0:

					to_del += 1
					row_to_del = to_rmv[to_del].astype(int)
					row_to_del[::-1].sort()

					for j in row_to_del:
						self.xlsx[i].delete_rows(j)


			self.xlsx.save(self.file_location)




	def store_option_prices(self, ticker, expiration):

			if len(np.where(self.stocks == ticker)[0]) != 0:

				expiration = datetime.datetime.strptime(expiration, "%Y-%m-%d")

				url = "https://www.barchart.com/stocks/quotes/" + ticker + "/options"

				browser = Chrome()
				browser.get(url)

				print("Webpage loaded.")
				print("")

				time.sleep(2)

				expiration_date_dropdown_menu = browser.find_element_by_class_name("expiration-name")
				date_options = expiration_date_dropdown_menu.find_elements_by_tag_name('option')

				expiration_dates = []

				for i in date_options:

					expiration_dates.append(i.text)

				date_diff = []

				for i in expiration_dates:

					date_diff.append(np.abs((expiration - datetime.datetime.strptime(i, "%Y-%m-%d")).total_seconds()))

				date_diff = np.array(date_diff)
				best = date_diff.argmin()

				choice = expiration_dates[best]

				print("Closest expiration date:")
				print("")

				print(choice)

				url = url + "?expiration=" + choice

				browser.get(url)

				print("Webpage loaded.")
				print("")

				time.sleep(2)


				content = browser.find_element_by_css_selector("body")
				content.send_keys(Keys.CONTROL + "a")
				content.send_keys(Keys.CONTROL + "c")


				win32clipboard.OpenClipboard()
				clipboard_data = win32clipboard.GetClipboardData()
				win32clipboard.EmptyClipboard()
				win32clipboard.CloseClipboard()

				browser.close()
				now = datetime.datetime.now()

				string_list = np.array(clipboard_data.split("\r\n"))

				calls_index = np.where(string_list == "Calls")[0][0]
				puts_index = np.where(string_list == "Puts")[0][0]	
				expiration_index = calls_index - 1


				call_table_string = []
				n = int((puts_index - calls_index)/12)

				frm = calls_index + 2
				to = frm + 11

				for i in range(n):
					call_table_string.append(string_list[frm:to])
					frm = to + 1
					to = frm + 11


				put_table_string = []
				n = 0
				reached_end = False

				while not reached_end:
					reached_end = "Put Volume Tota" in string_list[puts_index + n]
					n += 1

				n = int(n / 12)

				frm = puts_index + 2
				to = frm + 11

				for i in range(n):
					put_table_string.append(string_list[frm:to])
					frm = to + 1
					to = frm + 11

				call_table_columns = string_list[calls_index + 1].split("\t")[0:11]
				put_table_columns = string_list[puts_index + 1].split("\t")[0:11]

				frames_dict = {"Call" : pd.DataFrame(call_table_string, columns = call_table_columns),
								"Put" : pd.DataFrame(put_table_string, columns = put_table_columns)}
		
				n_rows = self.xlsx[ticker + "Calls"].max_row + 1

				for i in range(np.shape(frames_dict["Call"])[0]):
					for j in range(np.shape(frames_dict["Call"])[1]):

						try:
							self.xlsx[ticker + "Calls"][get_column_letter(j+1) + str(i + n_rows)].value = float(frames_dict["Call"].iloc[i,j])
						except:
							self.xlsx[ticker + "Calls"][get_column_letter(j+1) + str(i + n_rows)].value = frames_dict["Call"].iloc[i,j]

				for i in range(np.shape(frames_dict["Call"])[0]):

					self.xlsx[ticker + "Calls"][get_column_letter(np.shape(frames_dict["Call"])[1] + 1) + str(i + n_rows)].value = choice
					self.xlsx[ticker + "Calls"][get_column_letter(np.shape(frames_dict["Call"])[1] + 2) + str(i + n_rows)].value = now.strftime("%Y-%d-%m")

				n_rows = self.xlsx[ticker + "Puts"].max_row + 1

				for i in range(np.shape(frames_dict["Put"])[0]):
					for j in range(np.shape(frames_dict["Put"])[1]):

						try:
							self.xlsx[ticker + "Puts"][get_column_letter(j+1) + str(i + n_rows)].value = float(frames_dict["Put"].iloc[i,j])
						except:
							self.xlsx[ticker + "Puts"][get_column_letter(j+1) + str(i + n_rows)].value = frames_dict["Put"].iloc[i,j]

				for i in range(np.shape(frames_dict["Put"])[0]):

					self.xlsx[ticker + "Puts"][get_column_letter(np.shape(frames_dict["Put"])[1] + 1) + str(i + n_rows)].value = choice
					self.xlsx[ticker + "Puts"][get_column_letter(np.shape(frames_dict["Put"])[1] + 2) + str(i + n_rows)].value = now.strftime("%Y-%d-%m")		


				for j in range(np.shape(frames_dict["Call"])[1]):
					self.xlsx[ticker + "Calls"][get_column_letter(j+1) + str(1)].value = frames_dict["Call"].columns.values[j]

				self.xlsx[ticker + "Calls"][get_column_letter(np.shape(frames_dict["Call"])[1] + 1) + str(1)].value = "Expiration"
				self.xlsx[ticker + "Calls"][get_column_letter(np.shape(frames_dict["Call"])[1] + 2) + str(1)].value = "Date"

				for j in range(np.shape(frames_dict["Put"])[1]):
					self.xlsx[ticker + "Puts"][get_column_letter(j+1) + str(1)].value = frames_dict["Put"].columns.values[j]

				self.xlsx[ticker + "Puts"][get_column_letter(np.shape(frames_dict["Put"])[1] + 1) + str(1)].value = "Expiration"
				self.xlsx[ticker + "Puts"][get_column_letter(np.shape(frames_dict["Put"])[1] + 2) + str(1)].value = "Date"


				self.xlsx.save(self.file_location)

				print("DATA TABLES FOR " + ticker + ":")
				print("")
				print("CALL:")
				print("")
				print(frames_dict["Call"])
				print("")
				print("PUT:")
				print("")
				print(frames_dict["Put"])
				print("")
				print("")


			else:
				print(ticker + " not found within the portfolio.")



	def store_option_prices_list(self, tickers, expirations):

		for i in range(len(tickers)):

			try:
				self.store_option_prices(tickers[i], expirations[i])

			except:
				print("Couldn't get option price for " + tickers[i] + "...")







	def get_call_put_data(self, ticker, option):

		index = np.where(self.stocks == ticker)[0]

		if len(index) != 0 and (option == "Call" or option == "Put"):

			index = index[0]

			n_rows = self.xlsx[ticker + option + "s"].max_row

			if n_rows > 1:

				i = 0

				while  self.xlsx[ticker + option + "s"]["A" + str(n_rows - i)].value > self.xlsx[ticker + option + "s"]["A" + str(n_rows - i - 1)].value and n_rows > i + 3 :

					i += 1

				indices = np.arange(n_rows - i - 1, n_rows+1)

				n_cols = self.xlsx[ticker + option + "s"].max_column

				frame_list = []

				for i in indices:

					row = []

					for j in range(n_cols):

						row.append(self.xlsx[ticker + option + "s"][get_column_letter(j+1) + str(i)].value)

					frame_list.append(row)

				columns = []

				for j in range(n_cols):
					columns.append(self.xlsx[ticker + option + "s"][get_column_letter(j+1) + str(1)].value)

				frame = pd.DataFrame(data = frame_list, columns = columns)

				print("Most recent " + option + " Price Table:")
				print("")
				print(frame)

				return frame

			else:
				print("Error: invalid ticker and/or option (option = Call or Put)")

				return 0


		else:
			print("Error: invalid ticker and/or option (option = Call or Put)")

			return 0


	def buy_call_put_from_data(self, ticker, row, n_purchase, option_type):

		index = np.where(self.stocks == ticker)[0]

		if len(index) != 0:

			index = index[0]

			frame = self.get_call_put_data(ticker, option_type)

			if type(frame) != int:			

				opt = frame.iloc[row]
				option = opt[:-1]


				n = self.xlsx[ticker].max_row + 1

				for i in range(len(option)):

					try:
						self.xlsx[ticker][get_column_letter(i+1) + str(n)].value = float(option[i])

					except:
						self.xlsx[ticker][get_column_letter(i+1) + str(n)].value = option[i]
				
				self.xlsx[ticker][get_column_letter(len(option)+1) + str(n)] = float(n_purchase)
				self.xlsx[ticker][get_column_letter(len(option)+2) + str(n)] = option_type

				self.assets["Debt"] -= float(option["Ask"]) * int(n_purchase)
				self.xlsx["Accounts"]["B1"].value -= float(option["Ask"]) * int(n_purchase)
				self.assets["Net assets"] -= float(option["Ask"]) * int(n_purchase)

				self.xlsx.save(self.file_location)

				print("")
				print("")
				print("")
				print("Purchase completed.")
				print("")
				print("Current Assets:")
				print("")
				print(self.assets)



	def update_historical_data(self, start_date, end_date):

		output = dict()	

		for i in self.stocks:

			try:

				output[i] = data.DataReader(i, 'yahoo', start_date, end_date)
				self.xlsx.remove(self.xlsx[i + "Historical_data"])
				self.xlsx.create_sheet(i + "Historical_data")

				for k in range(np.shape(output[i])[0]):
					for j in range(np.shape(output[i])[1]):

						self.xlsx[i + "Historical_data"][get_column_letter(j+1) + str(k+2)].value = output[i].iloc[k,j]

				for j in range(np.shape(output[i])[1]):
					self.xlsx[i + "Historical_data"][get_column_letter(j+1) + str(1)].value = output[i].columns.values[j]

				m = self.xlsx[i + "Historical_data"].max_column + 1

				for k in range(np.shape(output[i])[0]):
					self.xlsx[i + "Historical_data"][get_column_letter(m) + str(k+2)].value = output[i].index[k]

				self.xlsx[i + "Historical_data"][get_column_letter(m) + str(1)].value = "Date"




			except:

				print(i + " couldn't be loaded...")

		self.xlsx.save(self.file_location)


	def get_historical_close(self):

		frame_list = []
		col_names = []

		for i in self.stocks:

			n_rows = self.xlsx[i + "Historical_data"].max_row

			if n_rows > 1:

				col = np.zeros(n_rows - 1)

				for j in range(n_rows - 1):
					col[j] = self.xlsx[i + "Historical_data"]["D" + str(j + 2)].value

				frame_list.append(col)
				col_names.append(i)

		sizes = np.zeros(len(frame_list))

		j = 0

		for i in frame_list:
			sizes[j] = len(i)
			j += 1

		n = int(np.min(sizes))
		m = len(sizes)

		output = np.zeros(shape = (n,m))

		j = 0
		for i in frame_list:
			output[:,j] = frame_list[j][:n]
			j += 1

		output = pd.DataFrame(data = output, columns = col_names)

		return output



	def get_covariance_matrix_from_close(self):

		return self.get_historical_close().cov()




	def get_historical_returns(self):

		frame_list = []
		col_names = []

		for i in self.stocks:

			n_rows = self.xlsx[i + "Historical_data"].max_row

			if n_rows > 1:

				col = np.zeros(n_rows - 1)

				for j in range(n_rows - 1):
					col[j] = self.xlsx[i + "Historical_data"]["D" + str(j + 2)].value

				frame_list.append(col)
				col_names.append(i)

		sizes = np.zeros(len(frame_list))

		j = 0

		for i in frame_list:
			sizes[j] = len(i)
			j += 1

		n = int(np.min(sizes))
		m = len(sizes) 

		output = np.zeros(shape = (n-1,m))

		j = 0
		for i in frame_list:
			output[:,j] = np.diff(frame_list[j][:n])
			j += 1

		output = pd.DataFrame(data = output, columns = col_names)

		return output




	def get_covariance_matrix_from_returns(self):

		return self.get_historical_returns().cov()


	def get_lowest_variance_pf_weights(self):

		v_cov = self.get_historical_returns().cov()

		l, v = np.linalg.eig(v_cov)

		return v[len(v)-1] / sum(np.abs(v[len(v)-1]))



	def get_eigen_pf_weights(self):

		v_cov = self.get_historical_returns().cov()

		l, v = np.linalg.eig(v_cov)

		return v[0] / sum(np.abs(v[0]))


	def get_sharpe_weights(self):

		data = self.get_historical_returns()

		v_cov = data.cov()

		returns = np.zeros(np.shape(data)[1])

		for i in range(len(returns)):

			returns[i] = np.sum(data.iloc[:,i])

		w = np.zeros(len(returns))

		for i in range(len(w)):
			w[i] = 1

		w /= len(w)

		sharpe_1 = (np.dot(returns, w) / np.sqrt(np.dot(np.transpose(w), np.matmul(v_cov, w))))

		count = 1

		t = 0.1

		for i in range(50*len(w)):

			g = returns - (np.dot(returns, w) / np.dot(np.transpose(w), np.matmul(v_cov, w))) * np.matmul(v_cov, w)
			g /= np.sqrt(np.sum(g**2))

			w += t*g
			t *= 0.99

			if t < 10**(-10):
				break

		t = 0.01

		for i in range(50*len(w)):

			g = returns - (np.dot(returns, w) / np.dot(np.transpose(w), np.matmul(v_cov, w))) * np.matmul(v_cov, w)
			g /= np.sqrt(np.sum(g**2))

			w += t*g
			t *= 0.98

			if t < 10**(-10):
				break	

		return w / sum(np.abs(w))



	def plot_sharpe_pf(self):

		w = self.get_sharpe_weights()

		prices = self.get_historical_close()

		pf = np.matmul(prices, w)
		pf = pd.DataFrame(data = pf, columns = ["Sharpe PF Close Price"])

		print(pf.plot.line())


	def plot_lowest_variance_pf(self):

		w = self.get_lowest_variance_pf_weights()

		prices = self.get_historical_close()

		pf = np.matmul(prices, w)
		pf = pd.DataFrame(data = pf, columns = ["Lowest Variance PF Close Price"])

		print(pf.plot.line())


	def plot_eigen_pf(self):

		w = self.get_eigen_pf_weights()

		prices = self.get_historical_close()

		pf = np.matmul(prices, w)
		pf = pd.DataFrame(data = pf, columns = ["Eigen PF Close Price"])

		print(pf.plot.line())




	def get_Mahalanobis_distances(self):

		data = self.get_historical_returns()

		pca = PCA(n_components=1, svd_solver='full')

		def scaler(matrix):

			for i in range(0, np.shape(matrix)[1]):

				matrix[:,i] = matrix[:,i] / np.std(matrix[:,i])

			return(matrix)


		pca = PCA(svd_solver='full')

		matrix = scaler(pca.fit_transform(np.transpose(data)))

		n = np.shape(matrix)[0]

		adj_matrix = np.empty(shape = (n,n))

		for j in range(0,n):

			for i in range(0,j+1):

				adj_matrix[i,j] = np.linalg.norm(matrix[i,:] - matrix[j,:])
				adj_matrix[j,i] = adj_matrix[i,j]

		adj_matrix = pd.DataFrame(data = adj_matrix, columns = data.columns.values)
		adj_matrix.index = data.columns.values

		return(adj_matrix)







		








	





































