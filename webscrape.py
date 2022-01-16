from traceback import print_tb
import requests
from bs4 import BeautifulSoup 
#Fetching Data from Websites
headers = { "user-Agent": 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36'}
amazon_url = 'https://www.amazon.in/Fire-Boltt-Smartwatch-Workout-Tracking-extensive/dp/B09NM8S1V4/ref=psdc_5605728031_t1_B0972BQ2RS'

page = requests.get(amazon_url,headers= headers)
soup = BeautifulSoup(page.content,"lxml")
amazon_title = soup.find( 'span', id = "productTitle").get_text()
print("Amazon Product : ",amazon_title)

amazon_price = soup.find('span', class_= "a-offscreen").get_text()
print("Amazon Product Price: ",amazon_price)

flipkart_product_url = "https://www.flipkart.com/fire-boltt-ninja-touch-wake-spo2-smartwatch/p/itm05b47a90c7c32"

page = requests.get(flipkart_product_url, headers=headers)
soup= BeautifulSoup(page.content,"lxml")
flipkart_title = soup.find('span',class_ = "B_NuCI").get_text()
print("Flipkart Product : ",flipkart_title)

flipkart_price = soup.find('div',class_ = "_30jeq3 _16Jk6d").get_text()
print("FlipKart Product Price: ", flipkart_price)

vijaysales_url = "https://www.vijaysales.com/amazfit-bip-u-pro-smart-watch-green-with-spo2-measurement-24-x-7-heart-rate-tracking-60-built-in-sport-modes/16805"
# vijaysales_url = "https://www.vijaysales.com/vivo-y21-4-gb-ram-64-gb-rom-diamond-glow/17179"


#Fetching data Vijay_Sales
page = requests.get(vijaysales_url, headers= headers)
soup = BeautifulSoup(page.content, "lxml")
vijaysalestitle= soup.find(id="ContentPlaceHolder1_h1ProductTitle", itemprop="name").get_text()
print("Vijay Sales Product : ",vijaysalestitle)

vijaysales_price = soup.find('div',class_="priceMRP").select("span:nth-child(3) > span")[0].text
print("Vijay Sales Product Price: "+ "₹" +vijaysales_price)