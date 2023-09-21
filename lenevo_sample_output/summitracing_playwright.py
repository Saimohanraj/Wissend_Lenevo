import os
import re
import csv
import json
import asyncio
import pandas as pd
from datetime import datetime
from parsel import Selector
from playwright.async_api import async_playwright



async def main():  
    with open("Input.csv", mode="r", encoding="utf-8") as file:
        csvFile = list(csv.DictReader(file))
    for url in csvFile:
        item = {}
        item["brand"] = url.get("Brand")
        item["sku"] = url.get("SKU")
        keyword = item['brand'] + ' '+ item['sku']
        try:
            async with async_playwright() as playwright:
                chromium = playwright.firefox
                browser = await chromium.launch(
                    args=["--start-maximized"], headless=False
                )
                page = await browser.new_page(no_viewport=True,java_script_enabled=True)
                page.set_default_timeout(0)
                await page.goto(f"https://www.summitracing.com/search?SortBy=BestKeywordMatch&SortOrder=Ascending&keyword={keyword}",wait_until="load")
                res1 = Selector(text=await page.content())
                value = res1.xpath('//script[@id="google-analytics-4-products"]/text()').get('').strip()
                if value:
                    json_data = json.loads(value)
                    for i in json_data:
                        item_id = i.get('item_id')
                        item_brand = i.get('item_brand')                   
                        if item['brand'] in item_brand and item['sku'] in item_id:
                            checking = res1.xpath(f'//div[@data-prodid="{item_id}"]')
                            if checking:
                                item['product_url']='https://www.summitracing.com'+res1.xpath(f'//div[@data-prodid="{item_id}"]//h2/a/@href').get('').strip()
                                item['title']=res1.xpath(f'//div[@data-prodid="{item_id}"]//h2/a/text()').get('').strip()
                            else:
                                item['product_url']=''
                                item['title']=''
                            item['input_brand'] = item['brand']
                            item['input_sku'] = item['sku']
                            item['domain_url'] = 'https://www.summitracing.com'
                            sales=i.get('price','')
                            item['Sale Price'] ='$'+i.get('price') if sales else ''
                            
                            item['Scrape Price'] = item['Sale Price']
                            item['Status'] = 'Found'
                            item['Crawl Timestamp']=datetime.now().strftime('%d-%m-%Y %I:%M %p')
                            
                            if not os.path.exists("html_file"):
                                os.makedirs("html_file")
                            with open(f'./html_file/{item["sku"]}.html','w')as f:
                                f.write(await page.content())
                            item['file_path'] = f'{item["sku"]}.html'
                            print('>>>>Output------->',item)
                            df = pd.DataFrame([item])
                            if not os.path.isfile("summitracing.csv"):
                                df.to_csv(
                                    "summitracing.csv",
                                    index=False,
                                    mode="a",
                                    header=True,
                                    encoding="utf_8_sig",
                                )
                            else:  # else it exists so append without writing the header
                                df.to_csv(
                                    "summitracing.csv",
                                    mode="a",
                                    header=False,
                                    index=False,
                                    encoding="utf_8_sig",
                                )
                else:
                    with open('not_found.txt','a') as f:
                        f.write(str(keyword) +'\n')    
                
                await page.close()
                await browser.close()
                await playwright.stop()
                

        except Exception as e:
            print(e)


asyncio.run(main())

