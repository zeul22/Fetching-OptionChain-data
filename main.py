import requests
import pandas_datareader
import pandas as pd
import xlwings as xw
import time


url="https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY"
headers={
 "accept-encoding": "gzip, deflate, br",
"accept-language": "en-US,en;q=0.7",
"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
}

session=requests.Session()
data=session.get(url,headers=headers).json()["records"]["data"]
# print(data)
ocData=[]
for i in data:
    for j,k in i.items():
        if j=="CE" or j=="PE":
            info=k
            info["instrumentType"]=j
            ocData.append(info)
# print(ocData)

df=pd.DataFrame(ocData)
# print(df)
wb=xw.Book("optionChain.xlsx")
st=wb.sheets("Bn")
st.range("A1").value=df
print("Update")
wb.save()
time.sleep(5)