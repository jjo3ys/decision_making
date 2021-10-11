import pandas as pd
import numpy as np

TEAM = ['Decide out', '결정체', '결정해듀오', '기맥', '산수', '일구', '토마토맛토마토', '해조']

def main():
    
    week = int(input("week?:"))
    filename = 'Weekly_Results_주간 (Week{0}).xlsx'.format(week)

    om = pd.read_excel(filename, sheet_name='Other Measures', header=None)
    ms = pd.read_excel(filename, sheet_name='Market Share', header=None)

    demand = pd.read_excel(filename, sheet_name='Price', header=None).iloc[5:8, week+1].to_numpy()#Price 시트의 수요량

    adms = ms.iloc[69:92, week+2].to_numpy()#Market Share시트의 Adjusted market Shares
    price = ms.iloc[6:29, week+2].to_numpy()#Market Share시트의 각 모듈 Price

    apple_satisf = om.iloc[59:67, week+1].to_numpy()#Other Measures시트의 apple 충족률         충족률 = 판매량/수요량
    sony_satisf = om.iloc[70:78, week+1].to_numpy()#Other Measures시트의 sony 충족률
    o_c_s = om.iloc[48:56, week+1].to_numpy()#Other Measures시트의 판매/생산 비
    comprod = om.iloc[37:45, week+1].to_numpy()#Other Measures시트의 component주문량/제품생산량 비

    for i in range(len(TEAM)):
        apple_demand = round(demand[0]*450*adms[i*3])#i번째 팀이 apple에 할당할 수 있는 량  (apple 수요량)
        sony_demand = round(demand[1]*450*adms[i*3+1])#i번째 팀이 sony에 할당할 수 있는 량  (sony 수요량)

        apple_sales = round(apple_demand*apple_satisf[i])
        sony_sales = round(sony_demand*sony_satisf[i])

        capacity = round((apple_sales + sony_sales)/o_c_s[i])

        revenue = round(apple_sales*price[i*3] + sony_sales*price[i*3+1])

        components = capacity*comprod[i]

        print("{0}팀의\napple 모듈 판매량:{1}, sony 모듈 판매량:{2}, 최대 생산량:{3}\n최대 생산량을 바탕으로 한 component 주문량:{4}\nRevenue:${5}\n"
        .format(TEAM[i],
                apple_sales, 
                sony_sales, 
                capacity, 
                components, 
                revenue)
                )
    
main()
