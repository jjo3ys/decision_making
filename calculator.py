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
    o_c_s = om.iloc[48:56, week+1].to_numpy()#Other Measures시트의 판매량/생산가능량 비
    comprod = om.iloc[37:45, week+1].to_numpy()#Other Measures시트의 component주문량/제품생산량 비
    inventory = om.iloc[26:34, week+1].to_numpy()#Other Measures시트의 총 재고비용/Revenue

    for i in range(len(TEAM)):
        apple_demand = round(demand[0]*450*adms[i*3])#i번째 팀이 apple에 할당할 수 있는 량  (apple 수요량)
        sony_demand = round(demand[1]*450*adms[i*3+1])#i번째 팀이 sony에 할당할 수 있는 량  (sony 수요량)

        apple_sales = round(apple_demand*apple_satisf[i])
        sony_sales = round(sony_demand*sony_satisf[i])

        max_capacity = round((apple_sales + sony_sales)/o_c_s[i])

        revenue = round(apple_sales*price[i*3] + sony_sales*price[i*3+1])

        

        inventory_cost = round(revenue*inventory[i]/33.33, 3)

        if inventory_cost == 0:
            manufactured = apple_sales+sony_sales
            components = manufactured*comprod[i]

            print("{0}팀의\napple 모듈 판매량:{1}개, sony 모듈 판매량:{2}개, 이번주 생산량:{3}개, 최대 생산량:{4}개\ncomponent 주문량:{5}개\nRevenue:${6}, 총 재고비용:${7}\n"
            .format(TEAM[i],
                    apple_sales, 
                    sony_sales, 
                    manufactured, 
                    max_capacity,
                    round(components), 
                    revenue,
                    inventory_cost)
                    )

        elif inventory_cost%(demand[2]*0.1) == 0:
            manufactured = apple_sales+sony_sales
            components = manufactured*comprod[i]
            inventory_count = round(inventory_cost/(demand[2]*0.1*0.03))

            print("{0}팀의\napple 모듈 판매량:{1}개, sony 모듈 판매량:{2}개, 이번주 생산량:{3}개, 최대 생산량:{4}\ncomponent 재고량:{5}개, 주문량:{6}개\nRevenue:${7}, 총 재고비용:${8}\n"
            .format(TEAM[i],
                    apple_sales, 
                    sony_sales, 
                    manufactured, 
                    max_capacity,
                    inventory_count,
                    round(components-inventory_count), 
                    revenue,
                    inventory_cost)
                    )
        else:

            print("{0}팀의\napple 모듈 판매량:{1}개, sony 모듈 판매량:{2}개, 최대 생산량:{3}개\n최대 생산량을 바탕으로 한 component 주문량+재고량:{4}개\nRevenue:${5} 총 재고비용:${6}\n"
            .format(TEAM[i],
                    apple_sales, 
                    sony_sales, 
                    max_capacity, 
                    round(max_capacity*comprod[i]), 
                    revenue,
                    inventory_cost)
                    )
    
main()
