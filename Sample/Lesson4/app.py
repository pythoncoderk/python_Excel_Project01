import math
import streamlit as st
import pandas as pd
import altair as alt

st.title('予想利益シミュレーション')
st.sidebar.write('## 入力フォーム')
ad_cost = st.sidebar.slider('広告宣伝費（万円）', 1000, 9000)*1.0E+04
fixed_cost = 1000*1.0E+04
cost = ad_cost+fixed_cost

def calc_earnings(ad_cost):
    earnings = 2.87E+07*math.log(ad_cost)-4.44E+08
    return int(earnings)

def calc_profit(earnings, cost):
    profit = earnings-cost
    return int(profit)

earnings = calc_earnings(ad_cost)
profit = calc_profit(earnings, cost)
profit_ratio = int((profit/earnings)*100)

data_ad_cost = list(range(1000, 9001, 1))
data_earnings = [calc_earnings(ad_cost*1.0E+04) for ad_cost in data_ad_cost]
data_profit = [calc_profit(earnings, ad_cost*1.0E+04+fixed_cost) 
               for earnings, ad_cost in zip(data_earnings, data_ad_cost)]

max_profit = max(data_profit)
best_ad_cost = data_ad_cost[data_profit.index(max_profit)]

col1, col2, col3 = st.columns(3)
col1.metric('費用', f'{int(cost/1.0E+04)} 万円')
col1.metric('最適な広告投下費用', f'{best_ad_cost} 万円')
col2.metric('予想売上', f'{int(earnings/1.0E+04)} 万円')
col3.metric('予想利益', f'{int(profit/1.0E+04)} 万円', f'{profit_ratio}%')
col3.metric('予想最大利益', f'{int(max_profit/1.0E+04)} 万円')

df_earnings = pd.DataFrame()
df_earnings['ad_cost'] = data_ad_cost
df_earnings['value'] = data_earnings
df_earnings['indicator'] = '売上'

df_profit = pd.DataFrame()
df_profit['ad_cost'] = data_ad_cost
df_profit['value'] = data_profit
df_profit['indicator'] = '利益'

df = pd.concat([df_earnings, df_profit])
df['value'] = df['value']/1.0E+04

chart = alt.Chart(df).mark_line().encode(
    alt.X('ad_cost', title='広告宣伝費（万円）'),
    alt.Y('value', title='売上 & 利益（万円）'),
    color='indicator'
).configure_axis(
    labelFontSize=12,
    titleFontSize=16,
).configure_legend(
    titleFontSize=12,
    labelFontSize=16,
) 

st.write('## 広告宣伝費に応じた予測売上・利益の推移')
st.altair_chart(chart, use_container_width=True)