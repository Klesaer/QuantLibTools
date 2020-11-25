import  QuantLib as ql
import pandas as pd
import datetime as dt

# input_file
inputfile = r"D:\PythonProject\Quantlib\Input.xlsx "

input_dataframe = pd.read_excel(inputfile, sheet_name="EuropeanOption")
# load those inputs
model_name = input_dataframe.iloc[0]['Value']
OptionType = input_dataframe.iloc[1]['Value']
val_date_str = input_dataframe.iloc[2]['Value']
maturity_str = input_dataframe.iloc[3]['Value']
underlying = input_dataframe.iloc[4]['Value']
strike = input_dataframe.iloc[5]['Value']
risk_free_rate = input_dataframe.iloc[6]['Value']
implied_vola = input_dataframe.iloc[7]['Value']

# change format and load dates:
val_date_d = dt.datetime.strptime(val_date_str, "%d.%m.%Y")
val_date = ql.Date(val_date_d.day, val_date_d.month, val_date_d.year)
ql.Settings_instance().evaluationDate = val_date

maturity_d = dt.datetime.strptime(maturity_str, "%d.%m.%Y")
maturity = ql.Date(maturity_d.day, maturity_d.month, maturity_d.year)
# set other inputs:
u = ql.SimpleQuote(underlying)
r = ql.SimpleQuote(risk_free_rate)
sigma = ql.SimpleQuote(implied_vola)

# set OptionType
if OptionType == "Call":
    option = ql.EuropeanOption(ql.PlainVanillaPayoff(ql.Option.Call, strike), ql.EuropeanExercise(maturity))
elif OptionType == "Put":
    option = ql.EuropeanOption(ql.PlainVanillaPayoff(ql.Option.Put, strike), ql.EuropeanExercise(maturity))
else:
    print("OptionType is not defined correct")
    exit()

# set model
if model_name != "BS":
    print("Model is not setup yet. Please check.")
    exit()

# set discount curve:
riskFreeCurve = ql.FlatForward(0, ql.TARGET(), ql.QuoteHandle(r), ql.Actual360())
volatility = ql.BlackConstantVol(0, ql.TARGET(), ql.QuoteHandle(sigma), ql.Actual360())

process = ql.BlackScholesProcess(ql.QuoteHandle(u),
                                 ql.YieldTermStructureHandle(riskFreeCurve),
                                 ql.BlackVolTermStructureHandle(volatility))
engine = ql.AnalyticEuropeanEngine(process)
option.setPricingEngine(engine)

# fill results:
input_dataframe.iloc[8]['Value'] = option.NPV()
input_dataframe.iloc[9]['Value'] = option.delta()
input_dataframe.iloc[10]['Value'] = option.vega()
input_dataframe.iloc[11]['Value'] = option.gamma()
input_dataframe.iloc[12]['Value'] = option.rho()
input_dataframe.iloc[13]['Value'] = option.theta()
input_dataframe.iloc[14]['Value'] = option.itmCashProbability()

track_time = dt.datetime.now().strftime("%d-%m-%Y (%H-%M-%S)")
result = "D:\\PythonProject\\Quantlib\\" + track_time + "_results_european option.xlsx"
input_dataframe.to_excel(result, index=None)


## output format
#Items	Value
#model	BS
#OptionType	Put
#Val_date	25.11.2020
#Maturity	19.03.2021
#Underlying	13289.80
#Strike	13100.00
#RiskFreeRate	-0.0045600
#implied_vola	0.2080950
#NPV
#delta
#vega
#gamma
#rho
#theta
#itmCashProbability
##
