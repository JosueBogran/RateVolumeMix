import pandas as pd

#Name of Excel workbook
workbook = "ProcessedSalesData.xlsx"

#Import workbook into dataframe
df = pd.read_excel(workbook, converters={'Volume':float,'Revenue':float})  

#Initial look of data.
df.head()

#Determine the actual rate (or average price) for period in question, for the item in question.
df['Actual Rate'] = (df['Revenue']  / df['Volume']).fillna(0)

#Sum the total amount of units(volume) and rate for period. Used for calculation purposes.
RateVolTotals = df.groupby(['Period'])[['Volume','Revenue']].sum().reset_index()

#Merge RateVolTotals unto df and do some cleanup
df = df.merge(RateVolTotals, on='Period', how='left')
df = df.rename(columns={'Volume_x':'Actual Volume','Volume_y':'PeriodTotalVolume','Revenue_x':'Revenue','Revenue_y':'PeriodTotalRevenue'})

#Calculate the actual mix for the product/period in question.
df['Actual Mix'] = (df['Actual Volume'] / df['PeriodTotalVolume']).fillna(0)

#Calculate the aggreate rate. Used for Mix calculation.
df['Aggregate Rate'] = (df['PeriodTotalRevenue'] / df['PeriodTotalVolume']).fillna(0)

#Set the targets. In this case, we are comparing month over month variance. Some are used strictly for other calculations.
df['Target Rate'] = (df.groupby(['Product'])['Actual Rate'].shift()).fillna(0)
df['Target Volume'] = (df.groupby(['Product'])['Actual Volume'].shift()).fillna(0)
df['Target Mix'] = (df.groupby(['Product'])['Actual Mix'].shift()).fillna(0)
df['Target Revenue'] = (df.groupby(['Product'])['Revenue'].shift()).fillna(0)
df['Target Aggregate Rate'] = (df.groupby(['Product'])['Aggregate Rate'].shift()).fillna(0)

#Calculate the variance attributed to rate changes
df['Rate Variance'] = ((df['Actual Rate'] - df['Target Rate'])*df['Actual Volume']).fillna(0)

#Calculate the variance attributed to product mix changes
df['Mix Variance'] = df['PeriodTotalVolume']*(df['Target Rate'] - df['Target Aggregate Rate']) * (df['Actual Mix'] - df['Target Mix']).fillna(0)

#Calculate the variance attributed to volume changes
df['Volume Variance'] = (df['Target Rate'] * (df['Actual Volume'] - df['Target Volume']) - df['Mix Variance']).fillna(0)

#Calculations to add clarity. Not using Mix Change in the final product, but could be helpful if knowing % change in mix was desired.
df['Revenue Change'] = df['Revenue'] - df['Target Revenue']
df['Volume Change'] = df['Actual Volume'] - df['Target Volume']
df['Mix Change'] = df['Actual Mix'] - df['Target Mix']

#Cleanup 
df = df[['Period', 'Product', 'Category', 'Revenue', 'Target Revenue', 'Revenue Change', 'Actual Volume', 'Target Volume', 'Volume Change', 'Rate Variance', 'Volume Variance', 'Mix Variance']]

#Preview
df.head()

#Export/aggregate/etc
