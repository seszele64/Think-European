# import
import pandas as pd
import plotly.express as px

# read data, skip first two rows
df = pd.read_excel('Crude Ural.xlsx', sheet_name='Table Data', skiprows=2, index_col=0)

# index to datetime
df.index = pd.to_datetime(df.index)

# rename index to 'Date'
df.index.name = 'Date'

# rename first column to 'Date', second to 'Close Price'
df.columns = ['Close Price']
print(df.head(10))

# create chart
fig = px.line(df, x=df.index, y= df['Close Price'], title='Crude Ural Price - last 10 years')

# change hover tooltip to "Date" DD/MM/YYYY and "Close Price" $0.00
fig.update_traces(hovertemplate='<b>Date</b>: %{x|%d/%m/%Y}<br><b>Close Price</b>: $%{y:.2f}')

# calculate 10 year average
df['10 Year Average'] = df['Close Price'].mean()

# add 10 year average line, color red, show its value above the line at the left, font color
fig.add_hline(y=df['10 Year Average'].mean(), line_dash="dash", line_color="red", annotation_text="10 Year Average: {}".format(df['10 Year Average'][0].round(2)), annotation_position="top left")

# show the value of 10 year average on the chart

print(df['10 Year Average'])

# add "data source" annotation
fig.add_annotation(text="Data Source: Refinitiv", xref="paper", yref="paper", x=1, y=0, showarrow=False)

# add logo annotation on the left bottom, font = EB Garamond, size = 20
fig.add_annotation(text = "thinkeuropean.org", xref="paper", yref="paper", x=0, y=0, showarrow=False, font=dict(family="EB Garamond", size=15))

# show chart
fig.show()

# export chart to html, use cdn for offline use
fig.write_html('Crude Ural.html', include_plotlyjs='cdn')