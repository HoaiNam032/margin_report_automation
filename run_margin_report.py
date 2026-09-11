import plotly.express as px
import pandas as pd

df = pd.DataFrame({
    "x": [1, 2, 3],
    "y": [10, 20, 15]
})

fig = px.line(df, x="x", y="y", title="Plotly Offline Test")
fig.write_html("test_plotly.html", auto_open=True)