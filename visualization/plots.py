import matplotlib.pyplot as plt
import numpy as np

def cluster_plot(df):

    fig = plt.figure(figsize=(5,4))

    ax = fig.add_subplot(111)

    numeric = df.select_dtypes(include=np.number)

    if numeric.shape[1] < 2:
        ax.text(0.3,0.5,"Not enough numeric columns for plotting")
        return fig

    x = numeric.iloc[:,0]
    y = numeric.iloc[:,1]

    clusters = df["cluster"]

    scatter = ax.scatter(x, y, c=clusters, cmap="viridis")

    ax.set_xlabel(numeric.columns[0])
    ax.set_ylabel(numeric.columns[1])
    ax.set_title("CBM Production Pattern Clusters")

    ax.grid(True)

    return fig
