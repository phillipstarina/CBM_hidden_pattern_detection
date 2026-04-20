import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D


def reservoir_3d(df):

    fig = plt.figure(figsize=(5, 4))

    ax = fig.add_subplot(111, projection="3d")

    numeric = df.select_dtypes(include="number")

    if numeric.shape[1] < 3:
        ax.text(0.2, 0.5, 0.5, "Need 3 numeric columns")
        return fig

    x = numeric.iloc[:, 0]
    y = numeric.iloc[:, 1]
    z = numeric.iloc[:, 2]

    clusters = df["cluster"]

    ax.scatter(x, y, z, c=clusters, cmap="viridis")

    ax.set_xlabel(numeric.columns[0])
    ax.set_ylabel(numeric.columns[1])
    ax.set_zlabel(numeric.columns[2])

    ax.set_title("3D Reservoir Visualization")

    return fig
