import matplotlib.pyplot as plt


def production_curve(df):

    fig = plt.figure(figsize=(5, 4))

    for cluster in df["cluster"].unique():

        subset = df[df["cluster"] == cluster]

        plt.plot(subset.index, subset.iloc[:, 0], label=f"Cluster {cluster}")

    plt.title("Well Production Curves")

    plt.xlabel("Time")

    plt.ylabel("Production")

    plt.legend()

    plt.grid(True)

    return fig
