from sklearn.decomposition import PCA
import matplotlib.pyplot as plt


def pca_cluster_plot(X, clusters):

    pca = PCA(n_components=2)

    components = pca.fit_transform(X)

    fig = plt.figure(figsize=(5, 4))

    plt.scatter(
        components[:, 0],
        components[:, 1],
        c=clusters,
        cmap="viridis"
    )

    plt.title("PCA Cluster Visualization")

    plt.xlabel("PC1")

    plt.ylabel("PC2")

    plt.grid(True)

    return fig
