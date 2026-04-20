from sklearn.cluster import KMeans


def find_optimal_clusters(X):

    distortions = []

    for k in range(1, 10):

        model = KMeans(n_clusters=k)

        model.fit(X)

        distortions.append(model.inertia_)

    best_k = distortions.index(min(distortions[1:])) + 2

    return best_k
