from sklearn.neighbors import LocalOutlierFactor


def detect_hidden_patterns(X):

    lof = LocalOutlierFactor(n_neighbors=20)

    labels = lof.fit_predict(X)

    anomaly_wells = (labels == -1).sum()

    rate = anomaly_wells / len(labels)

    return labels, rate
