
import pickle
import os
from sklearn.cluster import KMeans
from model.preprocess import preprocess_data


def train_model(df, clusters):

    X, feature_cols = preprocess_data(df)

    model = KMeans(n_clusters=clusters)

    labels = model.fit_predict(X)

    os.makedirs("models", exist_ok=True)

    model_data = {
        "model": model,
        "features": feature_cols
    }

    with open("models/kmeans_model.pkl", "wb") as f:
        pickle.dump(model_data, f)

    df["cluster"] = labels

    return df
