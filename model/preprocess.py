import numpy as np
from sklearn.preprocessing import StandardScaler


def preprocess_data(df):

    df = df.select_dtypes(include=np.number)

    df = df.fillna(df.mean())

    feature_cols = df.columns

    scaler = StandardScaler()

    X = scaler.fit_transform(df)

    return X, feature_cols
