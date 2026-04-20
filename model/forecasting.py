
import numpy as np
from sklearn.linear_model import LinearRegression

def forecast_production(df):

    numeric = df.select_dtypes(include="number")

    if numeric.shape[1] == 0:
        return np.zeros(30)

    production = numeric.iloc[:,0]

    X = np.arange(len(production)).reshape(-1,1)

    model = LinearRegression()

    model.fit(X, production)

    future_steps = np.arange(len(production), len(production)+30).reshape(-1,1)

    prediction = model.predict(future_steps)

    return prediction