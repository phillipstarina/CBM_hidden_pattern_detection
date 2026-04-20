import matplotlib.pyplot as plt
import numpy as np

def forecast_plot(actual, forecast):

    fig = plt.figure(figsize=(5,4))

    actual = np.array(actual)

    plt.plot(actual, label="Actual Production")

    future_index = np.arange(len(actual), len(actual)+len(forecast))

    plt.plot(future_index, forecast, label="Forecast Production")

    plt.title("CBM Production Forecast")

    plt.xlabel("Time")

    plt.ylabel("Production")

    plt.legend()

    plt.grid(True)

    return fig
