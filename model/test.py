def test_model(df):
    """
    Optional post-processing step after clustering.
    Keeps dataframe unchanged but ensures cluster column exists.
    """

    if "cluster" not in df.columns:
        return df

    # Example: compute cluster statistics
    cluster_summary = df.groupby("cluster").mean(numeric_only=True)

    # Print summary for debugging
    print("\nCluster Summary:")
    print(cluster_summary)

    return df
