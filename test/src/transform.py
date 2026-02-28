import pandas as pd
import pstats

def time_complexity_to_df(path="time_complexity.prof"):

    p = pstats.Stats(path)
    
    rows = []
    for func, stat in p.stats.items():
        filename, lineno, funcname = func
        rows.append({
            "func": funcname,
            "file": filename.split("/")[-1],
            "tottime": stat[2],
            "cumtime": stat[3],
            "ncalls": stat[0],
        })
    
    return pd.DataFrame(rows).sort_values("tottime", ascending=False)

