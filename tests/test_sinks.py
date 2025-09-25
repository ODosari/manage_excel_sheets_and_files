import pandas as pd

from excelmgr.core.sinks import csv_sink


def test_csv_sink_aligns_columns(tmp_path):
    out = tmp_path / "out.csv"
    df1 = pd.DataFrame({"A": [1], "B": [2]})
    df2 = pd.DataFrame({"B": [3], "A": [4]})

    with csv_sink(str(out)) as sink:
        sink.append(df1)
        sink.append(df2)

    contents = out.read_text().strip().splitlines()
    assert contents[0] == "A,B"
    assert contents[1] == "1,2"
    assert contents[2] == "4,3"
