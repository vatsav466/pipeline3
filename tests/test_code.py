import os
import pandas as pd
from code import generate_report

def test_generate_report_creates_csv():
    output = generate_report(
        "data/file1.csv",
        "data/file2.csv",
        "test_report.csv"
    )

    assert os.path.exists(output)

    df = pd.read_csv(output)
    assert not df.empty

    os.remove(output)

