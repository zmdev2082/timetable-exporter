import pandas as pd
from timetable_exporter import user_extensions
from timetable_exporter.user_extensions.datetime import combine_date_time, extrapolate_date_ranges
def test_combine_date_time():
    # Create a sample DataFrame
    data = {
        "date": ["2025-04-01", "2025-04-02", "2025-04-03"],
        "time": ["12:00:00", "14:30:00", "16:45:00"]
    }
    df = pd.DataFrame(data)
    datetime_col = "datetime"
    # Call the combine_date_time method
    # combined = df.timetable.combine_date_time("date", "time", datetime_col, tz="UTC")
    combined = getattr(df.timetable,"combine_date_time")("date", "time", datetime_col, tz="UTC")

    # Expected output
    expected_df = pd.DataFrame({
        datetime_col: pd.to_datetime(
            ["2025-04-01 12:00:00", "2025-04-02 14:30:00", "2025-04-03 16:45:00"]
        ).tz_localize("UTC")
    })

    pd.testing.assert_frame_equal(combined, expected_df)

    print("Test passed: combine_date_time works as expected.")

def test_extrapolate_date_ranges():
    date_range_str = "6/3-17/4, 1/5-29/5"
    expected_dates_str = [
        "06/03/2025", "13/03/2025", "20/03/2025", "27/03/2025",
        "03/04/2025", "10/04/2025", "17/04/2025", "01/05/2025",
        "08/05/2025", "15/05/2025", "22/05/2025", "29/05/2025"
    ]
    expected_dates = pd.to_datetime(expected_dates_str, format="%d/%m/%Y").tolist()
    # Call the extrapolate_date_ranges function
    weekly_dates = extrapolate_date_ranges(date_range_str, year=2025, format="%d/%m", frequency="7D")
    
    # Check if the result matches the expected output
    assert weekly_dates == expected_dates, f"Expected {expected_dates}, but got {weekly_dates}"
    
    fortnightly_dates = extrapolate_date_ranges(date_range_str, year=2025, format="%d/%m", frequency="14D")
    print(fortnightly_dates)
    assert fortnightly_dates == expected_dates[0:7:2]+ expected_dates[7::2], f"Expected {expected_dates}, but got {fortnightly_dates}"


    print("Test passed: extrapolate_date_ranges works as expected.")

def test_expand_dates():
    # Create a sample DataFrame
    data = {
        "dates": ["6/3-17/4", "1/5-29/5", "5/5"],
        "other_col": [1, 2, 3]
    }
    df = pd.DataFrame(data)
    
    # Call the expand_dates method via getattr method to test dynamic input from user
    # expanded_df = df.timetable.expand_dates("dates", year=2025, format="%d/%m")
    expanded_df = getattr(df.timetable,"expand_dates")("dates", year=2025, format="%d/%m")

    # Expected output
    expected_data = {
        "dates": [
            pd.Timestamp("2025-03-06"), pd.Timestamp("2025-03-13"),
            pd.Timestamp("2025-03-20"), pd.Timestamp("2025-03-27"),
            pd.Timestamp("2025-04-03"), pd.Timestamp("2025-04-10"),
            pd.Timestamp("2025-04-17"), pd.Timestamp("2025-05-01"),
            pd.Timestamp("2025-05-08"), pd.Timestamp("2025-05-15"),
            pd.Timestamp("2025-05-22"), pd.Timestamp("2025-05-29"),
            pd.Timestamp("2025-05-05")
        ],
        "other_col": [1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3]
    }
    expected_df = pd.DataFrame(expected_data)
    
    # Check if the result matches the expected output
    pd.testing.assert_frame_equal(expanded_df, expected_df)
    print("Test passed: expand_dates works as expected.")


def test_exclude_filters_contains():
    df = pd.DataFrame(
        {
            "summary": [
                "AMME2500-S1C-ND-CC/Practical-IGN/64",
                "AMME2500-S1C-ND-CC/Practical/63",
                "MTRX2700-S1C-ND-CC/Practical_MP_P1/03",
            ]
        }
    )

    out = df.timetable.exclude({"summary": "IGN"}, exact_match=False)
    assert len(out) == 2
    assert out["summary"].str.contains("IGN").sum() == 0
