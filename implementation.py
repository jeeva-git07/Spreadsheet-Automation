import pandas as pd

file_path = "file_name with extension"
raw=pd.read_csv(file_path)
# Load data
data = pd.read_csv(file_path, skiprows=2)

# Filter Teams and IT Center
def csat_survey(summary):
    fil = data[
    (data["Team"].isin(["Endpoint", "Network Services", "Server And Datacenter"])) &
    (data["IT Center"].isin(summary))
    ]

    # Pivot table: counts of Survey RatingRest
    pivot = fil.pivot_table(
        index="Team",
        columns="Survey RatingRest",
        aggfunc="size",
        fill_value=0
    )

    # Add row totals
    pivot["Grand Total"] = pivot.sum(axis=1)
    grnd = pivot.copy()

    # Add column totals (Excel-style)
    grand_total = pivot.sum(axis=0).to_frame().T
    grand_total.index = ["Grand Total"]

    # Final pivot (row + col totals)
    pivot = pd.concat([pivot, grand_total])

    # print(pivot)

    # Ensure column names are strings
    pivot.columns = pivot.columns.astype(str)

    # Identify numeric rating columns
    rating_cols = [c for c in pivot.columns if c.isnumeric()]

    # Build summary table
    summary = pd.DataFrame(index=pivot.index)

    # --- SUM = weighted total (count × rating value) ---
    summary["SUM"] = pivot[rating_cols].apply(
        lambda row: sum(int(c) * row[c] for c in rating_cols),
        axis=1
    )

    # Feedback Received = sum of rating counts
    summary["Feedback Received"] = pivot[rating_cols].sum(axis=1)

    # Feedback Requested = Grand Total - Feedback Received
    summary["Feedback Requested"] = grnd["Grand Total"]

    # ✅ Response % = Feedback Received / (Feedback Requested + Feedback Received)
    denom = summary["Feedback Requested"] + summary["Feedback Received"]
    summary["Response %"] = (
        (summary["Feedback Received"] / denom.replace({0: pd.NA})) * 100
    ).round(1).fillna(0)

    # --- Rating = weighted average ---
    def weighted_avg(row):
        total = 0
        count = 0
        for c in rating_cols:
            val = row[c]
            total += val * int(c)
            count += val
        return round(total / count, 1) if count else 0

    summary["Rating"] = pivot.apply(weighted_avg, axis=1)

    # Add constants
    summary["Low Rating"] = 4
    summary["Low Response"] = 15

    # --- Build final summary ---
    # Drop the first "Grand Total" row from pivot; recalc new one from summary
    body = summary.drop(index="Grand Total", errors="ignore")

    # Grand Total row (Excel-style)
    grand = pd.DataFrame(index=["Grand Total"])
    grand["SUM"] = body["SUM"].sum()
    grand["Feedback Received"] = body["Feedback Received"].sum()
    grand["Feedback Requested"] = body["Feedback Requested"].sum()

    denom = grand["Feedback Requested"] + grand["Feedback Received"]
    grand["Response %"] = (grand["Feedback Received"] / denom * 100).round(1)
    grand["Rating"] = (body["SUM"].sum() / body["Feedback Received"].sum()).round(1) if body["Feedback Received"].sum() else 0
    grand["Low Rating"] = 4
    grand["Low Response"] = 15

    # Final concat
    final_summary = pd.concat([body, grand])

    # Reorder columns
    final_summary = final_summary[
        ["SUM", "Response %", "Rating", "Low Rating", "Low Response", "Feedback Requested", "Feedback Received"]
    ]

    # Clean types
    int_cols = ["SUM", "Feedback Requested", "Feedback Received"]
    for col in int_cols:
        final_summary[col] = final_summary[col].astype(int)

    # print(final_summary.reset_index(drop=True))
    return final_summary.reset_index(drop=True),pivot

#Creating the pivot table and summary of the sites
summary_asia,pivot_asia=csat_survey(["IN","JP","KR","TH"])
summary_china,pivot_china=csat_survey(["CN"])
with pd.ExcelWriter("CSAT_survey.xlsx",engine="xlsxwriter") as writer:
    raw.to_excel(writer,sheet_name="survey",index=False)
    pivot_asia.to_excel(writer,sheet_name="asia pivot_table",startrow=0,startcol=0)
    summary_asia.to_excel(writer,sheet_name="asia pivot_table",startrow=0,startcol=len(pivot_asia.columns)+4,index=False)
    pivot_china.to_excel(writer,sheet_name="china pivot_table",startrow=0,startcol=0)
    summary_china.to_excel(writer,sheet_name="china pivot_table",startrow=0,startcol=len(pivot_china.columns)+4,index=False)