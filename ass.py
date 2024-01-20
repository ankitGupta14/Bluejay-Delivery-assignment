import pandas as pd
from datetime import datetime, timedelta
data = pd.read_excel('Assignment_Timecard.xlsx')
df = pd.DataFrame(data)

df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')
df['Shift Duration'] = df['Time Out'] - df['Time']

# a) Print employees who have worked for 7 consecutive days
consecutive_days = 7
for name, group in df.groupby('Employee Name'):
    consecutive_count = group['Time'].diff().dt.days.eq(1).cumsum()
    if consecutive_count.max() >= consecutive_days - 1:
        print(f"Employee {name} has worked for {consecutive_days} consecutive days.")

# b) Print employees who have less than 10 hours of time between shifts but greater than 1 hour
min_hours_between_shifts = 1
max_hours_between_shifts = 10
for name, group in df.groupby('Employee Name'):
    time_between_shifts = group['Time'].diff().dt.total_seconds() / 3600
    mask = (time_between_shifts < max_hours_between_shifts) & (time_between_shifts > min_hours_between_shifts)
    if mask.any():
        print(f"Employee {name} has less than {max_hours_between_shifts} hours but more than {min_hours_between_shifts} hours between shifts.")

# c) Print employees who have worked for more than 14 hours in a single shift
min_hours_single_shift = 14
for name, group in df.groupby('Employee Name'):
    mask = group['Shift Duration'].dt.total_seconds() / 3600 > min_hours_single_shift
    if mask.any():
        print(f"Employee {name} has worked for more than {min_hours_single_shift} hours in a single shift.")