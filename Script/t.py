from datetime import datetime, timedelta
 
import xlsxwriter
 
#import secrets
 
import random
def generate_time_intervals(start_time_str, num_files):
 
    # Parse the start time string into a datetime object
 
    start_time = datetime.strptime(start_time_str, "%I:%M %p")
 
    time_intervals = []
 
    for i in range(num_files):
 
        # Generate a random interval between 2 and 4 minutes
 
        interval_minutes = random.randint(25, 28)
 
        # Calculate the end time by adding the interval
 
        end_time = start_time + timedelta(minutes=interval_minutes)
 
        # Append the start and end times as a tuple
 
        time_intervals.append((start_time.strftime('%I:%M %p'), end_time.strftime('%I:%M %p')))
 
        # Update the start time for the next interval
 
        start_time = end_time
 
    return time_intervals
def main():
 
    # Get user input
 
    start_time_str = input("Enter the start time (e.g., 12:00 PM): ")
 
    num_files = int(input("Enter the number of files to generate: "))
 
    #r1=int(input("Enter the range 1: "))
 
    #r2=int(input("Enter the range 2: "))
 
    #interval_minutes = random.randint(r1, r2)
 
    #interval_minutes = secrets.randbelow(r2 - r1 + 1) + r1
 
    # Generate time intervals
 
    time_intervals = generate_time_intervals(start_time_str, num_files)
 
    # Create an Excel workbook and add a worksheet
 
    workbook = xlsxwriter.Workbook('time_intervals.xlsx')
 
    worksheet = workbook.add_worksheet()
 
    # Write headers
 
    worksheet.write('A1', 'Start Time')
 
    worksheet.write('B1', 'End Time')
 
    # Write data to the worksheet
 
    for row_num, (start, end) in enumerate(time_intervals, start=1):
 
        worksheet.write(row_num, 0, start)
 
        worksheet.write(row_num, 1, end)
 
    # Close the workbook
 
    workbook.close()
 
    print("Time intervals saved to time_intervals.xlsx")
if __name__ == "__main__":
 
    main()