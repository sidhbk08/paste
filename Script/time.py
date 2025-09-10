from datetime import datetime, timedelta

def generate_time_intervals(start_time_str, interval_minutes, num_files):
    # Parse the start time string into a datetime object
    start_time = datetime.strptime(start_time_str, "%I:%M %p")
    
    time_intervals = []
    
    for i in range(num_files):
        # Calculate the end time by adding the interval
        end_time = start_time + timedelta(minutes=interval_minutes)
        
        # Format the times back to the desired string format
        time_interval = f"{start_time.strftime('%I:%M %p')} {end_time.strftime('%I:%M %p')}"
        time_intervals.append(time_interval)
        
        # Update the start time for the next interval
        start_time = end_time
    
    return time_intervals

def main():
    # Get user input
    start_time_str = input("Enter the start time (e.g., 12:00 PM): ")
    interval_minutes = int(input("Enter the interval in minutes: "))
    num_files = int(input("Enter the number of files to generate: "))
    
    # Generate time intervals
    time_intervals = generate_time_intervals(start_time_str, interval_minutes, num_files)
    
    # Print the results
    for i, interval in enumerate(time_intervals, start=1):
        print(f"{interval}")

if __name__ == "__main__":
    main()