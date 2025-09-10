from datetime import datetime, timedelta

def generate_schedule(start_time, end_time, num_files, num_breaks):
    total_duration = (end_time - start_time).total_seconds() / 60  # total minutes
    break_time = num_breaks * 5  # each break is 5 minutes
    work_time = total_duration - break_time

    if work_time <= 0:
        return [("Invalid", "Break time exceeds or equals total time.")]

    chunk_time = work_time / num_files
    schedule = []
    current_time = start_time

    for i in range(num_files):
        next_time = current_time + timedelta(minutes=chunk_time)
        if next_time > end_time:
            next_time = end_time
        schedule.append((current_time.strftime("%I:%M %p"), next_time.strftime("%I:%M %p")))
        current_time = next_time

    # Add remaining time if any
    if current_time < end_time:
        schedule.append((current_time.strftime("%I:%M %p"), end_time.strftime("%I:%M %p")))

    return schedule

def parse_time(input_str):
    try:
        return datetime.strptime(input_str.strip(), "%I:%M %p")
    except ValueError:
        print("Invalid time format. Please use format like '08:00 AM'")
        return None

def main():
    print("\n📅 Work Schedule Generator (Now Supports Overnight Schedules)\n")

    while True:
        start_str = input("Enter start time (e.g., 08:00 AM): ")
        start_time = parse_time(start_str)
        if start_time:
            break

    while True:
        end_str = input("Enter end time (e.g., 05:00 PM): ")
        end_time = parse_time(end_str)
        if end_time:
            # If end is earlier than start, assume it's next day
            if end_time <= start_time:
                end_time += timedelta(days=1)
            break

    while True:
        try:
            num_files = int(input("Enter number of files (tasks): "))
            if num_files > 0:
                break
        except:
            pass
        print("Please enter a positive integer.")

    while True:
        try:
            num_breaks = int(input("Enter number of breaks: "))
            if num_breaks >= 0:
                break
        except:
            pass
        print("Please enter a non-negative integer.")

    print("\n🕒 Generated Schedule:\n")
    schedule = generate_schedule(start_time, end_time, num_files, num_breaks)
    for i, (start, end) in enumerate(schedule, 1):
        print(f"{i:02d}. {start}  -  {end}")


if __name__ == "__main__":
    main()