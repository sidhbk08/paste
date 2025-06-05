from datetime import datetime, timedelta
import random

def distribute_durations(total_minutes, num_files, min_dur=0.5):
    base = total_minutes / num_files
    durations = [base] * num_files
    for _ in range(num_files * 2):
        i, j = random.sample(range(num_files), 2)
        if durations[i] > min_dur:
            delta = min(durations[i] - min_dur, 0.5)
            durations[i] -= delta
            durations[j] += delta
    return durations

def parse_time(t_str):
    return datetime.strptime(t_str.strip(), "%I:%M %p")

def adjust_time(base_date, t):
    if t.time() >= base_date.time():
        return datetime.combine(base_date.date(), t.time())
    else:
        return datetime.combine(base_date.date(), t.time()) + timedelta(days=1)

def generate_schedule(start_time_str, end_time_str, num_files, breaks_input):
    base_date = datetime.today()
    start_time = parse_time(start_time_str)
    end_time = parse_time(end_time_str)

    start_time = adjust_time(base_date, start_time)
    end_time = adjust_time(start_time, end_time)  # handle overnight

    # Parse breaks into datetime and duration minutes
    parsed_breaks = []
    for b_start_str, b_end_str in breaks_input:
        b_start = adjust_time(start_time, parse_time(b_start_str))
        b_end = adjust_time(b_start, parse_time(b_end_str))
        parsed_breaks.append((b_start, b_end, (b_end - b_start).total_seconds() / 60))

    parsed_breaks.sort(key=lambda x: x[0])  # Sort breaks by start time

    total_time = (end_time - start_time).total_seconds() / 60
    total_break_minutes = sum(b[2] for b in parsed_breaks)
    work_minutes = total_time - total_break_minutes

    if work_minutes <= 0:
        raise ValueError("Total break time exceeds total available time!")
    if work_minutes < num_files * 0.5:
        raise ValueError("Not enough time for given number of files.")

    durations = distribute_durations(work_minutes, num_files)
    schedule = []
    current_time = start_time
    break_index = 0
    last_break_len = 0  # To hold break length that will be shown with next file label

    for i, dur in enumerate(durations, 1):
        # Skip past any breaks that ended before current_time
        while break_index < len(parsed_breaks) and current_time >= parsed_breaks[break_index][1]:
            break_index += 1

        # If next break exists and current work overlaps break start
        if break_index < len(parsed_breaks):
            b_start, b_end, b_len = parsed_breaks[break_index]
            if current_time < b_start and current_time + timedelta(minutes=dur) > b_start:
                # Work only until break start for this file
                work_end = b_start
                actual_dur = (work_end - current_time).total_seconds() / 60

                schedule.append({
                    'file_num': i,
                    'start': current_time.strftime("%I:%M %p"),
                    'end': work_end.strftime("%I:%M %p"),
                    'label': "Work"
                })

                current_time = b_end  # jump over break
                last_break_len = int(b_len)  # store break duration for next file label
                break_index += 1
                continue  # Repeat scheduling same file (i) after break with remaining duration

        # If we had a break just before this file, add break info in label
        label = "Work"
        if last_break_len > 0:
            label = f"Work + Break ({last_break_len} mins)"
            last_break_len = 0  # reset after using

        f_start = current_time
        f_end = f_start + timedelta(minutes=dur)

        schedule.append({
            'file_num': i,
            'start': f_start.strftime("%I:%M %p"),
            'end': f_end.strftime("%I:%M %p"),
            'label': label
        })

        current_time = f_end

    return schedule


# 🧠 Runtime Input
if __name__ == "__main__":
    try:
        print("Enter time in 12-hour format (e.g. 10:30 AM, 02:00 PM)")
        start_time = input("Enter start time: ")
        end_time = input("Enter end time: ")
        num_files = int(input("Enter number of files: "))
        num_breaks = int(input("Enter number of breaks: "))

        breaks = []
        for i in range(num_breaks):
            print(f"\nBreak {i + 1}:")
            b_start = input("  Start time: ")
            b_end = input("  End time: ")
            breaks.append((b_start, b_end))

        schedule = generate_schedule(start_time, end_time, num_files, breaks)

        print("\n🗓️ Generated Schedule:")
        for item in schedule:
            print(f"File {item['file_num']:03d}: {item['start']} - {item['end']} : {item['label']}")

    except Exception as e:
        print(f"\n❌ Error: {e}")
