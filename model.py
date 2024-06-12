from operator import itemgetter
import pandas
import pulp

from example_inputs import (
    periods,
)

AM_QUARTERS = 45 # length of amount of workers needed 

# workers_data = {
#   worker1: {
#       "period_avail": [0, 1, 1, 1, 0, 0,........, 1, 1, 0, 0]. Length 45,
#   }
# }


def model_problem():

    workerdf = pandas.read_excel("workers.xlsx", header=0) 
    quarters = pandas.read_excel("./quarter.xlsx", header=0).loc[0].tolist()
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)
    workers_data = {}

    for index, row in workerdf.iterrows(): 
        name = row["Worker"] 
        min_hours = row['Min']
        max_hours = row['Max']
        workers_data[name] = {
            "min_hours": min_hours,
            "max_hours": max_hours,
            "period_avail": []
        } 
        
        # checks worker excel file and gives 1 if available and 0 if not, check workers excel for availabilty
        for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']: 
            start_times = row[f'{day} Start']
            end_times = row[f'{day} End']

            # Check if values are comma-separated or single integers
            if isinstance(start_times, str) and ',' in start_times:
                start_times = [int(x) for x in start_times.split(',')]
            else:
                start_times = [int(start_times)]

            if isinstance(end_times, str) and ',' in end_times:
                end_times = [int(x) for x in end_times.split(',')]
            else:
                end_times = [int(end_times)]

            availability = [0] * 9  # 9 periods from 8 AM to 5 PM
            for start, end in zip(start_times, end_times):
                start_time = int(start)
                end_time = int(end)
                for period in range(9):
                    period_start = 8 + period
                    period_end = period_start + 1
                    if period_start >= start_time and period_end <= end_time:
                        availability[period] = 1
            workers_data[name]["period_avail"].extend(availability)

    for worker, data in workers_data.items():
        data["worked_periods"] = [
            pulp.LpVariable(f"x_{worker}_{period}", cat=pulp.LpBinary)
            for period in range(AM_QUARTERS)
        ]
    
    # respect worker unavailablity
    for worker, data in workers_data.items():
        for period in range(AM_QUARTERS):
            if data['period_avail'][period] == 0:
                problem += data['worked_periods'][period] == 0

    # ensures correct amount of staff no under/over staffing
    objective_function = None
    for worker in workers_data.keys():
        objective_function += sum(workers_data[worker]["worked_periods"])
    
    problem += objective_function

    # ensures workers meet demand each quarter, look in quater excel file for numbers
    for quarter in range(AM_QUARTERS):
        problem += (
            pulp.lpSum(workers_data[worker]['worked_periods'][quarter] for worker in workers_data) == quarters[quarter],
            f"Demand_Q{quarter}"
        )

    # minimum and maximum hours constraints for workers, look in workers excel for numbers
    for worker, data in workers_data.items():
        problem += (pulp.lpSum(data["worked_periods"]) >= data["min_hours"], f"{worker}_min_hours")
        problem += (pulp.lpSum(data["worked_periods"]) <= data["max_hours"], f"{worker}_max_hours")

    try:
        problem.solve()
    except Exception as e:
        print("Can't solve problem: {}".format(e))

    for worker in workers_data.keys(): 
        workers_data[worker]["schedule"] = [] 
        for element in range(len(workers_data[worker]["worked_periods"])):
            if workers_data[worker]["worked_periods"][element].varValue == 1:
                workers_data[worker]["schedule"].append(periods[element])

    return problem, workers_data

def parse_schedule(schedule, name):
    # Dictionary to hold the parsed schedule
    schedule_dict = {}
    
    # Splitting each time slot in the schedule string
    time_slots = schedule.split(', ')
    for slot in time_slots:
        day_time = slot.split(' ')
        day = day_time[0]
        hour = day_time[1]
        
        # Add the worker to the corresponding day and time in the dictionary
        if day in schedule_dict:
            if hour in schedule_dict[day]:
                schedule_dict[day][hour].append(name)
            else:
                schedule_dict[day][hour] = [name]
        else:
            schedule_dict[day] = {hour: [name]}
    
    return schedule_dict

def hour_sort_key(hour):
    start, end = hour.split('-')
    return int(start), int(end)

def create_nice_schedule(schedule_df):
    # Initialize an empty dictionary to store the new format
    new_schedule = {day: {} for day in ["Mon", "Tue", "Wed", "Thu", "Fri"]}
    hours = set()
    
    # Loop through each row in the original DataFrame to parse and accumulate the schedule
    for index, row in schedule_df.iterrows():
        name = row['Worker']
        worker_schedule = parse_schedule(row['Schedule'], name)
        
        # Add each worker's schedule to the new format
        for day, times in worker_schedule.items():
            for hour, names in times.items():
                hours.add(hour)
                if hour in new_schedule[day]:
                    new_schedule[day][hour] += names
                else:
                    new_schedule[day][hour] = names
    
    # Create a DataFrame from the new_schedule dictionary
    hours_sorted = sorted(list(hours), key=hour_sort_key)
    nice_schedule = pandas.DataFrame(index=hours_sorted, columns=["Mon", "Tue", "Wed", "Thu", "Fri"])
    
    for day in ["Mon", "Tue", "Wed", "Thu", "Fri"]:
        for hour in hours_sorted:
            if hour in new_schedule[day]:
                nice_schedule.at[hour, day] = ', '.join(new_schedule[day][hour])
    
    nice_schedule.fillna('', inplace=True)
    nice_schedule.reset_index(inplace=True)
    nice_schedule.rename(columns={'index': 'Hour'}, inplace=True)
    return nice_schedule

if __name__ == "__main__":
    problem, workers_data = model_problem()

    # Creating DataFrame from workers_data
    schedule_df = pandas.DataFrame([{**{'Worker': worker}, **{'Schedule': ', '.join(data['schedule'])}} for worker, data in workers_data.items()])
    # Writing to Excel file
    schedule_df.to_excel("./ignore.xlsx", index=False)

    schedule_df = pandas.read_excel("./ignore.xlsx")
    nice_schedule = create_nice_schedule(schedule_df)

    # Save to Excel
    output_path = "./niceSchedule.xlsx"
    nice_schedule.to_excel(output_path, index=False)
