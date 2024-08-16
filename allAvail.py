import pandas as pd
import pulp

periods = [ 
    "{} {}-{}".format(day, hour, hour+1) 
    for day in ["Mon", "Tue", "Wed", "Thu", "Fri"] 
    for hour in range(8, 17)
]

WORKING_HOURS = 45

def add_availability_with_names():
    workerdf = pd.read_excel("workersAvailability.xlsx", header=0)
    hoursdf = pd.read_excel("workingHours.xlsx", header=0).loc[0].tolist()

    workers_data = {}
    total_availability = [0] * WORKING_HOURS
    employee_availability = [[] for _ in range(WORKING_HOURS)]

    for index, row in workerdf.iterrows(): 
        name = row["Worker"]
        workers_data[name] = {"hour_avail": []}
        
        for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']: 
            start_times = row[f'{day} Start']
            end_times = row[f'{day} End']

            if isinstance(start_times, str) and ',' in start_times:
                start_times = [int(x) for x in start_times.split(',')]
            else:
                start_times = [int(start_times)]

            if isinstance(end_times, str) and ',' in end_times:
                end_times = [int(x) for x in end_times.split(',')]
            else:
                end_times = [int(end_times)]

            availability = [0] * 9
            for start, end in zip(start_times, end_times):
                start_time = int(start)
                end_time = int(end)
                for period in range(9):
                    period_start = 8 + period
                    period_end = period_start + 1
                    if period_start >= start_time and period_end <= end_time:
                        availability[period] = 1
            workers_data[name]["hour_avail"].extend(availability)
    
    for period in range(WORKING_HOURS):
        total_availability[period] = sum(
            workers_data[worker]["hour_avail"][period] for worker in workers_data.keys()
        )
        employee_availability[period] = [
            worker for worker in workers_data.keys() if workers_data[worker]["hour_avail"][period] == 1
        ]

    availability_df = pd.DataFrame({
        "Period": periods,
        "Total Availability": total_availability,
        "Employees Available": [", ".join(employees) for employees in employee_availability]
    })

    return availability_df, workers_data, hoursdf

def model_problem(workers_data, hoursdf):
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)

    for worker, data in workers_data.items():
        data["worked_hours"] = [
            pulp.LpVariable(f"x_{worker}_{period}", cat=pulp.LpBinary)
            for period in range(WORKING_HOURS)
        ]
    
    # Respect worker unavailability
    for worker, data in workers_data.items():
        for period in range(WORKING_HOURS):
            if data['hour_avail'][period] == 0:
                problem += data['worked_hours'][period] == 0

    # Ensures correct amount of staff no under/over staffing
    objective_function = pulp.lpSum(workers_data[worker]["worked_hours"][period] for worker in workers_data.keys() for period in range(WORKING_HOURS))
    problem += objective_function

    # Ensures workers meet demand each quarter
    for hour in range(WORKING_HOURS):
        if hoursdf[hour] == 0:
            problem += (
                pulp.lpSum(workers_data[worker]['worked_hours'][hour] for worker in workers_data) == 0,
                f"NoDemand_Q{hour}"
            )
        if hoursdf[hour] == 2:
            problem += (
                pulp.lpSum(workers_data[worker]['worked_hours'][hour] for worker in workers_data) == 2,
                f"ExactDemand_Q{hour}"
            )

        elif hoursdf[hour] > 2:
            problem += (
                pulp.lpSum(workers_data[worker]['worked_hours'][hour] for worker in workers_data) <= hoursdf[hour] + 1,
                f"MaxDemand_Q{hour}"
            )
            problem += (
                pulp.lpSum(workers_data[worker]['worked_hours'][hour] for worker in workers_data) >= hoursdf[hour] - 1,
                f"MinDemand_Q{hour}"
            )

    # Minimum and maximum hours constraints for workers
    for worker, data in workers_data.items():
        problem += (pulp.lpSum(data["worked_hours"]) >= data["min_hours"], f"{worker}_min_hours")
        problem += (pulp.lpSum(data["worked_hours"]) <= data["max_hours"], f"{worker}_max_hours")

    try:
        problem.solve()
    except Exception as e:
        print("Can't solve problem: {}".format(e))

    for worker in workers_data.keys(): 
        workers_data[worker]["schedule"] = [] 
        for element in range(len(workers_data[worker]["worked_hours"])):
            if workers_data[worker]["worked_hours"][element].varValue == 1:
                workers_data[worker]["schedule"].append(periods[element])

    return problem, workers_data

def create_nice_schedule(availability_df):
    # Initialize an empty dictionary to store the new format
    new_schedule = {day: {} for day in ["Mon", "Tue", "Wed", "Thu", "Fri"]}
    hours = set()

    # Loop through the availability DataFrame to accumulate the schedule
    for index, row in availability_df.iterrows():
        period = row['Period']
        employees = row['Employees Available']
        day, hour = period.split(' ')
        
        if day in new_schedule:
            new_schedule[day][hour] = employees
        hours.add(hour)
    
    # Create a DataFrame from the new_schedule dictionary
    hours_sorted = sorted(list(hours), key=lambda x: int(x.split('-')[0]))
    nice_schedule = pd.DataFrame(index=hours_sorted, columns=["Mon", "Tue", "Wed", "Thu", "Fri"])
    
    for day in ["Mon", "Tue", "Wed", "Thu", "Fri"]:
        for hour in hours_sorted:
            if hour in new_schedule[day]:
                nice_schedule.at[hour, day] = new_schedule[day][hour]
    
    nice_schedule.fillna('', inplace=True)
    nice_schedule.reset_index(inplace=True)
    nice_schedule.rename(columns={'index': 'Hour'}, inplace=True)
    return nice_schedule

if __name__ == "__main__":
    availability_df, workers_data, hoursdf = add_availability_with_names()
    problem, workers_data = model_problem(workers_data, hoursdf)

    # Creating DataFrame from workers_data
    schedule_df = pd.DataFrame([{**{'Worker': worker}, **{'Schedule': ', '.join(data['schedule'])}} for worker, data in workers_data.items()])
    # Writing to Excel file
    schedule_df.to_excel("schedule.xlsx", index=False)

    schedule_df = pd.read_excel("schedule.xlsx")
    nice_schedule = create_nice_schedule(availability_df)

    # Save to Excel
    output_path = "nice_schedule.xlsx"
    nice_schedule.to_excel(output_path, index=False)
