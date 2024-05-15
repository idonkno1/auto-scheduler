# shift-scheduling
### Shift Scheduling for workforce

This shift planner takes data from Excel files (quarter.xlsx and workers.xlsx) and returns an Excel file with weekly shifts for each worker.

### Problem Description

This was created to help the scheduling issue at Chem Stores and Inorganic BYU and make it more streamlined.

Taking into account each worker's hourly availability and minimum and maximum preferred hours.


### Output

According to the constraints, this program returns the turns for each worker during the week in an Excel file called niceSchedule.xlsl.

### Execution

To run, you have to install Pandas and PuLP.
Then, in shell:

    python model.py
    
It will take around a minute to solve, depending on the computer.
Then, we will have, for every worker in worker_data, a dictionary called "schedule", where it tells which period corresponds to each worker.

