    # Employee Scheduling Program

## Overview

This project is an employee scheduling program that assigns workers to shifts based on their availability and the company's staffing needs. The program reads data from two Excel files: `workers.xlsx` and `quarter.xlsx`.

- **workers.xlsx** contains the availability of each worker for each day of the week.
- **quarter.xlsx** specifies the number of employees needed for each hour of the workweek.

The program uses linear programming to find an optimal schedule that meets the staffing requirements while respecting each worker's availability.

## Features

- Handles single and multiple time ranges for worker availability within a single day.
- Ensures that staffing requirements for each hour of the workweek are met.
- Minimizes the total number of work periods assigned to workers.
- Respects worker availability constraints.

## File Descriptions

### workers.xlsx

Contains the availability and work hour limits for each worker. Each worker's availability can be specified as single or multiple time ranges for each day. The format is as follows:

| Worker  | Mon Start | Mon End | Tue Start | Tue End | ... | Min | Max |
|---------|-----------|---------|-----------|---------|-----|-----|-----|
| worker1 | 11,15     | 17,17   | 15        | 17      | ... | 15  | 30  |
| worker2 | 14        | 17      | 0         | 0       | ... | 8   | 11  |
| ...     | ...       | ...     | ...       | ...     | ... | ... | ... |

### quarter.xlsx

Specifies the number of employees needed for each hour of the workweek. The format is as follows:

| Mon 8 | Mon 9 | Mon 10 | Mon 11 | ... | Fri 16 |
|-------|-------|--------|--------|-----|--------|
| 2     | 2     | 2      | 2      | ... | 3      |

## Running the Program

1. Ensure you have Python and the necessary libraries installed. You can install the required libraries using:
    ```sh
    pip install pandas pulp
    ```

2. Place the `workers.xlsx` and `quarter.xlsx` files in the same directory as the script.

3. Run the `model.py` script:
    ```sh
    python model.py
    ```

4. The script will output the optimal schedule for each worker, indicating which periods they should work.### Output

According to the constraints, this program returns the turns for each worker during the week in an Excel file called niceSchedule.xlsl.

### Execution

To run, you have to install Pandas and PuLP.
Then, in shell:

    python model.py

or in Mac:

    python3 model.py
    
It will take around a minute to solve, depending on the computer.
Then, we will have, for every worker in worker_data, a dictionary called "schedule", where it tells which period corresponds to each worker.

