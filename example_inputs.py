
periods = [
    "{} {}-{}".format(day, hour, hour+1) 
    for day in ["Mon", "Tue", "Wed", "Thu", "Fri"] 
    for hour in range(8, 17)
]
