excel_coordinates = ((15, 20), (20, 25), (25, 30), (30, 35), (35, 40), (40, 45))
count = 1
for j in excel_coordinates:
    for i in range(*j):
        print(i)
