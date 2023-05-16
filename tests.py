from Final import ScheduleMaker

#these assert statements take some assumptions to work. You must use a fresh schedule and May_11 excel file
#and you cannot add anyone or it could fail

def test_ScheduleMakerClass():
    example = ScheduleMaker()

    # Checking headers
    example.write_header()
    assert example.sheet1.cell(row=1, column=2).value == "Monday"
    assert example.sheet1.cell(row=1, column=3).value == "Tuesday"
    assert example.sheet1.cell(row=1, column=4).value == "Wednesday"
    assert example.sheet1.cell(row=1, column=5).value == "Thursday"
    assert example.sheet1.cell(row=1, column=6).value == "Friday"
    
    # Checking time slots
    example.write_time_slots()
    assert example.sheet1.cell(row=2, column=1).value == "Times"
    assert example.sheet1.cell(row=3, column=1).value == "9:00"
    assert example.sheet1.cell(row=4, column=1).value == "10:00"
    assert example.sheet1.cell(row=5, column=1).value == "11:00"
    assert example.sheet1.cell(row=6, column=1).value == "12:00"
    assert example.sheet1.cell(row=7, column=1).value == "1:00"
    assert example.sheet1.cell(row=8, column=1).value == "2:00"
    assert example.sheet1.cell(row=9, column=1).value == "3:00"
    assert example.sheet1.cell(row=10, column=1).value == "4:00"
    assert example.sheet1.cell(row=11, column=1).value == "5:00"

    # Checking name added correctly for schedule
    example.write_schedule()
    assert example.sheet1.cell(row=4, column=2).value == "Mathew"
    assert example.sheet1.cell(row=9, column=2).value == "Jennifer"
    assert example.sheet1.cell(row=9, column=3).value == "Jennifer"
    assert example.sheet1.cell(row=6, column=4).value == "Dan"
    assert example.sheet1.cell(row=8, column=5).value == "Arianna"
    assert example.sheet1.cell(row=9, column=6).value == "Arianna"
    assert example.sheet1.cell(row=4, column=6).value == "Mathew"

test_ScheduleMakerClass()
