import os.path
from Final import ScheduleMaker

def test_ScheduleMakerClass():
    example = ScheduleMaker()

    # checking headers
    example.write_header()
    assert example.sheet1.cell(row=1, column=2).value == "Monday"
    assert example.sheet1.cell(row=1, column=3).value == "Tuesday"
    assert example.sheet1.cell(row=1, column=4).value == "Wednesday"
    assert example.sheet1.cell(row=1, column=5).value == "Thursday"
    assert example.sheet1.cell(row=1, column=6).value == "Friday"
    
    # checking time slots
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


    # checking name added correctly for schedule
    example.write_schedule()
    assert example.sheet1["B4"].value == "Jayla"
