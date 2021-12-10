import pandas
import plotly.express as px
import json
import xlwt
from xlwt import Workbook

#declaring variables
run = True

#printing statements
print("To enter the marks of of a particular subject type 'mark'")
print("then choose the subject name and write for the st's accordingly")
print("To view a plotted version of your marks type 'plot marks'")
print("------------------------------------------------------------------------")

#plotting marks in bar chart format
def write_marks():
    subject_inp = input("Enter the subject: ").lower()

    global data  
    opening = open("marks.json")

    data = json.load(opening)
    
    
    for i in range(len(data["subjects"])):
        if subject_inp == data["subjects"][i]["name"]:
            index_value = i

            st_inp  = input("To type the marks of 'st1' or 'st2' type st1 or st2 respectively: ")

            if st_inp == "st1":
                st = "st1"
                global mark_inp

                try: 
                    mark_inp = float(input("Enter the marks of " + st + ": "))
                except ValueError:
                    print("please enter a valid value")
                    mark_inp = float(input("Enter the marks of " + st + ": "))

                if mark_inp >= 0 and mark_inp <= 20:
                    with open("marks.json", 'r+') as file:
                        file_data = json.load(file)

                        file_data["subjects"][index_value]["marks"].append({"st1": mark_inp})
                        file.seek(0)
                        json.dump(file_data, file, indent=4)
                else:
                    print("Please enter a value between from 0 to 20")

            elif st_inp == "st2":
                st = "st2"
                global mark_inp2

                try: 
                    mark_inp2 = float(input("Enter the marks of " + st + ": "))
                except ValueError:
                    print("please enter a valid value")
                    mark_inp2 = float(input("Enter the marks of " + st + ": "))
            
                if mark_inp2 >=0 and mark_inp2 <= 20:

                    with open("marks.json", 'r+') as file:
                        file_data = json.load(file)

                        file_data["subjects"][index_value]["marks"].append({"st2": mark_inp2})
                        file.seek(0)
                        json.dump(file_data, file, indent=4)
                else:
                    print("please enter a valid value from 0 to 20")


def average_marks():
    global data, average

    opening = open("marks.json")
    data = json.load(opening)
  

    #calculating average marks
    for i in range(len(data["subjects"])):
        average = int(data["subjects"][i]["marks"][0]["st1"] + data["subjects"][i]["marks"][1]["st2"])/2

        with open("marks.json", 'r+') as file:
            file_data = json.load(file)

            file_data["subjects"][i]["marks"].append({"average": average})
            file.seek(0)
            json.dump(file_data, file, indent=4)


def write_to_excel():
   #creating workbook
   workbook = xlwt.Workbook()

   sheet = workbook.add_sheet("marks")

   #loading file
   file = open("marks.json")

   data = json.load(file)

   #writing values to excel file
   columns = ["subjects", "st1", "st2", "average"]

    #writing column heads
   for i in range(len(columns)):
       sheet.write(0, i, columns[i])
    
    #writing subject names
   for i in range(len(data["subjects"])):
       sheet.write(i + 1, 0, data["subjects"][i]["name"])

   for i in range(len(data["subjects"])):
       for j in range(len(data["subjects"][i]["marks"][0])):
           sheet.write(i + 1, 1, data["subjects"][i]["marks"][0]["st1"])

   for i in range(len(data["subjects"])):
       for j in range(len(data["subjects"][i]["marks"][0])):
           sheet.write(i + 1, 2, data["subjects"][i]["marks"][1]["st2"])

   for i in range(len(data["subjects"])):  
       for j in range(len(data["subjects"][i]["marks"][-1])):
           sheet.write(i + 1, 3, data["subjects"][i]["marks"][-1]["average"])

   workbook.save("averages.xls")

   plotting_data = pandas.read_excel("averages.xls")

   graph_input = input("what type of a graph would you like to see(bar, pie, scatter): ")

   if graph_input == "bar":
    fig = px.bar(plotting_data, x="subjects", y="average", color="subjects", title="average of marks")
    fig.show()
   elif graph_input == "pie":
    fig = px.pie(plotting_data, values="average", color="subjects", title="average of marks")
    fig.show()
   elif graph_input == "scatter":
    fig = px.scatter(plotting_data, x="subjects", y="average", color="subjects", size="average")
    fig.show()



#main loop
while run:
    inp = input("")

    if inp == "mark":
        write_marks()
    elif inp == "average":
        average_marks()
    elif inp == "plot marks":
        write_to_excel()

        #reading file
     

        