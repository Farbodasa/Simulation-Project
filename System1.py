import random
import math
import pandas as pd
import xlsxwriter
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib as mpl
def create_main_header(state, data):
    # This function creates the main part of header (returns a list)
    # A part of header which is used for future events will be created in create_excel()

    # Header consists of ...
    # 1. Step, Clock, Event Type and Event Customer
    header = ['Step', 'Clock', 'Event Type', 'Event Customer Type', 'Event Customer Number']
    # 2. Names of the state variables
    header.extend(list(state.keys()))
    # 3. Names of the cumulative stats
    header.extend(['VIP Percent Without Waiting', 'Normal Percent Without Waiting',
                   'VIP Mean Waiting Time', 'Normal Mean Waiting Time',
                   'VIP Lost Customers', 'Normal Lost Customers'])
    for Key in list(data['Cumulative Stats']['VIP']['Max'].keys()):
         header.extend(['VIP Max '+Key])
    for Key in list(data['Cumulative Stats']['VIP']['Mean'].keys()):
         header.extend(['VIP Mean '+Key])
    for Key in list(data['Cumulative Stats']['Normal']['Max'].keys()):
         header.extend(['Normal Max '+Key])
    for Key in list(data['Cumulative Stats']['Normal']['Mean'].keys()):
         header.extend(['Normal Mean '+Key])
    for Key in list(data['Cumulative Stats']['Productivity'].keys()):
         header.extend([Key+" Server Productivity"])
    return header

def justify(table):
    # This function adds blanks to short rows in order to match their lengths to the maximum row length

    # Find maximum row length in the table
    row_max_len = 0
    for row in table:
        if len(row) > row_max_len:
            row_max_len = len(row)

    # For each row, add enough blanks
    for row in table:
        row.extend([""] * (row_max_len - len(row)))

def create_row(step, current_event, state, data, future_event_list):
    # This function will create a list, which will eventually become a row of the output Excel file
    sorted_fel = sorted(future_event_list, key=lambda x: x['Event Time'])
    # What should this row contain?
    # 1. Step, Clock, Event Type and Event Customer
    row = [step, current_event['Event Time'], current_event['Event Type'], current_event['Customer Type'], current_event['Customer Number']]
    # 2. All state variables
    row.extend(list(state.values()))
    # 3. All Cumulative Stats
    row.extend([data['Cumulative Stats']['VIP']['Percent Without Waiting'], data['Cumulative Stats']['Normal']['Percent Without Waiting'],
    data['Cumulative Stats']['VIP']['Mean Waiting Time'], data['Cumulative Stats']['Normal']['Mean Waiting Time'],
    data['Cumulative Stats']['VIP']['Lost Customers'], data['Cumulative Stats']['Normal']['Lost Customers']])
    row.extend(list(data['Cumulative Stats']['VIP']['Max'].values()))
    row.extend(list(data['Cumulative Stats']['VIP']['Mean'].values()))
    row.extend(list(data['Cumulative Stats']['Normal']['Max'].values()))
    row.extend(list(data['Cumulative Stats']['Normal']['Mean'].values()))
    row.extend(list(data['Cumulative Stats']['Productivity'].values()))
    # 4. All events in fel
    for event in sorted_fel:
        row.append(event['Event Time'])
        row.append(event['Event Type'])
        row.append(event['Customer Type'])
        row.append(event['Customer Number'])
    return row

def get_col_widths(dataframe):
    # Copied from https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def create_excel(table, header):
    # This function creates and fine-tunes the Excel output file
    # Find length of each row in the table
    row_len = len(table[0])
    # Find length of header (header does not include cells for fel at this moment)
    header_len = len(header)
    # row_len exceeds header_len by (max_fel_length * 3) (Event Type, Event Time & Customer for each event in FEL)
    # Extend the header with 'Future Event Time', 'Future Event Type', 'Future Event Customer'
    # for each event in the fel with maximum size
    i = 1
    for col in range((row_len - header_len) // 4):
        header.append('Future Event Time ' + str(i))
        header.append('Future Event Type ' + str(i))
        header.append('Future Event Customer Type ' + str(i))
        header.append('Future Event Customer Number ' + str(i))
        i += 1
    # Dealing with the output
    # First create a pandas DataFrame
    df = pd.DataFrame(table, columns=header, index=None)
    # Create a handle to work on the Excel file
    writer = pd.ExcelWriter('System1.xlsx', engine='xlsxwriter')
    # Write out the Excel file to the hard drive
    df.to_excel(writer, sheet_name='Call-Center Output', header=False, startrow=1, index=False)
    # Use the handle to get the workbook (just library syntax, can be found with a simple search)
    workbook = writer.book
    # Get the sheet you want to work on
    worksheet = writer.sheets['Call-Center Output']
    # Create a cell-formatter object (this will be used for the cells in the header, hence: header_formatter!)
    header_formatter = workbook.add_format()
    # Define whatever format you want
    header_formatter.set_align('center')
    header_formatter.set_align('vcenter')
    header_formatter.set_font('Times New Roman')
    header_formatter.set_bold('True')
    # Write out the column names and apply the format to the cells in the header row
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_formatter)
    # Auto-fit columns
    # Copied from https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter
    for i, width in enumerate(get_col_widths(df)):
        worksheet.set_column(i - 1, i - 1, width)
    # Create a cell-formatter object for the body of excel file
    main_formatter = workbook.add_format()
    main_formatter.set_align('center')
    main_formatter.set_align('vcenter')
    main_formatter.set_font('Times New Roman')
    # Apply the format to the body cells
    for row in range(1, len(df) + 1):
        worksheet.set_row(row, None, main_formatter)
    # Save your edits
    writer.save()

def starting_state():
    data = dict()
    # To track each customer, saving their arrival time, Number of customers, etc.
    data["Customer"] = dict()
    data["Customer"]['VIP'] = dict()
    data["Customer"]['Normal'] = dict()
    data["Customer"]['VIP']['Arrival Time'] = dict()
    data["Customer"]['Normal']['Arrival Time'] = dict()
    data['Customer']["Normal"]['Number'] = 0
    data['Customer']["VIP"]['Number'] = 0
    data['Customer']["Normal"]['Out'] = 0
    data['Customer']["VIP"]['Out'] = 0
    data['Customer']['VIP']['Data'] = dict()
    data['Customer']['Normal']['Data'] = dict()

    # To track each queue, Customer Arrival in queue Time, used to find first customer in queue ,etc.
    data["Queue"] = dict()
    data["Queue"]['VIP'] = dict()
    data["Queue"]["Normal"] = dict()
    data['Queue']['Amateur']=[]
    data['Queue']['Professional']=[]
    data['Queue']['Normal']['Technical']=[]
    data['Queue']["VIP"]['Technical'] = []
    data['Queue']['VIP']['Without Waiting in Queue'] = []
    data['Queue']['Normal']['Without Waiting in Queue'] = []
    # Needed to calculate area under queue length curve
    data['Last time'] = 0
    # Cumulative Stats
    data['Cumulative Stats'] = dict()
    data['Cumulative Stats']['VIP'] = dict()
    data['Cumulative Stats']['Normal'] = dict()
    data['Cumulative Stats']['VIP']['Max'] = dict()
    data['Cumulative Stats']['Normal']['Max'] = dict()
    data['Cumulative Stats']['VIP']['Area'] = dict()
    data['Cumulative Stats']['Normal']['Area'] = dict()
    data['Cumulative Stats']['VIP']['Mean'] = dict()
    data['Cumulative Stats']['Normal']['Mean'] = dict()
    data['Cumulative Stats']['Busy Time'] = dict()
    data['Cumulative Stats']['Productivity'] = dict()
    data['Cumulative Stats']['VIP']['Technical Need'] = 0
    data['Cumulative Stats']['Normal']['Technical Need'] = 0
    data['Cumulative Stats']['VIP']['Lost Customers'] = 0
    data['Cumulative Stats']['Normal']['Lost Customers'] = 0
    data['Cumulative Stats']['VIP']['Without Waiting'] = 0
    data['Cumulative Stats']['Normal']['Without Waiting'] = 0
    data['Cumulative Stats']['VIP']['Waiting Time'] = 0
    data['Cumulative Stats']['Normal']['Waiting Time'] = 0
    data['Cumulative Stats']['VIP']['Percent Without Waiting'] = 0
    data['Cumulative Stats']['Normal']['Percent Without Waiting'] = 0
    data['Cumulative Stats']['VIP']['Mean Waiting Time'] = 0
    data['Cumulative Stats']['Normal']['Mean Waiting Time'] = 0
    data['Cumulative Stats']['VIP']['Max']['Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Max']['Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Max']['Technical Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Max']['Technical Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Area']['Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Area']['Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Area']['Technical Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Area']['Technical Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Mean']['Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Mean']['Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Mean']['Technical Queue Length'] = 0
    data['Cumulative Stats']['Normal']['Mean']['Technical Queue Length'] = 0
    data['Cumulative Stats']['VIP']['Max']['Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Max']['Queue Wait'] = 0
    data['Cumulative Stats']['VIP']['Max']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Max']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['VIP']['Area']['Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Area']['Queue Wait'] = 0
    data['Cumulative Stats']['VIP']['Area']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Area']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['VIP']['Mean']['Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Mean']['Queue Wait'] = 0
    data['Cumulative Stats']['VIP']['Mean']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['Normal']['Mean']['Technical Queue Wait'] = 0
    data['Cumulative Stats']['Busy Time']['Amateur'] = 0
    data['Cumulative Stats']['Busy Time']['Professional'] = 0
    data['Cumulative Stats']['Busy Time']['Technical'] = 0
    data['Cumulative Stats']['Productivity']['Amateur'] = 0
    data['Cumulative Stats']['Productivity']['Professional'] = 0
    data['Cumulative Stats']['Productivity']['Technical'] = 0
    data['Waiting Time'] = dict()
    data['Waiting Time']['VIP'] = []
    data['Waiting Time']['Normal'] = []
    # State variables
    state = dict()
    state['Shift'] = 1 # 1,2,3
    state['LQa'] = 0
    state['LQp'] = 0
    state['LQtn'] = 0
    state['LQtv'] = 0
    state['La'] = 0 # 0,1,2,3
    state['Lp'] = 0 # 0,1,2
    state['Lt'] = 0 # 0,1,2
    state['Day'] = 1 # 1,2,...,30
    # Starting FEL
    future_event_list = list()
    future_event_list.append({'Event Type': 'Customer Arrival', 'Event Time':P1()})  # This is an Event Notice
    future_event_list.append({'Event Type': 'Shift Change', 'Event Time': ST(), 'Customer Type':'-', 'Customer Number':'-'})
    return state, future_event_list, data
#function to Amateur time service variable with D1 distribution
def S1():
    return exponential(1/3)
#function to Proffesional time service variable with D2 distribution
def S2():
    return exponential(1/7)
#function to Technical time service variable with D3 distribution
def S3():
    return exponential(1/10)
#function to Shift 1 Arrival time with exponential distrbiution
def P1():
    return exponential(1/1.1)
#function to Shift 2 Arrival time with exponential distrbiution
def P2():
    return exponential(1/1.1)
#function to Shift 3 Arrival time with exponential distrbiution
def P3():
    return exponential(1/1.1)
#function to create Tierd and Departure time
def T(Queue_Length):
    return uniform(5,max(25,Queue_Length))
#function to create Customer Type that Arrivaled in system
def Customer_type():
    if R()<0.4:
        return "VIP"
    else:
        return "Normal"
#function to generate random number with uniform distribution between 0,1
def R():
    return random.random()
#time of shift change
def ST():
    return 480
#function to generate random number by exponential distribution with parameter λ
def exponential(lambd):
    r = random.random()
    return -(1 / lambd) * math.log(r)
#function to generate random number with uniform distribution between a,b
def uniform(a, b):
    r = random.random()
    return a + (b - a) * r
#function for Assign a Number to customers
def Customer_number(Customer_type, data):
    data['Customer'][Customer_type]['Number'] += 1
    return data['Customer'][Customer_type]['Number']
#function to remove an Tired and Departure event from FEL for specific customer
def Delete_event(future_event_list, Customer_Type, Customer_Number):
    for event in future_event_list:
        if event['Event Type'] == 'Customer Tierd and Departure' :
            if  event['Customer Type'] == Customer_Type and event['Customer Number'] == Customer_Number:
                future_event_list.remove(event)
#function to create event for Arrival in FEL
def fel_maker_Customer_Arrival(future_event_list, state, clock):
    event_time = 0
    if state['Shift']== 1 :
        event_time = clock + P1()
    elif state['Shift']== 2 :
        event_time = clock + P2()
    elif state['Shift']== 3 :
        event_time = clock + P3()
    new_event = {'Event Type': 'Customer Arrival', 'Event Time': event_time, 'Customer Type':"-", 'Customer Number':"-"}
    future_event_list.append(new_event)
#function to create event for Shift Change in FEL
def fel_maker_Shift_Change(future_event_list, clock):
    event_time = clock + ST()
    new_event = {'Event Type': 'Shift Change', 'Event Time': event_time, 'Customer Type':"-" ,"Customer Number":"-"}
    future_event_list.append(new_event)
#function to create event for spesific event in FEL
def fel_maker(future_event_list, event_type, clock, Customer_type, Customer_number, Queue_Length = None):
    event_time=0
    if event_type == 'Professional Task End':
        event_time = clock + S1()
    elif event_type == 'Amateur Task End':
        event_time = clock + S2()
    elif event_type == 'Technical Team Task End':
        event_time = clock + S3()
    elif event_type =='Customer Tierd and Departure':
        event_time = clock + T(Queue_Length)
    new_event = {'Event Type': event_type, 'Event Time': event_time,'Customer Type':Customer_type, 'Customer Number':Customer_number}
    future_event_list.append(new_event)
    # This function does not have a return value

def Customer_Arrival(future_event_list, state, clock, data, system):
    #generate Customer type and number for who came in
    Customer_Type = Customer_type()
    Customer_Number = Customer_number(Customer_Type, data)
    future_event_list[0]['Customer Type'] = Customer_Type
    future_event_list[0]['Customer Number'] = Customer_Number
    data['Customer'][Customer_Type]['Arrival Time'][Customer_Number] = clock
    data['Customer'][Customer_Type]['Data'][Customer_Number] = dict()
    data['Customer'][Customer_Type]['Data'][Customer_Number]['Arrival'] = clock
    if Customer_Type == "VIP" :
        if state['Lp'] == system['Professional']:
            if state['LQp'] > 4:
                state['LQp'] += 1
                # Calculate Maximum Queue Length
                if data['Cumulative Stats']['VIP']['Max']['Queue Length']< state['LQp'] :
                    data['Cumulative Stats']['VIP']['Max']['Queue Length'] = state['LQp']
                #adding cutomer in the professional queue
                data['Queue']['Professional'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
                if R() <= 0.15: #if the possiblity of a customer getting tired is extant
                    fel_maker(future_event_list,'Customer Tierd and Departure', clock, Customer_Type, Customer_Number, state['LQp'])
            else:
                state['LQp'] += 1
                # Calculate Maximum Queue Length
                if data['Cumulative Stats']['VIP']['Max']['Queue Length'] < state['LQp']:
                    data['Cumulative Stats']['VIP']['Max']['Queue Length'] = state['LQp']
                # adding cutomer in the professional queue
                data['Queue']['Professional'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
                if R() <= 0.15:#if the possiblity of a customer getting tired is extant
                    fel_maker(future_event_list, 'Customer Tierd and Departure', clock, Customer_Type, Customer_Number, state['LQp'])
        else:
            #Customer without waiting in Queue
            data['Queue']['VIP']['Without Waiting in Queue'].append(Customer_Number)
            data['Customer']['VIP']['Data'][Customer_Number]['Start Service'] = clock
            data['Waiting Time']["VIP"].append(0)
            state['Lp'] += 1
            fel_maker(future_event_list,'Professional Task End', clock,Customer_Type,Customer_Number)
    else:
        if state['La'] == system['Amateur']:
            if state['Lp'] == system['Professional']:
                if state["LQa"] > 4:
                    state['LQa'] += 1
                    # Calculate Maximum Queue Length
                    if data['Cumulative Stats']['Normal']['Max']['Queue Length'] < state['LQa'] :
                        data['Cumulative Stats']['Normal']['Max']['Queue Length'] = state['LQa']
                    # adding cutomer in the Amateur queue
                    data['Queue']['Amateur'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
                    if R() <= 0.15:#if the possiblity of a customer getting tired is extant
                        fel_maker(future_event_list, 'Customer Tierd and Departure', clock, Customer_Type, Customer_Number, state['LQa'])
                else:
                    state['LQa'] += 1
                    if data['Cumulative Stats']['Normal']['Max']['Queue Length'] < state['LQa']:
                        data['Cumulative Stats']['Normal']['Max']['Queue Length'] = state['LQa']
                    # adding cutomer in the Amateur queue
                    data['Queue']['Amateur'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
                    if R() <= 0.15:#if the possiblity of a customer getting tired is extant
                        fel_maker(future_event_list, 'Customer Tierd and Departure', clock, Customer_Type, Customer_Number, state['LQa'])
            else:
                #beginning of Professional service
                state['Lp'] += 1
                data['Queue']['Normal']['Without Waiting in Queue'].append(Customer_Number)
                data['Waiting Time']["Normal"].append(0)
                data['Customer']['Normal']['Data'][Customer_Number]['Start Service'] = clock
                fel_maker(future_event_list,'Professional Task End', clock, Customer_Type, Customer_Number)
        else:
            # beginning of Amateur service
            state['La'] += 1
            data['Queue']['Normal']['Without Waiting in Queue'].append(Customer_Number)
            data['Waiting Time']["Normal"].append(0)
            data['Customer']['Normal']['Data'][Customer_Number]['Start Service'] = clock
            fel_maker(future_event_list,'Amateur Task End', clock, Customer_Type, Customer_Number)
    fel_maker_Customer_Arrival(future_event_list,state,clock)

def Professional_Task_End(future_event_list, state, clock, data, system):
    Customer_Type = future_event_list[0]['Customer Type']
    Customer_Number = future_event_list[0]['Customer Number']
    if R() <= 0.15 : #if customer need technical service
        data['Cumulative Stats'][Customer_Type]['Technical Need'] += 1
        if state['Lt'] == system['Technical'] :
            #remove the customer who didn't wait in first queue but wait in technical queue
            if Customer_Number in data['Queue'][Customer_Type]['Without Waiting in Queue']:
                data['Queue'][Customer_Type]['Without Waiting in Queue'].remove(Customer_Number)
            if Customer_Type == "VIP" :
                state['LQtv'] += 1
                # Calculate Maximum VIP Technical Queue Length
                if data['Cumulative Stats']['VIP']['Max']['Technical Queue Length'] < state['LQtv'] :
                    data['Cumulative Stats']['VIP']['Max']['Technical Queue Length'] = state['LQtv']
                # Adding customer in technical queue
                data['Queue']["VIP"]['Technical'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
            else:
                state['LQtn'] += 1
                # Calculate Maximum Normal Technical Queue Length
                if data['Cumulative Stats']['Normal']['Max']['Technical Queue Length'] < state['LQtn'] :
                    data['Cumulative Stats']['Normal']['Max']['Technical Queue Length'] = state['LQtn']
                # Adding customer in technical queue
                data['Queue']['Normal']['Technical'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
        else:
            state['Lt'] += 1
            fel_maker(future_event_list, 'Technical Team Task End', clock, Customer_Type, Customer_Number)
    else:
        data['Customer'][Customer_Type]['Out'] += 1
        #calculate Number of customers whitout waiting in Queue
        if Customer_Number in data['Queue'][Customer_Type]['Without Waiting in Queue'] :
            data['Queue'][Customer_Type]['Without Waiting in Queue'].remove(Customer_Number)
            data['Cumulative Stats'][Customer_Type]['Without Waiting'] += 1
        data['Cumulative Stats'][Customer_Type]['Waiting Time'] += (clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
        data['Customer'][Customer_Type]['Data'][Customer_Number]['End Service'] = clock
        del data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]
    if state['LQp'] > 0 :
        state['LQp'] -= 1
        fel_maker(future_event_list, 'Professional Task End', clock,
                  data['Queue']['Professional'][0]["Customer Type"], data['Queue']['Professional'][0]["Customer Number"])
        #remove Tierd and Departure of customer who start service
        Delete_event(future_event_list, data['Queue']['Professional'][0]["Customer Type"], data['Queue']['Professional'][0]["Customer Number"])
        #Calculate VIP Area Queue wait
        data['Cumulative Stats']["VIP"]['Area']['Queue Wait'] += (clock - data['Queue']['Professional'][0]["Customer Arrival"])
        #Calculate Maximum VIP Queue wait Time
        if data['Cumulative Stats']["VIP"]['Max']['Queue Wait'] < clock - data['Queue']['Professional'][0]["Customer Arrival"] :
            data['Cumulative Stats']["VIP"]['Max']['Queue Wait'] = clock - data['Queue']['Professional'][0]["Customer Arrival"]
        data['Waiting Time']["VIP"].append((clock - data['Queue']['Professional'][0]["Customer Arrival"]))
        data['Customer']['VIP']['Data'][data['Queue']['Professional'][0]["Customer Number"]]['Start Service'] = clock
        #Customer left queue
        data['Queue']['Professional'].remove(data['Queue']['Professional'][0])
    else:
        if state['LQa'] > 0 :
            state['LQa'] -= 1
            fel_maker(future_event_list, 'Professional Task End', clock,
                      data['Queue']['Amateur'][0]["Customer Type"], data['Queue']['Amateur'][0]["Customer Number"])
            # remove Tierd and Departure of customer who start service
            Delete_event(future_event_list, data['Queue']['Amateur'][0]["Customer Type"], data['Queue']['Amateur'][0]["Customer Number"])
            # Calculate Normal Area Queue wait
            data['Cumulative Stats']['Normal']['Area']['Queue Wait'] += (clock - data['Queue']['Amateur'][0]["Customer Arrival"])
            # Calculate Maximum Normal Queue wait Time
            if data['Cumulative Stats']['Normal']['Max']['Queue Wait'] < clock - data['Queue']['Amateur'][0]["Customer Arrival"]:
                data['Cumulative Stats']['Normal']['Max']['Queue Wait'] = clock - data['Queue']['Amateur'][0]["Customer Arrival"]
            data['Waiting Time']["Normal"].append((clock - data['Queue']['Amateur'][0]["Customer Arrival"]))
            data['Customer']['Normal']['Data'][data['Queue']['Amateur'][0]["Customer Number"]]['Start Service'] = clock
            # Customer left queue
            data['Queue']['Amateur'].remove(data['Queue']['Amateur'][0])
        else:
            state['Lp'] -= 1


def Amateur_Task_End(future_event_list, state, clock, data, system):
    Customer_Type = future_event_list[0]['Customer Type']
    Customer_Number = future_event_list[0]['Customer Number']
    if R() <= 0.15 : #if customer need technical service
        data['Cumulative Stats'][Customer_Type]['Technical Need'] += 1
        if state['Lt'] == system['Technical'] :
            #remove the customer who didn't wait in first queue but wait in technical queue
            if Customer_Number in data['Queue'][Customer_Type]['Without Waiting in Queue']:
                data['Queue'][Customer_Type]['Without Waiting in Queue'].remove(Customer_Number)
            state['LQtn'] += 1
            # Calculate Maximum Normal Technical Queue Length
            if data['Cumulative Stats']['Normal']['Max']['Technical Queue Length'] < state['LQtn']:
                data['Cumulative Stats']['Normal']['Max']['Technical Queue Length'] = state['LQtn']
            #Adding customer in technical queue
            data['Queue']['Normal']['Technical'].append({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": clock})
        else:
            state['Lt'] +=1
            fel_maker(future_event_list, 'Technical Team Task End', clock, Customer_Type, Customer_Number)
    else:
        data['Customer'][Customer_Type]['Out'] += 1
        #calculate Number of customers whitout waiting in Queue
        if Customer_Number in data['Queue'][Customer_Type]['Without Waiting in Queue'] :
            data['Queue'][Customer_Type]['Without Waiting in Queue'].remove(Customer_Number)
            data['Cumulative Stats'][Customer_Type]['Without Waiting'] += 1
        data['Cumulative Stats'][Customer_Type]['Waiting Time'] += (clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
        data['Customer'][Customer_Type]['Data'][Customer_Number]['End Service'] = clock
        del data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]
    if state['LQa'] > 0 :
        state['LQa'] -= 1
        fel_maker(future_event_list, 'Amateur Task End', clock,
                  data['Queue']['Amateur'][0]["Customer Type"], data['Queue']['Amateur'][0]["Customer Number"])
        #remove Tierd and Departure of customer who start service
        Delete_event(future_event_list, data['Queue']['Amateur'][0]["Customer Type"], data['Queue']['Amateur'][0]["Customer Number"])
        #Calculate Normal Area Queue wait
        data['Cumulative Stats']['Normal']['Area']['Queue Wait'] += (clock - data['Queue']['Amateur'][0]["Customer Arrival"])
        #Calculate Normal Maximum Queue wait
        if data['Cumulative Stats']['Normal']['Max']['Queue Wait'] < clock - data['Queue']['Amateur'][0]["Customer Arrival"] :
            data['Cumulative Stats']['Normal']['Max']['Queue Wait'] = clock - data['Queue']['Amateur'][0]["Customer Arrival"]
        data['Waiting Time']["Normal"].append((clock - data['Queue']['Amateur'][0]["Customer Arrival"]))
        data['Customer']['Normal']['Data'][data['Queue']['Amateur'][0]["Customer Number"]]['Start Service'] = clock
        #Customer left queue
        data['Queue']['Amateur'].remove(data['Queue']['Amateur'][0])
    else:
        state['La'] -= 1

def Technical_Team_Task_End(future_event_list, state, clock ,data, system):
    Customer_Type = future_event_list[0]['Customer Type']
    Customer_Number = future_event_list[0]['Customer Number']
    if state['LQtv'] > 0 :
        state['LQtv'] -= 1
        fel_maker(future_event_list, 'Technical Team Task End', clock ,data['Queue']["VIP"]['Technical'][0]["Customer Type"], data['Queue']["VIP"]['Technical'][0]["Customer Number"])
        # Calculate VIP Area Technical Queue wait
        data['Cumulative Stats']["VIP"]['Area']['Technical Queue Wait'] += clock - data['Queue']["VIP"]['Technical'][0]["Customer Arrival"]
        # Calculate VIP Maximum Technical Queue wait
        if data['Cumulative Stats']["VIP"]['Max']['Technical Queue Wait'] < clock - data['Queue']["VIP"]['Technical'][0]["Customer Arrival"] :
            data['Cumulative Stats']["VIP"]['Max']['Technical Queue Wait'] = clock - data['Queue']["VIP"]['Technical'][0]["Customer Arrival"]
        # Customer left Technical queue
        data['Queue']["VIP"]['Technical'].remove(data['Queue']["VIP"]['Technical'][0])
    else:
        if state['LQtn'] > 0 :
            state['LQtn'] -= 1
            fel_maker(future_event_list, 'Technical Team Task End', clock,
                      data['Queue']['Normal']['Technical'][0]["Customer Type"], data['Queue']['Normal']['Technical'][0]["Customer Number"])
            # Calculate Normal Area Technical Queue wait
            data['Cumulative Stats']['Normal']['Area']['Technical Queue Wait'] += clock - data['Queue']['Normal']['Technical'][0]["Customer Arrival"]
            # Calculate Normal Maximum Technical Queue wait
            if data['Cumulative Stats']['Normal']['Max']['Technical Queue Wait'] < clock - data['Queue']['Normal']['Technical'][0]["Customer Arrival"]:
                data['Cumulative Stats']['Normal']['Max']['Technical Queue Wait'] = clock - data['Queue']['Normal']['Technical'][0]["Customer Arrival"]
            # Customer left Technical queue
            data['Queue']['Normal']['Technical'].remove(data['Queue']['Normal']['Technical'][0])
        else:
            state['Lt'] -= 1
    #calculate customer waiting time in system
    data['Cumulative Stats'][Customer_Type]['Waiting Time'] += (clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
    data['Customer'][Customer_Type]['Data'][Customer_Number]['End Service'] = clock
    del data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]
    data['Customer'][Customer_Type]['Out'] += 1
    # calculate Number of customers whitout waiting in Queue
    if Customer_Number in data['Queue'][Customer_Type]['Without Waiting in Queue']:
        data['Queue'][Customer_Type]['Without Waiting in Queue'].remove(Customer_Number)
        data['Cumulative Stats'][Customer_Type]['Without Waiting'] += 1

def Customer_Tierd_and_Departure(future_event_list, state, clock, data, system):
    Customer_Type = future_event_list[0]['Customer Type']
    Customer_Number = future_event_list[0]['Customer Number']
    if Customer_Type == "VIP" :
        state['LQp'] -= 1
        # Calculate Lost Customers
        data['Cumulative Stats']['VIP']['Lost Customers'] +=1
        # Calculate VIP Area Queue wait
        data['Cumulative Stats']["VIP"]['Area']['Queue Wait'] += (clock - data['Queue']['Professional'][0]["Customer Arrival"])
        # Calculate VIP Maximume Queue wait
        if data['Cumulative Stats']["VIP"]['Max']['Queue Wait'] < clock - data['Queue']['Professional'][0]["Customer Arrival"] :
            data['Cumulative Stats']["VIP"]['Max']['Queue Wait'] = clock - data['Queue']['Professional'][0]["Customer Arrival"]
        data['Waiting Time']["VIP"].append((clock - data['Queue']['Professional'][0]["Customer Arrival"]))
        data['Customer']['VIP']['Data'][Customer_Number]['Start Service'] = clock
        # Customer left queue
        data['Queue']['Professional'].remove({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": data['Customer'][Customer_Type]['Arrival Time'][Customer_Number] })
    else:
        state['LQa'] -= 1
        # Calculate Lost Customers
        data['Cumulative Stats']['Normal']['Lost Customers'] += 1
        # Calculate Normal Area Queue wait
        data['Cumulative Stats']['Normal']['Area']['Queue Wait'] += (clock - data['Queue']['Amateur'][0]["Customer Arrival"])
        # Calculate Normal Maximume Queue wait
        if data['Cumulative Stats']['Normal']['Max']['Queue Wait'] < clock - data['Queue']['Amateur'][0]["Customer Arrival"] :
            data['Cumulative Stats']['Normal']['Max']['Queue Wait'] = clock - data['Queue']['Amateur'][0]["Customer Arrival"]
        data['Waiting Time']["Normal"].append((clock - data['Queue']['Amateur'][0]["Customer Arrival"]))
        data['Customer']['Normal']['Data'][Customer_Number]['Start Service'] = clock
        # Customer left queue
        data['Queue']['Amateur'].remove({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]})
    # calculate customer waiting time in system
    data['Cumulative Stats'][Customer_Type]['Waiting Time'] += (clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
    data['Customer'][Customer_Type]['Data'][Customer_Number]['End Service'] = clock
    del data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]
    data['Customer'][Customer_Type]['Out'] += 1
#funcion to change shift
#1->2 , 2->3 , 3->1
def Shift_Change(future_event_list, state, clock):
    if state['Shift'] == 3 :
        state['Shift'] = 1
        if state['Day'] == 30:
            state['Day'] = 1
        else:
            state['Day'] += 1
    else:
        state['Shift'] += 1
    fel_maker_Shift_Change(future_event_list, clock)

def simulation(simulation_time, system):
    state, future_event_list, data = starting_state()
    clock = 0
    table = []
    step = 1
    future_event_list.append({'Event Type': 'End of Simulation', 'Event Time': simulation_time, 'Customer Type':"-" ,"Customer Number":"-"})
    while clock < simulation_time:
        future_event_list = sorted(future_event_list, key=lambda x: x['Event Time'])
        current_event = future_event_list[0]  # Find imminent event
        clock = current_event['Event Time']  # Advance time
        #Calculating Area for Cumulative states before runing last FEL
        data['Cumulative Stats']['Busy Time']['Amateur'] += state['La'] * (clock - data['Last time'])
        data['Cumulative Stats']['Busy Time']['Professional'] += state['Lp'] * (clock - data['Last time'])
        data['Cumulative Stats']['Busy Time']['Technical'] += state['Lt'] * (clock - data['Last time'])
        data['Cumulative Stats']["VIP"]['Area']['Queue Length'] += state['LQp'] * (clock - data['Last time'])
        data['Cumulative Stats']['Normal']['Area']['Queue Length'] += state['LQa'] * (clock - data['Last time'])
        data['Cumulative Stats']["VIP"]['Area']['Technical Queue Length'] += state['LQtv'] * (clock - data['Last time'])
        data['Cumulative Stats']['Normal']['Area']['Technical Queue Length'] += state['LQtn'] * (clock - data['Last time'])
        data['Last time']= clock
        if clock < simulation_time:  # if current_event['Event Type'] != 'End of Simulation'
            if current_event['Event Type'] == 'Customer Arrival':
                Customer_Arrival(future_event_list, state, clock, data, system)
            elif current_event['Event Type'] == 'Professional Task End':
                Professional_Task_End(future_event_list, state, clock, data, system)
            elif current_event['Event Type'] == 'Amateur Task End':
                Amateur_Task_End(future_event_list, state, clock, data, system)
            elif current_event['Event Type'] == 'Technical Team Task End':
                Technical_Team_Task_End(future_event_list, state, clock, data, system)
            elif current_event['Event Type'] == 'Shift Change':
                Shift_Change(future_event_list, state, clock)
            elif current_event['Event Type'] == 'Customer Tierd and Departure':
                Customer_Tierd_and_Departure(future_event_list, state, clock, data, system)
            future_event_list = sorted(future_event_list, key=lambda x: x['Event Time'])
            future_event_list.remove(current_event)
        # Calculating Cumulative states
        else:
            future_event_list.clear()
        try:
            data['Cumulative Stats']['VIP']['Percent Without Waiting'] = 100 * (data['Cumulative Stats']['VIP']['Without Waiting'] / data['Customer']["VIP"]['Out'])
        except:
            data['Cumulative Stats']['VIP']['Percent Without Waiting'] = '-'
        try:
            data['Cumulative Stats']['Normal']['Percent Without Waiting'] = 100 * (data['Cumulative Stats']['Normal']['Without Waiting'] / data['Customer']["Normal"]['Out'])
        except:
            data['Cumulative Stats']['Normal']['Percent Without Waiting'] = '-'
        try:
            data['Cumulative Stats']['VIP']['Mean Waiting Time'] = data['Cumulative Stats']['VIP']['Waiting Time'] / data['Customer']["VIP"]['Out']
        except:
            data['Cumulative Stats']['VIP']['Mean Waiting Time'] = '-'
        try:
            data['Cumulative Stats']['Normal']['Mean Waiting Time'] = data['Cumulative Stats']['Normal']['Waiting Time'] / data['Customer']["Normal"]['Out']
        except:
            data['Cumulative Stats']['Normal']['Mean Waiting Time'] = '-'
        try:
            data['Cumulative Stats']['VIP']['Mean']['Queue Wait'] = data['Cumulative Stats']['VIP']['Area']['Queue Wait'] / (data['Customer']["VIP"]['Number'] - state["LQp"])
        except:
            data['Cumulative Stats']['VIP']['Mean']['Queue Wait'] = '-'
        try:
            data['Cumulative Stats']['Normal']['Mean']['Queue Wait'] = data['Cumulative Stats']['Normal']['Area']['Queue Wait'] / (data['Customer']["Normal"]['Number'] - state['LQa'])
        except:
            data['Cumulative Stats']['Normal']['Mean']['Queue Wait'] = '-'
        try:
            data['Cumulative Stats']['VIP']['Mean']['Technical Queue Wait'] = data['Cumulative Stats']['VIP']['Area']['Technical Queue Wait'] / (data['Cumulative Stats']['VIP']['Technical Need'] - state['LQtv'])
        except:
            data['Cumulative Stats']['VIP']['Mean']['Technical Queue Wait'] = '-'
        try:
            data['Cumulative Stats']['Normal']['Mean']['Technical Queue Wait'] = data['Cumulative Stats']['Normal']['Area']['Technical Queue Wait'] / (data['Cumulative Stats']['Normal']['Technical Need'] - state['LQtn'])
        except:
            data['Cumulative Stats']['Normal']['Mean']['Technical Queue Wait'] = '-'
        data['Cumulative Stats']['VIP']['Mean']['Queue Length'] = data['Cumulative Stats']['VIP']['Area']['Queue Length'] / clock
        data['Cumulative Stats']['Normal']['Mean']['Queue Length'] = data['Cumulative Stats']['Normal']['Area']['Queue Length'] / clock
        data['Cumulative Stats']['VIP']['Mean']['Technical Queue Length'] = data['Cumulative Stats']['VIP']['Area']['Technical Queue Length'] / clock
        data['Cumulative Stats']['Normal']['Mean']['Technical Queue Length'] = data['Cumulative Stats']['Normal']['Area']['Technical Queue Length'] / clock
        data['Cumulative Stats']['Productivity']['Amateur'] = data['Cumulative Stats']['Busy Time']['Amateur'] / (clock * system['Amateur'])
        data['Cumulative Stats']['Productivity']['Professional'] = data['Cumulative Stats']['Busy Time']['Professional'] / (clock * system['Professional'])
        data['Cumulative Stats']['Productivity']['Technical'] = data['Cumulative Stats']['Busy Time']['Technical'] / (clock * system['Technical'])
        table.append(create_row(step, current_event, state, data, future_event_list))
        step += 1
    #Creating Output excel
    #excel_main_header = create_main_header(state, data)
    #justify(table)
    #create_excel(table, excel_main_header)
    return data
system = dict()
system['Amateur'] = 3
system['Professional'] = 2
system['Technical'] = 2
L = []
replications = 2000
sim_time = 60*24*3

"""for i in range(replications):
    L.append(simulation(sim_time , system)['Waiting Time'])
    print(i)
    if i == 0:
        NN = len(L[0]['Normal'])
        NV = len(L[0]['VIP'])
    else:
        if NN > len(L[-1]['Normal']):
            NN = len(L[-1]['Normal'])
        if NV > len(L[-1]['VIP']):
            NV = len(L[-1]['VIP'])
WV = [0] * NV
WN = [0] * NN
for i in range(NV):
    for j in range(replications):
        WV[i] += L[j]['VIP'][i]
for i in range(NN):
    for j in range(replications):
        WN[i] += L[j]['Normal'][i]
for i in range(NV):
    WV[i] = WV[i] / replications
for i in range(NN):
    WN[i] = WN[i] / replications
plt.plot(WV)
plt.plot(WN)
plt.show()"""

# Initialize parameters
num_of_replications = 1000
num_of_days = 25
frame_length = 300
window_size = 9
tick_spacing = 5

# Set font and font size
mpl.rc('font', family='Times New Roman')
mpl.rc('font', size=12)

# Create an empty figure with two subplots
fig, ax = plt.subplots(nrows=2, ncols=1, figsize=(8, 6))

# Set up a data structure to save required outputs in each replication
finishing_customers_frame_count = dict()  # keys are replications
waiting_time_frame_aggregate = dict()  # keys are replications


# Function to calculate moving average of a list over a sliding window of length m.
def moving_average(input_list, m):
    output_list = []
    n = len(input_list)
    for i in range(n):
        output_list.append(sum(input_list[max(i - m // 2, 2 * i - n + 1, 0):min(i + m // 2 + 1, 2 * i + 1, n)]) / (
                min(i + m // 2, 2 * i, n - 1) - max(i - m // 2, 2 * i - n + 1, 0) + 1))
    return output_list


# Function to calculate the number of customers who finish getting service in one time-frame
# frame: [start_time, end_time]
def calculate_number_of_finishing_customers(start_time, end_time, data):
    number_of_finishing_customers = 0
    for customer in customers_data:
        if start_time < data[customer]['End Service'] <= end_time:
            number_of_finishing_customers += 1
        elif customers_data[customer]['End Service'] > end_time:
            break

    return number_of_finishing_customers


def calculate_aggregate_queue_waiting_time(start_time, end_time, customers_data):

    cumulative_waiting_time = 0

    for customer in customers_data:
        # if the customer has arrived in this time-frame ...
        if start_time <= customers_data[customer]['Arrival'] < end_time:
            # if the customer starts getting service in this time-frame...
            if customers_data[customer]['Start Service'] < end_time:
                cumulative_waiting_time += customers_data[customer]['Start Service'] - \
                                           customers_data[customer]['Arrival']
            # else if the customer will start getting service after this time-frame...
            else:
                cumulative_waiting_time += end_time - \
                                           customers_data[customer]['Arrival']
        # if the customer has arrived before the beginning of this time-frame
        # but starts getting service during this time-frame...
        elif start_time < customers_data[customer]['Start Service'] < end_time:
            cumulative_waiting_time += customers_data[customer]['Start Service'] - \
                                       start_time
        # There might be another (very rare) corner case. What is it? Handle it if you want.

        elif customers_data[customer]['Arrival'] > end_time:
            break

    return cumulative_waiting_time


simulation_time = num_of_days * 1440
# Just use the frames with full information (drop last 2 frames)
num_of_frames = simulation_time // frame_length - 2
x = [i for i in range(1, num_of_frames + 1)]

for replication in range(1, num_of_replications + 1):

    simulation_data = simulation(num_of_days * 1440,system)
    # print(simulation_data)
    customers_data = simulation_data['Customer']['Normal']['Data']
    print(replication)
    finishing_customers_frame_count[replication] = []
    waiting_time_frame_aggregate[replication] = []

    # do calculations frame by frame
    for time in range(0, num_of_frames * frame_length, frame_length):
        finishing_customers_frame_count[replication].append(
            calculate_number_of_finishing_customers(time, time + frame_length, customers_data))

        waiting_time_frame_aggregate[replication].append(
            calculate_aggregate_queue_waiting_time(time, time + frame_length, customers_data))

    # plot outputs in each replication
    # ax[0].plot(x, finishing_customers_frame_count[replication], alpha=0.5)
    # ax[1].plot(x, waiting_time_frame_aggregate[replication], alpha=0.5)

finishing_customers_replication_average = []
waiting_time_replication_average = []

for i in range(num_of_frames):
    average_finishing_customers = 0
    average_waiting_time = 0

    for replication in range(1, num_of_replications + 1):
        average_finishing_customers += finishing_customers_frame_count[replication][i] * (1 / num_of_replications)
        average_waiting_time += waiting_time_frame_aggregate[replication][i] * (1 / num_of_replications)

    finishing_customers_replication_average.append(average_finishing_customers)
    waiting_time_replication_average.append(average_waiting_time)

finishing_customers_moving_replication_average = moving_average(finishing_customers_replication_average, window_size)
waiting_time_moving_replication_average = moving_average(waiting_time_replication_average, window_size)

fig.suptitle(f'Warm-up analysis over {num_of_replications} replications')

ax[0].plot(x, finishing_customers_replication_average, 'r', linewidth=5, label="Average across replications")
ax[0].plot(x, finishing_customers_moving_replication_average, 'k', label=f'Moving average (m = {window_size})')
ax[0].set_title('Number of Finishing Customers')
ax[0].set_xlabel('Frame No.')
ax[0].set_ylabel('Number of Finishing Customers')
ax[0].xaxis.set_major_locator(ticker.MultipleLocator(tick_spacing))
ax[0].legend()

ax[1].plot(x, waiting_time_replication_average, 'r', linewidth=5, label="Average across replications")
ax[1].plot(x, waiting_time_moving_replication_average, 'k', label=f'Moving average (m = {window_size})')
ax[1].set_title('Aggregate Waiting Time')
ax[1].set_xlabel('Frame No.')
ax[1].set_ylabel('Aggregate Waiting Time')
ax[1].xaxis.set_major_locator(ticker.MultipleLocator(tick_spacing))
ax[1].legend()

fig.tight_layout()
fig.show()
fig.savefig('Single server queue - Warm-up analysis - Time-Frame Approach')



