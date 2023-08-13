import random
import math
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib as mlp
import xlsxwriter

def create_main_header (data):
    header = ['Step']
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

def create_row(step, data):
    row = [step]
    row.extend([data['Cumulative Stats']['VIP']['Percent Without Waiting'], data['Cumulative Stats']['Normal']['Percent Without Waiting'],
    data['Cumulative Stats']['VIP']['Mean Waiting Time'], data['Cumulative Stats']['Normal']['Mean Waiting Time'],
    data['Cumulative Stats']['VIP']['Lost Customers'], data['Cumulative Stats']['Normal']['Lost Customers']])
    row.extend(list(data['Cumulative Stats']['VIP']['Max'].values()))
    row.extend(list(data['Cumulative Stats']['VIP']['Mean'].values()))
    row.extend(list(data['Cumulative Stats']['Normal']['Max'].values()))
    row.extend(list(data['Cumulative Stats']['Normal']['Mean'].values()))
    row.extend(list(data['Cumulative Stats']['Productivity'].values()))
    return row

def get_col_widths(dataframe):
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def create_excel(table, header):
    df = pd.DataFrame(table, columns=header, index=None)
    writer = pd.ExcelWriter('data_S2.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Call-Center Output', header=False, startrow=1, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Call-Center Output']
    header_formatter = workbook.add_format()
    header_formatter.set_align('center')
    header_formatter.set_align('vcenter')
    header_formatter.set_font('Times New Roman')
    header_formatter.set_bold('True')
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_formatter)
    for i, width in enumerate(get_col_widths(df)):
        worksheet.set_column(i - 1, i - 1, width)
    main_formatter = workbook.add_format()
    main_formatter.set_align('center')
    main_formatter.set_align('vcenter')
    main_formatter.set_font('Times New Roman')
    for row in range(1, len(df) + 1):
        worksheet.set_row(row, None, main_formatter)
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
    return exponential(1/2.7)
#function to Proffesional time service variable with D2 distribution
def S2():
    return exponential(1/5.8)
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
                fel_maker(future_event_list,'Professional Task End', clock, Customer_Type, Customer_Number)
        else:
            # beginning of Amateur service
            state['La'] += 1
            data['Queue']['Normal']['Without Waiting in Queue'].append(Customer_Number)
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
        data['Waiting Time'][Customer_Type].append(clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
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
        data['Waiting Time'][Customer_Type].append(clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
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
    data['Waiting Time'][Customer_Type].append(clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
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
        # Customer left queue
        data['Queue']['Amateur'].remove({"Customer Type": Customer_Type, "Customer Number": Customer_Number, "Customer Arrival": data['Customer'][Customer_Type]['Arrival Time'][Customer_Number]})
    # calculate customer waiting time in system
    data['Cumulative Stats'][Customer_Type]['Waiting Time'] += (clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
    data['Waiting Time'][Customer_Type].append(clock - data['Customer'][Customer_Type]['Arrival Time'][Customer_Number])
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

def data_gathering(data, state, clock, system):
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

def simulation(simulation_time, system, run, table):
    for x in range(run):
        state, future_event_list, data = starting_state()
        clock = 0
        future_event_list.append({'Event Type': 'End of Simulation', 'Event Time': simulation_time, 'Customer Type':"-" ,"Customer Number":"-"})
        while clock < simulation_time:
            future_event_list = sorted(future_event_list, key=lambda x: x['Event Time'])
            current_event = future_event_list[0]  # Find imminent event
            clock = current_event['Event Time']  # Advance time
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
            else:
                future_event_list.clear()
        data_gathering(data ,state ,clock ,system)
        print(x)
        table.append(create_row(x+1, data))

system = dict()
system['Amateur'] = 2
system['Professional'] = 2
system['Technical'] = 2
def estimate(Amateur, Professional, Technical, run, simulation_time):
    state, future_event_list, data = starting_state()
    excel_main_header = create_main_header(data)
    system = dict()
    table = []
    system['Amateur'] = Amateur
    system['Professional'] = Professional
    system['Technical'] = Technical
    simulation(simulation_time, system, run, table)
    create_excel(table, excel_main_header)

estimate(2,2,2,1000,60*24*30)

