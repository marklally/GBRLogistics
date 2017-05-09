'''
Created on 2 May 2017

@author: markl
'''
import os
#from openpyxl.reader.excel import load_workbook
import pandas
import copy
import datetime
from pandas.core.frame import DataFrame

LegNum = 1
LegDay = 2
LegStartLoc = 3
Distance = 4
DropOffTime = 7
LegStartTime = 8
DrivingTime = 9
DropperSenior = 14
DropperVets = 15
RunnerSenior = 10
RunnerVets = 11
CollectorSenior = 18
CollectorVets = 19

PACE = 6 #minutes per mile!!

def get_runner_schedule(name, row, RunnerIndex, DropperIndex, CollectorIndex):
    start = datetime.datetime(2000, 1, 1, hour = row[8].hour, minute = row[8].minute, second = row[8].second)
    end = str((start + datetime.timedelta(minutes=PACE * row[Distance])).time())
    
    if(row[DropperIndex] == name):
        #print(str(row[7]) + ": Leg " + str(row[1]) + ": " + row[RunnerIndex] + " drives from " + str(row[16]) + " to " + row[3] + " for start of their leg" )
        print(str(row[DropOffTime]) + ": Leg " + str(row[1]) + ": " + row[RunnerIndex] + " drives to " + row[3] + " for start of their leg" )
    else:
        #print(str(row[7]) + ": Leg " + str(row[1]) + ": " +row[RunnerIndex] + " is driven from " + str(row[16]) + " to " + row[3] + " by " + row[DropperIndex] )
        print(str(row[DropOffTime]) + ": Leg " + str(row[1]) + ": " +row[RunnerIndex] + " is driven to " + row[3] + " by " + row[DropperIndex] )
        
    print(str(row[LegStartTime]) + ": Leg " + str(row[1]) + ": " +row[RunnerIndex] + " starts running")
        
    if(row[CollectorIndex]== name):
        print(end + ": Leg " + str(row[1]) + ": " + row[RunnerIndex] + " picks up their car at the end of the leg" )
    else:
        print(end + ": Leg " + str(row[1]) + ": " +row[RunnerIndex] + " is picked up by " + row[CollectorIndex] + " at the end of their run " )
        
def get_schedule(name, leg_data):
    for row in leg_data.itertuples():
        if(row[RunnerSenior] == name ):
            get_runner_schedule(name, row, RunnerIndex = RunnerSenior, DropperIndex= DropperSenior, CollectorIndex= CollectorSenior )
        if(row[RunnerVets] == name ):
            get_runner_schedule(name, row, RunnerIndex = RunnerVets, DropperIndex= DropperVets, CollectorIndex= CollectorVets )



def create_single_drop_event(rowdata, dropevent, dropper, runner):
    seniordropevent = copy.deepcopy(dropevent)
    seniordropevent["Details"] = "{}: {} drops {} to the start of Stage {} ".format(
        rowdata["Drop Off Time"], rowdata[dropper], rowdata[runner], rowdata["Stage"])
    seniordropevent["People"] = set([rowdata[dropper], rowdata[runner]])
    return seniordropevent

def create_single_collect_event(rowdata, collect_time, collector, runner):
    collect_event = initialise_event(collect_time)
    collect_event["Details"] = "{}: {} arrives to pick up {} at the end of Stage {} ".format(
        collect_time, rowdata[collector], rowdata[runner], rowdata["Stage"])
    collect_event["People"] = set([rowdata[collector], rowdata[runner]])
    return collect_event

def create_run_event(rowdata, runner):
    runevent = {"Time": rowdata["Start Time"]}
    runevent["Details"] = "{}: {} starts running Stage {} ".format(
        rowdata["Start Time"],  rowdata[runner], rowdata["Stage"])
    runevent["People"] = set([rowdata[runner]])
    return runevent


def create_drop_events(rowdata):
    dropevents = []
    
    dropevent = {"Time": rowdata["Drop Off Time"]}
    
    if (rowdata["Dropper Senior"] == rowdata["Dropper Vets"]):
        bothdropevent = dropevent
        bothdropevent["Time"] = rowdata["Drop Off Time"]
        bothdropevent["Details"] = "{}: {} drops {} and {} to the start of Stage {} ".format(
            rowdata["Drop Off Time"], rowdata["Dropper Senior"], rowdata["Senior Runner"], rowdata["Veteran Runner"], rowdata["Stage"])
        bothdropevent["People"] = set([rowdata["Dropper Senior"], rowdata["Senior Runner"], rowdata["Veteran Runner"]])
        dropevents.append(bothdropevent)
    else:
        dropevents.append(create_single_drop_event(rowdata, dropevent, "Dropper Vets", "Veteran Runner")) 
        dropevents.append(create_single_drop_event(rowdata, dropevent, "Dropper Senior", "Senior Runner"))      

    return dropevents

def initialise_event(time):
    return {"Time": time}

def create_collect_events(rowdata):
    collect_events = []
    
    #early collect time based on quick runners - could make this better with actual runner paces
    start_time = rowdata["Start Time"]
    start_date_time = datetime.datetime(2000, 1, 1, hour = start_time.hour, minute = start_time.minute, second = start_time.second)
    collect_time = str((start_date_time + datetime.timedelta(minutes=PACE * rowdata["Distance"])).time())
    
    if (rowdata["Collector Senior"] == rowdata["Collector Vet"]):
        bothdropevent = initialise_event(collect_time)
        bothdropevent["Details"] = "{}: {} arrives to pick up {} and {} at the end of Stage {} ".format(
            collect_time, rowdata["Collector Senior"], rowdata["Senior Runner"], rowdata["Veteran Runner"], rowdata["Stage"])
        bothdropevent["People"] = set([rowdata["Collector Senior"], rowdata["Senior Runner"], rowdata["Veteran Runner"]])
        collect_events.append(bothdropevent)
    else:
        collect_events.append(create_single_collect_event(rowdata, collect_time, "Collector Vet", "Veteran Runner")) 
        collect_events.append(create_single_collect_event(rowdata, collect_time, "Collector Senior", "Senior Runner"))      

    return collect_events

def get_timeline(leg_data):
    timeline = []
        
    for row in leg_data.iterrows():
        rowdata = row[1]
        timeline = timeline + create_drop_events(rowdata) 
        timeline.append(create_run_event(rowdata, "Veteran Runner")) 
        timeline.append(create_run_event(rowdata, "Senior Runner"))
        timeline = timeline + create_collect_events(rowdata) 
    
    df_timeline = DataFrame(timeline)    
    return(df_timeline)
  
def print_full_timeline(timeline):
    #timeline.sort_values("Time")
    for row in timeline.iterrows():
        print(row[1]["Details"])
        
def print_individual_timelines(timeline):
    people = set.union(*timeline["People"].tolist())
    for person in people: 
        print("\nSchedule for " + person)
        #person_timeline = timeline.query("@person in People")
        for row in timeline.iterrows():
            if person in row[1]["People"]:
                print(row[1]["Details"]) 
                          
def main():
    filename = "../data/GBR May 2017.xlsx"
    print(os.getcwd())
    
    gbr_data = pandas.read_excel(filename,sheetname = None) #, converters = {"Start Time": str} )
    #print(gbr_data.keys())
    leg_data = gbr_data['Legs']
    people_data = gbr_data['People Data']
    names = people_data['Name']
    #print(leg_data.keys())
    
    #===========================================================================
    # for name in names:
    #     print("\nSchedule for " + name)
    #     for day in ["Saturday" , "Sunday"]:
    #         print(day)
    #         get_schedule(name, leg_data.query('Day == @day'))
    #         get_timeline(leg_data)
    #     print("\n")
    #===========================================================================
        
    for day in ["Saturday" , "Sunday"]:
        print("\n" + day)
        timeline_day = get_timeline(leg_data.query('Day == @day'))
        print_full_timeline(timeline_day)
        print_individual_timelines(timeline_day)
        
    #===========================================================================
    # for row in legdata.rows:
    #     for cell in row:
    #         if cell.value != 'None':
    #             print(cell.value)
    #===========================================================================
if __name__ == "__main__":
    main()
