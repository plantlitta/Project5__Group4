import streamlit as st
import pandas as pd
import math as m
import random as r 
import matplotlib.pyplot as plt
from datetime import*
from collections import Counter
import contextlib
import io
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter, MultipleLocator
import docx
import xlsxwriter

# '''
# MAIN PAGE OF THE WEBSITE
# '''

#####################
st.write('# **SCHEDULE OPTIMAZER**')
st.write('*Made by Group 4 during project 5*')
st.divider()

#####################
# uploading file and dissplaying it
st.header('Upload files')
count = 0
# status = 0
file_names = [] # just in case we want to store names of files somewhere it can be done from here
all_data = []
data_dis_m = pd.DataFrame()
data_bus = pd.DataFrame()
data_timet = pd.DataFrame()

#####################

uploaded_file_distance_m = st.file_uploader('Excel file containing Distance Matrix', help='Choose an excel file from your device')

if uploaded_file_distance_m is not None:
    count += 1
    file_names.append(uploaded_file_distance_m.name)
    df = pd.read_excel(uploaded_file_distance_m)
    data_dis_m = pd.DataFrame(df)
    all_data.append(data_dis_m)
    st.write(data_dis_m.head())
    st.toast(f'Operation Done. File {uploaded_file_distance_m.name} is uploaded...')
else:
    st.warning('Upload a file with a type xlsx')

#####################

uploaded_file_busplanning = st.file_uploader('Excel file containing Bus Planning', help='Choose an excel file from your device')

if uploaded_file_busplanning is not None:
    count += 1
    file_names.append(uploaded_file_busplanning.name)
    df = pd.read_excel(uploaded_file_busplanning)
    data_bus = pd.DataFrame(df)
    all_data.append(data_bus)
    st.write(data_bus.head())
    st.toast(f'Operation Done. File {uploaded_file_busplanning.name} is uploaded...')
else:
    st.warning('Upload a file with a type xlsx')

#####################

uploaded_file_Timetable = st.file_uploader('Excel file containing Time Table', help='Choose an excel file from your device')

if uploaded_file_Timetable is not None:
    count += 1
    file_names.append(uploaded_file_Timetable.name)
    df = pd.read_excel(uploaded_file_Timetable)
    data_timet = pd.DataFrame(df)
    all_data.append(data_timet)
    st.write(data_timet.head())
    st.toast(f'Operation Done. File {uploaded_file_Timetable.name} is uploaded...')
    # status = 1
else:
    st.warning('Upload a file with a type xlsx')


if count != 3:
    st.write('Please upload all files.')
else:
    st.write(f'You have uploaded {count} files.')

    st.divider()

    st.write('# Preparation')
    # cleaning function

    def data_cleaning(timetable, distance_matrix, planning):
            #Time to minutes 
        def time_to_minutes(value):
            if pd.isna(value):
                return None
            try:
                if isinstance(value, (datetime, time, pd.Timestamp)):
                    return value.hour * 60 + value.minute

                value = str(value).strip()
                parts = value.split(':')
                hours = int(parts[0])
                minutes = int(parts[1])

                if hours > 23 or hours < 0 or minutes > 59 or minutes < 0:
                    return None

                return hours * 60 + minutes
            except (ValueError, TypeError, IndexError):
                return None


        #time to datetime
        def time_to_datetime(value, first_day):
            if pd.isna(value):
                return pd.NaT
            
            if isinstance(value, (datetime, pd.Timestamp)):
                return pd.Timestamp(value)

            minutes = time_to_minutes(value)
            if minutes is None:
                return pd.NaT
            return first_day + pd.Timedelta(minutes=minutes)

    
        timetable = timetable.copy()
        distance_matrix = distance_matrix.copy()
        planning = planning.copy()
        changes = []
        issues = []


        #FIRST checking if the excel has the collumns we need
        timetable_columns = ['start', 'end', 'departure_time', 'line']
        distance_columns = ['start', 'end', 'line', 'min_travel_time', 'distance_m']
        planning_columns = ['start location', 'end location', 'start time', 'end time', 'activity', 'energy consumption', 'bus', 'line']

        for column in timetable_columns:
            if column not in timetable.columns:
                raise ValueError("Missing timetable column: " + column)
        for column in distance_columns:
            if column not in distance_matrix.columns:
                raise ValueError("Missing distance column: " + column)
        for column in planning_columns:
            if column not in planning.columns:
                raise ValueError("Missing planning column: " + column)

        #we need this to see which rows have errors
        planning['_row'] = range(2, len(planning) + 2)
        timetable['_row'] = range(2, len(timetable) + 2)
        distance_matrix['_row'] = range(2, len(distance_matrix) + 2)


        #SECOND checking if theres something wrong with the names of everything
        text_columns = {
            'timetable': (timetable, ['start', 'end']),
            'distance_matrix': (distance_matrix, ['start', 'end']),
            'planning': (planning, ['start location', 'end location', 'activity'])
        }

        for file_name, (data, columns) in text_columns.items():
            for column in columns:
                for i in data.index:
                    value = data.at[i, column]
                    #check text
                    if not isinstance(value, str):
                        issues.append([file_name, data.at[i, '_row'], column, value, 'Not text'])
                        continue
                    #check extra spaces
                    if value != value.strip():
                        issues.append([file_name, data.at[i, '_row'], column, value, 'Extra spaces'])
                    #check activity names
                    if column == 'activity':
                        if value not in ['service trip', 'material trip', 'idle', 'charging']:
                            issues.append([file_name, data.at[i, '_row'], column, value, 'Check activity name'])

        #checking if the names of the stops are correct of if theres a typo anywehere
        valid_stops = set(distance_matrix['start'].dropna()) | set(distance_matrix['end'].dropna())

        for column in ['start location', 'end location']:
            for i in planning.index:
                value = planning.at[i, column]
                if not isinstance(value, str) or value not in valid_stops:
                    issues.append(['planning', planning.at[i, '_row'], column, value, 'Unknown bus stop'])


        #idk if its needed but checking duplicates
        duplicates = planning.drop(columns=['_row']).duplicated()

        for i in planning[duplicates].index:
            issues.append(['planning', planning.at[i, '_row'], 'all', None, 'Duplicate row'])


        #time into datetime
        first_day = pd.Timestamp('2026-10-08') #idk what to do help?

        original_start = planning['start time'].copy()
        original_end = planning['end time'].copy()

        planning['start time'] = planning['start time'].apply(lambda value: time_to_datetime(value, first_day))
        planning['end time'] = planning['end time'].apply(lambda value: time_to_datetime(value, first_day))
        timetable['departure_time'] = timetable['departure_time'].apply(lambda value: time_to_datetime(value, first_day))

        #checking wrong times
        for column, original in [('start time', original_start), ('end time', original_end)]:
            for i in planning[planning[column].isna()].index:
                issues.append(['planning', planning.at[i, '_row'], column, original.loc[i], 'Wrong time'])

        #checking wrong departure times
        departure_minutes = timetable['departure_time'].apply(time_to_minutes)

        for i in timetable[departure_minutes.isna()].index:
            issues.append(['timetable', timetable.at[i, '_row'], 'departure_time', timetable.at[i, 'departure_time'], 'Wrong time'])


        #checking the time bewten days BUT the plan needs to be in cronologic order *I COULDNT FIGURE THIS OUT THIS IS AI SHIT sorry :(*
        for bus in planning['bus'].dropna().unique():

            bus_rows = planning[planning['bus'] == bus]

            day = 0
            previous_end = None

            for i in bus_rows.index:

                start = planning.at[i, 'start time']
                end = planning.at[i, 'end time']

                if pd.isna(start) or pd.isna(end):
                    continue

                start_clock = start.hour * 60 + start.minute
                end_clock = end.hour * 60 + end.minute

                new_start = start + pd.Timedelta(days=day)
                new_end = end + pd.Timedelta(days=day)

                #if next day starts
                if previous_end is not None:

                    previous_clock = previous_end.hour * 60 + previous_end.minute

                    if 1260 <= previous_clock <= 1439 and 0 <= start_clock <= 180:

                        if new_start < previous_end:
                            day += 1
                            new_start += pd.Timedelta(days=1)
                            new_end += pd.Timedelta(days=1)

                #if activity crosses midnight
                if start_clock >= 1260 and end_clock <= 180:

                    if new_end < new_start:
                        new_end += pd.Timedelta(days=1)

                #save changes
                if new_start != start:
                    changes.append(['planning',planning.at[i, '_row'],'start time',start,new_start,'Next day'])

                if new_end != end:
                    changes.append(['planning',planning.at[i, '_row'],'end time',end,new_end,'Next day'])

                planning.at[i, 'start time'] = new_start
                planning.at[i, 'end time'] = new_end

                #save previous end time
                previous_end = new_end
                day = (new_start.normalize() - first_day).days

        #convert datetime to minutes
        planning['start_minutes'] = ((planning['start time'] - first_day).dt.total_seconds() / 60)

        planning['end_minutes'] = ((planning['end time'] - first_day).dt.total_seconds() / 60)

        timetable['departure_time_minutes'] = ((timetable['departure_time'] - first_day).dt.total_seconds() / 60)


        #calculate duration
        duration = planning['end_minutes'] - planning['start_minutes']
        #print(type(duration))
        #checking negative duration
        wrong_duration = duration < 0

        for i in planning[wrong_duration].index:
            issues.append(['planning', planning.at[i, '_row'], 'start time / end time', duration.loc[i], 'Negative duration'])

        #check idle with 0 minutes and taking them out
        zero_idle = (planning['activity'] == 'idle') & (duration == 0)

        for i in planning[zero_idle].index:
            changes.append(['planning', planning.at[i, '_row'], 'activity', 'idle', 'Deleted', 'idle with 0 min'])

        planning = planning[~zero_idle].copy()

        #calculate duration again
        duration = planning['end_minutes'] - planning['start_minutes']

    #-----------------I added this up because it will make it handy for the calculation 
        #distance lookup
        distance_lookup = {}

        for i in distance_matrix.index:
            start = distance_matrix.at[i, 'start']
            end = distance_matrix.at[i, 'end']
            line = distance_matrix.at[i, 'line']
            distance_m = distance_matrix.at[i, 'distance_m']

            if pd.isna(line):
                line = None

            distance_lookup[(start, end, line)] = distance_m


        #recalculate energy
        for i in planning.index:
            activity = planning.at[i, 'activity']
            duration_min = duration.loc[i]
            old_energy = planning.at[i, 'energy consumption']

            if pd.isna(duration_min) or duration_min < 0:
                continue

            #this is the new stuff
            if activity == 'service trip' or activity== 'material trip':
                start = planning.at[i, 'start location']
                end = planning.at[i, 'end location']
                line = planning.at[i, 'line']

                if pd.isna(line):
                    line = None
                distance_m = distance_lookup.get((start, end, line))

                if distance_m is None or pd.isna(distance_m):
                    issues.append(['planning', planning.at[i, '_row'],'energy consumption', old_energy,'Distance not found'])
                    continue

                new_energy = (float(distance_m) / 1000) * 1.78 #divide by 1000 to put in km 

            elif activity == 'idle':
                new_energy = duration_min / 60 * 5
            elif activity == 'charging':
                new_energy = -(duration_min / 60 * 450)
            else:
                continue

            #check old energy
            old_number = pd.to_numeric(old_energy, errors='coerce')

            #save energy changes
            if pd.isna(old_number) or abs(old_number - new_energy) > 0.001:
                changes.append(['planning',planning.at[i, '_row'],'energy consumption',old_energy,new_energy,'Energy recalculated'])

            planning.at[i, 'energy consumption'] = new_energy

        #make reports on what was wrong
        changes_report = pd.DataFrame(changes, columns=['file', 'row', 'column', 'old value', 'new value', 'reason'])
        issues_report = pd.DataFrame(issues, columns=['file', 'row', 'column', 'value', 'problem'])

        #remove the thing to help identify the rows that had problems 
        timetable = timetable.drop(columns=['_row'])
        distance_matrix = distance_matrix.drop(columns=['_row'])
        planning = planning.drop(columns=['_row'])

        print("Changes made:")
        print(changes_report)
        print("Issues found:")
        print(issues_report)

        return timetable, distance_matrix, planning, changes_report, issues_report

    # feasibility function

    def feasibility_check(input_data_planning):
        df = pd.DataFrame(input_data_planning).copy()

        #i had to put this bc it was giving an error 
        if 'start_minutes' not in df.columns and 'start time' in df.columns:
            df = df.rename(columns={'start time': 'start_minutes'})
        if 'end_minutes' not in df.columns and 'end time' in df.columns:
                df = df.rename(columns={'end time': 'end_minutes'})
        required = ['bus', 'start_minutes', 'end_minutes', 'end location', 'start location', 'activity', 'energy consumption']
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f'missing collumns {missing}')

        #to know where the errors are
        errors = []
        def report(bus, row, reason):
            errors.append((bus, row, reason))

        #separating and sorting
        for bus_id, group in df.groupby('bus', sort=False):
            group = group.sort_values(by=['start_minutes', 'end_minutes'], kind='stable')

            battery = 270.0  #ITS NOT 270*0.9!!! 300*0.9=270
            previous_end = None
            previous_location = None
            first_trip_found = False

            for row_id, activity in group.iterrows():
                start = activity['start_minutes']
                end = activity['end_minutes']
                energy = activity['energy consumption']
                kind = activity['activity']
                location_start = activity['start location']
                location_end = activity['end location']

                if pd.isna(start) or pd.isna(end) or end <= start:
                    report(bus_id, row_id, 'activity invalid')
                if previous_end is not None:
                    if start < previous_end - 1e-8:
                        report(bus_id, row_id, 'overlapping activities')
                    if location_start != previous_location:
                        report(bus_id, row_id, 'problem with location')

                if kind in ('idle', 'charging') and location_start != location_end:
                    report(bus_id, row_id, 'idle or charging at dif locations')
                if kind in ('service trip', 'material trip') and location_start == location_end:
                    report(bus_id, row_id, 'trip has the same location')
                if kind in ('service trip', 'material trip') and not first_trip_found:
                    first_trip_found = True
                    if kind != 'material trip' or location_start != 'ehvgar':
                        report(bus_id, row_id, 'first trip is not material')
                if pd.isna(energy):
                    report(bus_id, row_id, 'energy missing')

                else:
                    if kind == 'charging':
                        if location_start != 'ehvgar':
                            report(bus_id, row_id, 'charges outside of the garage')
                        if energy >= 0:
                            report(bus_id, row_id, 'while charging needs to be negative')
                        if pd.notna(start) and pd.notna(end) and end - start < 15: #we forgot this btw the bus needs to charge for at least 15 min
                            report(bus_id, row_id, 'charging time less then 15 min')
                        if pd.notna(start) and pd.notna(end) and abs(energy + (end-start)*450/60) > 0.1: #the 0.1 can be changed its just margin 
                            report(bus_id, row_id, 'charging time its different then the calculated')
                    elif energy < 0:
                        report(bus_id, row_id, 'energy is wrong, its not charging')
                    battery = min(270.0, battery - float(energy)) #updating the energy
                    if battery < 27:
                        report(bus_id, row_id, f'battery is below 27, ({battery:.1f})')
                previous_end, previous_location = end, location_end
            if previous_location != 'ehvgar':
                report(bus_id, group.index[-1], 'bus does not return to the garage') #JULIEN change 

        if errors:
            print(f'not feasible- {len(errors)} errors- {len(set(e[0] for e in errors))} busses')
            for bus, row, reason in errors[:30]:
                print(f'bus {bus} - row {row} - {reason}')
            if len(errors) > 30:
                print(f'... and {len(errors)-30} more errors')
        else:
            print('its feasible under assumptions')
        return pd.DataFrame(errors, columns=['bus', 'row', 'reason'])

    # displaying cleaned data
    st.write('Check your files before optimization for better results')
    clean_b = st.button('Check Files', help='Before you optimize your files and see the schedule, please, use this button to clean and check your files for feasibility.')

    clean_timet, clean_dis_m, clean_bus, clean_changes_report, clean_issues_report = data_cleaning(data_timet, data_dis_m, data_bus)

    feasible_og_planning_bus = feasibility_check(clean_bus) #clean_bus is the planning


    if clean_b:
        df1 = pd.DataFrame(clean_dis_m)
        df2 = pd.DataFrame(clean_bus) # thiss clean_bus after it became feasible. so its a cleaning + feasibility check in 1
        df3 = pd.DataFrame(clean_timet)
        st.write('Number of buses current planning is using: ', clean_bus['bus'].nunique())
        st.write('Checked Distance Matrix (short version)')
        st.write(df1.head())
        st.write('Checked Planning (short version)')
        st.write(df2.head())
        st.write('Checked Time Table (short version)')
        st.write(df3.head())
        st.toast('Done!')

    st.write('If you would want to review all changes made please use button below.')

    with st.expander('Push to see cleaning report'):
        st.write("Changes made:")
        st.write(clean_changes_report)
        st.write(feasible_og_planning_bus)
        st.write("Issues found:")
        st.write(clean_issues_report) 


    ########

    st.write('# Optimization') 

    garage = "ehvgar"
    capacity = 270.0
    min_state_charge = 27.0
    driving_kwh_km = 1.78
    idle_kw = 5.0
    charge_kw = 450.0
    min_charge_time = 15.0
    tolerance = 1e-7
    cols = ["start location", "end location", "start time","end time", "activity", "line","energy consumption", "bus"]
    boarding = 5 

    #for the routes trips to give us the time and energy nec to perform a trip
    def route(routes, start, end, line=None):
        if start == end:
            return 0., 0. 
        return routes.get((start, end, line), (m.inf, m.inf)) #(min, energy)

    def add_activity(schedule, start_location, end_location,
                    start_time, end_time, activity, energy, bus, line=None):
        schedule.append([start_location,end_location, start_time, end_time,activity,line,energy,bus ])


    def rebuild_bus(indice, trips, routes, bus=1):
        if not indice:
            return []
        indice = sorted(indice, key=lambda i: (trips[i]['departure'], trips[i]['arrival'])) #soorting first by time departure if the same then arrival
        rows = []

        soc, position, ready = capacity, garage, -m.inf

        #a=local start b=local end
        def add(a, b, start_hour, end_hour, kind, energy, line=None):
            nonlocal soc, ready, position
            if start_hour < ready-tolerance:
                return False
            if end_hour + tolerance < start_hour: #check act ends before starting
                return False
            if kind == 'idle' and end_hour - start_hour < tolerance: #idle 0
                return True
            new_soc = soc - energy
            if new_soc < min_state_charge - tolerance or new_soc > capacity + tolerance:
                return False
            rows.append([a, b, start_hour, end_hour, kind, line, energy, bus])
        #uodate the location and battery
            soc = new_soc
            ready = end_hour
            position = b
            return True #give true if its in the limits

        for trip_id in indice:
            tr = trips[trip_id]
            dep, arr, start, end, service_energy = (tr[k] for k in ['departure', 'arrival', 'start', 'end', 'energy'])
            dt, de = route(routes, position, start)
            #if it gives us a number or the departure time - the time that needs to arrive at the next trip < time that the bus is available then cant ride
            if rows:
                extra_time = boarding #Мама, вызывай такси!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            else:
                extra_time=0 

            if not m.isfinite(dt) or dep - dt-extra_time < ready - tolerance:
                return None
            if service_energy + min_state_charge > capacity + tolerance:
                return None

            #checking if the bus has enough battery fo the next trip without going to charge
            wait = max(0., dep - dt - ready) if rows else 0. #if its the first act is 0
            direct_end_soc = soc - (wait * idle_kw / 60) - de - service_energy #battery after srvice trip if the bus doesnt charge
            direct_ok = direct_end_soc >= min_state_charge - tolerance #cehcking if we dont suprass the battery

            #can the bus go to the garage to charge b4 next trip
            gt, ge = route(routes, position, garage) #gtge means garage time and garage energy
            st, se = route(routes, garage, start) #st time to the start point and energy
            can_charge = (m.isfinite(gt) and m.isfinite(st) and rows and soc - ge >= min_state_charge - tolerance)
            charge_duration = 0.
            if can_charge:
                if rows:
                    garage_arrival = ready + gt
                else:
                    garage_arrival=dep-st - min_charge_time

                availab_time = dep - st - garage_arrival #time we have to charge
                max_charge = min(capacity - (soc - ge), max(0., availab_time) * charge_kw / 60)
                if max_charge >= min_charge_time * charge_kw / 60 - tolerance:
                    charge_duration = max_charge / charge_kw * 60
                    #idle after charging and b4 next trip consume energy
                    remain = max(0., availab_time - charge_duration)
                    bat_after_trip = soc - ge + max_charge - remain * idle_kw / 60 - se - service_energy
                    if bat_after_trip < min_state_charge - tolerance:
                        charge_duration = 0.

            #checking if the bus can charge for at least 15 min
            if charge_duration >= min_charge_time - tolerance and (not direct_ok or soc < capacity - 15):
                if rows and gt > tolerance and not add(position, garage, ready, ready + gt, 'material trip', ge): #to the garage
                    return None
                if not rows: #dont charge b4 first trip if the battery is full
                    charge_duration = 0.
                else:
                    start_hour = ready
                    added = charge_duration * charge_kw / 60 #add energy and time
                    if not add(garage, garage, start_hour, start_hour + charge_duration, 'charging', -added):
                        return None

                    gap = dep- st -ready
                    if gap > tolerance and not add(garage, garage, ready,ready + gap,'idle', gap * idle_kw / 60): #adding time in the garage
                        return None

                    if st > tolerance and not add(garage, start, dep - st, dep, 'material trip', se): #from garage to next tri
                        return None

                    if not add(start, end, dep, arr, 'service trip', service_energy, tr['line']): #adding service
                        return None
                    continue

            #so now i wnat to check if it has enough battery without charging
            if not direct_ok:
                return None
            if rows and wait > tolerance and not add(position,position,ready, ready + wait,'idle', wait * idle_kw / 60): #addinh idle b4 next
                return None
            if dt > tolerance and not add(position, start, dep - dt, dep, 'material trip', de): #start of the next trip
                return None
            if not add(start, end, dep, arr, 'service trip', service_energy, tr['line']): #add service t
                return None

        back_t, back_e = route(routes, position, garage) #after all go back to the garage

        if not m.isfinite(back_t) or soc - back_e < min_state_charge - tolerance: #based on tge battery checking if it still can go back
            return None

        if back_t > tolerance and not add(position, garage, ready, ready + back_t, 'material trip', back_e): #adding the material trip to gargae
            return None
        return rows



    def score(rows):
        total_energy = sum(max(0, r[6]) for r in rows) #total energy used
        idle_time = sum(r[3] - r[2] for r in rows if r[4] == 'idle') #total time bus stopped
        return (total_energy, idle_time) 


    def construct(trips, routes, seed=0, randomized=False):
        random_generator = r.Random(seed) #random to try different solutiosn
        busses_trips = {}
        schedules = {}

        #first going through all service trips
        for trip_index in range(len((trips))):
            options = []
            #i want to try to add trip to an existing bus to reduce the number
            for bus, indices in busses_trips.items():
                new_schedule = rebuild_bus(indices + [trip_index], trips, routes, bus)
                if new_schedule is not None: #if possible 
                    old_energy, old_idle = score(schedules[bus])
                    new_energy, new_idle = score(new_schedule)
                    extra_energy = new_energy - old_energy
                    extra_idle = new_idle - old_idle
                    options.append((extra_energy, extra_idle, random_generator.random(), bus, new_schedule))

            if options:
                options.sort() #from best to worst
                if randomized and random_generator.random() < 0.2: #20% 
                    number_of_options = min(3, len(options))
                    selected = random_generator.randrange(number_of_options)
                    choice = options[selected] #chosing randomly between the 3 best ones
                else:
                    choice = options[0] #the best one

                bus = choice[3]
                schedule = choice[4]

            else:
                bus = len(busses_trips) + 1 #if no existing bus can do trip creating new one
                schedule = rebuild_bus([trip_index], trips, routes, bus)
                if schedule is None: #checking if bus can perform
                    raise ValueError(f'bus cannot complete trip {trip_index}')

            busses_trips.setdefault(bus, []).append(trip_index)
            schedules[bus] = schedule

        return busses_trips, schedules


    def to_dataframe(schedules):
        data = [row for bus_rows in schedules.values() for row in bus_rows] #put everything on one list
        if not data:
            return pd.DataFrame(columns=cols)
        return pd.DataFrame(data, columns=cols).sort_values(['bus', 'start time', 'end time'], kind='stable').reset_index(drop=True) #soorting by bus time

    def prepare_data(timetable, distance_matrix):
        routes = {} #for travel time and energy
        for _, row in distance_matrix.iterrows():
            if pd.isna(row['line']): #to get the line
                line= None
            else:
                line= int(row['line'])

            duration = (float(row['min_travel_time']) + float(row['max_travel_time'])) / 2 #average travel time
            energy = float(row['distance_m']) / 1000 * driving_kwh_km #energy needed
            routes[(row['start'], row['end'], line)] = (duration, energy)

        trips = [] #for service trips
        for _, row in timetable.iterrows():
            departure = float(row['departure_time_minutes'])
            # trips before 05:00 belong to the next day
            if departure < 300:
                departure += 1440
            
            line = int(row['line'])
            key = (row['start'], row['end'], line)

            if key not in routes:
                raise ValueError(f'missing service route {key}') #checking if exists 

            duration, energy = routes[key]

            trips.append({
                'start': row['start'],
                'end': row['end'],
                'line': line,
                'departure': departure,
                'arrival': departure + duration,
                'energy': energy
            })
        trips.sort(key=lambda t: t['departure'])
        return trips, routes


    def service_coverage_ok(schedule, trips): #if everything is on the schedule
        expected = Counter((t['start'], t['end'], t['line'], round(t['departure'], 5)) for t in trips)
        services = schedule[schedule['activity'] == 'service trip'] #only the services trips
        actual = Counter((r['start location'], r['end location'], int(r['line']), round(r['start time'], 5))for _, r in services.iterrows()) #counts them
        return actual == expected


    def feasibility_check_silent(schedule): #just without giving the messages
        with contextlib.redirect_stdout(io.StringIO()):
            return feasibility_check(schedule)


    def create_simple_schedule(starts=1, passes=2):
        trips, routes = prepare_data(clean_timet, clean_dis_m)
        best = None #for the big one
        for attempt in range(starts): #trying to make different solutions
            busses_trips, schedules = construct(trips, routes, seed=42+attempt, randomized=attempt > 0) #the first schedule
            #busses_trips, schedules = improve(busses_trips, schedules, trips, routes,seed=42+attempt, passes=passes) #we improve 
            result = to_dataframe(schedules) #toa table
            result['start_minutes'] = result['start time']
            result['end_minutes'] = result['end time']
            errors = feasibility_check_silent(result) #if feasible

            coverage = service_coverage_ok(result, trips)
            #print('feasible errors', len(errors))
            #print('coverage', coverage)
            #print('number of buses', result['bus'].nunique())
            if not errors.empty or not coverage:
                print(f'attempt {attempt+1} invalid result')
                continue

            quality = (result['bus'].nunique(), *score(result[cols].values.tolist())) #compare the solutions
            if best is None or quality < best[0]:
                best = (quality, result) #we save th best one
        if best is None:
            raise RuntimeError('No feasible solution found!!!')
        return best[1]

    simple_planning = create_simple_schedule(starts=1, passes=2)

    kpi_opener = 0

    st.write('Use button below to implement optimization algorithm and create your schedule.')
    opt_button = st.button('Optimize', help='Push this button to optimize your files and create new schedule.')
    if opt_button:
        kpi_opener = 1
        st.write('Number of buses in the new schedule: ', simple_planning['bus'].nunique())
        check = feasibility_check(simple_planning)
        st.write('Schedule was successuly optimized. You can now create a visualisation and download new schedule.')
        with st.expander('Push to see feasibility check report'):
            st.write("Changes made:")
            st.write(check)

    # take a variable in which you have saved your optimized schedule and make it a df to display with df.head()
    st.divider()

    #####################

    st.write('Graph of your new schedule:')

    #colours for each activity
    colors = {
        'service trip': '#527F8B',
        'material trip': '#B18AFF',
        'idle': '#D5D5D5',
        'charging': '#BDB17D'
    }


    fig, ax = plt.subplots()

    for bus, activities in simple_planning.groupby('bus'):
        for _, activity in activities.iterrows():
            start_time = activity['start time']
            end_time = activity['end time']
            duration = end_time - start_time
            activity_type = activity['activity']

            plt.barh(
                bus,
                duration,
                left=start_time,
                height=0.7,
                color=colors[activity_type],
                edgecolor='black',
                linewidth=0.01
            )

    buses = sorted(simple_planning['bus'].unique())
    bus_labels = []
    for bus in buses:
        bus_labels.append(f'Bus {int(bus)}')

    plt.yticks(buses,bus_labels,fontsize=7)

    ax = plt.gca()
    ax.xaxis.set_major_locator(MultipleLocator(60))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: f'{int(x // 60) % 24:02d}:{int(x % 60):02d}'))#prettier to convert minutes into hours

    plt.xlabel('Time', fontsize=10, labelpad=10)
    plt.ylabel('Bus number', fontsize=10, labelpad=1)
    plt.title('Bus Schedule',fontsize=20, fontweight='bold', pad=25)

    #legend for the activities
    legend = []
    for activity, color in colors.items():
        legend.append(Patch(facecolor=color, edgecolor='black', label=activity.title()))

    plt.legend(
        handles=legend,
        title='Activities',
        fontsize=5,
        title_fontsize=7,
        loc='upper left',
        bbox_to_anchor=(1.01, 1)
    )

    plt.xticks(fontsize=5.5)
    plt.grid(
        axis='x',
        linestyle='--',
        alpha=0.5,
        linewidth=0.6
    )
    for spine in ax.spines.values():
        spine.set_linewidth(1.0)


    display_graph_button = st.button('Show visual', help='Push this button to see visualisations of your new schedule.')
    if display_graph_button:
        kpi_opener = 1
        st.pyplot(fig, clear_figure=True) 
        st.toast('Done!')

    st.write('If you are satisfied with the result of optimization you can see below (short version). Feel free to download your new schedule using the button below.')

    st.write(simple_planning)

    # download_button = st.button('Download new schedule', use_container_width= True, help='Push this button to download an excel document with new planning onto your device.')

    # def download_this():
    #     temp = simple_planning.copy()
    #     temp = temp.to_excel('Optimized_Planning.xlsx')
    #     return temp

    # if download_button:
    #     kpi_opener = 1
    #     temp = download_this()
    #     st.toast('New Optimized_Planning.xlsx is downloaded!')

    document1 = simple_planning.copy()
    document1 = document1.to_excel('Optimized_Planning.xlsx')

    bio = io.BytesIO()
    document1.save(bio)
    user_manual = st.sidebar.download_button('new schedule', data=bio.getvalue(), file_name= 'newschedule.xlsx',use_container_width= True, help='Push this button to download user manual.')


# '''
# SIDEBAR
# '''

st.sidebar.write('Instruction on how to use this schedule optimazer can be downloaded using button below.')

document = docx.Document('User manual template 2.docx')

bio = io.BytesIO()
document.save(bio)
user_manual = st.sidebar.download_button('DOWNLOAD USER MANUAL', data=bio.getvalue(), file_name= 'USER_MANUAL.docx',use_container_width= True, help='Push this button to download user manual.')


st.sidebar.header('Key Performance Calculator')
st.sidebar.write('See here Key Performance Indicator (KPI) of your pre-optimazed files.')

if count ==3:
    st.sidebar.write('**Original KPI**')

    def calculate_bus_KPI(planningg):
        Original_Number_of_buses = len(set(clean_bus['bus']))
        distinct_busses = len(set(planningg['bus']))
        buss_K_P_I = 100 * distinct_busses / Original_Number_of_buses
        return buss_K_P_I

    def calculate_idle_KPI(planningg):
        total_idle = 0
        total_time = 0

        for i in planningg.index:
            start = planningg['start_minutes'].loc[i]
            end = planningg['end_minutes'].loc[i]
            activity = planningg['activity'].loc[i]
            duration = end - start
            if activity == "idle":
                total_idle += duration
            total_time += duration
        idle_K_P_I = 100 * total_idle / total_time
        return idle_K_P_I

    def calculate_energy_KPI(planningg):
        original_energy_consumption = clean_bus[clean_bus['energy consumption'] > 0]['energy consumption'].sum() #makes it positive  
        planning_charge = planningg[planningg['energy consumption'] > 0]
        total_energy_consumption = planning_charge['energy consumption'].sum() # i took the abs out bc now should be all positive
        energy_KPI = 100* total_energy_consumption / original_energy_consumption
        return energy_KPI

    bus_KPI = calculate_bus_KPI(clean_bus)
    idle_KPI = calculate_idle_KPI(clean_bus)
    energy_KPI = calculate_energy_KPI(clean_bus)


    st.sidebar.write("Bus KPI:", round(bus_KPI, 2), ', that is: ', clean_bus['bus'].nunique(), 'buses in total.')
    st.sidebar.write("Idle KPI:", round(idle_KPI, 2))
    st.sidebar.write("Energy KPI:", round(energy_KPI, 2))
    # put here function to calculate and display KPI
    st.sidebar.divider()

    #####################

    if kpi_opener:
        st.sidebar.write('**KPI after optimization**')

        bus_KPI_optimized = calculate_bus_KPI(simple_planning)
        idle_KPI_optimized = calculate_idle_KPI(simple_planning)
        energy_KPI_optimized = calculate_energy_KPI(simple_planning)

        st.sidebar.write('Bus KPI:', round(bus_KPI_optimized, 2), ', that is: ', simple_planning['bus'].nunique(), 'buses in total.')
        st.sidebar.write('Idle KPI:', round(idle_KPI_optimized, 2))
        st.sidebar.write('Energy KPI:', round(energy_KPI_optimized, 2))

        st.sidebar.divider()

        st.sidebar.write('**Difference of KPI**')

        bus_KPI_improvement = bus_KPI_optimized - bus_KPI 
        idle_KPI_improvement = idle_KPI_optimized - idle_KPI 
        energy_KPI_improvement = energy_KPI_optimized - energy_KPI

        st.sidebar.write('KPI for buses changed with ', round(bus_KPI_improvement, 2), ' percent point.')
        st.sidebar.write('KPI for idle time changed with ', round(idle_KPI_improvement, 2), ' percent point.')
        st.sidebar.write('KPI for energy changed with ', round(energy_KPI_improvement, 2), ' percent point.')


# ''' 
# COLOUR CONFIGURATIOIN 
# '''

page_header_clr = """
<style>
[data-testid='stHeader'] {
background-color: #26464F;
opacity: 1.0;
}
</style>
"""
st.markdown(page_header_clr,unsafe_allow_html=True)