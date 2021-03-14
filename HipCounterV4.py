import pyodbc
import PySimpleGUI as sg
#V4 - Plan-finish add for prelims to finals, expose track size
#V3 - added display window for results 
#V2 - added GUI, refined query using 'T' selector for track and generated printed output
# Create layout for PySimpleGUI window
layout = [[sg.T('Use this many numbers:'),sg.Input(1,key='-M1-',size=(1,1)),sg.T("If Distance >"), sg.Input(401, key='-CO-', size=(4,1)),sg.T('then use:'),sg.Input(2,key='-M2-',size=(1,1)),sg.T('Lanes:'),sg.Input(8,key='-LANES-',size=(1,1))],
          [sg.Text("Choose a folder: "), sg.Input(key="-IN2-" ,change_submits=True), sg.FileBrowse(key="-IN-")],
          [sg.Button("Submit")]]
# Create instance of window
window = sg.Window('Select Generic Database', layout)
# Define query to get Data from Access Generic Hytek Database
# Filter for only Track events and where seeded lane is not zero - meaning has a lane assigned
q1="SELECT Entries.Event_name, Entries.Event_code, Entries.Event_dist, Entries.Seeded_heat, Entries.Seeded_lane, Entries.Rnd_ltr, Entries.Event_gender \
FROM Entries \
WHERE Entries.Seeded_lane >0  AND Entries.Trk_Field = 'T' \
ORDER BY Entries.Seeded_lane"
# Query for Meet Info
q2="SELECT Meet_name, Meet_start, Meet_end  FROM Meet;"
#Initialize some working variables
debug = 'N'
output = ''
Meet_name , Meet_start, Meet_end = '','',''
basefont = 'Courier New'
basefsize = 9
def DisplayResults(results):
    layout = [[sg.Multiline(results,size=(60,30), key="new", font=(basefont,basefsize))]]
    window = sg.Window("Results Window", layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
    window.close()
    
def Get_MeetInfo(filename):
    global Meet_name, Meet_start, Meet_end
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+filename+';')
    cursor = conn.cursor()
    cursor.execute(q2)
    for row in cursor.fetchall():
        Meet_name, Meet_start, Meet_end = row
        
def Calc_hips(filename):
    global output
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+filename+';')
    cursor = conn.cursor() 
    cursor.execute(q1)
    hips = dict(H1=0)   
    Prelim_cnt = 0
    Prelim_Unique = []
    hips_total = 0
    for row in cursor.fetchall():
        # Get row data into variables
        Name, Code, Dist, Heat, Lane, Rnd, Gender = row
        if Dist > int(cutoff) or Dist == 1 or Dist == 2:
            multiple = int(Mult2)
        else:
            multiple = int(Mult1)
        # Detect Preliminary Rounds - Unque by Gender-Code-Distance
        if Rnd == 'P' :
            if Gender+Code+str(Dist)+str(multiple) not in Prelim_Unique:
                Prelim_Unique.append(Gender+Code+str(Dist)+str(multiple))
                Prelim_cnt += 1
        HLane = 'H'+str(Lane)
        if HLane in hips :
            #print (' Key found : ' + HLane)
            c_cnt = hips[HLane]
        else:
            #print (' Key not present ' + str(HLane) )
            hips[HLane] = 0
            c_cnt = 0
            
        c_cnt = hips[HLane]
        #print (' Current Value for' + HLane + ' is : ' + str(c_cnt))
        
        hips[HLane] = c_cnt + multiple
        hips_total = hips_total + multiple
    # Now do fix up for Preliminary found. Can only do for one prelimanry setup.
    for prace in Prelim_Unique:
        iLane = 1
        if debug == 'Y' :
            print(prace)
        while iLane <=  int(Lane_cnt):
            HLane = 'H'+str(iLane)
            c_cnt = hips[HLane]
            hips[HLane] = c_cnt + int(prace[-1])
            if debug == 'Y':
                print (str(iLane)+' '+HLane+' '+str(hips[HLane]))
            iLane += 1
            hips_total = hips_total + int(prace[-1])
    if debug == 'Y' :
        print ('Meet Name: ' + Meet_name +' Start Date: '+ str(Meet_start)[0:10])
        print ("Calculated HIP Number requirements based on seeding.")
        print ("Races in P/F format include estimate for finals.\nNote: 1 and 2 Mile use second number.")
        print ("There were "+str(Prelim_cnt) +" prelimnary races detected.")
        print ('_________________________________________________')
        print("\n".join("{}\t{}".format(k, v) for k, v in hips.items()))
        print ('_________________________________________________')
        print ('Total Hip count was '+ str(hips_total))
    
    output = 'Meet Name: ' + Meet_name +' Start Date: '+ str(Meet_start)[0:10] +'\n'
    output = output + "Calculated HIP Number requirements based on seeding. \n"
    output = output + "Races in P/F format include estimate for Finals.\nNote: 1 and 2 Mile use second number.\n"
    output = output + "There were "+str(Prelim_cnt) +" prelimnary races detected.\n"
    output = output + '___________________________________\n'
    output = output + "\n".join("{}\t{}".format(k, v) for k, v in hips.items())
    output = output + '\n___________________________________\n'
    output = output + 'Total Hip count was '+ str(hips_total) + '\n'
    
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Submit":
        fileselected = values["-IN-"]
        if fileselected == '':
            sg.popup("No file seleced!")
            continue
        cutoff = values["-CO-"]
        Mult1 = values["-M1-"]
        Mult2 = values["-M2-"]
        Lane_cnt = values["-LANES-"]
        Get_MeetInfo( fileselected )
        Calc_hips( fileselected)
        DisplayResults(output)
        
