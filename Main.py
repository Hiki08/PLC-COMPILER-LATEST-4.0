#%%
from Imports import *
import PiMachineManager
import CsvWriter
import ColumnCreator

import DateAndTimeManager
from FilesReader import *

import EventLogging
import ProcessCsvManager

#GUI Variables
root = ""

frame1 = ""
frame2 = ""

autoRun = False
autoRunButton = ""
calendarPicker = ""
compileButton = ""

on = ""
off = ""

time_picker = ""

def showGui():
    global root

    global frame1
    global frame2

    global autoRunButton
    global calendarPicker
    global compileButton

    global on
    global off

    global time_picker

    #Fixing Blur
    windll.shcore.SetProcessDpiAwareness(1)

    root = tk.Tk()
    root.title('FC1 Compiler')
    root.iconbitmap('Icons/HiblowLogo.ico')
    root.geometry('600x650+50+50')
    root.resizable(False, False)

    on = PhotoImage(file = "Icons/on.png")
    off = PhotoImage(file = "Icons/off.png")

    #Frames
    frame1 = tk.Frame(root)
    frame1.pack()

    frame2 = tk.Frame(root)
    frame2.pack_forget()

    # configure the grid
    frame1.columnconfigure(0, weight=1)
    frame1.columnconfigure(1, weight=1)

    #FRAME1

    # place a label on the root window
    message = tk.Label(frame1, text="FC1 Compiler", font=("Arial", 12, "bold"))
    message.grid(column=0, row=0, columnspan=2, padx=220)

    # button
    compileButton = tk.Button(frame1, text='COMPILE', font=("Arial", 12), command = StartProgram, width=15, height=1)
    compileButton.grid(column=0, row=1, ipadx=5, ipady=5, pady=10, columnspan=2)
    compileButton.config(bg="lightgreen", fg="black")

    autoRunLabel = tk.Label(frame1, text="Auto Run", font=("Arial", 12, "bold"))
    autoRunLabel.grid(column=0, row=2)

    autoRunButton = tk.Button(frame1, image = off, bd = 0, font=("Arial", 12), command = ToggleAutoRun)
    autoRunButton.grid(column=1, row=2, ipadx=5, ipady=5, pady=10)

    configureButton = tk.Button(frame1, text='CONFIGURE', font=("Arial", 8), command = Configure, width=10, height=1)
    configureButton.grid(column=0, row=3, ipadx=5, ipady=5, pady=10, columnspan=2)
    configureButton.config(bg="lightgreen", fg="black")

    calendarPicker = DateEntry(frame1, width=16, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
    calendarPicker.grid(column=0, row=4, pady=20, padx=(200, 0))

    #FRAME2

    message = tk.Label(frame2, text="Configure", font=("Arial", 12, "bold"))
    message.grid(column=0, row=1, columnspan=2, padx=220)

    # button
    backButton = tk.Button(frame2, text='BACK', font=("Arial", 8), command = Back, width=10, height=1)
    backButton.grid(column=0, row=0, ipadx=5, ipady=5, sticky=W)
    backButton.config(bg="lightgreen", fg="black")

    time_picker = AnalogPicker(frame2)
    time_picker.grid(column = 0, row = 4, columnspan = 2)
    theme = AnalogThemes(time_picker)
    theme.setNavyBlue()

    root.protocol("WM_DELETE_WINDOW", StopProgram)
    root.mainloop()

def Loading():
    global compileButton
    compileButton.config(text= "LOADING...")
    compileButton.config(state= "disabled")

def FinishedLoading():
    global compileButton
    compileButton.config(text= "COMPILE")
    compileButton.config(state= "normal")

def Configure():
    global frame1
    global frame2

    frame1.pack_forget()
    frame2.pack()

def Back():
    global frame1
    global frame2

    frame1.pack()
    frame2.pack_forget()

def AutoRun():
    global autoRun
    global time_picker

    coolDown = False

    while autoRun:
        print("Auto Run Activated")
        DateAndTimeManager.GetTimeNow()
        print(DateAndTimeManager.timeNow)
        
        hour = time_picker.hours()
        minutes = time_picker.minutes()
        period = time_picker.period()

        timeSet = f"{hour}:{minutes} {period}"
        timeSet = datetime2.strptime(timeSet, "%I:%M %p")
        timeSet = timeSet.strftime("%H:%M")

        print(timeSet)

        if DateAndTimeManager.timeNow == timeSet and not coolDown:
            coolDown = True
            StartProgram()
            time.sleep(70)
            coolDown = False
        time.sleep(1)

def toggleAutoRun():
    global autoRunButton
    global autoRun
    global compileButton
    global calendarPicker

    if not autoRun:
        autoRunButton.config(image = on)
        compileButton.config(state = "disabled")
        calendarPicker.config(state = "disabled")
        autoRun = True
        setDate()
        AutoRun()
    else:
        autoRunButton.config(image = off)
        compileButton.config(state = "normal")
        calendarPicker.config(state = "normal")
        autoRun = False
        setDate()

def ToggleAutoRun():
    threading.Thread(target=toggleAutoRun).start()

def setDate():
    global calendarPicker
    autoRun = False

    if autoRun:
        DateAndTimeManager.GetDateToday()
        DateAndTimeManager.dateToRead = DateAndTimeManager.dateToday
        DateAndTimeManager.dateToReadDashFormat = DateAndTimeManager.dateToRead.replace("/", "-")
        print(f"Date To Read: {DateAndTimeManager.dateToReadDashFormat}")
    else:
        selectedDate = calendarPicker.get_date()
        selectedDate = selectedDate.strftime("%Y/%m/%d")

        DateAndTimeManager.dateToRead = selectedDate
        DateAndTimeManager.dateToReadDashFormat = DateAndTimeManager.dateToRead.replace("/", "-")
        print(f"Date To Read: {DateAndTimeManager.dateToReadDashFormat}")

def StopProgram():
    global root
    global autoRun

    autoRun = False
    root.destroy()
        
def ErrorPopUp(error):
    showerror(title='Error Status', message=error)

def StartProgram():
    threading.Thread(target=start).start()

def start():
    #Setting Date From Calendar Picker
    setDate()

    #Loading GUI
    Loading()

    #Creating Empty Data Frame
    ColumnCreator.createEmptyColumn()
    PiMachineManager.compiledFrame = pd.DataFrame(columns=ColumnCreator.emptyColumn)

    #Reading Date Today
    DateAndTimeManager.GetDateToday()

    #READING ALL FILES USING FILES READER
    filesreader = filesReader()
    filesreader.readingYearStored = DateAndTimeManager.yearNow
    filesreader.ReadAllFiles()

    try:
        #Resetting Variables
        # PiMachineManager.ResetVariables()
        #Checking For Master Pump Data And Running Data
        PiMachineManager.CheckPICsv()

        #Checking If Master Pump Or Running Data Existed
        if PiMachineManager.canCompilePI:
            PiMachineManager.CompilePICsv()
            CsvWriter.WriteCsv(PiMachineManager.compiledFrame)
        else:
            ErrorPopUp("Error Creating MasterPump Or Running")

        EventLogging.logEvent("Creating MasterPump Or Running Successfully")
    except:
        ErrorPopUp("Error Creating MasterPump Or Running")
        EventLogging.logEvent("Creating MasterPump Or Running Failed")

    
    try:
        #Reading VT CSV Files
        isCsvReaded = False
        while not isCsvReaded:
            try:
                ProcessCsvManager.ResetVariables()
                ProcessCsvManager.ReadCsv()
                isCsvReaded = True
            except:
                print("Cannot Read Csv Retrying In 1 Seconds")
                isCsvReaded = False
                time.sleep(1)
    except:
        pass

    while ProcessCsvManager.programRunning:
        ProcessCsvManager.CsvOrganize()
        if ProcessCsvManager.canCompile:
            ProcessCsvManager.CompileCsv()
        
    CsvWriter.WriteCsv(PiMachineManager.compiledFrame)

    FinishedLoading()

showGui()

# %%
