#%%
from Imports import *
import PiMachineManager
import CsvWriter
import ColumnCreator

import DateAndTimeManager
from FilesReader import *

import EventLogging
import ProcessCsvManager

# piDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\PICompiled')
# os.chdir(piDirectory)

# dfPi = pd.read_csv(f'PICompiled2024-11-04.csv', encoding='latin1')

# previousTempDfPiRow = dfPi.iloc[[1 - 1], :]

# # previousTime = pd.to_datetime(previousTempDfPiRow["TIME"].values).time()
# previousDate = previousTempDfPiRow["TIME"].values[0]
# previousDate = datetime.strptime(previousDate, "%H:%M:%S")
# previousDate = previousDate + timedelta(seconds=1)
# previousDate = previousDate.strftime("%H:%M:%S")
# print(previousDate)


# %%
#Reading Date Today
DateAndTimeManager.GetDateToday()

#READING ALL FILES USING FILES READER
filesreader = filesReader()
filesreader.readingYearStored = DateAndTimeManager.yearNow
filesreader.ReadHPIQAQCFiles()
# %%
