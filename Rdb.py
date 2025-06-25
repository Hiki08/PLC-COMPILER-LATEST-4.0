#%%
from Imports import *
import DateAndTimeManager
from FilesReader import *

class rDB():
    rdbData = ""

    rdbModelCode = ""
    rdbLetterCode = ""
    rdbMonth = ""
    rdbLotNumber = ""
    rdbLotNumber3 = ""
    rdbYear = ""
    rdbYearFormat2 = ""

    rdbTeslaTotalAverage1 = ""
    rdbTeslaTotalAverage2 = ""
    rdbTeslaTotalAverage3 = ""
    rdbTeslaTotalAverage4 = ""

    rdbTeslaTotalMinimum1 = ""
    rdbTeslaTotalMinimum2 = ""
    rdbTeslaTotalMinimum3 = ""
    rdbTeslaTotalMinimum4 = ""

    rdbTeslaTotalMaximum1 = ""
    rdbTeslaTotalMaximum2 = ""
    rdbTeslaTotalMaximum3 = ""
    rdbTeslaTotalMaximum4 = ""

    rdbNoDataFound = ""

    readingYear = ""

    fileList = []
    fileFinishedReading = False

    rdbTotalAverage1 = []
    rdbTotalAverage2 = []
    rdbTotalAverage3 = []
    rdbTotalAverage4 = []
    rdbTotalAverage5 = []
    rdbTotalAverage6 = []
    rdbTotalAverage7 = []
    rdbTotalAverage8 = []
    rdbTotalAverage9 = []

    rdbTotalMinimum1 = []
    rdbTotalMinimum2 = []
    rdbTotalMinimum3 = []
    rdbTotalMinimum4 = []
    rdbTotalMinimum5 = []
    rdbTotalMinimum6 = []
    rdbTotalMinimum7 = []
    rdbTotalMinimum8 = []
    rdbTotalMinimum9 = []

    rdbTotalMaximum1 = []
    rdbTotalMaximum2 = []
    rdbTotalMaximum3 = []
    rdbTotalMaximum4 = []
    rdbTotalMaximum5 = []
    rdbTotalMaximum6 = []
    rdbTotalMaximum7 = []
    rdbTotalMaximum8 = []
    rdbTotalMaximum9 = []

    def __init__(self):
        pass

    def ReadCheckSheet(self, lotNumber, modelCode):
        RDB5200200Data = ""

        self.rdbModelCode = modelCode

        if self.rdbModelCode == "RDB5200200":
            self.rdbLetterCode = lotNumber[-1]

            #Removing Not Needed Values
            lotNumber = lotNumber.replace('-', '')
            lotNumber = lotNumber.replace(' ', '')

            #Removing The Last Two Values Of Lot Number
            lotNumber = lotNumber[:-1]

            self.rdbYear = lotNumber[:-4]
            self.rdbYearFormat2 = lotNumber[:-4]

            #Changing The Format Of Lot Number
            lotNumber = datetime2.strptime(lotNumber, "%Y%m%d")

            self.rdbMonth = lotNumber.strftime("%B")
            self.rdbLotNumber = lotNumber.strftime("%d/%m/%Y")

            pd.set_option('display.max_columns', None)
            pd.set_option('display.max_rows', None)

            targetValue = self.rdbMonth

            for file in RDB5200200CheckSheet:
                try:
                    # workbook = CalamineWorkbook.from_path(fileName)
                    for s in file.sheet_names:
                        if targetValue.lower() in s.lower():
                            RDB5200200Data = file.get_sheet_by_name(s).to_python(skip_empty_area=True)
                            RDB5200200Data = pd.DataFrame(RDB5200200Data)
                            RDB5200200Data = RDB5200200Data.replace(r'\s+', '', regex=True)
                            break

                    #Getting The RDB Code
                    rdbCode = RDB5200200Data.iloc[4, 9]
                    rdbCode = rdbCode[:-3]
                    rdbCode = rdbCode[27:]

                    #__________________________________
                    #Getting The RDB Code 2
                    rdbCode2 = RDB5200200Data.iloc[4, 9]
                    rdbCode2 = rdbCode2[27:]
                    #__________________________________
                    
                    #Getting The Lot Number (Supplier)
                    rdbLotNumber2 = RDB5200200Data.iloc[max(0 + 7, 0):min(len(RDB5200200Data), 0 + 999), RDB5200200Data.columns.get_loc(0):RDB5200200Data.columns.get_loc(0) + 11]
                    rdbLotNumber2 = rdbLotNumber2[(rdbLotNumber2[0].isin([self.rdbLotNumber])) & (rdbLotNumber2[8].isin([self.rdbLetterCode]))]
                    rdbProdDate = rdbLotNumber2[10].values[0]

                    rdbProdDate = rdbProdDate[:-1]

                    rdbLotNumber2 = rdbLotNumber2[9].values[0]

                    #Getting The Row, Column Location Of Lot Number
                    findLotNumber = [(index, column) for index, row in RDB5200200Data.iterrows() for column, value in row.items() if str(value) == str(self.rdbLotNumber)]
                    lotNumberRow = [index for index, _ in findLotNumber]
                    lotNumberColumn = [column for _, column in findLotNumber]

                    print("Row indices:", lotNumberRow)
                    print("Column names:", lotNumberColumn)

                    #Getting The Tesla Table
                    inspectionData = RDB5200200Data.iloc[max(0, lotNumberRow[0]):min(len(RDB5200200Data), lotNumberRow[0] + 7), RDB5200200Data.columns.get_loc(lotNumberColumn[0] + 21):RDB5200200Data.columns.get_loc(lotNumberColumn[0]) + 26]

                    self.rdbTeslaTotalAverage1 = inspectionData.iloc[0].mean()
                    self.rdbTeslaTotalAverage2 = inspectionData.iloc[2].mean()
                    self.rdbTeslaTotalAverage3 = inspectionData.iloc[4].mean()
                    self.rdbTeslaTotalAverage4 = inspectionData.iloc[6].mean()

                    self.rdbTeslaTotalMinimum1 = inspectionData.iloc[0].min()
                    self.rdbTeslaTotalMinimum2 = inspectionData.iloc[2].min()
                    self.rdbTeslaTotalMinimum3 = inspectionData.iloc[4].min()
                    self.rdbTeslaTotalMinimum4 = inspectionData.iloc[6].min()

                    self.rdbTeslaTotalMaximum1 = inspectionData.iloc[0].max()
                    self.rdbTeslaTotalMaximum2 = inspectionData.iloc[2].max()
                    self.rdbTeslaTotalMaximum3 = inspectionData.iloc[4].max()
                    self.rdbTeslaTotalMaximum4 = inspectionData.iloc[6].max()

                    print(f"Tesla Average 1:{self.rdbTeslaTotalAverage1}")
                    print(f"Tesla Average 2:{self.rdbTeslaTotalAverage2}")
                    print(f"Tesla Average 3:{self.rdbTeslaTotalAverage3}")
                    print(f"Tesla Average 4:{self.rdbTeslaTotalAverage4}")

                    print(f"Tesla Minimum 1:{self.rdbTeslaTotalMinimum1}")
                    print(f"Tesla Minimum 2:{self.rdbTeslaTotalMinimum2}")
                    print(f"Tesla Minimum 3:{self.rdbTeslaTotalMinimum3}")
                    print(f"Tesla Minimum 4:{self.rdbTeslaTotalMinimum4}")

                    print(f"Tesla Maximum 1:{self.rdbTeslaTotalMaximum1}")
                    print(f"Tesla Maximum 2:{self.rdbTeslaTotalMaximum2}")
                    print(f"Tesla Maximum 3:{self.rdbTeslaTotalMaximum3}")
                    print(f"Tesla Maximum 4:{self.rdbTeslaTotalMaximum4}")

                    break
                except:
                    pass

            #Checking Each Files In Files;
            for file in HPIQAQCData:
                try:
                    file['DATE RECEIVED'] = file['DATE RECEIVED'].astype(str).str.replace("-", "")
                    file = file[(file["DATE RECEIVED"].isin([str(rdbProdDate)])) & (file["ITEM CODE"].isin([str(rdbCode2)]))]
                    file = file[file['LOT NUMBER'].str.contains(rdbLotNumber2[:-3], na=False)]

                    self.rdbLotNumber3 = file["LOT NUMBER"].values[0]
                    self.rdbNoDataFound = False
                    break
                except:
                    pass

        else:
            self.rdbLotNumber = lotNumber

    def GettingData(self, itemCode):
        if itemCode == "RDB5200200":
            self.fileList = RD05200200Data
        elif itemCode == "RDB4200801":
            self.fileList = RDB4200801Data

        if not self.rdbNoDataFound:
            for fileNum in range(len(self.fileList)):
                self.rdbTotalAverage1 = []
                self.rdbTotalAverage2 = []
                self.rdbTotalAverage3 = []
                self.rdbTotalAverage4 = []
                self.rdbTotalAverage5 = []
                self.rdbTotalAverage6 = []
                self.rdbTotalAverage7 = []
                self.rdbTotalAverage8 = []
                self.rdbTotalAverage9 = []

                self.rdbTotalMinimum1 = []
                self.rdbTotalMinimum2 = []
                self.rdbTotalMinimum3 = []
                self.rdbTotalMinimum4 = []
                self.rdbTotalMinimum5 = []
                self.rdbTotalMinimum6 = []
                self.rdbTotalMinimum7 = []
                self.rdbTotalMinimum8 = []
                self.rdbTotalMinimum9 = []

                self.rdbTotalMaximum1 = []
                self.rdbTotalMaximum2 = []
                self.rdbTotalMaximum3 = []
                self.rdbTotalMaximum4 = []
                self.rdbTotalMaximum5 = []
                self.rdbTotalMaximum6 = []
                self.rdbTotalMaximum7 = []
                self.rdbTotalMaximum8 = []
                self.rdbTotalMaximum9 = []
                
                #Getting The Row, Column Location Of HIBLOW
                findHiblow = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
                hiblowRow = [index for index, _ in findHiblow]
                hiblowColumn = [column for _, column in findHiblow]

                print("Row indices:", hiblowRow)
                print("Column names:", hiblowColumn)

                # Get the Neighboring Data Of Hiblow
                hiblowFiltered = self.fileList[fileNum].iloc[max(0, hiblowRow[0] - 3):min(len(self.fileList[fileNum]), hiblowRow[0] + 10), self.fileList[fileNum].columns.get_loc(hiblowColumn[0]):self.fileList[fileNum].columns.get_loc(hiblowColumn[0]) + 999999]

                try:
                    if self.rdbModelCode == "RDB5200200":
                        #Getting The Row, Column Location Of Lot Number
                        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == self.rdbLotNumber3]
                        lotNumberRow = [index for index, _ in findLotNumber]
                        lotNumberColumn = [column for _, column in findLotNumber]
                    elif self.rdbModelCode == "RDB4200801":
                        #Getting The Row, Column Location Of Lot Number
                        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == self.rdbLotNumber]
                        lotNumberRow = [index for index, _ in findLotNumber]
                        lotNumberColumn = [column for _, column in findLotNumber]

                    print("Row indices:", lotNumberRow)
                    print("Column names:", lotNumberColumn)

                    for a in range(0, len(lotNumberColumn)):
                        # Get The Neighboring Data of Lot Number
                        inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 12), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                        if self.rdbModelCode == "RDB5200200":

                            average1 = inspectionData.iloc[3].mean()
                            average2 = inspectionData.iloc[4].mean()
                            average3 = inspectionData.iloc[5].mean()
                            average4 = inspectionData.iloc[6].mean()
                            average5 = inspectionData.iloc[7].mean()
                            average6 = inspectionData.iloc[8].mean()
                            average8 = inspectionData.iloc[10].mean()

                            minimum1 = inspectionData.iloc[3].min()
                            minimum2 = inspectionData.iloc[4].min()
                            minimum3 = inspectionData.iloc[5].min()
                            minimum4 = inspectionData.iloc[6].min()
                            minimum5 = inspectionData.iloc[7].min()
                            minimum6 = inspectionData.iloc[8].min()
                            minimum8 = inspectionData.iloc[10].min()

                            maximum1 = inspectionData.iloc[3].max()
                            maximum2 = inspectionData.iloc[4].max()
                            maximum3 = inspectionData.iloc[5].max()
                            maximum4 = inspectionData.iloc[6].max()
                            maximum5 = inspectionData.iloc[7].max()
                            maximum6 = inspectionData.iloc[8].max()
                            maximum8 = inspectionData.iloc[10].max()

                            self.rdbTotalAverage1.append(average1)
                            self.rdbTotalAverage2.append(average2)
                            self.rdbTotalAverage3.append(average3)
                            self.rdbTotalAverage4.append(average4)
                            self.rdbTotalAverage5.append(average5)
                            self.rdbTotalAverage6.append(average6)
                            self.rdbTotalAverage8.append(average8)

                            self.rdbTotalMinimum1.append(minimum1)
                            self.rdbTotalMinimum2.append(minimum2)
                            self.rdbTotalMinimum3.append(minimum3)
                            self.rdbTotalMinimum4.append(minimum4)
                            self.rdbTotalMinimum5.append(minimum5)
                            self.rdbTotalMinimum6.append(minimum6)
                            self.rdbTotalMinimum8.append(minimum8)

                            self.rdbTotalMaximum1.append(maximum1)
                            self.rdbTotalMaximum2.append(maximum2)
                            self.rdbTotalMaximum3.append(maximum3)
                            self.rdbTotalMaximum4.append(maximum4)
                            self.rdbTotalMaximum5.append(maximum5)
                            self.rdbTotalMaximum6.append(maximum6)
                            self.rdbTotalMaximum8.append(maximum8)

                        elif self.rdbModelCode == "RDB4200801":
                            average2 = inspectionData.iloc[4].mean()
                            average4 = inspectionData.iloc[6].mean()
                            average5 = inspectionData.iloc[7].mean()
                            average6 = inspectionData.iloc[8].mean()
                            average7 = inspectionData.iloc[9].mean()
                            average8 = inspectionData.iloc[10].mean()
                            average9 = inspectionData.iloc[11].mean()

                            minimum2 = inspectionData.iloc[4].min()
                            minimum4 = inspectionData.iloc[6].min()
                            minimum5 = inspectionData.iloc[7].min()
                            minimum6 = inspectionData.iloc[8].min()
                            minimum7 = inspectionData.iloc[9].min()
                            minimum8 = inspectionData.iloc[10].min()
                            minimum9 = inspectionData.iloc[11].min()

                            maximum2 = inspectionData.iloc[4].max()
                            maximum4 = inspectionData.iloc[6].max()
                            maximum5 = inspectionData.iloc[7].max()
                            maximum6 = inspectionData.iloc[8].max()
                            maximum7 = inspectionData.iloc[9].max()
                            maximum8 = inspectionData.iloc[10].max()
                            maximum9 = inspectionData.iloc[11].max()

                            self.rdbTotalAverage2.append(average2)
                            self.rdbTotalAverage4.append(average4)
                            self.rdbTotalAverage5.append(average5)
                            self.rdbTotalAverage6.append(average6)
                            self.rdbTotalAverage7.append(average7)
                            self.rdbTotalAverage8.append(average8)
                            self.rdbTotalAverage9.append(average9)

                            self.rdbTotalMinimum2.append(minimum2)
                            self.rdbTotalMinimum4.append(minimum4)
                            self.rdbTotalMinimum5.append(minimum5)
                            self.rdbTotalMinimum6.append(minimum6)
                            self.rdbTotalMinimum7.append(minimum7)
                            self.rdbTotalMinimum8.append(minimum8)
                            self.rdbTotalMinimum9.append(minimum9)

                            self.rdbTotalMaximum2.append(maximum2)
                            self.rdbTotalMaximum4.append(maximum4)
                            self.rdbTotalMaximum5.append(maximum5)
                            self.rdbTotalMaximum6.append(maximum6)
                            self.rdbTotalMaximum7.append(maximum7)
                            self.rdbTotalMaximum8.append(maximum8)
                            self.rdbTotalMaximum9.append(maximum9)

                    if self.rdbModelCode == "RDB5200200":

                        self.rdbTotalAverage1 = statistics.mean(self.rdbTotalAverage1)
                        self.rdbTotalAverage2 = statistics.mean(self.rdbTotalAverage2)
                        self.rdbTotalAverage3 = statistics.mean(self.rdbTotalAverage3)
                        self.rdbTotalAverage4 = statistics.mean(self.rdbTotalAverage4)
                        self.rdbTotalAverage5 = statistics.mean(self.rdbTotalAverage5)
                        self.rdbTotalAverage6 = statistics.mean(self.rdbTotalAverage6)
                        self.rdbTotalAverage8 = statistics.mean(self.rdbTotalAverage8)

                        self.rdbTotalMinimum1 = min(self.rdbTotalMinimum1)
                        self.rdbTotalMinimum2 = min(self.rdbTotalMinimum2)
                        self.rdbTotalMinimum3 = min(self.rdbTotalMinimum3)
                        self.rdbTotalMinimum4 = min(self.rdbTotalMinimum4)
                        self.rdbTotalMinimum5 = min(self.rdbTotalMinimum5)
                        self.rdbTotalMinimum6 = min(self.rdbTotalMinimum6)
                        self.rdbTotalMinimum8 = min(self.rdbTotalMinimum8)

                        self.rdbTotalMaximum1 = max(self.rdbTotalMaximum1)
                        self.rdbTotalMaximum2 = max(self.rdbTotalMaximum2)
                        self.rdbTotalMaximum3 = max(self.rdbTotalMaximum3)
                        self.rdbTotalMaximum4 = max(self.rdbTotalMaximum4)
                        self.rdbTotalMaximum5 = max(self.rdbTotalMaximum5)
                        self.rdbTotalMaximum6 = max(self.rdbTotalMaximum6)
                        self.rdbTotalMaximum8 = max(self.rdbTotalMaximum8)

                        self.rdbTotalAverage1 = f"{self.rdbTotalAverage1:.2f}"
                        self.rdbTotalAverage2 = f"{self.rdbTotalAverage2:.2f}"
                        self.rdbTotalAverage3 = f"{self.rdbTotalAverage3:.2f}"
                        self.rdbTotalAverage4 = f"{self.rdbTotalAverage4:.2f}"
                        self.rdbTotalAverage5 = f"{self.rdbTotalAverage5:.2f}"
                        self.rdbTotalAverage6 = f"{self.rdbTotalAverage6:.2f}"
                        self.rdbTotalAverage8 = f"{self.rdbTotalAverage8:.2f}"

                        self.rdbTotalMinimum1 = f"{self.rdbTotalMinimum1:.2f}"
                        self.rdbTotalMinimum2 = f"{self.rdbTotalMinimum2:.2f}"
                        self.rdbTotalMinimum3 = f"{self.rdbTotalMinimum3:.2f}"
                        self.rdbTotalMinimum4 = f"{self.rdbTotalMinimum4:.2f}"
                        self.rdbTotalMinimum5 = f"{self.rdbTotalMinimum5:.2f}"
                        self.rdbTotalMinimum6 = f"{self.rdbTotalMinimum6:.2f}"
                        self.rdbTotalMinimum8 = f"{self.rdbTotalMinimum8:.2f}"

                        self.rdbTotalMaximum1 = f"{self.rdbTotalMaximum1:.2f}"
                        self.rdbTotalMaximum2 = f"{self.rdbTotalMaximum2:.2f}"
                        self.rdbTotalMaximum3 = f"{self.rdbTotalMaximum3:.2f}"
                        self.rdbTotalMaximum4 = f"{self.rdbTotalMaximum4:.2f}"
                        self.rdbTotalMaximum5 = f"{self.rdbTotalMaximum5:.2f}"
                        self.rdbTotalMaximum6 = f"{self.rdbTotalMaximum6:.2f}"
                        self.rdbTotalMaximum8 = f"{self.rdbTotalMaximum8:.2f}"

                        self.rdbTotalAverage7 = "None"
                        self.rdbTotalMinimum7 = "None"
                        self.rdbTotalMaximum7 = "None"

                        self.rdbTotalAverage9 = "None"
                        self.rdbTotalMinimum9 = "None"
                        self.rdbTotalMaximum9 = "None"

                        break

                    elif self.rdbModelCode == "RDB4200801":

                        self.rdbTotalAverage2 = statistics.mean(self.rdbTotalAverage2)
                        self.rdbTotalAverage4 = statistics.mean(self.rdbTotalAverage4)
                        self.rdbTotalAverage5 = statistics.mean(self.rdbTotalAverage5)
                        self.rdbTotalAverage6 = statistics.mean(self.rdbTotalAverage6)
                        self.rdbTotalAverage7 = statistics.mean(self.rdbTotalAverage7)
                        self.rdbTotalAverage8 = statistics.mean(self.rdbTotalAverage8)
                        self.rdbTotalAverage9 = statistics.mean(self.rdbTotalAverage9)

                        self.rdbTotalMinimum2 = min(self.rdbTotalMinimum2)
                        self.rdbTotalMinimum4 = min(self.rdbTotalMinimum4)
                        self.rdbTotalMinimum5 = min(self.rdbTotalMinimum5)
                        self.rdbTotalMinimum6 = min(self.rdbTotalMinimum6)
                        self.rdbTotalMinimum7 = min(self.rdbTotalMinimum7)
                        self.rdbTotalMinimum8 = min(self.rdbTotalMinimum8)
                        self.rdbTotalMinimum9 = min(self.rdbTotalMinimum9)

                        self.rdbTotalMaximum2 = max(self.rdbTotalMaximum2)
                        self.rdbTotalMaximum4 = max(self.rdbTotalMaximum4)
                        self.rdbTotalMaximum5 = max(self.rdbTotalMaximum5)
                        self.rdbTotalMaximum6 = max(self.rdbTotalMaximum6)
                        self.rdbTotalMaximum7 = max(self.rdbTotalMaximum7)
                        self.rdbTotalMaximum8 = max(self.rdbTotalMaximum8)
                        self.rdbTotalMaximum9 = max(self.rdbTotalMaximum9)

                        self.rdbTotalAverage2 = f"{self.rdbTotalAverage2:.2f}"
                        self.rdbTotalAverage4 = f"{self.rdbTotalAverage4:.2f}"
                        self.rdbTotalAverage5 = f"{self.rdbTotalAverage5:.2f}"
                        self.rdbTotalAverage6 = f"{self.rdbTotalAverage6:.2f}"
                        self.rdbTotalAverage7 = f"{self.rdbTotalAverage7:.2f}"
                        self.rdbTotalAverage8 = f"{self.rdbTotalAverage8:.2f}"
                        self.rdbTotalAverage9 = f"{self.rdbTotalAverage9:.2f}"

                        self.rdbTotalMinimum2 = f"{self.rdbTotalMinimum2:.2f}"
                        self.rdbTotalMinimum4 = f"{self.rdbTotalMinimum4:.2f}"
                        self.rdbTotalMinimum5 = f"{self.rdbTotalMinimum5:.2f}"
                        self.rdbTotalMinimum6 = f"{self.rdbTotalMinimum6:.2f}"
                        self.rdbTotalMinimum7 = f"{self.rdbTotalMinimum7:.2f}"
                        self.rdbTotalMinimum8 = f"{self.rdbTotalMinimum8:.2f}"
                        self.rdbTotalMinimum9 = f"{self.rdbTotalMinimum9:.2f}"

                        self.rdbTotalMaximum2 = f"{self.rdbTotalMaximum2:.2f}"
                        self.rdbTotalMaximum4 = f"{self.rdbTotalMaximum4:.2f}"
                        self.rdbTotalMaximum5 = f"{self.rdbTotalMaximum5:.2f}"
                        self.rdbTotalMaximum6 = f"{self.rdbTotalMaximum6:.2f}"
                        self.rdbTotalMaximum7 = f"{self.rdbTotalMaximum7:.2f}"
                        self.rdbTotalMaximum8 = f"{self.rdbTotalMaximum8:.2f}"
                        self.rdbTotalMaximum9 = f"{self.rdbTotalMaximum9:.2f}"

                        self.rdbTotalAverage1 = "None"
                        self.rdbTotalMinimum1 = "None"
                        self.rdbTotalMaximum1 = "None"

                        self.rdbTotalAverage3 = "None"
                        self.rdbTotalMinimum3 = "None"
                        self.rdbTotalMaximum3 = "None"

                        break

                except Exception as e:
                    print(f"Error Cannot Get RDB Data {e}")

                print(f"RDB Total Average 1: {self.rdbTotalAverage1}")
                print(f"RDB Total Average 2: {self.rdbTotalAverage2}")
                print(f"RDB Total Average 3: {self.rdbTotalAverage3}")
                print(f"RDB Total Average 4: {self.rdbTotalAverage4}")
                print(f"RDB Total Average 5: {self.rdbTotalAverage5}")
                print(f"RDB Total Average 6: {self.rdbTotalAverage6}")
                print(f"RDB Total Average 7: {self.rdbTotalAverage7}")
                print(f"RDB Total Average 8: {self.rdbTotalAverage8}")
                print(f"RDB Total Average 9: {self.rdbTotalAverage9}")

                print(f"RDB Total Minimum 1: {self.rdbTotalMinimum1}")
                print(f"RDB Total Minimum 2: {self.rdbTotalMinimum2}")
                print(f"RDB Total Minimum 3: {self.rdbTotalMinimum3}")
                print(f"RDB Total Minimum 4: {self.rdbTotalMinimum4}")
                print(f"RDB Total Minimum 5: {self.rdbTotalMinimum5}")
                print(f"RDB Total Minimum 6: {self.rdbTotalMinimum6}")
                print(f"RDB Total Minimum 7: {self.rdbTotalMinimum7}")
                print(f"RDB Total Minimum 8: {self.rdbTotalMinimum8}")
                print(f"RDB Total Minimum 9: {self.rdbTotalMinimum9}")

                print(f"RDB Total Maximum 1: {self.rdbTotalMaximum1}")
                print(f"RDB Total Maximum 2: {self.rdbTotalMaximum2}")
                print(f"RDB Total Maximum 3: {self.rdbTotalMaximum3}")
                print(f"RDB Total Maximum 4: {self.rdbTotalMaximum4}")
                print(f"RDB Total Maximum 5: {self.rdbTotalMaximum5}")
                print(f"RDB Total Maximum 6: {self.rdbTotalMaximum6}")
                print(f"RDB Total Maximum 7: {self.rdbTotalMaximum7}")
                print(f"RDB Total Maximum 8: {self.rdbTotalMaximum8}")
                print(f"RDB Total Maximum 9: {self.rdbTotalMaximum9}")

            print(f"Selected RDB Total Average 1: {self.rdbTotalAverage1}")
            print(f"Selected RDB Total Average 2: {self.rdbTotalAverage2}")
            print(f"Selected RDB Total Average 3: {self.rdbTotalAverage3}")
            print(f"Selected RDB Total Average 4: {self.rdbTotalAverage4}")
            print(f"Selected RDB Total Average 5: {self.rdbTotalAverage5}")
            print(f"Selected RDB Total Average 6: {self.rdbTotalAverage6}")
            print(f"Selected RDB Total Average 7: {self.rdbTotalAverage7}")
            print(f"Selected RDB Total Average 8: {self.rdbTotalAverage8}")
            print(f"Selected RDB Total Average 9: {self.rdbTotalAverage9}")

            print(f"Selected RDB Total Minimum 1: {self.rdbTotalMinimum1}")
            print(f"Selected RDB Total Minimum 2: {self.rdbTotalMinimum2}")
            print(f"Selected RDB Total Minimum 3: {self.rdbTotalMinimum3}")
            print(f"Selected RDB Total Minimum 4: {self.rdbTotalMinimum4}")
            print(f"Selected RDB Total Minimum 5: {self.rdbTotalMinimum5}")
            print(f"Selected RDB Total Minimum 6: {self.rdbTotalMinimum6}")
            print(f"Selected RDB Total Minimum 7: {self.rdbTotalMinimum7}")
            print(f"Selected RDB Total Minimum 8: {self.rdbTotalMinimum8}")
            print(f"Selected RDB Total Minimum 9: {self.rdbTotalMinimum9}")

            print(f"Selected RDB Total Maximum 1: {self.rdbTotalMaximum1}")
            print(f"Selected RDB Total Maximum 2: {self.rdbTotalMaximum2}")
            print(f"Selected RDB Total Maximum 3: {self.rdbTotalMaximum3}")
            print(f"Selected RDB Total Maximum 4: {self.rdbTotalMaximum4}")
            print(f"Selected RDB Total Maximum 5: {self.rdbTotalMaximum5}")
            print(f"Selected RDB Total Maximum 6: {self.rdbTotalMaximum6}")
            print(f"Selected RDB Total Maximum 7: {self.rdbTotalMaximum7}")
            print(f"Selected RDB Total Maximum 8: {self.rdbTotalMaximum8}")
            print(f"Selected RDB Total Maximum 9: {self.rdbTotalMaximum9}")
            
#%%
#READING ALL FILES USING FILES READER
# DateAndTimeManager.GetDateToday()

# filesreader = filesReader()
# filesreader.readingYearStored = DateAndTimeManager.yearNow
# filesreader.ReadAllFiles()

# rdb = rDB()
# rdb.readingYear = int(DateAndTimeManager.yearNow)
# rdb.ReadCheckSheet("20241018-F", "RDB5200200")
# rdb.GettingData("RDB5200200")



# rdb.ReadCheckSheet("3P00015758-3", "RDB4200801")
# rdb.ReadRDB5200200()
# # rdb.fileList[0]
# # print(len(rdb.fileList))

# %%
