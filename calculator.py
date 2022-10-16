import random
import os
import openpyxl


class FPCalculator:
    def __init__(self, lastfilename, Start):
        self.UFPdict = {
            "Inputs": [3, 4, 6],
            "Outputs": [4, 5, 7],
            "Inquiries": [3, 4, 6],
            "Files": [7, 10, 15],
            "ExternalFiles": [5, 7, 10],
        }
        self.table = {
            "Organic": {
                "a": 2.4,
                "b": 1.05,
                "c": 2.5,
                "d": 0.38,
            },
            "Semi-detached": {
                "a": 3.0,
                "b": 1.12,
                "c": 2.5,
                "d": 0.35,
            },
            "Embedded": {
                "a": 3.6,
                "b": 1.20,
                "c": 2.5,
                "d": 0.32,
            },
        }
        self.EAFtable = {
            "Organic": {
                "a": 3.2,
                "b": 1.05,
                "c": 2.5,
                "d": 0.38,
            },
            "Semi-detached": {
                "a": 3.0,
                "b": 1.12,
                "c": 2.5,
                "d": 0.35,
            },
            "Embedded": {
                "a": 2.8,
                "b": 1.20,
                "c": 2.5,
                "d": 0.32,
            },
        }
        self.driver = [
                    [0.75, 0.88, 1.0, 1.15, 1.4,1],
                    [1, 0.94, 1.0, 1.08, 1.16,1],
                    [0.7, 0.85, 1.0, 1.15, 1.3, 1.65],
                    [1,1,1.0, 1.11, 1.3, 1.66],
                    [1,1,1.0, 1.06, 1.21, 1.56],
                    [1,0.87, 1.0, 1.15, 1.30, 1],
                    [1,0.87, 1.0, 1.07, 1.15, 1],
                    [1.46, 1.19, 1.0, 0.86, 0.71, 1],
                    [1.29, 1.13, 1.0, 0.91, 0.82, 1],
                    [1.42, 1.17, 1.0, 0.86, 0.7, 1],
                    [1.21, 1.1, 1.0, 0.9, 1 , 1],
                    [1.14, 1.07, 1.0, 0.95, 1, 1],
                    [1.24, 1.1, 1.0, 0.91, 0.82,1],
                    [1.24, 1.1, 1.0, 0.91, 0.83,1],
                    [1.23, 1.08, 1.0, 1.04, 1.1,1]
]
        self.costDrivers = ['Required SW reliability', 'Size of application DB', 'Complexity of the product', 'Run-time performance constraints', 'Memory constraints', 'Volatility of the virtual machine environment', 'Required turnabout time', 'Analyst capability', 'Applications experience', 'Software engineer capability', 'Virtual machine experience', 'Programming language experience', 'Application of software engineering methods', 'Use of software tools', 'Required development schedule']
        self.KLOC = 0
        self.selected = None
        self.selectedI = None
        self.LOC = 0
        self.UFP = 0
        self.VAF = 0
        self.AFP = 0
        self.EAF = 0
        self.E = 0
        self.D = 0
        self.P = 0
        self.IE = 0
        self.ID = 0
        self.IP = 0
        self.lastSaveName = lastfilename
        self.CurrentXL = None
        self.CurrentCol = Start[0]
        self.CurrentRow = int(Start[1])

    def calculateUFP(self):
        a = 0
        types = ["Simple", "Average", "Complex"]
        for i in self.UFPdict:
            temp = 0
            print(f"{'-' * 25}{i}{'-' * 25}")
            for j in self.UFPdict[i]:
                val = input(f"{i} {types[a]} : ")
                if not val:
                    val = 0
                temp += j * int(val)
                a += 1
            a = 0
            self.UFP += temp

    def calculateVAF(self):
        print("\n")
        vals = [0] * 14
        valString = ['1. Does the system require reliable backup and recovery ?', '2. Is data communication required ?',
                     '3. Are there distributed processing functions ?', '4. Is performance critical ?',
                     '5. Will the system run in an existing heavily utilized operational environment ?',
                     '6. Does the system require on line data entry ?',
                     '7. Does the on line data entry require the input transaction to be built over multiple screens or',
                     '8. Are the master files updated on line ? operations ?',
                     '9. Is the inputs, outputs, files, or inquiries complex ?',
                     '10. Is the internal processing complex ?', '11. Is the code designed to be reusable ?',
                     '12. Are conversion and installation included in the design ?',
                     '13. Is the system designed for multiple installations in different organizations ?',
                     '14. Is the application designed to facilitate change and ease of use by the user ?']
        for i in range(len(vals)):
            val = input(f"{valString[i]} (0-5): ")
            if not val:
                val = 0
            vals[i] = int(val)
        self.VAF = 0.65 + (sum(vals) / 100)

    def calculateAFP(self):
        self.AFP = self.UFP * self.VAF

    def calculateCOCOMOS(self):
        self.LOC = (self.AFP * 128)
        self.KLOC = self.LOC / 1000
        if self.KLOC <= 50:
            self.selected = self.table["Organic"]
            self.selectedI = self.EAFtable["Organic"]
        elif 50 < self.KLOC <= 300:
            self.selected = self.table["Semi-detached"]
            self.selectedI = self.EAFtable["Semi-detached"]
        elif 300 < self.KLOC:
            self.selected = self.table["Embedded"]
            self.selectedI = self.EAFtable["Embedded"]
        self.E = self.selected["a"] * (pow(self.KLOC, self.selected["b"]))
        self.D = self.selected["c"] * (pow(self.E, self.selected["d"]))
        self.P = self.E / self.D

    def calculateEAF(self):
        temp = 1
        for i in range(len(self.driver)):
            inp = input(f"{self.costDrivers[i]} (0-5) : ")
            if not inp:
                continue
            inp = int(inp)
            if inp <= 0:
                temp *= self.driver[i][0]
            elif inp >= len(self.driver[i])-1:
                temp *= self.driver[i][-1]
            else:
                temp *= self.driver[i][inp]
        self.EAF = temp
        self.IE = (self.selectedI["a"] * (pow(self.KLOC, self.selectedI["b"]))) * self.EAF
        self.ID = self.selectedI["c"] * (pow(self.IE, self.selectedI["d"]))
        self.IP = self.IE / self.ID

    def calculateall(self):
        file = open("last.txt", "w")
        wb = openpyxl.load_workbook(self.lastSaveName)
        wa = wb.active
        self.calculateUFP()
        self.calculateVAF()
        self.calculateAFP()
        self.calculateCOCOMOS()
        self.calculateEAF()
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.UFP
        print(f"UFP : {self.UFP}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.VAF
        print(f"VAF : {self.VAF}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.AFP
        print(f"AFP : {self.AFP}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.AFP * 128
        print(f"LOC C : {self.AFP * 128}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.AFP * 30
        print(f"LOC OO : {self.AFP * 30}")
        print(f"KLOC : {self.KLOC}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.E
        print(f"Effort : {self.E} Effort/PM")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.D
        print(f"Duration : {self.D} Month/s")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.P
        print(f"People : {self.P} Person/s")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.EAF
        print(f"EAF : {self.EAF}")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.IE
        print(f"IE : {self.IE} EFFORT/PM")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.ID
        print(f"IDuration : {self.ID} Month/s")
        self.CurrentCol = chr(ord(self.CurrentCol) + 1)
        wa[f"{self.CurrentCol}{self.CurrentRow}"] = self.IP
        print(f"IPeople : {self.IP} Person/s")
        name = f"fp_changed{str(random.randrange(1, 1000))}.xlsx"
        file.write(name)
        os.system(f"del {self.lastSaveName}")
        print(f"New XL file Created : {name}")
        file.close()
        wb.save(name)

        
#change the second input in FPCalculator to move to different row

calc = FPCalculator("fp.xlsx", "D3")
calc.calculateall()
