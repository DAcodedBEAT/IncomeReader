import os
import os.path
import re
import xlrd


class Weekly():
    def __init__(self):
        self.pledge_list = []
        self.special_list = []
        self.plate = 0
        self.sunday_school = 0
        self.total_pledge = 0
        self.total_special = 0

        self.num_pledge = 0
        self.num_special = 0
        self.num_sunday_school = 0
        self.num_plate = 0

        self.worksheet_data = None

    def parse(self, weekly_file):
        if os.path.isfile(weekly_file):
            workbook = xlrd.open_workbook(weekly_file)
            worksheet = workbook.sheet_by_index(0)
            num_rows = worksheet.nrows
            num_cols = worksheet.ncols
            # retrieve worksheet data
            wrksht = [[worksheet.cell_value(row, col) for col in range(num_cols)] for row in range(num_rows)]
            self.worksheet_data = wrksht

            # find and parse sunday school data
            sunday_school_row = -1
            test_cell = wrksht[31][0]
            if isinstance(test_cell, str) and ('Sunday School:' in test_cell):
                sunday_school_row = 31
            else:
                for r in range(num_rows-1, 0, -1):
                    test_cell = wrksht[r][0]
                    if isinstance(test_cell, str) and ('Sunday School:' in test_cell):
                        sunday_school_row = r

            if sunday_school_row != -1:
                test_cell = wrksht[sunday_school_row][2]
                if isinstance(test_cell, int) or isinstance(test_cell, float):
                    self.sunday_school = test_cell
                else:
                    for col in range(num_cols-1, 0, -1):
                        test_cell = wrksht[sunday_school_row][col]
                        if test_cell == '':
                            continue
                        elif isinstance(test_cell, int) or isinstance(test_cell, float):
                            self.sunday_school = test_cell
                        else:   # 'Sunday School:' not in wrksht[sunday_school_row][col]
                            self.sunday_school = 0
            if self.sunday_school != 0:
                self.num_sunday_school = 1

            # find and parse plate data
            plate_row = -1
            test_cell = wrksht[30][0]
            if isinstance(test_cell, str) and ('Plate:' in test_cell):
                plate_row = 30
            elif isinstance(wrksht[sunday_school_row-1][0], str) and ('Plate:' in wrksht[sunday_school_row-1][0]):
                plate_row = sunday_school_row-1
            else:
                for r in range(num_rows-1, 0, -1):
                    if isinstance(wrksht[r][0], str) and ('Plate:' in wrksht[r][0]):
                        plate_row = r

            if plate_row != -1:
                test_cell = wrksht[plate_row][2]
                if isinstance(test_cell, int) or isinstance(test_cell, float):
                    self.plate = test_cell
                else:
                    for col in range(num_cols-1, 0, -1):
                        test_cell = wrksht[plate_row][col]
                        if test_cell == '':
                            continue
                        elif isinstance(test_cell, int) or isinstance(test_cell, float):
                            self.plate = test_cell
                        else:   # 'Plate:' not in wrksht[plate_row][col]
                            self.plate = 0
            if self.plate != 0:
                self.num_plate = 1

            # get the start of useful income data
            start_row = -1
            if "A/C" in wrksht[3][0]:
                start_row = 4
            else:
                for row in range(num_rows):
                    if "A/C" in wrksht[row][0]:
                        start_row = row + 1

            if start_row != -1:     # if start row is found, continue
                # add pledge data to pledge_list
                for pledge_row in range(start_row, plate_row):
                    a = wrksht[pledge_row][0]
                    b = wrksht[pledge_row][1]
                    c = wrksht[pledge_row][2]
                    if (re.sub(r'\s+', '', str(a)) == '') and \
                            (re.sub(r'\s+', '', str(b)) == '') and \
                            (re.sub(r'\s+', '', str(c)) == ''):
                        break
                    else:
                        if isinstance(a, float):
                            a = int(a)
                        elif a == '':
                            a = 0
                        if isinstance(b, float):
                            b = str(int(b))
                        self.total_pledge += c
                        self.num_pledge += 1
                        pledge_tuple = (a, b, c)
                        self.pledge_list.append(pledge_tuple)
                self.pledge_list.sort()

                # add special data to special_list
                special_col_offset = 3
                for special_row in range(start_row, plate_row):
                    w = wrksht[special_row][special_col_offset]
                    x = wrksht[special_row][special_col_offset + 1]
                    y = wrksht[special_row][special_col_offset + 2]
                    z = wrksht[special_row][special_col_offset + 3]
                    if (re.sub(r'\s+', '', str(w)) == '') and \
                            (re.sub(r'\s+', '', str(x)) == '') and \
                            (re.sub(r'\s+', '', str(y)) == '') and \
                            (re.sub(r'\s+', '', str(z)) == ''):
                        break
                    else:
                        if isinstance(w, float):
                            w = str(int(w))
                        if isinstance(y, float):
                            y = str(int(y))
                        self.total_special += z
                        self.num_special += 1
                        special_tuple = (w, x, y, z)
                        self.special_list.append(special_tuple)
                self.special_list.sort()

    def finder(self, num=None, keywords=None):
        pledge = []
        special = []
        if num is not None:
            pledge += [x for x in self.pledge_list if x[0] == num]
            special += [x for x in self.special_list if x[0] == str(num)]
        if keywords is not None:
            for word in keywords:
                special += [x for x in self.special_list if word in x[1].lower()]
        if len(pledge) > 0 or len(special) > 0:
            print("")
            if len(pledge) > 0:
                print("SPECIFIED PLEDGE (" + num + "): ")
                for e in pledge:
                    print(e)
            if len(special) > 0:
                print("SPECIFIED SPECIAL (" + num + "): ")
                for e in special:
                    print(e)

    def print_summary(self):
        print("\nSunday School: " + str(self.sunday_school))
        print("Plate: " + str(self.plate))
        print("Amount of pledge: " + str(self.num_pledge))
        print("Amount of special: " + str(self.num_special))
        print("Total Pledge: " + str(self.total_pledge))
        print("Total Special: " + str(self.total_special))
        #print("\n\nOutside Incomes:")
        print("\n\nGrand Total: " + str(self.total_pledge+self.total_special+self.plate+self.sunday_school)+"\n")


class Annual():
    pledge_list = []
    special_list = []
    plate = []
    sunday_school = []

    total_pledge = 0
    total_special = 0
    total_plate = 0
    total_sunday_school = 0

    num_pledge = 0
    num_special = 0
    num_sunday_school = 0
    num_plate = 0

    def insert_weekly_data(self,weekly):
        self.pledge_list.extend(weekly.pledge_list)
        self.special_list.extend(weekly.special_list)
        if weekly.plate > 0:
            self.plate.append(weekly.plate)
        if weekly.sunday_school > 0:
            self.sunday_school.append(weekly.sunday_school)

        self.total_pledge += weekly.total_pledge
        self.total_special += weekly.total_special
        self.total_plate += weekly.plate
        self.total_sunday_school += weekly.sunday_school

        self.num_plate += weekly.num_plate
        self.num_sunday_school += weekly.num_sunday_school
        self.num_pledge += weekly.num_pledge
        self.num_special += weekly.num_special

    def get_data(self, path=None, file_list=None):
        if path is not None:
            if os.path.isdir(path):
                if file_list is not None:
                    # print(file_list)
                    # print(os.path.join(path, file_list))
                    path_and_list = os.path.join(path, file_list)
                    if os.path.isfile(path_and_list):
                        # print("im in")
                        # file = open(path_and_list, 'r')
                        with open(path_and_list, 'r') as f:
                            weekly_list = f.read().splitlines()
                        for weekly_file in weekly_list:
                            print("--------------------------------------------------")
                            print("\n"+weekly_file)
                            weekly = Weekly()
                            weekly.parse(os.path.join(path, weekly_file))
                            weekly.print_summary()
                            self.insert_weekly_data(weekly)
                            weekly = None
                    else:
                        print("File List \"" + str(file_list) + "\" is not a file!")
                else:
                    for weeklyfile in os.listdir(path):
                        if weeklyfile.endswith(".xls") and ("Income 2014" not in weeklyfile):
                            path_and_weekly = os.path.join(path, weeklyfile)
                            if os.path.isfile(path_and_weekly) is True:
                                weekly = Weekly()
                                weekly.parse(path_and_weekly)
                                weekly.print_summary()
                                self.insert_weekly_data(weekly)
                                weekly = None
                            else:
                                print("File \"" + str(weeklyfile) + "\" is not a file!")
                        else:
                            print("File \"" + str(weeklyfile) + "\" is not a valid file!")
            else:
                print("Path \"" + str(path) + "\" is not a directory!")
        else:
            if file_list is not None:
                if os.path.isfile(file_list):
                    # file = open(file_list, 'r')
                    with open(file_list, 'r') as f:
                        weekly_list = f.read().splitlines()
                    for weekly_file in weekly_list:
                        weekly = Weekly()
                        weekly.parse(weekly_file)
                        weekly.print_summary()
                        self.insert_weekly_data(weekly)
                        weekly = None
                else:
                    print("File List \"" + str(file_list) + "\" is not a file!")

    def finder(self, num=None, keywords=None):
        pledge = []
        special = []
        if num is not None:
            pledge += [x for x in self.pledge_list if x[0] == num]
            special += [x for x in self.special_list if x[0] == str(num)]
        if keywords is not None:
            for word in keywords:
                special += [x for x in self.special_list if word in x[1].lower()]
        if len(pledge) > 0 or len(special) > 0:
            print("")
            sum1 = 0
            sum2 = 0
            if len(pledge) > 0:
                print("TOTAL SPECIFIED PLEDGE LIST ("+num+"): ")
                for e in pledge:
                    print(e)
                    sum1 += e[2]
                print("TOTAL SPECIFIED PLEDGE ("+num+"): " + str(sum1))
            if len(special) > 0:
                print("TOTAL SPECIFIED SPECIAL LIST ("+num+"): ")
                for e in special:
                    print(e)
                    sum2 += e[3]
                print("TOTAL SPECIFIED SPECIAL ("+num+"): " + str(sum2))

    def print_summary(self):
        print("--------------------------------------------------")
        print("                  Annual Summary                  ")
        print("--------------------------------------------------")
        print("Amount of Plate: " + str(self.num_plate))
        print("Amount of Sunday School: " + str(self.num_sunday_school))
        print("Amount of Pledge: " + str(self.num_pledge))
        print("Amount of Special: " + str(self.num_special))

        print("Total Plate: " + str(self.total_plate))
        print("Total Sunday School: " + str(self.total_sunday_school))
        print("Total Pledge: " + str(self.total_pledge))
        print("Total Special: " + str(self.total_special))
        # print("\n\nOutside Incomes:")
        print("\n\nGrand Total For the Year: " +
              str(self.total_pledge+self.total_special+self.total_plate+self.total_sunday_school)+"\n")