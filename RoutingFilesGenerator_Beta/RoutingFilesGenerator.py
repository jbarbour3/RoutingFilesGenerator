import tkinter as tk
from tkinter import filedialog
import xlrd
import xlsxwriter
import random
import math
import csv
import traceback
from tkinter import messagebox


class RoutingFileGenerator:
    def __init__(self, root):
        self.root = root
        self.frames = {}
        self.apply_defaults()

        # last resort for ugly variable locations
        self.design_file = tk.StringVar()
        self.sheet_names = ("Select a design file first",)
        self.sheet_name = tk.StringVar(self.root)
        self.sheet_name.set(self.sheet_names[0])
        self.scatter_radius = tk.DoubleVar()
        self.scatter_radius.set(5.0)
        self.total_stops_in_stop_file = 0
        self.day = {"Monday":'01',
                    "Tuesday":'02',
                    "Wednesday":'03',
                    "Thursday":'04',
                    "Friday":'05',
                    "Saturday":'06',
                    "Sunday":'07'}


        self.stop_data = self.initialize_stop_data()
        self.truck_data = self.initialize_truck_data()



        # build UI
        self.win = self.build_frame()

        self.top = self.build_top(self.win) #<--! code 3/z for full screen
        self.middle = self.build_middle(self.win) #<--! code 4/z for full screen
        self.bottom = self.build_bottom(self.win) #<--! code 5/z for full screen

    def apply_defaults(self):
        #self.root.geometry('{0}x{1}'.format(root.winfo_screenwidth(), root.winfo_screenheight())) <--! code 1/z for full screen
        #self.root.state('zoomed') <--! code 2/z for full screen


        #self.root.geometry('{0}x{1}'.format(300, 200))

        self.root.title("Routing File Generator")
        self.header = ("Verdana bold", 8)
        self.header2 = ("Verdana", 8)
        self.bg1 = "Coral2"
        self.bg2 = "silver"
        self.bg3 = "grey"
        self.simulation_start_monday_date = "10/01/2018" # monday

    def build_frame(self):
        win = tk.Frame(self.root, bg=self.bg1)
        win.pack(expand=True, fill=tk.BOTH, pady=2, padx=2)
        return win

    def build_top(self, aFrame):
        top = tk.Frame(aFrame)
        top.pack(fill=tk.X)
        tk.Entry(top, textvariable = self.design_file, width=98).pack(side=tk.LEFT)

        top2 = tk.Frame(aFrame)
        top2.pack(fill=tk.X)
        tk.Button(top2, text="Browse For Design File", command=self.browse_for_design_file, font=self.header).pack(side=tk.LEFT)
        tk.Label(top2, text=" ").pack(side=tk.LEFT)

        self.sheet_Options = tk.OptionMenu(top2, self.sheet_name, *self.sheet_names)
        self.sheet_Options.pack(side=tk.LEFT)

        tk.Label(top2, text="Scatter Radius").pack(side=tk.LEFT)
        tk.Entry(top2, textvariable=self.scatter_radius, width=4, justify=tk.RIGHT).pack(side=tk.LEFT)
        tk.Label(top2, text="mi").pack(side=tk.LEFT)

        return top

    def build_middle(self, aFrame):
        middle = tk.Frame(aFrame, bg=self.bg2)
        middle.pack(expand=True, fill=tk.BOTH)

        #StopFrame
        stopFrame = tk.Frame(middle, bd=2, bg = self.bg2)
        stopFrame.pack(anchor=tk.W)

        #TruckFrame
        truckFrame = tk.Frame(middle, bd=2, bg = self.bg2)
        truckFrame.pack(anchor=tk.W)

        for frame, data in [(stopFrame, self.stop_data), (truckFrame, self.truck_data)]:
            for section in data.keys():
                if section != "Other_Required":
                    # build the frame
                    self.frames[section] = tk.Frame(frame, bd=0)

                    self.frames[section].pack(side=tk.LEFT, anchor=tk.N, padx=2)

                    tk.Label(self.frames[section], text=section, bg=self.bg1, font=self.header).pack(expand=True, fill=tk.X)

        self.build_pcs_dist_widget(self.frames["Pieces Distribution"])
        self.build_volume_widget(self.frames["Volume"])
        self.build_time_windows_widget(self.frames["Time Windows"])
        self.build_days_of_service_widget(self.frames["Days Of Service"])

        self.build_capacity_widget(self.frames["Capacity"])
        self.build_facility_widget(self.frames["Facility"])
        self.build_costs_layover_widget(self.frames["Costs/Layover"])
        self.build_work_rules_widget(self.frames["Work Rules"])

        return middle

    def build_pcs_dist_widget(self, aFrame):
        header = tk.Frame(aFrame, bg=self.bg2)
        header.pack(fill=tk.X)
        tk.Label(header, text="         #", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text=" "*2, font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text="   % ", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        for i in range(8):
            key = str(i+1)
            variable = self.stop_data["Pieces Distribution"][key]
            frame = tk.Frame(aFrame)
            frame.pack()
            tk.Label(frame,text=key, width=3).pack(side=tk.LEFT)
            tk.Entry(frame, state= tk.DISABLED, relief=tk.FLAT, width=1).pack(side=tk.LEFT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=variable, width=5).pack(side=tk.LEFT)
        return None

    def build_volume_widget(self, aFrame):
        for key in self.stop_data["Volume"].keys():
            header = tk.Frame(aFrame, bg=self.bg2)
            header.pack(fill=tk.X, expand=True, anchor=tk.W)
            tk.Label(header, text=key, font=self.header, justify=tk.LEFT, bg=self.bg2).pack(anchor=tk.W)
            for key2 in self.stop_data["Volume"][key].keys():
                frame = tk.Frame(aFrame)
                frame.pack()
                tk.Label(frame, text="  ").pack(side=tk.LEFT)
                tk.Entry(frame, textvariable=self.stop_data["Volume"][key][key2], justify=tk.RIGHT, width=5).pack(side=tk.LEFT)
                tk.Label(frame, text=key2, width=10, justify=tk.LEFT).pack(side=tk.LEFT, expand=True, fill=tk.X)
        return None

    def build_time_windows_widget(self, aFrame):
        header = tk.Frame(aFrame, bg=self.bg2)
        header.pack(fill=tk.X,anchor=tk.W)
        tk.Label(header, text=" " * 2, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text="From", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text=" " * 2, font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text="  To", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text=" " * 2, font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text="    Pattern", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text=" ",bg=self.bg2).pack(side=tk.LEFT)

        for i in range(8):
            key = str(i+1)
            frame = tk.Frame(aFrame, bg=self.bg2)
            frame.pack()
            tk.Label(frame, text=key).pack(side=tk.LEFT)
            tk.Entry(frame, textvariable=self.stop_data["Time Windows"][key]["From"], justify=tk.RIGHT, width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" " * 4).pack(side=tk.LEFT)
            tk.Entry(frame, textvariable=self.stop_data["Time Windows"][key]["To"], justify=tk.RIGHT, width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" " * 4).pack(side=tk.LEFT)
            tk.Entry(frame, textvariable=self.stop_data["Time Windows"][key]["Pattern"], justify=tk.RIGHT, width=10).pack(side=tk.LEFT)
            tk.Label(frame, text=" ").pack(side=tk.LEFT)


        return None

    def build_days_of_service_widget(self, aFrame):
        header = tk.Frame(aFrame, bg=self.bg2)
        header.pack(fill=tk.X, expand=True, anchor=tk.W)
        tk.Label(header, text="     Day", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text=" " , font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        tk.Label(header, text="% Volume ", font=self.header, bg=self.bg2).pack(side=tk.LEFT)
        for key in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]:
            variable = self.stop_data["Days Of Service"][key]
            frame = tk.Frame(aFrame)
            frame.pack(expand=True, fill=tk.X, anchor=tk.W)
            tk.Label(frame, text=key, justify=tk.LEFT).pack(side=tk.LEFT, expand=True, fill=tk.X)
            tk.Label(frame, text="  ").pack(side=tk.RIGHT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=variable, width=5).pack(side=tk.RIGHT)

        return None

    def build_capacity_widget(self, aFrame):
        for key in self.truck_data["Capacity"]:
            frame = tk.Frame(aFrame, bg=self.bg2)
            frame.pack()
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=self.truck_data["Capacity"][key], width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Label(frame, text=key, justify=tk.LEFT, width=10).pack(side=tk.LEFT)
        return None

    def build_facility_widget(self, aFrame):
        for key in self.truck_data["Facility"]:
            frame = tk.Frame(aFrame, bg=self.bg2)
            frame.pack()
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=self.truck_data["Facility"][key], width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Label(frame, text=key, justify=tk.LEFT, width=9).pack(side=tk.LEFT)
        return None

    def build_costs_layover_widget(self, aFrame):
        for key in self.truck_data["Costs/Layover"]:
            frame = tk.Frame(aFrame, bg=self.bg2)
            frame.pack()
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=self.truck_data["Costs/Layover"][key], width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Label(frame, text=key, justify=tk.LEFT, width=19).pack(side=tk.LEFT)

        return None

    def build_work_rules_widget(self, aFrame):
        for key in self.truck_data["Work Rules"]:
            frame = tk.Frame(aFrame, bg=self.bg2)
            frame.pack()
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Entry(frame, justify=tk.RIGHT, textvariable=self.truck_data["Work Rules"][key], width=5).pack(side=tk.LEFT)
            tk.Label(frame, text=" ").pack(side=tk.LEFT)
            tk.Label(frame, text=key, justify=tk.LEFT, width=11).pack(side=tk.LEFT)
        return None

    def build_bottom(self, aFrame):
        bottom = tk.Frame(aFrame, bg=self.bg2)
        bottom.pack(side=tk.BOTTOM,fill=tk.X)

        tk.Button(bottom, text="Build Routing Files", command=self.build_routing_files).pack(padx=1, side=tk.RIGHT)
        tk.Button(bottom, text="Reset All Variables", command=self.reset_all_variables).pack(padx=1, side=tk.RIGHT)
        tk.Button(bottom, text="Close", command=self.root.destroy).pack(padx=2,side=tk.RIGHT)
        return bottom

    def reset_all_variables(self):
        print("Hey Hey Batman, look at that design file")
        return None

    def initialize_stop_data(self):
        stopdict = {"Pieces Distribution": {"1": tk.DoubleVar(),
                                            "2": tk.DoubleVar(),
                                            "3": tk.DoubleVar(),
                                            "4": tk.DoubleVar(),
                                            "5": tk.DoubleVar(),
                                            "6": tk.DoubleVar(),
                                            "7": tk.DoubleVar(),
                                            "8": tk.DoubleVar(),
                                            },

                    "Volume": {"Cube": {"Per Piece": tk.IntVar(), "Per Stop": tk.IntVar(), "Minimum": tk.IntVar()},
                               "Weight": {"Per Piece": tk.IntVar(), "Per Stop": tk.IntVar(), "Minimum": tk.IntVar()},
                               "Time": {"Per Piece": tk.IntVar(), "Per Stop": tk.IntVar(), "Minimum": tk.IntVar()},
                               },

                    "Time Windows": {"1": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "2": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "3": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "4": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "5": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "6": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "7": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     "8": {"From": tk.StringVar(), "To": tk.StringVar(), "Pattern": tk.StringVar()},
                                     },

                    "Days Of Service": {"Monday": tk.DoubleVar(),
                                        "Tuesday": tk.DoubleVar(),
                                        "Wednesday": tk.DoubleVar(),
                                        "Thursday": tk.DoubleVar(),
                                        "Friday": tk.DoubleVar(),
                                        "Saturday": tk.DoubleVar(),
                                        "Sunday": tk.DoubleVar(),
                                        }
                    }
        # default value loading
        stopdict["Pieces Distribution"]["1"].set(40.0)
        stopdict["Pieces Distribution"]["2"].set(30.0)
        stopdict["Pieces Distribution"]["3"].set(20.0)
        stopdict["Pieces Distribution"]["4"].set(10.0)

        stopdict["Time Windows"]["1"]["From"].set("0700")
        stopdict["Time Windows"]["1"]["To"].set("1500")
        stopdict["Time Windows"]["1"]["Pattern"].set("SMTWRFA")

        stopdict["Days Of Service"]["Monday"].set(10.0)
        stopdict["Days Of Service"]["Tuesday"].set(16.0)
        stopdict["Days Of Service"]["Wednesday"].set(16.0)
        stopdict["Days Of Service"]["Thursday"].set(16.0)
        stopdict["Days Of Service"]["Friday"].set(16.0)
        stopdict["Days Of Service"]["Saturday"].set(16.0)
        stopdict["Days Of Service"]["Sunday"].set(10.0)

        stopdict["Volume"]["Time"]["Per Piece"].set(5)
        stopdict["Volume"]["Time"]["Minimum"].set(15)

        return stopdict

    def build_stop_data(self):
        # create stop ids from the expected demand in an xlsx file
        stop_ids = self.transform_demand_into_stops(self.design_file.get(), self.sheet_name.get())

        # save total num of stop ids for use in allocating trucks
        self.total_stops_in_stop_file = len(stop_ids)

        # do preleminary data work we need to simulate orders
        # pieces dist
        pieces_cdf = self.get_discrete_dist("Pieces Distribution")

        # volume function params
        cube_parameters = self.get_volume_parameters("Cube")
        weight_parameters = self.get_volume_parameters("Weight")
        time_parameters = self.get_volume_parameters("Time")

        # time window definitions
        time_windows = self.get_time_windows()

        # get the days of service distribution
        days_of_service_cdf = self.get_discrete_dist("Days Of Service")

        # turn the days in the dos CDF into actual dates formatted MM/DD/YYYY
        dates_of_service_cdf = []
        start_date = self.simulation_start_monday_date
        for prob, value in days_of_service_cdf:
            days_to_add = self.day[value]
            date = start_date[:3] + days_to_add + start_date[5:]
            dates_of_service_cdf.append((prob, date))

        # get zip centroids to make available for scatter calculation
        zip_center = self.get_zip_centroid_dict()

        # get scatter radius
        radius = self.scatter_radius.get()

        # for every stop id, build a simulated stop with all the requested information
        stop_data = {}
        for stop_id in stop_ids:
            stop_data[stop_id] = {}
            stop = stop_data[stop_id]
            stop["ID1"] = stop_id  # create the stop id column and fill with id's
            stop["Zip"] = int(stop_id.split("_")[0])  # create the stop zip column and fill with zips
            stop["Pieces"] = int(self.discrete_dist_draw(pieces_cdf))
            pcs = stop["Pieces"]
            stop["FixedTime"] = self.calc_volume(time_parameters, pcs)
            stop["Weight"] = self.calc_volume(weight_parameters, pcs)
            stop["Cube"] = self.calc_volume(cube_parameters, pcs)
            for tw, opn, close, pattern in time_windows:
                stop["Open" + tw] = opn
                stop["Close" + tw] = close
                stop["Pattern" + tw] = pattern

            stop["Earliest Date"] = self.discrete_dist_draw(dates_of_service_cdf)
            stop["Latest Date"] = stop["Earliest Date"]

            # stop scatter logic
            dist_dir_vector = self.get_random_polar_vector_within_radius(radius)
            zip_lat = zip_center[stop["Zip"]]["Latitude"]
            zip_lon = zip_center[stop["Zip"]]["Longitude"]
            new_point = self.get_new_lat_lon_given(zip_lat, zip_lon, dist_dir_vector[0], dist_dir_vector[1])
            stop["Latitude"] = new_point[0]
            stop["Longitude"] = new_point[1]

        return stop_data

    def build_file(self, a_dict_of_data, a_file_name, a_file_type):
        """Writes a dicts of data out with keys as columns as either a .csv or .xlsx file.
        Note: Truck files must be written out as csv or appian will fail to load the trucks"""
        data = []
        build_header = True
        header = []
        for id in a_dict_of_data.keys():
            record = a_dict_of_data[id]
            if build_header:
                header = [key for key, value in record.items()]
                build_header = False
                data.append(header)
            line = []
            for key in header:
                line.append(record[key])
            data.append(line)

        if a_file_type == "xlsx":
            workbook = xlsxwriter.Workbook(a_file_name)
            worksheet = workbook.add_worksheet()

            # Iterate over the data and write it out row by row.
            for row in range(len(data)):
                for col in range(len(data[row])):
                    worksheet.write(row, col, data[row][col])
            workbook.close()

        if a_file_type == "csv":
            with open(a_file_name, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                for row in data:
                    writer.writerow(row)



        return None

    def initialize_truck_data(self):
        # Create variables, must be StringVar, IntVar, DoubleVar, and BooleanVar
        truckdict = {"Capacity":{"Cube":tk.IntVar(), "Pieces": tk.IntVar(), "Weight": tk.IntVar(),"Truck Count":tk.IntVar()},
                     "Facility":{"PreTrip": tk.IntVar(), "PostTrip": tk.IntVar(), "Zip": tk.StringVar()},
                     "Costs/Layover":{"FixedCost": tk.IntVar(), "MiCost": tk.DoubleVar(), "LayoverCost": tk.IntVar(), "MinLayover":tk.IntVar(), "MaxLayover": tk.IntVar(), "MaxDrvTmB4Layover": tk.IntVar()},
                     "Work Rules": {"EarStart": tk.IntVar(), "LatFinish": tk.IntVar(), "MaxWorkTm":tk.IntVar(), "MaxDrvTm": tk.IntVar(), "TargetWrkTm": tk.IntVar()},
                     "Other_Required": {"Longitude": tk.StringVar(), "Latitude": tk.StringVar(), "Available": tk.StringVar(), "OneWay": tk.StringVar(), "Redispatch": tk.StringVar(),"EDate": tk.IntVar(), "LDate":tk.IntVar(), },
                     }
        # Default Value Loading
        truckdict["Capacity"]["Cube"].set(1500)
        truckdict["Capacity"]["Pieces"].set(55)
        truckdict["Capacity"]["Weight"].set(9999)
        truckdict["Capacity"]["Truck Count"].set(0)


        truckdict["Facility"]["PreTrip"].set(60)
        truckdict["Facility"]["PostTrip"].set(15)
        truckdict["Facility"]["Zip"].set(0)

        truckdict["Costs/Layover"]["FixedCost"].set(550)
        truckdict["Costs/Layover"]["MiCost"].set(.320)
        truckdict["Costs/Layover"]["LayoverCost"].set(125)
        truckdict["Costs/Layover"]["MinLayover"].set(10)
        truckdict["Costs/Layover"]["MaxLayover"].set(16)
        truckdict["Costs/Layover"]["MaxDrvTmB4Layover"].set(11)

        truckdict["Work Rules"]["EarStart"].set(800)
        truckdict["Work Rules"]["LatFinish"].set(2100)
        truckdict["Work Rules"]["MaxWorkTm"].set(12)
        truckdict["Work Rules"]["MaxDrvTm"].set(10)
        truckdict["Work Rules"]["TargetWrkTm"].set(12)

        truckdict["Other_Required"]["Longitude"].set("")
        truckdict["Other_Required"]["Latitude"].set("")
        truckdict["Other_Required"]["Available"].set("TRUE")
        truckdict["Other_Required"]["OneWay"].set("FALSE")
        truckdict["Other_Required"]["Redispatch"].set("FALSE")
        truckdict["Other_Required"]["EDate"].set(0)
        truckdict["Other_Required"]["LDate"].set(6)

        return truckdict

    def build_truck_data(self):
        # generate the number of truck ids required (either fixed per day or = stops/5)
        if self.truck_data["Capacity"]["Truck Count"].get()>0:
            truck_count_is_fixed=True
        else:
            truck_count_is_fixed=False

        truck_ids = self.transform_capacity_into_trucks(truck_count_is_fixed)

        # build the constants for each truck
        constants = {}
        for key1 in self.truck_data.keys():
            for key2 in self.truck_data[key1].keys():
                constants[key2] = self.truck_data[key1][key2].get()

        # remove the truck count constant
        constants.pop("Truck Count")

        # generate the lat and lon for the facility zip
        zip_center = self.get_zip_centroid_dict()
        facility_zip = int(self.truck_data["Facility"]["Zip"].get())
        constants["Latitude"] = zip_center[facility_zip]["Latitude"]
        constants["Longitude"] = zip_center[facility_zip]["Longitude"]

        # write the last of the updates
        truck_data = {}
        for truck_id in truck_ids:
            truck_data[truck_id] = {}
            truck = truck_data[truck_id]
            truck["TrkID"]=truck_id
            for key in constants:
                truck[key] = constants[key]
            if truck_count_is_fixed:
                truck_day = truck_id.split("_")[1]
                truck["EDate"] = int(self.day[truck_day])-1
                truck["LDate"] = truck["EDate"]
        return truck_data

    def browse_for_design_file(self):
        self.design_file.set(tk.filedialog.askopenfilename(title="Select Design File Containing Expected Zip Demand"))
        m = self.sheet_Options["menu"]
        m.delete(0,"end")
        newvalues = tuple(xlrd.open_workbook(self.design_file.get()).sheet_names())

        for val in newvalues:
            m.add_command(label=val, command=lambda v=self.sheet_name, l=val: v.set(l))
        self.sheet_name.set(newvalues[0])
        return None

    def build_routing_files(self):
        passed = True
        try:
            stop_data = self.build_stop_data()
        except Exception as e:
            passed = False
            title = "Error Building Stop File Data"
            tk.messagebox.showerror(title, "{0}\n{1}".format(e,traceback.format_exc()))

        try:
            truck_data = self.build_truck_data()
        except Exception as e:
            passed = False
            title = "Error Building Truck File Data"
            tk.messagebox.showerror(title, "{0}\n{1}".format(e, traceback.format_exc()))

        if passed:
            self.build_file(stop_data, 'Stop_File_Demo.xlsx', 'xlsx')
            self.build_file(truck_data,'Truck_File_Demo.TRUCK','csv')
            tk.messagebox.showinfo("Routing Files Generator", "Stop and Truck File Generated Successfully.")

        return None

    def build_truck_file(self, truck_data_dict):
        truck_file = [[]]
        return truck_file

    def transform_demand_into_stops(self, excel_file_name, sheet_name):
        """using a sheet containing expected weekly demand by zip, creates a list of unique stops"""
        z = {}
        book = xlrd.open_workbook(excel_file_name, 'r')
        sheet = book.sheet_by_name(self.sheet_name.get())
        num_rows = sheet.nrows

        for row_index in range(1,num_rows, 1):
            zip_code = sheet.cell(row_index,0).value
            demand = sheet.cell(row_index,1).value
            z[zip_code] = demand

        demand = {}
        for key in z:
            stops = int(z[key])
            if random.random()<(z[key] - int(z[key])):
                stops+=1
            demand[key] = stops
        stop_ids = []
        for key in demand:
            count=demand[key]
            for stop_num in range(1,count,1):
                stop_id = "{}_{}".format(key, stop_num)
                stop_ids.append(stop_id)
        return stop_ids

    def transform_capacity_into_trucks(self, true_if_fixed):
        truck_ids = []
        if true_if_fixed:
            days = []
            for prob, day in self.get_discrete_dist("Days Of Service"):
                days.append(day)
            for day in days:
                for i in range(self.truck_data["Capacity"]["Truck Count"].get()):
                    truck_ids.append("Truck_{0}_{1}".format(day, i))
        else:
            for i in range(int(self.total_stops_in_stop_file/5)):
                truck_ids.append("Truck_{0}".format(i))
        return truck_ids


    def discrete_dist_draw(self, a_PDF_List):
        """Returns a realization of a random variable according to a piecewise PDF, ex. [(.25,1), (.75,2)]"""
        if len(a_PDF_List)==1:
            return a_PDF_List[0][1]
        else:
            draw = random.random()*100
            cumul = 0
            for prob, value in a_PDF_List:
                cumul += prob
                if draw<cumul:
                    return value

    def get_discrete_dist(self, aSection):
        "Returns the pdf entered by the user for the pieces category"
        discrete_dist = []
        probability_input = self.stop_data[aSection]
        for value in probability_input:
            probability = probability_input[value].get()
            if probability > 0.001:
                discrete_dist.append((probability, value))
        return discrete_dist

    def get_volume_parameters(self,a_dimension):
        """gets (minimmum_vol, vol_per_piece, vol_per_stop) from the UI for a given volume"""
        params = self.stop_data["Volume"][a_dimension]
        return (params["Minimum"].get(), params["Per Piece"].get(), params["Per Stop"].get())

    def calc_volume(self,a_tuple_of_values, a_piece_count):
        """Returns the volume of a stop according to
        inputs of ("Minimum Per Stop", "qty Per Piece", "qty Per Stop")"""
        mini, per_piece, per_stop = a_tuple_of_values

        if per_stop>0.0001:
            return per_stop

        vol = a_piece_count * per_piece
        if vol>mini:
            return vol

        return mini

    def get_time_windows(self):
        tws = []
        tw_input = self.stop_data["Time Windows"]
        for win in tw_input:
            open = tw_input[win]["From"].get()
            if len(open)>0:
                close = tw_input[win]["To"].get()
                pattern = tw_input[win]["Pattern"].get()
                tws.append((win, open, close, pattern))
        return tws

    def get_zip_centroid_dict(self):

        zip = {}
        for i in range(100000):
            zip[i] = {"Latitude": None,
                      "Longitude": None}
        with open("GeoZip_Center_Geo_Code.csv", 'r')as f:
            for line in f:
                line = line.strip('\n').split(',')
                zip[int(line[0])]["Latitude"] = float(line[1])
                zip[int(line[0])]["Longitude"] = float(line[2])

        return zip

    def get_random_polar_vector_within_radius(self, a_max_radius):
        distance = random.random()*a_max_radius
        direction = random.random()*360.0
        return (distance, direction)

    def get_new_lat_lon_given(self, origin_lat, origin_lon, distance_in_miles, direction_in_degrees):
        # this uses https://gis.stackexchange.com/questions/142326/calculating-longitude-length-in-miles
        # to return new point. It's not uniformly distributing the points but it's good enough for now

        # get the x,y components of the shift in miles
        x_mileage_shift = distance_in_miles * math.cos(direction_in_degrees)
        y_mileage_shift = distance_in_miles * math.sin(direction_in_degrees)

        # convert lat and lon diff
        lat_shift = x_mileage_shift*1.00/69.00
        lon_shift = y_mileage_shift*1.00/(math.cos(origin_lat)*69.172)

        lat = origin_lat + lat_shift
        lon = origin_lon + lon_shift

        return (lat, lon)

if __name__ == "__main__":

    root = tk.Tk()
    app = RoutingFileGenerator(root)
    app.root.mainloop()