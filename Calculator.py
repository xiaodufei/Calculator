import os
import time
import math
import datetime
import holidays
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from interval import Interval
from business_calendar import Calendar, MO, TU, WE, TH, FR
############################################################################################
cable_variety = [80, 40, 20, 10, 4, 2, 1]
time_variety = [40, 20, 10, 5, 2, 1]

dark_cable_sort = ['04 GIS DARK FIBER CABLE 20000', '04 GIS DARK FIBER CABLE 10000', '04 GIS DARK FIBER CABLE 5000',
                    '04 GIS DARK FIBER CABLE 2500', '04 GIS DARK FIBER CABLE 1000', '04 GIS DARK FIBER CABLE 500',
                    '04 GIS DARK FIBER CABLE 250']

own_cable_sort = ['04 GIS OWN CABLE 20000', '04 GIS OWN CABLE 10000', '04 GIS OWN CABLE 5000',
                  '04 GIS OWN CABLE 2500', '04 GIS OWN CABLE 1000', '04 GIS OWN CABLE 500', '04 GIS OWN CABLE 250']

work_time_sort = ['04 GIS Additional Work 40H', '04 GIS Additional Work 20H', '04 GIS Additional Work 10H',
                    '04 GIS Additional Work 05H', '04 GIS Additional Work 02H', '04 GIS Additional Work 01H']

class Calculator(object):

    def __init__(self, ticket_number, cable_type, initial_setup, cable_length, additional_work_time, region):
        self.ticket_number = ticket_number
        self.initial_setup = initial_setup
        self.df = pd.DataFrame(columns=('Ticket', 'Category_3', 'SLA', 'Start', 'End', 'Area'))
        self.cable_type = cable_type
        self.cable_length = cable_length
        self.region = region
        self.additional_work_time = additional_work_time
        self.sla_data = self.sla_calculate()
        self.date = self.date_calculate()
        self.area = self.region_calculate()

    def sla_calculate(self):
        if self.cable_length in Interval(0, 1, lower_closed=False):
            return {'SLA': 1, 'working_days': 9}
        elif self.cable_length in Interval(1, 5, lower_closed=False):
            return {'SLA': 2, 'working_days': 14}
        elif self.cable_length in Interval(5, 10, lower_closed=False):
            return {'SLA': 3, 'working_days': 19}
        elif self.cable_length in Interval(10, 20, lower_closed=False):
            return {'SLA': 4, 'working_days': 24}
        else:
            return {'SLA': 5, 'working_days': 29}

    def date_calculate(self):
        i = 1
        de_holidays = holidays.Germany(years=(2019, 2020, 2021))
        while True:
            next_working_day = datetime.date.today() + datetime.timedelta(days=i)
            if next_working_day.weekday() >= 5:
                i += 1
                continue
            elif next_working_day in de_holidays:
                i += 1
                continue
            else:
                break
        cal = Calendar(workdays=[MO, TU, WE, TH, FR], holidays=de_holidays)
        last_working_day = cal.addbusdays(next_working_day, self.sla_data['working_days'])
        start_date = next_working_day.strftime('%#m/%#d/%Y 8:00:00 AM')
        end_date = last_working_day.strftime('%#m/%#d/%Y 5:00:00 PM')
        return {'start': start_date, 'end': end_date}

    def region_calculate(self):
        if self.region[0] in ('O', 'W'):
            return {'Area': 'G-SAENTIS-TRANS-GIS-2'}
        else:
            return {'Area': 'G-SAENTIS-TRANS-GIS-1'}

    def init_setup_calculate(self):
        for i in range(self.initial_setup):
            if self.cable_type == 'Dark Fiber':
                self.df = self.df.append({'Ticket': self.ticket_number, 'Category_3': '04 GIS DARK FIBER CABLE Initial Setup',
                        'SLA': str(self.sla_data['SLA']), 'Start': self.date['start'], 'End': self.date['end'],
                                          'Area': self.area['Area']}, ignore_index=True)
            else:
                self.df = self.df.append({'Ticket': self.ticket_number, 'Category_3': '04 GIS OWN CABLE Initial Setup',
                     'SLA': str(self.sla_data['SLA']), 'Start': self.date['start'], 'End': self.date['end'],
                     'Area': self.area['Area']}, ignore_index=True)

    def length_calculate(self):
        remainder = math.ceil(self.cable_length / 0.25)
        for i in range(len(cable_variety)):
            result = divmod(remainder, cable_variety[i])
            for num in range(result[0]):
                if self.cable_type == 'Dark Fiber':
                    self.df = self.df.append({'Ticket': self.ticket_number, 'Category_3': dark_cable_sort[i], 'SLA': str(self.sla_data['SLA']),
                            'Start': self.date['start'], 'End': self.date['end'], 'Area': self.area['Area']}, ignore_index=True)
                else:
                    self.df = self.df.append({'Ticket': self.ticket_number, 'Category_3': own_cable_sort[i], 'SLA': str(self.sla_data['SLA']),
                            'Start': self.date['start'], 'End': self.date['end'], 'Area': self.area['Area']}, ignore_index=True)
            remainder = result[1]
            if remainder == 0:
                break

    def time_calculate(self):
        remainder = self.additional_work_time
        for i in range(len(time_variety)):
            result = divmod(remainder, time_variety[i])
            for num in range(result[0]):
                self.df = self.df.append({'Ticket': self.ticket_number, 'Category_3': work_time_sort[i], 'SLA': str(self.sla_data['SLA']),
                        'Start': self.date['start'], 'End': self.date['end'], 'Area': self.area['Area']}, ignore_index=True)
            remainder = result[1]
            if remainder == 0:
                break

class User_Interface(Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry('480x180')
        self.master.iconbitmap('huawei_logo.ico')
        self.master.title('  Calculation - Cable Length')
        self.create_widgets()

    def create_widgets(self):
        self.file = Label(self.master, text='File:')
        self.file.place(x=30, y=70, anchor='nw')
        self.choose_file = Button(self.master, text='choose file', command=self.open_file)
        self.choose_file.place(x=380, y=65, anchor='nw')
        self.file_path_display = Entry(self.master, width=45)
        self.file_path_display.config(state=DISABLED)
        self.file_path_display.place(x=70, y=70, anchor='nw')
        self.run = Button(self.master, text='Calculate !', command=self.start_calculate)
        self.run.place(x=70, y=110, anchor='nw')

    def open_file(self):
        try:
            self.file_path = filedialog.askopenfile(title='please choose Excel-file').name
        except AttributeError:
            pass
        else:
            self.file_path_display.config(state=NORMAL)
            self.file_path_display.delete(0, END)
            self.file_path_display.insert(INSERT, self.file_path.split('/')[-1])
            self.file_path_display.config(state=DISABLED)

    def start_calculate(self):
        if self.file_path == 'inital':
            messagebox.showwarning(title=' Warning', message='Please select the Excel-file first!')
        else:

            input_data = pd.read_excel(io=self.file_path, sheet_name='Calculate')
            output = pd.DataFrame(columns=('Ticket', 'Category_3', 'SLA', 'Start', 'End', 'Area'))
            for index, row in input_data.iterrows():
                cable_calculator = Calculator(row.Ticket, row.Cable_Type, row.Initial_Setup, row.Cable_Length, row.Additional_Time, row.Region)
                cable_calculator.init_setup_calculate()
                cable_calculator.length_calculate()
                cable_calculator.time_calculate()
                output = output.append(cable_calculator.df, sort=False)
            output.index = np.arange(1, len(output) + 1)
            input_file = os.path.dirname(self.file_path) + '/Cable_Ticketing_' + str(datetime.date.today()) + '.xlsx'
            try:
                output.to_excel(input_file, sheet_name='Ticketing', index_label='Index')
            except PermissionError:
                messagebox.showwarning(title='  Warning', message=' Excel file is opened ! please closed the file first!' )
            else:
                messagebox.showinfo(title='Prompt', message=' Calculation is completed !')
                time.sleep(1)
                self.master.destroy()

###################################################################################################

def main():
    root = Tk()
    interface = User_Interface(master=root)
    interface.mainloop()

if __name__ == '__main__':
    main()