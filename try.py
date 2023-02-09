import xlwings as xw
from datetime import date
from calendar import monthrange
from datetime import datetime
import openpyxl
import os
import glob


class Reader:

    def __init__(self):
        self.structure = []     # the final data structure to be written in the compare gile
        self.ws = xw.Book(self.get_file()).sheets['Sheet0']     # connect to the current report
        self.campains, self.lines = self.get_campains()
        self.budgets = self.get_values('E')
        self.prices = self.get_values('L')
        self.lids = self.get_values('K')
        self.last_price_per_lid = {}    # price per lid from the last comparation file
        self.get_last_compare_file()
        self.formatter()

    def get_file(self):
        '''

        :return: the path to the current report
        '''
        today = date.today().strftime("%d-%m-%Y")
        return f"C:\\Users\\tamir\Downloads\\{today}doh.xlsx"

    def get_last_compare_file(self):
        '''

        :return: find the las compare filr, deliver and call to get_last_price_for_lid
        '''
        list_of_files = glob.glob(os.getcwd() + "\compares\\*")  # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)  #the latest compare file
        ws = xw.Book(latest_file).sheets['Sheet']
        self.get_last_price_for_lid(ws)

    def get_last_price_for_lid(self, ws):
        '''

        :param ws: worksheet of the last compare file
        :return: update the "last_price_per_lid" var from the last compare file
        '''
        campain = ws.range(f"A2").value
        line = 2
        while campain:
            self.last_price_per_lid[campain] = ws.range(f"H{line}").value
            line += 1
            campain = ws.range(f"A{line}").value

    def get_campains(self):
        '''

        :return: all the relevant campains from the report, their lines
        '''
        campains = []
        campain = self.ws.range(f"B4").value
        lines = []
        if self.ws.range(f"L4").value != 0:
            campains.append(campain)
        line = 5
        while campain:
            if self.ws.range(f"J{line}").value != 0:
                campain = self.ws.range(f"B{line}").value
                campains.append(campain)
                lines.append(line)
            line += 1
        return campains[:-1], lines[:-1]

    def get_values(self, collum):
        '''

        :param collum: collum on the workSheet
        :return: all the values of (line,collum)
        '''
        ret = []
        for line in self.lines:
            ret.append(self.ws.range(f"{collum}{line}").value)
        return ret

    def get_budgets(self, daily_bud):
        '''

        :param daily_bud: daily budget of campain
        :return: the monthly, weekly budget
        '''
        year = int(datetime.now().strftime('%y'))
        month = int(datetime.now().strftime('%m'))
        num_days = monthrange(year, month)[1]  # num_days = 28

        monthly_bud = num_days * daily_bud
        weekly_bud = 7 * daily_bud
        return monthly_bud, weekly_bud

    def formatter(self):
        '''

        :return: formatt all the data to tuple that fit the new file
        '''
        ind = 0
        for campain in self.campains:
            monthly_bud, weekly_bud = self.get_budgets(self.budgets[ind])   # get the relevamt budgets
            bud_usage = str(int((weekly_bud / self.prices[ind]) * 100)) + "%"   # clculate the budget usage
            price_for_lid = '--'
            last_price_for_lid = self.last_price_per_lid[campain]
            price_for_lid_difference = ""
            if self.lids[ind] != 0:
                price_for_lid = self.prices[ind] / self.lids[ind]   # calculate the price for lid
            if price_for_lid != "--" and last_price_for_lid:
                price_for_lid_difference = str(int((price_for_lid / last_price_for_lid - 1) * 100)) + "%"   # calculate the price for lid difference
            self.structure.append((campain, monthly_bud, weekly_bud, bud_usage, self.budgets[ind], self.prices[ind], self.lids[ind], price_for_lid, last_price_for_lid, price_for_lid_difference))
            ind += 1


def writer(lst):
    '''

    :param lst: list of the data srtucture
    :return: create the compare file and write all the data to it
    '''
    today = date.today().strftime("%d-%m-%Y")

    headers = ("קמפיין", "תקציב חודשי", "תקציב לתקופה", "ניצול תקציב", "תקציב יומי", "עלות", "תוצאה", "עלות לתוצאה",
               "עלות קודמת לתוצאה", "שינוי")
    wb = openpyxl.Workbook()
    ws = wb.active
    data = [headers]

    for tup in lst:
        data.append(tup)
    data = tuple(data)

    for i in data:
        ws.append(i)
    wb.save(f'{os.getcwd()}\compares\{today}compare.xlsx')


if __name__ == '__main__':

    r = Reader()
    lines = r.structure
    writer(lines)




