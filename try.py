# Python3 code to select
# data from excel
import xlwings as xw
from datetime import date
from calendar import monthrange
from datetime import datetime
import openpyxl


class Reader:

    def __init__(self):
        self.campains_dic = []
        self.ws = xw.Book(self.get_file()).sheets['Sheet0']
        self.campains, self.lines = self.get_campains()
        self.budgets = self.get_values('E')
        self.prices = self.get_values('L')
        self.lids = self.get_values('K')
        self.campains_creator()

    def get_file(self):
        today = date.today().strftime("%d-%m-%Y")
        return f"C:\\Users\\tamir\Downloads\\{today}doh.xlsx"

    def get_campains(self):
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
        ret = []
        for line in self.lines:
            ret.append(self.ws.range(f"{collum}{line}").value)
        return ret

    def get_budgets(self, daily_bud):
        year = int(datetime.now().strftime('%y'))
        month = int(datetime.now().strftime('%m'))
        num_days = monthrange(year, month)[1]  # num_days = 28

        monthly_bud = num_days * daily_bud
        weekly_bud = 7 * daily_bud
        return monthly_bud, weekly_bud

    def campains_creator(self):
        ind = 0
        print(self.lids)
        for campain in self.campains:
            monthly_bud, weekly_bud = self.get_budgets(self.budgets[ind])
            bud_usage = str(int((weekly_bud / self.prices[ind]) * 100)) + "%"
            price_for_lid = '--'
            if self.lids[ind] != 0:
                price_for_lid = self.prices[ind] / self.lids[ind]
            self.campains_dic.append((campain, monthly_bud, weekly_bud, bud_usage, self.budgets[ind], self.prices[ind], self.lids[ind], price_for_lid))
            ind += 1


def get_budgets(daily_bud):
    year = int(datetime.now().strftime('%y'))
    month = int(datetime.now().strftime('%m'))
    num_days = monthrange(year, month)[1]  # num_days = 28

    monthly_bud = num_days * daily_bud
    weekly_bud = 7 * daily_bud
    return monthly_bud, weekly_bud


def writer(lst):
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
    wb.save('C:\lemesh\output.xlsx')


if __name__ == '__main__':

    r = Reader()
    lines = r.campains_dic
    writer(lines)




