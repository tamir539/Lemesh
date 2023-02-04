# Python3 code to select
# data from excel
import xlwings as xw
from datetime import date
from calendar import monthrange
from datetime import datetime


class Reader:

    def __init__(self):
        self.campains_dic = {}
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
            if self.ws.range(f"L{line}").value != 0:
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

    def campains_creator(self):
        ind = 0
        for campain in self.campains:
            self.campains_dic[campain] = [self.budgets[ind], self.prices[ind], self.lids[ind]]
            ind += 1
        print(self.campains_dic)


def add_campain(daily_bud, price, lids):
    today = date.today().strftime("%d-%m-%Y")
    monthly_bud, weekly_bud = get_budgets(daily_bud)
    budgets_usage = weekly_bud/price
    price_for_lid = price / lids
    with open(f"compare{today}", 'w+') as file:
        pass


def get_budgets(daily_bud):
    year = int(datetime.now().strftime('%y'))
    month = int(datetime.now().strftime('%m'))
    num_days = monthrange(year, month)[1]  # num_days = 28

    monthly_bud = num_days * daily_bud
    weekly_bud = 7 * daily_bud
    return monthly_bud, weekly_bud

def formater():
    pass

if __name__ == '__main__':

    Reader()




