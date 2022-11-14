import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


class SAPTask:
    def __init__(self):
        self.path = 'data/58735501.xlsx'
        self.resultPath = 'data/58735500.xlsx'
        self.book = pd.ExcelFile(self.path)
        self.sheet = self.book.parse('исх данные')
        #self.firstTask()
        self.secondTask()

    @staticmethod
    def cm_to_inch(self, cm):
        return cm/2.54

    @staticmethod
    def plotbars(self,x, y):
        minx = np.min(x)
        miny = np.min(y)
        maxx = np.max(x)
        maxy = np.max(y)
        plt.figure(figsize=(self.cm_to_inch(50), self.cm_to_inch(50)))
        fig, ax = plt.subplots()
        ax.bar(x, y, width=1, edgecolor="white", linewidth=0.7)
        ax.set(xlim=(minx-1, maxx+1), xticks=np.arange(minx-1, maxx+1),
               ylim=(miny-1, maxy+1), yticks=np.arange(miny-1, maxy+1))
        plt.show()

    @staticmethod
    def plotpies(self,name, labels, sizes):

        explode = (0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        plt.title(name)
        plt.show()

    def firstTask(self):
        date_column = self.sheet['Дата'].to_numpy()
        unique, counts = np.unique(date_column, return_counts=True)
        self.plotbars(pd.DatetimeIndex(unique).day, counts)
        ckg = self.sheet['ЦКГ'].to_numpy().astype('str')
        deliveryPrice = self.sheet['Сумма за доставку'].to_numpy()
        unique, total = np.unique(ckg, return_counts=True)
        unique = dict(zip(unique, total))
        counter = 0
        for i in unique.keys():
            unique[i] = [unique[i],0]
        for price, group in zip(deliveryPrice, ckg):
            unique[group][1] += 1 if price == 0 else 0
        for i in unique.keys():
            free = unique[i][1]*100/unique[i][0]
            notfree = (unique[i][0] - unique[i][1])*100/unique[i][0]
            sizes = [free, notfree]
            labels = ['Бесплатно', 'Платно']
            self.plotpies(i,labels, sizes)

    def secondTask(self):
        dataA = {'Грузополучатель':[], 'ЦКГ':[], 'N заявки':[], 'Менеджер ЭТМ':[]}
        dataB = {'Грузополучатель': [], 'ЦКГ': [], 'N заявки': [], 'Менеджер ЭТМ': []}
        dataC = {'Грузополучатель': [], 'ЦКГ': [], 'N заявки': [], 'Сумма за товар': []}
        dataA = self.secondTaskA(dataA)
        dataB = self.secondTaskB(dataB)
        dataC = self.secondTaskС(dataC)
        with pd.ExcelWriter(self.resultPath) as writer:
            dataA.to_excel(writer, sheet_name="2a")
            dataB.to_excel(writer, sheet_name="2b")
            dataC.to_excel(writer, sheet_name="2c")


    def secondTaskA(self, data):
        delivery = self.sheet['Вид перевозки'].to_numpy()
        vip = self.sheet['VIP признак'].to_numpy().astype('str')
        for i, item in enumerate(delivery):
            if item.find('VIP')+1:
                if vip[i] == 'nan':
                    for j in data.keys():
                        data[j].append(self.sheet[j][i])
        return pd.DataFrame(data)

    def secondTaskB(self, data):
        violation = self.sheet['Нарушение  барьерной цены'].to_numpy().astype('str')
        for i, item in enumerate(violation):
            if item == 'да':
                for j in data.keys():
                    data[j].append(self.sheet[j][i])
        return pd.DataFrame(data)

    def secondTaskС(self, data):
        delivery = self.sheet['Вид перевозки'].to_numpy().astype('str')
        for i, item in enumerate(delivery):
            if item.find('Бесплатная')+1:
                if self.sheet['Сумма за товар'][i]>50000:
                    for j in data.keys():
                        data[j].append(self.sheet[j][i])
        return pd.DataFrame(data)
if __name__ == "__main__":
    SAPTask = SAPTask()

