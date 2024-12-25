import xlwings as xw
import pandas as pd 
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.widgets import Cursor

import datetime

def main():
    wb = xw.Book.caller()

    select_range=wb.selection
    sht=select_range.sheet
    rawdata=select_range.value
    rawdata1=pd.DataFrame(rawdata,columns=["time","data"])
    datetime_convert=pd.to_datetime(rawdata1["time"])
    rawdata1["time"]=datetime_convert
    
    singleDataList=[]
    x_DateList=[]
    yearMark=rawdata1["time"][0].year
    fig, ax = plt.subplots()


    for i in rawdata1[:].itertuples():
        i_datetime=i.time.to_pydatetime()
        timestruct=i_datetime.timetuple()
        t_year=timestruct.tm_year
        if t_year==yearMark:
        

            singleDataList.append(i.data)
            x_DateList.append(i_datetime.replace(year=1972))
        else:
            ax.plot(x_DateList,singleDataList,label=yearMark)
            x_DateList=[]
            singleDataList=[]
            singleDataList.append(i.data)
            x_DateList.append(i_datetime.replace(year=1972))
            yearMark=timestruct.tm_year

    if x_DateList!=[]:
        ax.plot(x_DateList,singleDataList,label=yearMark)
    date_formatter = mdates.DateFormatter('%m-%d') 
    ax.xaxis.set_major_formatter(date_formatter)
    start_date = datetime.datetime(1972, 1, 1)
    end_date = datetime.datetime(1972, 12, 31)
    ax.set_xlim(start_date, end_date)

    ax.grid()
    ax.legend()
    fig.autofmt_xdate()
    cursor = Cursor(ax, useblit=True, color='red', linewidth=2)


#1 cursor 十字线的问题
#2 如何将X轴转换为时间并且要考虑 闰年的问题
    sht.pictures.add(fig, name='MyPlot', update=True,left=sht.range('B5').left, top=sht.range('B5').top)
    #plt.show()
if __name__ == "__main__":
    xw.Book("tmp1.xlsm").set_mock_caller()
    main()
