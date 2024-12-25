import xlwings as xw
import xlwings as xw
import pandas as pd 
import matplotlib.pyplot as plt
from matplotlib.widgets import Cursor

def main():
    # 暂时性的可用
    #wb = xw.Book.caller()
    # 在单独的连续单元格指定event 时间
    wb= xw.books.active
    
    select_range=wb.selection
    #
    sht=select_range.sheet

    #指定数据行
    rawdata=sht.range("E4:F772").value
    #
    rawdata1=pd.DataFrame(rawdata,columns=["time","data"])

    back_days=-100
    forward_days=100

    for c in select_range:


        zero_index=rawdata1[rawdata1["time"]==c.value].index[0]

        singleDataList=[]
        x_DateList=[]
        
        for i in range(back_days,forward_days,1):
            x_DateList.append(i)
            #singleDataList.append(rawdata1["data"].iloc[zero_index+i])
            #这里使用百分比幅度来界定
            singleDataList.append((rawdata1["data"].iloc[zero_index+i]/rawdata1["data"].iloc[zero_index]-1)*100)
        plt.plot(x_DateList,singleDataList,label=rawdata1["time"].iloc[zero_index])


    plt.legend()
    plt.show()

if __name__ == "__main__":
    #xw.Book("test.xlsm").set_mock_caller()
    main()
