import xlrd
import xlsxwriter
import numpy as np
import pandas as pd
from functools import reduce

def main():
    pd.set_option('display.max_rows', None) 
    pd.set_option('display.max_columns', None) 
    pd.set_option('display.max_colwidth', None)

    
    df_excel_file=pd.ExcelFile('Таблица автозакуп-для работы-СПБ ОСНОВНАЯ. КОПИИ НЕ ДЕЛАТЬ!.xlsx')
    df_excel=pd.read_excel(df_excel_file, sheet_name='Tillypad XL')
    df1_excel=pd.read_excel(df_excel_file, sheet_name='Поставщики', usecols=['Продукт', 'Поставщик'])


    df=pd.DataFrame(df_excel)    
    df1=pd.DataFrame(df1_excel)
    df['С учетом прироста','С учетом убывания']=' '#добавляем столбцы
    

    df_append=reduce(lambda x, y: pd.merge(x, y, on = ['Продукт']), [df, df1]) #объединяем таблицы, последовательно через reduce  
    
    #df_filter=df_append.loc[(df_append['Группа продуктов'] == 'Бар') & (df_append['Склад'] == 'ФК "Луначарского-80"')]# если выборку делать сразу
    

    def func_growth_decrease(period, kf_gr, kf_dec): #расчитываем с учетом прироста, либо убывания
        for i in range(len(df_append)):
            res=(df_append['Объем']/21)*period
            df_append['Объем']=res.round(2)
            if (kf_gr != 0) & (kf_gr != ' '):
                res1=(res*kf_gr+res).round(2)            
                df_append['С учетом прироста']=res1
            else:
                df_append['С учетом прироста']=' '
            if (kf_dec != 0) & (kf_dec != ' '):
                res1=(res-res*kf_dec).round(2)            
                df_append['С учетом убывания']=res1
            else:
                df_append['С учетом убывания']=' '
            i+=1        
            return df_append[['Продукт', 'Группа продуктов','Склад','Объем','С учетом прироста','С учетом убывания','Цена по себестоимости', 'Поставщик']]  
        
    try:
        a=float(input('Введите период, нажмите Enter: '))
        b=float(input('Введите коэффициент прироста (либо нажите 0), нажмите Enter: '))
        c=float(input('Введите коэффициент убывания (либо нажмите 0), нажмите Enter: '))
        res=func_growth_decrease(a,b,c)        
        with pd.ExcelWriter('file.xlsx') as writer:
            res.to_excel(writer)
    except ValueError:
        print('Вы ввели некорректные данные')        
    else:
        print('Ваш запрос обрабатывается, результат в файле xlsx')
    
    
main()  
