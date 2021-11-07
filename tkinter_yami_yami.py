import openpyxl
import numpy as np
import pandas as pd
from functools import reduce
import tkinter as tk
import xlsxwriter, xlrd
from tkinter.filedialog import askopenfilename, asksaveasfilename 
from tkinter import messagebox, ttk
    
    
  
class Tknew():
    def __init__(self, window):       
          self.window = window   
          self.window.title('Отчет')
          self.txt_edit = tk.Text(self.window)
          self.txt_edit.place(height=400, width=1000)

          self.file_frame=tk.LabelFrame(self.window)
          self.file_frame.place(height=270, width=1000, rely=0.62, relx=0)


          self.fr_buttons = tk.Frame(self.file_frame, background="seagreen1", relief=tk.RAISED)
          self.fr_buttons.place(height=270, width=500,relx=0)
          self.fr_buttons_second=tk.Frame(self.file_frame, background="seagreen2", relief=tk.RAISED)
          self.fr_buttons_second.place(height=270, width=500, relx=0.51)

          self.label = tk.Label(self.fr_buttons,text="Период:", bg='seagreen3')
          self.label1 = tk.Label(self.fr_buttons,text="Коэффициент прироста или '0':",bg='seagreen3')
          self.label2 = tk.Label(self.fr_buttons,text="Коэффициент убывания или '0':",bg='seagreen3')
          self.label_fb = tk.Label(self.fr_buttons_second,text="Склад, выберите из списка:",bg='seagreen3')
          self.label_product = tk.Label(self.fr_buttons_second,text="Продукт:",bg='seagreen3')
          self.label_grproduct = tk.Label(self.fr_buttons_second,text="Группа продуктов:",bg='seagreen3')
          self.label_provider = tk.Label(self.fr_buttons_second,text="Поставщик:",bg='seagreen3')



          self.label.place(rely=0.10, relx=0.31)
          self.label1.place(rely=0.25, relx=0.05)
          self.label2.place(rely=0.40, relx=0.05)
          self.label_fb.place(rely=0.10, relx=0.05)
          self.label_product.place(rely=0.25, relx=0.25)
          self.label_grproduct.place(rely=0.40, relx=0.15)
          self.label_provider.place(rely=0.55, relx=0.22)



          self.entry_period = tk.Entry(self.fr_buttons)
          self.entry_period.insert(0, "0")
          self.entry_growth = tk.Entry(self.fr_buttons)
          self.entry_growth.insert(0, "0.0")
          self.entry_decrease = tk.Entry(self.fr_buttons)
          self.entry_decrease.insert(0, "0.0")
         
          self.choose_fb=ttk.Combobox(self.fr_buttons_second, value=('ФК "Беговая-3"','ФК "Белорусская-6"','ФК "Белы Куна-16"','ФК "Богатырский-10"','ФК "В.О.Большой пр-кт-83"','ФК "Гагарина-42"',
                                                                     'ФК "Глухая Зеленина-6"','ФК "Зайцева-41"','ФК "Кораблестроителей-30"','ФК "Купчинская-1"','ФК "Лиговский-289" ( ООО МОРЕ)', 
                                                                     'ФК "Луначарского-64"','ФК "Луначарского-80"','ФК "Маршала Блюхера-9"','ФК "Науки-25" (УМЕЛЫЕ РУКИ)','ФК "Печатников-21"',
                                                                     'ФК "Подвойского-34"','ФК "10-ая Советская-15 27"','ФК "Тамбасова-32"  Вахрамеев','ФК "Туристская-18"','ФК "Уральская-6"',
                                                                    'ФК "Шуваловский-37"(ООО Шуваловский)','ФК "Шоссе Революции-31"' ), takefocus=0)          
          self.entry_product=tk.Entry(self.fr_buttons_second)                                                                    
          self.entry_grproduct=tk.Entry(self.fr_buttons_second)
          self.entry_provider=tk.Entry(self.fr_buttons_second)

          self.entry_period.place(rely=0.10, relx=0.45)
          self.entry_growth.place(rely=0.25, relx=0.45)
          self.entry_decrease.place(rely=0.40, relx=0.45)
          self.choose_fb.place(rely=0.10, relx=0.40)
          self.entry_product.place(rely=0.25, relx=0.40)
          self.entry_grproduct.place(rely=0.40, relx=0.40)
          self.entry_provider.place(rely=0.55, relx=0.40)
       
        
          self.tv=ttk.Treeview(self.txt_edit)        
          self.tv.place(relheight=1, relwidth=1)

          self.treescrolly = ttk.Scrollbar(self.txt_edit, orient="vertical", command=self.tv.yview) 
          self.treescrollx = ttk.Scrollbar(self.txt_edit, orient="horizontal", command=self.tv.xview)
          self.tv.configure(xscrollcommand=self.treescrollx.set, yscrollcommand=self.treescrolly.set)
          self.treescrollx.pack(side="bottom", fill="x") 
          self.treescrolly.pack(side="right", fill="y")

          self.btn_growth_decrease=tk.Button(self.fr_buttons, text="Рассчитать и вывести результат",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1',cursor="arrow", command=self.result)
          self.btn_filter = tk.Button(self.fr_buttons_second, text="Отфильтровать",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1', cursor="arrow", command=self.search)
          self.btn_save = tk.Button(self.fr_buttons, text="Сохранить файл как...",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1',cursor="arrow", command=self.save_file)
          self.btn_open = tk.Button(self.fr_buttons, text="Открыть сохраненный файл",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1',cursor="arrow", command=self.open_file)
          self.btn_delete=tk.Button(self.fr_buttons, text="Отменить",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1',cursor="arrow", command=self.clear_data)          
          self.btn_delete2=tk.Button(self.fr_buttons_second, text="Отменить",bg="black",fg="gold",activebackground='gray25',activeforeground='seagreen1',cursor="arrow", command=self.clear_data_entry)
          self.btn_delete.place(rely=0.55, relx=0.80)
          self.btn_delete2.place(rely=0.70, relx=0.75)
          self.btn_growth_decrease.place(rely=0.55, relx=0.25)
          self.btn_save.place(rely=0.70, relx=0.36)
          self.btn_open.place(rely=0.70, relx=0.65)
          self.btn_filter.place(rely=0.70, relx=0.45)
          
      
       
            
    
class BD(Tknew):
    def __init__(self):
        super().__init__(root)
        pd.set_option('display.max_rows', None) 
        pd.set_option('display.max_columns', None) 
        pd.set_option('display.max_colwidth', None)

        df_excel_file=pd.ExcelFile('Таблица автозакуп-для работы-СПБ ОСНОВНАЯ. КОПИИ НЕ ДЕЛАТЬ!.xlsx')
        df_excel=pd.read_excel(df_excel_file, sheet_name='Tillypad XL')
        df1_excel=pd.read_excel(df_excel_file, sheet_name='Поставщики', usecols=['Продукт', 'Поставщик'])
        self.df_res=pd.DataFrame(df_excel)    
        df1=pd.DataFrame(df1_excel)
        self.df_res['С учетом прироста','С учетом убывания']=' '#добавляем столбцы  

        self.df_res=reduce(lambda x, y: pd.merge(x, y, on = ['Продукт']), [self.df_res, df1]) #объединяем таблицы, последовательно через reduce
            
    def result(self):
        try:
           a=self.entry_period.get()
           b=self.entry_growth.get()
           c=self.entry_decrease.get()
           a=float(a)
           b=float(b)
           c=float(c)
           self.func_growth_decrease(a,b,c)
        except ValueError:
            tk.messagebox.showerror("Информация",'Вы ввели некорректные данные')   
            return None
        

    def func_growth_decrease(self,period, kf_gr, kf_dec): #расчитываем с учетом прироста, либо убывания 
           
         for i in range(len(self.df_res)):
                    res=(self.df_res['Объем']/21)*period
                    self.df_res['Объем']=res.round(2)
                    if (kf_gr != 0) & (kf_gr != ' '):
                        res1=(res*kf_gr+res).round(2)            
                        self.df_res['С учетом прироста']=res1
                    else:
                        self.df_res['С учетом прироста']=' '
                    if (kf_dec != 0) & (kf_dec != ' '):
                        res1=(res-res*kf_dec).round(2)            
                        self.df_res['С учетом убывания']=res1
                    else:
                        self.df_res['С учетом убывания']=' '
                    i+=1        
                    self.res=self.df_res[['Продукт', 'Группа продуктов','Склад','Объем','С учетом прироста','С учетом убывания','Цена по себестоимости', 'Поставщик']]  
                    self.clear_data()
                    self.tv["column"] = list(self.res.columns)
                    self.tv["show"] = "headings"
                    for column in self.tv["columns"]:
                             self.tv.heading(column, text=column) 
                             self.tv.column(column,stretch=tk.NO, minwidth=120, width=200)
                    bd_rows = self.res.to_numpy().tolist() 
                    for row in bd_rows:
                             self.tv.insert("", "end", values=row) 
                    return None    
         

    def open_file(self):   
            filepath = askopenfilename(
                filetypes=[("Excel файлы","*.xlsx"), ("Все файлы", "*.*")]
            )
            if not filepath:
                return            
            try:
                excel_filename = r"{}".format(filepath)
                if excel_filename[-4:] == ".csv":
                    bd = pd.read_csv(excel_filename)
                else:
                    bd = pd.read_excel(excel_filename)
                self.clear_data()
                self.read_file(bd)
            except ValueError:
                tk.messagebox.showerror("Информация", "Файл не может быть открыт")
                return None
            except FileNotFoundError:
                tk.messagebox.showerror("Информация", f"Файл не найден")
                return None 
            

    def read_file(self,lst):
            self.tv["column"] = list(lst.columns)
            self.tv["show"] = "headings"
            for column in self.tv["columns"]:
                     self.tv.heading(column, text=column) 
                     self.tv.column(column,stretch=tk.NO, minwidth=120, width=200)
            bd_rows = lst.to_numpy().tolist() 
            for row in bd_rows:
                     self.tv.insert("", "end", values=row) 
            return None 



    def clear_data(self):
            self.tv.delete(*self.tv.get_children())            
            return None

    def clear_data_entry(self):
          self.tv.delete(*self.tv.get_children())           
          self.choose_fb.delete('0',tk.END)      
          self.entry_product.delete('0',tk.END)
          self.entry_grproduct.delete('0',tk.END)
          self.entry_provider.delete('0',tk.END)
          self.result()
          return None

    def save_file(self): 
           filepath = asksaveasfilename(
               defaultextension="xlsx",
               filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
                )
           if not filepath:
               return
           try:
               writer = pd.ExcelWriter(filepath) 
               data=[]             
               for Parent in self.tv.get_children():
                    data.append(self.tv.item(Parent)["values"])
                    for child in self.tv.get_children(Parent):
                            data.append(self.tv.item(child)["values"])              
               data_df=pd.DataFrame(data)               
               data_df.columns=self.tv['column']               
               data_df.to_excel(writer,'Отчет')
               writer.save()
               tk.messagebox.showinfo("Информация", "Файл сохранен")
               return None
           except ValueError:
                tk.messagebox.showerror("Информация", "Неверный формат файла")
                return None
           except:
                tk.messagebox.showerror("Информация", "Файл невозможно сохранить")
                return None

    def search(self):     
             prod=self.entry_product.get()
             grprod=self.entry_grproduct.get()
             provider=self.entry_provider.get()
             choose=self.choose_fb.get()
         
             self.selections=[]                             
             for child in self.tv.get_children():            
                     if (provider == "" or provider.lower() in self.tv.item(child)['values'][7].lower()) and (choose == "" or choose in self.tv.item(child)['values'][2]) and (prod == "" or prod.lower() in self.tv.item(child)['values'][0].lower()) and (grprod == "" or grprod.lower() in self.tv.item(child)['values'][1].lower()):
                          self.selections.append(self.tv.item(child)['values'][0:])   
             self.bd=pd.DataFrame(self.selections, columns = ['Продукт', 'Группа продуктов','Склад','Объем','С учетом прироста','С учетом убывания','Цена по себестоимости', 'Поставщик'])
             self.clear_data()
             self.tv["column"] = list(self.bd.columns)
             self.tv["show"] = "headings"
             for column in self.tv["columns"]:
                     self.tv.heading(column, text=column) 
                     self.tv.column(column,stretch=tk.NO, minwidth=120, width=200)
             bd_rows = self.bd.to_numpy().tolist() 
             for row in bd_rows:
                     self.tv.insert("", "end", values=row) 
             return None 

         


                
if __name__ == "__main__":
    root = tk.Tk()     
    root.geometry("1000x650")
    root.pack_propagate(False)
    root.resizable(0,0)  
  
    app = BD()
    root.mainloop()
 

