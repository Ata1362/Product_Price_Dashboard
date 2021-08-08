import pandas as pd
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

def cal():
    entries = read_entries()

    '''1.025 is the safety coefficent for Sell and buy of USD dollor'''
    total_bulk_costs = (entries.Forwarding + (entries.OriginPortToDestinationPort * (
                entries.ExchangeRate * (1 + (entries.ExchangeRateSafety / 100)))) + entries.FactoryToPort +
                        entries.FarmToFactory)

    total_packing_costs = entries.PackingCost * entries.GrossWeight

    total_costs = total_bulk_costs + total_packing_costs

    unit_cost = (total_costs * (1 + (entries.CostSafetyRate / 100))) / entries.GrossWeight

    raw_price = entries.RawPrice + entries.FactoryMargin

    final_price = raw_price + unit_cost

    ir_price = int(round(final_price))

    '''Current code will not accept a float entry for the Exchange Rate Safety due to the code "int" '''
    usd_price = round((ir_price * (1 + (int(entries.ExchangeRateSafety) / 100))) / float(entries.ExchangeRate), 4)

    rial = tk.Label(win, text=" قیمت ریالی  :", font=('Arial', 20))
    rial.grid(row=31, column=0)
    text = tk.Label(win, text=str(f"{ir_price:,}"), font=('Helvetica bold', 20))
    text.grid(row=31, column=1)

    dollar = tk.Label(win, text="USD Price: ", font=('times new roman', 20))
    dollar.grid(row=32, column=0)
    text = tk.Label(win, text=str(round(usd_price, 4)), font=('Times New Roman', 20))
    text.grid(row=32, column=1)

    entries.IR_Price = [ir_price]
    entries.USD_Price = [usd_price]
    print(entries)
    '''Here we are loading the database, add the new recorde and save the database again'''

    recording = pd.read_excel('Records.xls',
                              sheet_name="records",
                              header=0,
                              index_col=False)

    new = pd.concat([recording, entries])

    new.to_excel('Records.xls', sheet_name="records", index=False)

    # text.config(font=('Helvetica bold',40))
    return


def read_entries():
    quoted_party_input = quoted_party.get()
    quoted_party_input = quoted_party_input.strip().upper()

    product_type_input = product_type.get()
    product_type_input = product_type_input.strip().upper()

    product_subtype_input = product_subtype.get()
    product_subtype_input = product_subtype_input.strip().upper()

    product_size_input = product_size.get()
    # product_size_input= product_size_input.strip().upper()

    prodcut_price_input = prodcut_price.get()
    # prodcut_price_input= prodcut_price_input.strip().upper()

    gross_weight_input = gross_weight.get()
    # gross_weight_input= gross_weight_input.strip().upper()

    Destination_Port_input = Destination_Port.get()
    Destination_Port_input = Destination_Port_input.strip().upper()

    Farm_to_Factory_input = Farm_to_Factory.get()
    # Farm_to_Factory_input= Farm_to_Factory_input.strip().upper()

    Factory_to_Port_input = Factory_to_Port.get()
    # Factory_to_Port_input= Factory_to_Port_input.strip().upper()

    Origin_Port_to_Destination_Port_input = Origin_Port_to_Destination_Port.get()
    # Origin_Port_to_Destination_Port_input= Origin_Port_to_Destination_Port_input.strip().upper()

    Factory_Margin_input = Factory_Margin.get()
    # Factory_Margin_input= Factory_Margin_input.strip().upper()

    Forwarding_input = Forwarding.get()
    # Forwarding_input= Forwarding_input.strip().upper()

    Safety_on_Costs_input = Safety_on_Costs.get()
    # Safety_on_Costs_input= Safety_on_Costs_input.strip().upper()

    packing_cost_input = packing_cost.get()
    # packing_cost_input= packing_cost_input.strip().upper()

    exchange_rate_input = exchange_rate.get()
    # exchange_rate_input= exchange_rate_input.strip().upper()

    exchange_rate_safety_input = exchange_rate_safety.get()

    Date = datetime.today().strftime('%Y-%m-%d-%H:%M:%S')

    data_entry = pd.DataFrame(
        {'Date': Date, 'Company': quoted_party_input, 'Type': product_type_input, 'Item': product_subtype_input,
         'Size': product_size_input,
         'RawPrice': float(prodcut_price_input), 'PackingCost': float(packing_cost_input),
         'GrossWeight': float(gross_weight_input), 'FarmToFactory': float(Farm_to_Factory_input),
         'FactoryToPort': float(Factory_to_Port_input),
         'OriginPortToDestinationPort': float(Origin_Port_to_Destination_Port_input),
         'DestinationPortName': Destination_Port_input, 'FactoryMargin': float(Factory_Margin_input),
         'Forwarding': float(Forwarding_input),
         'CostSafetyRate': float(Safety_on_Costs_input), 'ExchangeRate': float(exchange_rate_input),
         'ExchangeRateSafety': float(exchange_rate_safety_input),
         "IR_Price": 0, "USD_Price": 0}, index=[0])

    return data_entry


def refresh():
    entries = read_entries()

    spacer1 = tk.Label(win, text=str(f"{int(entries.RawPrice):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=6, column=1)

    spacer1 = tk.Label(win, text=str(f"{int(entries.GrossWeight):,}" + " کیلوگرم "), font=('Helvetica bold', 11))
    spacer1.grid(row=6, column=3)

    spacer1 = tk.Label(win, text=str(f"{int(entries.FarmToFactory):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=8, column=3)

    spacer1 = tk.Label(win, text=str(f"{int(entries.FactoryToPort):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=10, column=1)

    spacer1 = tk.Label(win, text=str(f"{int(entries.OriginPortToDestinationPort):,}" + " دلار "),
                       font=('Helvetica bold', 11))
    spacer1.grid(row=10, column=3)

    spacer1 = tk.Label(win, text=str(f"{int(entries.FactoryMargin):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=12, column=1)

    spacer1 = tk.Label(win, text=str(f"{int(entries.Forwarding):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=12, column=3)

    spacer1 = tk.Label(win, text=str(f"{int(entries.CostSafetyRate):,}" + " درصد "), font=('Helvetica bold', 11))
    spacer1.grid(row=14, column=1)

    spacer1 = tk.Label(win, text=str(f"{int(entries.PackingCost):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=14, column=3)

    spacer1 = tk.Label(win, text=str(f"{int(entries.ExchangeRate):,}" + " ریال "), font=('Helvetica bold', 11))
    spacer1.grid(row=16, column=1)

    spacer1 = tk.Label(win, text=str(f"{int(entries.ExchangeRateSafety):,}" + " درصد "), font=('Helvetica bold', 11))
    spacer1.grid(row=16, column=3)

    return




win = tk.Tk()
win.geometry("800x600")
win.title("TAHER PRICE CALCULATION")

'''This function will load the Excel dataset in this function we try to use only Pandas to manipulate our data'''
recording = pd.DataFrame()
address = 'Records.xls'
'''Specifying the headers and index column actually will help alot to search the file easier.'''
recording = pd.read_excel(address,
                         sheet_name="records",
                         header=0,
                         index_col=False)

'''To have a dynamic Options Menu we can read the available options all the time and create a better experience'''
options = pd.read_excel('prices.xls',
                         sheet_name="ref",
                         header=0,
                         index_col= False)

''' The Menu lists for different optins. '''
type_list = options.Type.drop_duplicates().to_list()
subtype_list = options.Item.drop_duplicates().to_list()
size_list = options.Size.drop_duplicates().to_list()

''' Create all objects needed in the main window. '''
message = "لطفا اطلاعات را با توجه به نوع داده در محل مناسب وارد نمائید"
header_message = tk.Label(win, text=message)

label1 = tk.Label(win, text="نام شرکت درخواست کننده")
quoted_party = tk.Entry(win)
quoted_party.insert(0, "Lambo")


label2 = tk.Label(win, text="محصول مورد نظر")
product_type = tk.StringVar(win)
product_type.set("انتخاب کنید") # default value
'''Adding the "*" was a key solution to make the list vertical for the OptionMenu'''
product_type_menu = tk.OptionMenu(win, product_type, *type_list)


label3 = tk.Label(win, text="زیر مجموعه محصول")
product_subtype = tk.StringVar(win)
product_subtype.set("انتخاب کنید") # default value
'''Adding the "*" was a key solution to make the list vertical for the OptionMenu'''
product_subtype_menu = tk.OptionMenu(win, product_subtype, *subtype_list)


label4 = tk.Label(win, text="سایز محصول مورد درخواست")
product_size = tk.StringVar(win)
product_size.set("انتخاب کنید") # default value
'''Adding the "*" was a key solution to make the list vertical for the OptionMenu'''
product_size_menu = tk.OptionMenu(win, product_size, *size_list)


label5 = tk.Label(win, text="قیمت ماده خام")
prodcut_price = tk.Entry(win)
prodcut_price.insert(0, "1900000")

label6 = tk.Label(win, text="وزن کل بار درخواستی")
gross_weight = tk.Entry(win)
gross_weight.insert(0, "12000")

label7 = tk.Label(win, text='نام بندر مقصد')
Destination_Port = tk.Entry(win)
Destination_Port.insert(0, "Port Klang")

label8 = tk.Label(win, text="هزینه ارسال به کارخانه")
Farm_to_Factory = tk.Entry(win)
Farm_to_Factory.insert(0, "60000000")

label9 = tk.Label(win, text="هزینه ارسال از کارخانه به بندرعباس")
Factory_to_Port = tk.Entry(win)
Factory_to_Port.insert(0, "120000000")

label10 = tk.Label(win, text="هزینه دلاری ارسال از بندر به مقصد")
Origin_Port_to_Destination_Port = tk.Entry(win)
Origin_Port_to_Destination_Port.insert(0, "2700")

label11 = tk.Label(win, text="سود کارخانه")
Factory_Margin = tk.Entry(win)
Factory_Margin.insert(0, "50000")

label12 = tk.Label(win, text="هزینه تیم گمرک")
Forwarding = tk.Entry(win)
Forwarding.insert(0, "68000000")

label13 = tk.Label(win, text="ضریب حاشیه امنیت هزینه ها %")
Safety_on_Costs = tk.Entry(win)
Safety_on_Costs.insert(0, "25")

label14 = tk.Label(win, text = "هزینه بسته بندی برای هر کیلو")
packing_cost = tk.Entry(win)
packing_cost.insert(0, "12000")


label15 = tk.Label(win, text="نرخ تبدیل دلار به ریال ")
exchange_rate = tk.Entry(win)
exchange_rate.insert(0, "245000")

label16 = tk.Label(win, text= 'ضریب حمایتی قیمت دلار')
exchange_rate_safety = tk.Entry(win)
exchange_rate_safety.insert(0, '2')


'''Positioning all obejcts  and enquiry lables.'''

header_message.grid(row=0, column=2)


label1.grid(row=1, column=0, sticky=tk.E)
quoted_party.grid(row=1, column=1)

label2.grid(row=1, column=2, sticky=tk.E)
product_type_menu.grid(row=1, column=3, sticky="ew")


spacer1 = tk.Label(win, text="")
spacer1.grid(row=2, column=0)




label3.grid(row=3, column=0, sticky=tk.E)
product_subtype_menu.grid(row=3, column=1, sticky="ew")

label4.grid(row=3, column=2, sticky=tk.E)
product_size_menu.grid(row=3, column=3, sticky="ew")


spacer1 = tk.Label(win, text="")
spacer1.grid(row=4, column=0)




label5.grid(row=5, column=0, sticky=tk.E)
prodcut_price.grid(row=5, column=1)

label6.grid(row=5, column=2, sticky=tk.E)
gross_weight.grid(row=5, column=3)


spacer1 = tk.Label(win, text="")
spacer1.grid(row=6, column=1)




label7.grid(row= 7, column=0, sticky=tk.E)
Destination_Port.grid(row= 7, column=1)

label8.grid(row=7, column=2, sticky=tk.E)
Farm_to_Factory.grid(row=7, column=3)

spacer1 = tk.Label(win, text="")
spacer1.grid(row=8, column=1)




label9.grid(row=9, column=0, sticky=tk.E)
Factory_to_Port.grid(row=9, column=1)

label10.grid(row=9, column=2, sticky=tk.E)
Origin_Port_to_Destination_Port.grid(row=9, column=3)

spacer1 = tk.Label(win, text="")
spacer1.grid(row=10, column=0)




label11.grid(row = 11, column= 0, sticky= tk.E)
Factory_Margin.grid(row= 11, column=1)

label12.grid(row = 11, column= 2, sticky = tk.E)
Forwarding.grid(row= 11, column= 3)


spacer1 = tk.Label(win, text="")
spacer1.grid(row=12, column=1)




label13.grid(row= 13, column= 0, sticky= tk.E)
Safety_on_Costs.grid(row= 13, column= 1)

label14.grid(row= 13, column= 2, sticky= tk.E)
packing_cost.grid(row= 13, column= 3)

spacer1 = tk.Label(win, text="")
spacer1.grid(row=14, column=1)



label15.grid(row= 15, column= 0, sticky= tk.E)
exchange_rate.grid(row= 15, column=1)

label16.grid(row= 15, column= 2, sticky= tk.E)
exchange_rate_safety.grid(row= 15, column=3)

spacer1 = tk.Label(win, text="")
spacer1.grid(row=16, column=1)

''' Design and positioning the Buttons'''

calculate = tk.Button(win, text="محاسبه قیمت دلاری", pady=15, padx=25, bg='green', fg='white',
                         command=cal)#find_count(authentic_code.get(), mainlist, database))
calculate.grid(row=30, column=2)

shutdown = tk.Button(win, text="خروج", pady=15, padx=35, bg='red', fg='white', command=win.destroy)#save_database(database))
shutdown.grid(row=30, column=3)

Refresher = tk.Button(win, text="بررسی ارقام", pady=15, padx=35, bg='Yellow', fg='Black', command=refresh)
Refresher.grid(row=30, column=1)


win.mainloop()