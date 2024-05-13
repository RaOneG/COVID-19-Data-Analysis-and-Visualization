"""
File : COVID-19 Data Visualization
"""

import requests         # library of url requests
import xlrd          # library of spreadsheets (aka. excel sheets)
import datetime         # library for datetime to transform date/time into python date/time format
from tkinter import *
from tkinter.ttk import *
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg   # figure canvas: allows us to draw a graph using matplot, navigation: allows us to move the graph, zoom the graph...etc.
from matplotlib import style        # style of graph sheet
from pandas import DataFrame

matplotlib.use("TkAgg")
style.use("Solarize_Light2")


def main():
    cases = []
    dates = []
    deaths = []
    countries = []
    no_row = 1
    title = "Welcome to COVID-19 World Data Analysis"
    intro_msg = "This program analysis the COVID-19 confirmed cases and deaths and visualizes\n" \
                "them in the form of graphs vs. time (since the outbreak of the virus)* for each country,\n" \
                "to view the visualized data please choose a country then press OK."
    note = "*This Data is updated Daily"

    # request to get the url and download the spreadsheet of COVID-19 data in given location
    url = "https://www.ecdc.europa.eu/sites/default/files/documents/COVID-19-geographic-disbtribution-worldwide.xlsx"
    file = requests.get(url, allow_redirects=True)
    open('D:/Courses+Camps/Code In Place/Final Project/Final/file'
         '/COVID-19-geographic-disbtribution-worldwide.xlsx', 'wb').write(file.content)

    # open spreadsheet
    workbook = xlrd.open_workbook('file/COVID-19-geographic-disbtribution-worldwide.xlsx')
    # open sheet of index 1 (first sheet)
    worksheet = workbook.sheet_by_index(0)

    # store countries in a list
    countries = get_countries_list(no_row, worksheet, countries)

    # makes a canvas of 600-pixels wide
    # and 600-pixels tall with title "COVID-19 Data"
    window = make_window("COVID-19 World Data Analysis")

    # create label on welcome screen
    welcome_msg1 = Label(window, justify=LEFT, text=title, font="Arial 25 bold")
    welcome_msg1.pack()
    welcome_msg2 = Label(window, justify=CENTER, text=intro_msg, font="Bahnschrift 20")
    welcome_msg2.pack()

    # create drop down menu of list countries01285552026Oo
    # ask user to choose a country
    variable = StringVar(window)
    variable.set("Countries")  # default value
    drop_down_menu = OptionMenu(window, variable, *countries)
    drop_down_menu.pack()

    def ok():
        user_selection = variable.get()
        # empty the window from all widgets
        widget_list = window.winfo_children()
        for item in widget_list:
            item.pack_forget()
        window.quit()
        return user_selection

    ok_button = Button(window, text="OK", command=ok)
    ok_button.pack()

    welcome_msg3 = Label(window, justify=CENTER, text=note, font="Arial 11 italic")
    welcome_msg3.pack()

    window.mainloop()

########################################################################################################################

    country = ok()

    # look for country
    while worksheet.cell(no_row, 6).value != xlrd.empty_cell.value:
        if country == worksheet.cell(no_row, 6).value:
            break
        no_row += 1

    # loop on the country's data
    # get range of dates for this country
    # get range of cases corresponding range of dates
    while worksheet.cell(no_row, 6).value == country:
        deaths = get_deaths(no_row, deaths, worksheet)
        cases = get_case(no_row, cases, worksheet)
        dates = get_date(no_row, dates, worksheet)
        no_row += 1
    deaths.reverse()
    cases.reverse()
    dates.reverse()
    # convert list of dates into python date/time format to bea ble to display it correctly
    dates = [datetime.datetime.strptime(d, "%m/%d/%Y").date() for d in dates]

    # make canvas for graphs
    canvas = window

    # make data frame for cases and dates
    cases_corresponding_dates = {"Date": dates, "Cases": cases}
    cases_dates_data_frame = DataFrame(cases_corresponding_dates, columns=["Date", "Cases"])

    # make data frame for deaths and dates
    deaths_corresponding_dates = {"Date": dates, "Deaths": deaths}
    deaths_dates_data_frame = DataFrame(deaths_corresponding_dates, columns=["Date", "Deaths"])

    # plot (cases vs. date range) frames on matplot graph
    cases_figure = plt.Figure(figsize=(6, 6), dpi=100)
    cases_axis = cases_figure.add_subplot(111)
    cases_graph = FigureCanvasTkAgg(cases_figure, canvas)
    cases_graph.get_tk_widget().pack(side=LEFT, fill=BOTH)
    cases_dates_data_frame = cases_dates_data_frame[["Date", "Cases"]].groupby("Date").sum()
    cases_dates_data_frame.plot(kind='line', legend=True, ax=cases_axis, color='y', fontsize=7)
    cases_axis.set_title(country+"'s COVID-19 confirmed cases")

    # plot (deaths vs. date range) frames on matplot graph
    deaths_figure = plt.Figure(figsize=(8, 6), dpi=100)
    deaths_axis = deaths_figure.add_subplot(111)
    deaths_graph = FigureCanvasTkAgg(deaths_figure, canvas)
    deaths_graph.get_tk_widget().pack(side=RIGHT, fill=BOTH)
    deaths_dates_data_frame = deaths_dates_data_frame[["Date", "Deaths"]].groupby("Date").sum()
    deaths_dates_data_frame.plot(kind='line', legend=True, ax=deaths_axis, color='r', fontsize=7)
    deaths_axis.set_title(country + "'s COVID-19 Deaths")

    canvas.mainloop()


def get_countries_list(no_row, worksheet, countries):
    country = worksheet.cell(no_row, 6).value
    countries.append(country)
    for n_row in range(1, worksheet.nrows):
        if country != worksheet.cell(n_row, 6).value:
            country = worksheet.cell(n_row, 6).value
            countries.append(country)
    return countries


def get_date(no_row, dates, worksheet):
    """
    get date at corresponding row and store it in dates list
    :param no_row: number of row
    :param dates: list of dates
    :param worksheet: the sheet containing the date, cases and country data
    :return: list of dates
    """
    day = int(worksheet.cell(no_row, 1).value)
    month = int(worksheet.cell(no_row, 2).value)
    year = int(worksheet.cell(no_row, 3).value)
    dates.append(str(month)+"/"+str(day)+"/"+str(year))
    return dates


def get_case(no_row, cases, worksheet):
    """
    get cases at corresponding row and store it in cases list
    :param no_row: number of row
    :param cases: list of cases
    :param worksheet: the sheet containing the date, cases and country data
    :return: list of cases
    """
    cases_cell = worksheet.cell(no_row, 4).value
    cases.append(cases_cell)
    return cases


def get_deaths(no_row, deaths, worksheet):
    """
    get deaths at corresponding row and store it in deaths list
    :param no_row: number of row
    :param deaths: list of deaths
    :param worksheet: the sheet containing the date, cases, deaths and country data
    :return: list of deaths
    """
    deaths_cell = worksheet.cell(no_row, 5).value
    deaths.append(deaths_cell)
    return deaths


def make_window(title):
    """
    function : make canvas
    creates a canvas of width and height
    as passed and gives it a title
    :param title: title of the canvas
    :return: returns the canvas with specified parameters
    """
    # creating a window so that we
    # could create a canvas on it
    window = Tk()
    # giving the window a title
    window.title(title)
    # makes it full screen
    window.geometry("{0}x{1}+0+0".format(window.winfo_screenwidth(), window.winfo_screenheight()))
    return window


if __name__ == '__main__':
    main()
