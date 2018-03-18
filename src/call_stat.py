#! python3

import datetime
import openpyxl
import pprint


def hours_minutes_seconds(td):
    hours, remainder = divmod(td.seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return datetime.time(hours, minutes, seconds)

# http://www.eprehledy.cz/predvolby_mobilnich_operatoru.php
# TODO: Add the other operators to the dict
CODES = {
    601:
    'O2',
    602:
    'O2',
    603:
    'T-Mobile',
    604:
    'T-Mobile',
    605:
    'T-Mobile',
    606:
    'O2',
    607:
    'O2',
    608:
    'Vodafone',
    702:
    'O2',
    721:
    'O2',
    722:
    'O2',
    723:
    'O2',
    724:
    'O2',
    725:
    'O2',
    726:
    'O2',
    727:
    'O2',
    728:
    'O2',
    729:
    'O2',
    731:
    'T-Mobile',
    732:
    'T-Mobile',
    733:
    'T-Mobile',
    734:
    'T-Mobile',
    735:
    'T-Mobile',
    736:
    'T-Mobile',
    737:
    'T-Mobile',
    738:
    'T-Mobile',
    739:
    'T-Mobile',
    770:
    'Vodafone',
    773:
    'Vodafone',
    774:
    'Vodafone',
    775:
    'Vodafone',
    776:
    'Vodafone',
    777:
    'Vodafone',
    778:
    'Vodafone'
}

WB = openpyxl.load_workbook(
    'C:\\Users\\ludek\\Downloads\\call_history_14_01_2018__08_10_edit.xlsx')
SHEET = WB.get_sheet_by_name('Call History To Excel')

call_stat = {}    # {'2017-01':{'O2':'11:41', }}

for calls in SHEET['B2':'F501']:  # + WB.get_highest_row()]:
    call = []
    for cellObj in calls:
        call.append(cellObj.value)
    name, mobile_number, call_type, call_duration, call_date = call

    if call_type == 'Outgoing':
        call_date_form = call_date.strftime('%Y-%m')
        call_stat.setdefault(call_date_form, {})

        code_prefix = int(str(mobile_number)[:3])
        if code_prefix in CODES.keys():
            provider = CODES[code_prefix]
        elif str(code_prefix)[0] == '8':
            continue
        else:
            provider = 'Others'
        month = call_stat[call_date_form]
        month.setdefault(provider, '00:00:00')
        t_new = datetime.timedelta(
            hours=call_duration.hour,
            minutes=call_duration.minute,
            seconds=call_duration.second)
        t_old = datetime.datetime.strptime(month[provider], '%H:%M:%S')
        t_old = datetime.timedelta(
            hours=t_old.hour, minutes=t_old.minute, seconds=t_old.second)
        month[provider] = hours_minutes_seconds(t_old + t_new).strftime(
            '%H:%M:%S')

FILE_STAT = open('stat.txt', 'w')
pprint.pprint(call_stat, FILE_STAT)
