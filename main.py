#!/usr/bin/env python
# -*- coding: windows-1250 -*-

from pprint import pprint
from lxml import etree as et
from decimal import Decimal
from os.path import abspath
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from calendar import monthrange

import csv
import configparser
import os
import codecs
import time
import datetime

#TODO obsługa headerów w cfg

__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))


def mkWiersz(index, val, stawka):
    global root

    tns_ns = 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/'
    etd_ns = 'http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2016/01/25/eD/DefinicjeTypy/'

    SprzedazWiersz = et.SubElement(root, et.QName(tns_ns, 'SprzedazWiersz'))

    LpSprzedazy = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'LpSprzedazy'))
    LpSprzedazy.text = str(index)

    NrKontrahenta = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'NrKontrahenta'))
    if len(val['NIP']) > 0:
        NrKontrahenta.text = val['NIP']
    else:
        NrKontrahenta.text = "brak"

    NazwaKontrahenta = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'NazwaKontrahenta'))
    NazwaKontrahenta.text = val['Klient']

    AdresKontrahenta = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'AdresKontrahenta'))
    AdresKontrahenta.text = val['Adres']

    DowodSprzedazy = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'DowodSprzedazy'))
    DowodSprzedazy.text = val['Numer faktury']

    DataWystawienia = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'DataWystawienia'))
    date = val['Data wystawienia'].split(' ')
    DataWystawienia.text = date[0]

    DataSprzedazy = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'DataSprzedazy'))
    date = val['Data sprzedaży'].split(' ')
    DataSprzedazy.text = date[0]

    if stawka == 23:
        K_19 = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'K_19'))
        K_19.text = val['Netto 23%'].replace(',', '.')

        K_20 = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'K_20'))
        K_20.text = val['VAT 23%'].replace(',', '.')

    if stawka == 8:
        K_17 = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'K_17'))
        K_17.text = val['Netto 8%'].replace(',', '.')

        K_18 = et.SubElement(SprzedazWiersz, et.QName(tns_ns, 'K_18'))
        K_18.text = val['VAT 8%'].replace(',', '.')


def mkheader(datex):
    global settings
    global root
    datex = datex.split('-')
    dateRange = monthrange(int(datex[0]), int(datex[1]))
    pprint(datex)
    pprint(dateRange)

    tns_ns = 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/'
    Naglowek = et.SubElement(root, et.QName(tns_ns, 'Naglowek'))

    KodFormularza = et.SubElement(Naglowek, et.QName(tns_ns, 'KodFormularza'))
    KodFormularza.text = 'JPK_VAT'
    KodFormularza.set('wersjaSchemy', '1-1')
    KodFormularza.set('kodSystemowy', 'JPK_VAT (3)')

    WariantFormularza = et.SubElement(Naglowek, et.QName(tns_ns, 'WariantFormularza'))
    WariantFormularza.text = '3'

    CelZlozenia = et.SubElement(Naglowek, et.QName(tns_ns, 'CelZlozenia'))
    CelZlozenia.text = settings.get('naglowek', 'celzlozenia')

    DataWytworzeniaJPK = et.SubElement(Naglowek, et.QName(tns_ns, 'DataWytworzeniaJPK'))
    DataWytworzeniaJPK.text = str(datetime.datetime.utcnow()).replace(' ', 'T') #yep, work around

    DataOd = et.SubElement(Naglowek, et.QName(tns_ns, 'DataOd'))
    DataOd.text = datetime.date(int(datex[0]), int(datex[1]), 1).strftime('%Y-%m-%d')

    DataDo = et.SubElement(Naglowek, et.QName(tns_ns, 'DataDo'))
    DataDo.text = datetime.date(int(datex[0]), int(datex[1]), dateRange[1]).strftime('%Y-%m-%d')

    NazwaSystemu = et.SubElement(Naglowek, et.QName(tns_ns, 'NazwaSystemu'))
    NazwaSystemu.text = settings.get('naglowek', 'nazwasystemu')

def mkpodmiot():
    tns_ns = 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/'
    Podmiot = et.SubElement(root, et.QName(tns_ns, 'Podmiot1'))

    NIP = et.SubElement(Podmiot, et.QName(tns_ns, 'NIP'))
    NIP.text = settings.get('podmiot', 'nip')

    PelnaNazwa = et.SubElement(Podmiot, et.QName(tns_ns, 'PelnaNazwa'))
    PelnaNazwa.text = settings.get('podmiot', 'pelna_nazwa')

    Email = et.SubElement(Podmiot, et.QName(tns_ns, 'Email'))
    Email.text = settings.get('podmiot', 'email')



def run():
    rows = []  # wszystkie rekordy z .csv
    podatek = Decimal() # suma całego należnego podatku
    global gui
    global root
    csvFilename = gui.csvFilename.get()

    try:
        with open(csvFilename, 'r', encoding='windows-1250') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            for row in reader:
                rows.append(row)
    except Exception as e:
        messagebox.showerror('Błąd', e)
        return -1

    namespaces = {
        'tns': 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/',
        'etd': 'http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2016/01/25/eD/DefinicjeTypy/'
    }

    et.register_namespace('tns', 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/')
    et.register_namespace('etd', 'http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2016/01/25/eD/DefinicjeTypy/')

    tns_ns = 'http://jpk.mf.gov.pl/wzor/2017/11/13/1113/'
    etd_ns = 'http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2016/01/25/eD/DefinicjeTypy/'

    namespaces = {
        'tns': tns_ns,
        'etd': etd_ns
    }

    root = et.Element("{http://jpk.mf.gov.pl/wzor/2017/11/13/1113/}JPK", nsmap=namespaces)
    mkheader(rows[1]['Data wystawienia'].split(' ')[0])
    mkpodmiot()

    index = 0
    for key, val in enumerate(rows):
        if val['Netto 23%'] != '0,00':
            index += 1
            mkWiersz(index, val, 23)

        if val['Netto 8%'] != '0,00':
            index += 1
            mkWiersz(index, val, 8)

            podatek += Decimal(val['Netto 8%'].replace(',', '.'))
            podatek += Decimal(val['Netto 23%'].replace(',', '.'))

    SprzedazCtrl = et.SubElement(root, et.QName(tns_ns, 'SprzedazCtrl'))
    LiczbaWierszySprzedazy = et.SubElement(SprzedazCtrl, et.QName(tns_ns, 'LiczbaWierszySprzedazy'))
    # LiczbaWierszySprzedazy.text = str(len(rows))
    LiczbaWierszySprzedazy.text = str(index)
    PodatekNalezny = et.SubElement(SprzedazCtrl, et.QName(tns_ns, 'PodatekNalezny'))
    PodatekNalezny.text = str(podatek)

    tree = et.ElementTree(root)
    # tree.write('jpk 0618.xml', pretty_print=True)
    # tree.write('jpk 0618.xml', pretty_print=True, encoding='cp1250')
    tree.write('jpk.xml', pretty_print=True, encoding='UTF-8')

    messagebox.showinfo('Gotowe', 'Skrypt zakończył działanie, sprawdź poprawność pliku')


def browseCsv():
    global gui
    global settings, cfgfile

    path = settings.get('gui', 'lastpath')
    gui.csvFilename.set(filedialog.askopenfilename(initialdir=path, title="Select file",
                                                 filetypes=(("csv files", "*.csv"), ("all files", "*.*"))))

    path = os.sep.join(gui.csvFilename.get().split(os.sep)[0:-1])
    settings.set('gui', 'lastpath', path)
    savecfg()

    txt = gui.csvFilename.get().split(os.sep)
    CsvLabel.configure(text=txt[-1])


def savecfg():
    global settings
    with open(os.path.join(__location__, 'settings.ini'), 'w') as cfgfile:
        settings.write(cfgfile)


if __name__ == '__main__':
    gui = Tk()
    gui.title('PyJPK3')
    gui.csvFilename = StringVar()
    # gui.grid_columnconfigure(2, weight=1)

    # with open(os.path.join(__location__, 'settings.ini'), 'w') as cfgfile:
    settings = configparser.ConfigParser()
    # settings._interpolation = configparser.ExtendedInterpolation()
    settings.read(os.path.join(__location__, 'settings.ini'))

    gui.minsize(300, 200)

    CsvLabel = Label(gui, text=gui.csvFilename.get())
    CsvLabel.grid(row=0, column=1, sticky=E)

    loadCsv = Button(gui, text="Wczytaj plik CSV", command=browseCsv)
    loadCsv.grid(row=0, column=0, sticky=W)

    startBtn = Button(gui, text="Start", command=run)
    startBtn.grid(row=3, column=0, sticky=W)

    # try:
    #     config = configparser.ConfigParser()
    #     config.read_file(codecs.open(abspath('pyjpk.cfg'), encoding='windows-1250'))
    # except Exception as e:
    #     messagebox.showinfo("Błąd", e)

    mainloop()
