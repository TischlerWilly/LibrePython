# alles einfalten: strg+k --> strg+0
from __future__ import unicode_literals
from ast import Not
from genericpath import exists
from importlib.util import MAGIC_NUMBER
import os
from os.path import expanduser
from pathlib import Path
import string
import uno
import datetime
import time
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.sheet.CellInsertMode import RIGHT as INSERT_RE
from com.sun.star.sheet.CellInsertMode import DOWN as INSERT_UN
from com.sun.star.table.CellHoriJustify import LEFT as AUSRICHTUNG_HORI_Li
from com.sun.star.table.CellHoriJustify import CENTER as AUSRICHTUNG_HORI_MI
from com.sun.star.table.CellHoriJustify import RIGHT as AUSRICHTUNG_HORI_RE
from com.sun.star.sheet.CellDeleteMode import LEFT as DEL_LI
from com.sun.star.table.CellContentType import TEXT as CELLCONTENTTYP_TEXT
from com.sun.star.table import BorderLine
from com.sun.star.awt.FontWeight import NORMAL as FONT_NOT_BOLD
from com.sun.star.awt.FontWeight import BOLD as FONT_BOLD
from com.sun.star.awt.FontUnderline import SINGLE as FONT_UNDERLINED_SINGLE
from shutil import copyfile
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
"""
ToDo:
- 

"""
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------msgbox für LibreOffice definieren:
def msgbox(message, title='LibreOffice', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        MSG_BUTTONS => BUTTONS_OK = 1, BUTTONS_OK_CANCEL = 2, BUTTONS_YES_NO = 3, BUTTONS_YES_NO_CANCEL = 4, BUTTONS_RETRY_CANCEL = 5, BUTTONS_ABORT_IGNORE_RETRY = 6
        https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1MessageBoxButtons.html

        type_msg => MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
        https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html#ad249d76933bdf54c35f4eaf51a5b7965
    """
    CTX = XSCRIPTCONTEXT.getComponentContext()
    toolkit = CTX.ServiceManager.createInstance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()
    # Anwendung:
    # msgbox('Hallo Oliver', 'msgbox', 1, 'QUERYBOX')
    # """
#----------------------------------------------------------------------------------
def RGBTo32bitInt(r, g, b):
  return int('%02x%02x%02x' % (r, g, b), 16)
def erstelle_datei(full_path):
    # Pfadtrenner ist auf Windows das \\
    # Beispiel: "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    # full_path = "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    erfolg = True
    my_file = Path(full_path)
    if my_file.is_file():
        msg = "Die Datei existiert bereits und wird nicht überschrieben."
        titel = "erstelle_datei(full_path)"
        msgbox(msg, titel, 1, 'QUERYBOX')
        erfolg = False
    else:
        new_file = open(full_path, "w")
        new_file.close()
    return erfolg
def schreibe_in_datei(full_path, sText):
    # Pfadtrenner ist auf Windows das \\
    # Beispiel: "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    # full_path = "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    erfolg = True
    my_file = Path(full_path)
    if my_file.is_file():
        msg = "Die Datei existiert bereits und wird nicht überschrieben."
        msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        erfolg = False
    else:
        new_file = open(full_path, "w")
        new_file.write(sText)
        new_file.close()
    return erfolg
def schreibe_in_datei_entferne_bestehende(full_path, sText):
    # Pfadtrenner ist auf Windows das \\
    # Beispiel: "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    # full_path = "C:\\Users\\AV6\\Desktop\\Unbekannt.odt"
    datei_vorhanden = False
    my_file = Path(full_path)
    if my_file.is_file():
        datei_vorhanden = True    
    new_file = open(full_path, "w")
    new_file.write(sText)
    new_file.close()
    return datei_vorhanden
def get_userpath():
    return expanduser("~")
def istZiffer(c):
    erg = False
    if c is '0':
        erg = True
    elif c is '1':
        erg = True
    elif c is '2':
        erg = True
    elif c is '3':
        erg = True
    elif c is '4':
        erg = True
    elif c is '5':
        erg = True
    elif c is '6':
        erg = True
    elif c is '7':
        erg = True
    elif c is '8':
        erg = True
    elif c is '9':
        erg = True
    return erg
def findeDateien(name, pfad):
    alleDateien = os.listdir(pfad)
    dateien = []
    for i in alleDateien:
        if name in i:
            dateien += [os.path.join(pfad, i)]
    return dateien
#----------------------------------------------------------------------------------

class ol_tabelle:
    # Anwendung: t = ol_tabelle() # ein Objekt der Klasse slist anlegen
    # row = Zeile
    # column = spalte
    def __init__(self):
        self.context = XSCRIPTCONTEXT # globale Variable im sOffice-kontext
        self.doc = self.context.getDocument() #aktuelles Document per Methodenaufruf ! mit Klammern !
        self.sheets = self.doc.Sheets # ! Attributaufruf ohne Klammern !
        self.sheet = self.doc.getCurrentController().getActiveSheet() # die Tabelle die gerade den Fokus hat
        self.ctrlr = self.doc.CurrentController
        pass
    def get_selction(self):
        return self.ctrlr.getSelection()
    def get_selection_zeile_start(self):
        sel = self.get_selction()
        area = sel.getRangeAddress()
        return area.StartRow
    def get_selection_zeile_ende(self):
        sel = self.get_selction()
        area = sel.getRangeAddress()
        return area.EndRow
    def get_selection_spalte_start(self):
        sel = self.get_selction()
        area = sel.getRangeAddress()
        return area.StartColumn
    def get_selection_spalte_ende(self):
        sel = self.get_selction()
        area = sel.getRangeAddress()
        return area.EndColumn
    def set_tabindex(self, i): # ist optional
        self.sheet = self.sheets.getByIndex(i) # erstes Blatt per Index
        pass
        # Anwendung: t.set_tabindex(0) 
    def set_tabname(self, n): # ist optional
        self.sheet = self.sheets.getByName(n) # 'Tabelle2 per Namen
        pass
        # Anwendung: t.set_tabname('Tabelle1')  
    def get_tabname(self):
        return self.sheet.getName()
    #-----------------------------------------------------------------------------------------------
    # Seite:
    #-----------------------------------------------------------------------------------------------
    def set_seitenformat(self, sPapierformat, IstQuerformat, iRandLi, iRandRe, iRandOb, iRandUn, hatKopfzeile, hatFusszeile):
        pageStyle = self.doc.getStyleFamilies().getByName("PageStyles")
        page = pageStyle.getByName("Default")
        # Seitenränder:
        # 500 == 5mm
        page.LeftMargin = iRandLi
        page.RightMargin = iRandRe
        page.TopMargin = iRandOb
        page.BottomMargin = iRandUn 
        # Kopfzeile an/aus:
        page.HeaderOn = hatKopfzeile
        # Fußzeile an/aus:
        page.FooterOn = hatFusszeile
        # Seitenformat:
        if(sPapierformat == "A4"):
            if(IstQuerformat == False):
                # A4 hoch:
                page.IsLandscape = False
                page.Width = 21000
                page.Height = 29700
            else:
                # A4 quer:
                page.IsLandscape = False
                page.Width = 29700
                page.Height = 21000
        elif(sPapierformat == "A3"):
            if(IstQuerformat == False):
                # A3 hoch:
                page.IsLandscape = False
                page.Width = 29700
                page.Height = 42000
            else:
                # A3 quer:
                page.IsLandscape = True
                page.Width = 42000
                page.Height = 29700        
        pass
        # Anwendung: set_setenformat("A3", True, 500, 500, 500 , 500, False, False)
    def set_pageScaling(self, iSkaling):
        pageStyle = self.doc.getStyleFamilies().getByName("PageStyles")
        page = pageStyle.getByName("Default")
        # page.PageScale = 25 # 25%
        page.PageScale = iSkaling
        pass
    #-----------------------------------------------------------------------------------------------
    # Tabs:
    #-----------------------------------------------------------------------------------------------
    def tab_existiert(self, sTabname):
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_schon_vorhanden = 0
        for i in range (0, len(namen)):
            if namen[i] == sTabname:
                tab_schon_vorhanden = 1
                break #for i
        if tab_schon_vorhanden == 1:
            return True
        else: 
            return False
    def tab_anlegen(self, sTabname, iTabIndex):
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_schon_vorhanden = 0
        for i in range (1, len(namen)):
            if namen[i] == sTabname:
                tab_schon_vorhanden = 1
                break #for i
        if tab_schon_vorhanden == 1:
            msg = "Die Registerkarte \""
            msg += sTabname
            msg += "\" ist bereits vorhanden!"
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        else:
            self.doc.Sheets.insertNewByName(sTabname, iTabIndex)
        pass
    def tab_entfernen(self, sTabname):
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_schon_vorhanden = 0
        for i in range (0, len(namen)):
            if namen[i] == sTabname:
                tab_schon_vorhanden = 1
                break #for i
        if tab_schon_vorhanden == 1:
            self.doc.Sheets.removeByName(sTabname)
        pass
    def tab_kopieren(self, sNeuerTabName, iTabIndex):
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_schon_vorhanden = 0
        for i in range (1, len(namen)):
            if namen[i] == sNeuerTabName:
                tab_schon_vorhanden = 1
                break #for i
        if tab_schon_vorhanden == 1:
            msg = "Die Registerkarte \""
            msg += sNeuerTabName
            msg += "\" ist bereits vorhanden!"
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
            return 1
        else:
            sAlterTabName = self.get_tabname()
            self.doc.Sheets.copyByName(sAlterTabName, sNeuerTabName, iTabIndex)
            return 0
        return 0
    def tab_kopieren2(self, sAlterTabName, sNeuerTabName, iTabIndex):
        retwert = 0
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_alt_schon_vorhanden = 0
        for i in range (0, len(namen)):
            if namen[i] == sAlterTabName:
                tab_alt_schon_vorhanden = 1
                break #for i
        tab_neu_schon_vorhanden = 0
        for i in range (0, len(namen)):
            if namen[i] == sNeuerTabName:
                tab_neu_schon_vorhanden = 1
                break #for i
        if tab_alt_schon_vorhanden == 1:
            if tab_neu_schon_vorhanden == 0:
                self.doc.Sheets.copyByName(sAlterTabName, sNeuerTabName, iTabIndex)
            else:
                msg = "Die Registerkarte \""
                msg += sNeuerTabName
                titel = "tab_kopieren2(self, sAlterTabName, sNeuerTabName, iTabIndex)"
                msg += "\" ist schon vorhanden und kann desshalb nicht kopiert werden!"
                msgbox(msg, titel, 1, 'QUERYBOX')
                retwert = 2
        else:
            msg = "Die Registerkarte \""
            msg += sAlterTabName
            titel = "tab_kopieren2(self, sAlterTabName, sNeuerTabName, iTabIndex)"
            msg += "\" ist nicht vorhanden und kann desshalb nicht kopiert werden!"
            msgbox(msg, titel, 1, 'QUERYBOX')
            retwert = 1
        return retwert
    def tab_setName(self, sNeuerTabName):
        namen = []
        namen = self.doc.Sheets.ElementNames
        tab_schon_vorhanden = 0
        for i in range (1, len(namen)):
            if namen[i] == sNeuerTabName:
                tab_schon_vorhanden = 1
                break #for i
        if tab_schon_vorhanden == 1:
            msg = "Die Registerkarte \""
            msg += sNeuerTabName
            msg += "\" ist bereits vorhanden!"
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
            return 1
        else:
            self.sheet.Name = sNeuerTabName
            return 0
        return 0
    def set_tabfokus_s(self, sTabname):
        sheet = self.doc.Sheets[sTabname]
        self.doc.getCurrentController().setActiveSheet(sheet)
        self.set_tabname(sTabname)
        pass
    def set_tabfokus_i(self, iTabindex):
        sheet = self.doc.Sheets[iTabindex]
        self.doc.getCurrentController().setActiveSheet(sheet)
        pass
    #-----------------------------------------------------------------------------------------------
    # Zellen / Ranges(Bereiche):
    #-----------------------------------------------------------------------------------------------
    def get_zelle_i(self, zeile, spalte):
        return self.sheet.getCellByPosition(spalte, zeile)
        # Anwendung: text = t.get_zelle_i(1,1)
    def zelle_verschieben_i(self, iZeileVon, iSpalteVon, iZeileNach, iSpalteNach):
        source = self.sheet.getCellRangeByPosition(iSpalteVon, iZeileVon, iSpalteVon, iZeileVon)
        target = self.sheet.getCellByPosition(iSpalteNach, iZeileNach)
        self.sheet.moveRange(target.CellAddress, source.RangeAddress)
        pass
    def set_zelltext_s(self, sRange, text): # self muss immer als erster Parameter übergeben werden
        self.sheet.getCellRangeByName(sRange).String = text
        pass
        # Anwendung: t.set_zelltext_s('A1', 'Hallo 1')
    def set_zelltext_datum_s(self, sRange, jjjj, mm, tt): # self muss immer als erster Parameter übergeben werden        
        numberformats = self.doc.NumberFormats
        Locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        dateformat = numberformats.queryKey('TT.MM.JJJJ', Locale, True )
        if dateformat == -1:
            dateformat = numberformats.addNew('TT.MM.JJJJ', Locale)
        datum = datetime.date(int(jjjj), int(mm), int(tt))
        d = datum.isoformat()
        self.sheet.getCellRangeByName(sRange).Formula = d
        self.sheet.getCellRangeByName(sRange).NumberFormat = dateformat
        pass
    def set_zellformat_s(self, sRange, sFormatcode):
        numberformats = self.doc.NumberFormats
        Locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        myformat = numberformats.queryKey(sFormatcode, Locale, True )
        if myformat == -1:
            myformat = numberformats.addNew(sFormatcode, Locale)
        self.sheet.getCellRangeByName(sRange).NumberFormat = myformat
        pass
    def get_zelltext_s(self, zellname):
        return self.sheet.getCellRangeByName(zellname).String
        # Anwendung: text = t.get_zelltext_s("B2")
    def set_zelltext_i(self, zeile, spalte, text):
        self.sheet.getCellByPosition(spalte, zeile).String = text
        pass
        # Anwendung: t.set_zelltext_i(1, 0, 'Hallo 2')
    def get_zelltext_i(self, zeile, spalte):
        return self.sheet.getCellByPosition(spalte, zeile).String
        # Anwendung: text = t.get_zelltext_i(1,1)
    def get_zelltext_akt_auswahl(self):
        iZeile = 0
        iSpalte = 0
        iZeileStart = self.get_selection_zeile_start()
        iZeileEnde  = self.get_selection_zeile_ende()
        iSpalteStart = self.get_selection_spalte_start()
        iSpalteEnde = self.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):# wird gebraucht zur Typenumwandlung
            for s in range(iSpalteStart, iSpalteEnde+1):
                iZeile = z
                iSpalte = s
                break # nur 1 durchlauf erwünscht
            break # nur 1 durchlauf erwünscht
        return self.get_zelltext_i(iZeile, iSpalte)
    def set_zellformel_s(self, zellname, formel):
        self.sheet.getCellRangeByName(zellname).Formula = formel
        pass
        # Anwendung: t.set_zellformel_s('B1', '=1+1')
    def get_zellformel_s(self, zellname):
        return self.sheet.getCellRangeByName(zellname).Formula
        # Anwendung: text = t.get_zellformel_s("C1")
    def set_zellformel_i(self, zeile, spalte, formel):
        self.sheet.getCellByPosition(spalte, zeile).Formula = formel
        pass
        # Anwendung: t.set_zellformel_i(1, 1, '=2+2')
    def get_zellformel_i(self, zeile, spalte):
        return self.sheet.getCellByPosition(spalte, zeile).Formula
        # Anwendung: text = t.get_zellformel_i(0,2)
    def set_zellzahl_s(self, zellname, zahl):
        self.sheet.getCellRangeByName(zellname).Value = zahl
        pass
        # Anwendung: t.set_zellzahl_s('C1', '555')
    def get_zellzahl_s(self, zellname):
        return self.sheet.getCellRangeByName(zellname).Value
        # Anwendung: zahl = t.get_zellzahl_s("D1")
    def set_zellzahl_i(self, zeile, spalte, zahl):
        self.sheet.getCellByPosition(spalte, zeile).Value = zahl
        pass
        # Anwendung: t.set_zellzahl_i(1, 2, '678')
    def get_zellzahl_i(self, zeile, spalte):
        return self.sheet.getCellByPosition(spalte, zeile).Value
        # Anwendung: t.get_zellzahl_i(1, 2)
    def set_zellfarbe_s(self, sRange, farbe): # farbe ist ein long-wert
        self.sheet.getCellRangeByName(sRange).CellBackColor = farbe
        pass
        # Anwendung: t.set_zellfarbe_s("A2", farbe)
    def get_zellfarbe_s(self, zellname): # farbe ist ein long-wert
        return self.sheet.getCellRangeByName(zellname).CellBackColor
        # Anwendung: farbe = t.get_zellfarbe_s("A1")
    def set_zellfarbe_i(self, zeile, spalte, farbe): # farbe ist ein long-wert
        self.sheet.getCellByPosition(spalte, zeile).CellBackColor = farbe
        pass
        # Anwendung: t.set_zellfarbe_i(1, 0, farbe)
    def get_zellfarbe_i(self, zeile, spalte): # farbe ist ein long-wert
        return self.sheet.getCellByPosition(spalte, zeile).CellBackColor
        # Anwendung: farbe = t.get_zellfarbe_i(0,0)
    def set_zellausrichtungHori_s(self, sRange, sAusrichtung):
        oRange = self.sheet.getCellRangeByName(sRange)
        if sAusrichtung == "li":
            oRange.HoriJustify = AUSRICHTUNG_HORI_Li
        elif sAusrichtung == "mi":
            oRange.HoriJustify = AUSRICHTUNG_HORI_MI
        elif sAusrichtung == "re":
            oRange.HoriJustify = AUSRICHTUNG_HORI_RE
        pass
        #Anwendung: t.set_zellausrichtungHori_s("B2:C3", "re")
    def set_SchriftGroesse_s(self, sRange, iGroesse):
        self.sheet.getCellRangeByName(sRange).CharHeight = iGroesse
        pass
    def set_SchriftFett_s(self, sRange, bIstFett):
        if(bIstFett == True):
            self.sheet.getCellRangeByName(sRange).CharWeight = FONT_BOLD
        else:
            self.sheet.getCellRangeByName(sRange).CharWeight = FONT_NOT_BOLD
        pass
    def set_SchriftFarbe_s(self, sZelle, farbe): # farbe ist ein long-wert
        zelle = self.sheet.getCellRangeByName(sZelle) # Range ist nich möglich nur eine Zelle siehe nächste Zeile
        cursor = zelle.createTextCursor() # funktioniert nur mit je einer Zelle
        cursor.setPropertyValue( "CharColor", farbe )
        pass
    def set_Rahmen_komplett_s(self, sRange, iLinienbreite):
        tableBorder = self.sheet.getPropertyValue("TableBorder")
        borderLine  = BorderLine() # Objekt anlegen
        borderLine.OuterLineWidth = iLinienbreite # Linienbreite bestimmen
        tableBorder.VerticalLine = borderLine
        tableBorder.IsVerticalLineValid = True
        tableBorder.HorizontalLine = borderLine
        tableBorder.IsHorizontalLineValid = True
        tableBorder.LeftLine = borderLine
        tableBorder.IsLeftLineValid = True
        tableBorder.RightLine = borderLine
        tableBorder.IsRightLineValid = True
        tableBorder.TopLine = borderLine
        tableBorder.IsTopLineValid = True
        tableBorder.BottomLine = borderLine
        tableBorder.IsBottomLineValid = True
        self.sheet.getCellRangeByName(sRange).setPropertyValue("TableBorder", tableBorder)
        pass
    def set_Rahmen_s(self, sRange, iLinienbreite, rahmenfarbe):
        tableBorder = self.sheet.getPropertyValue("TableBorder")
        borderLine  = BorderLine() # Objekt anlegen
        borderLine.OuterLineWidth = iLinienbreite # Linienbreite bestimmen
        borderLine.Color = rahmenfarbe
        tableBorder.VerticalLine = borderLine
        tableBorder.IsVerticalLineValid = True
        tableBorder.HorizontalLine = borderLine
        tableBorder.IsHorizontalLineValid = True
        tableBorder.LeftLine = borderLine
        tableBorder.IsLeftLineValid = True
        tableBorder.RightLine = borderLine
        tableBorder.IsRightLineValid = True
        tableBorder.TopLine = borderLine
        tableBorder.IsTopLineValid = True
        tableBorder.BottomLine = borderLine
        tableBorder.IsBottomLineValid = True
        self.sheet.getCellRangeByName(sRange).setPropertyValue("TableBorder", tableBorder)
        pass
    #-----------------------------------------------------------------------------------------------
    # Spalten:
    #-----------------------------------------------------------------------------------------------
    def insert_spalte_re_i(self, iSpalte, iMenge):
        start_column = iSpalte
        end_column = iSpalte + iMenge - 1
        start_row = 0        
        end_row = 9999
        bereich = self.sheet.getCellRangeByPosition(start_column, start_row, end_column, end_row)
        self.sheet.insertCells(bereich.RangeAddress, INSERT_RE)
        pass
        # Anwendung: t.insert_spalte_re_i(2,2)
    def delelte_spalten_re_i(self, iSpalte, iMenge):
        start_column = iSpalte
        end_column = iSpalte + iMenge - 1
        start_row = 0        
        end_row = 9999
        bereich = self.sheet.getCellRangeByPosition(start_column, start_row, end_column, end_row)
        self.sheet.removeRange(bereich.RangeAddress, DEL_LI)
        pass
        # Anwendung: t.delete_spalten_re_i(2,2)
    def spalte_verschieben_i(self, iVon, iNach):
        start_column = iVon
        end_column = iVon
        start_row = 0        
        end_row = 9999
        source = self.sheet.getCellRangeByPosition(start_column, start_row, end_column, end_row)
        target = self.sheet.getCellByPosition(iNach, 0)
        self.sheet.moveRange(target.CellAddress, source.RangeAddress)
        pass
        # Anwendung: t.spalte_verschieben_i(2, 5)
    def optimale_spaltenbreiten(self):
        oSpalten = self.sheet.getColumns()
        oSpalten.OptimalWidth = True
        pass
        # Anwendung: t.optimale_spaltenbreiten()
    def set_spaltenausrichtung_i(self, spalte, sAusrichtung):
        oSpalte = self.sheet.getColumns().getByIndex(spalte)
        if sAusrichtung == "li":
            oSpalte.HoriJustify = AUSRICHTUNG_HORI_Li
        elif sAusrichtung == "mi":
            oSpalte.HoriJustify = AUSRICHTUNG_HORI_MI
        elif sAusrichtung == "re":
            oSpalte.HoriJustify = AUSRICHTUNG_HORI_RE
        pass
        pass
        # Anwendung: t.optimale_spaltenbreiten()
    def optimale_spaltenbreite_i(self, iSpalte):
        oSpalten = self.sheet.getColumns()
        oSpalte = oSpalten.getByIndex(iSpalte)
        oSpalte.OptimalWidth = True
        pass
        # Anwendung: t.optimale_spaltenbreite_i(5)
    def set_spaltenbreite_i(self, iSpalte, iBreite): # 100 == 1mm
        oSpalten = self.sheet.getColumns()
        oSpalte = oSpalten.getByIndex(iSpalte)
        oSpalte.Width = iBreite
        pass
        #Anwendung: t.set_spaltenbreite_i(5, 500)
    def get_spaltenbreite_i(self, iSpalte): # 100 == 1mm
        oSpalten = self.sheet.getColumns()
        oSpalte = oSpalten.getByIndex(iSpalte)
        return oSpalte.Width
        #Anwendung: t.get_spaltenbreite_i(5)
    def set_spalte_sichtbar_i(self, iSpalte, boolSichtbar):
        oSpalten = self.sheet.getColumns()
        oSpalte = oSpalten.getByIndex(iSpalte)
        oSpalte.IsVisible = boolSichtbar
        pass
        # Anwendung: t.set_spalte_sichtbar_i(2, False)
    #-----------------------------------------------------------------------------------------------
    # Zeilen:
    #-----------------------------------------------------------------------------------------------
    def insert_zeile_un_i(self, iZeile, iMenge):
        start_column = 0
        end_column = 999
        start_row = iZeile      
        end_row = iZeile + iMenge - 1
        bereich = self.sheet.getCellRangeByPosition(start_column, start_row, end_column, end_row)
        self.sheet.insertCells(bereich.RangeAddress, INSERT_UN)
        pass
        # Anwendung: t.insert_zeile_un_i(2,2)
    def set_zeilenhoehen(self, iHoehe): # 100 == 1mm
        oZeilen = self.sheet.getRows()
        oZeilen.Height = iHoehe
        pass
        #Anwendung: t.set_zeilenhoehen(1000)
    def set_zeilenhoehe_i(self, iZeile, iHoehe): # 100 == 1mm
        oZeilen = self.sheet.getRows()
        oZeile = oZeilen.getByIndex(iZeile)
        oZeile.Height = iHoehe
        pass
        #Anwendung: t.set_zeilenhoehe_i(5, 1000)
    #-----------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
class ol_textdatei:
    def __init__(self):
        #self.doc = XSCRIPTCONTEXT.getDocument()
        #self.text = self.doc.getText()
        #self.text.setString('Hello World in Python in Writer')
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.text = self.doc.getText()
        self.desktop = XSCRIPTCONTEXT.getDesktop()
        self.model = self.desktop.getCurrentComponent() 
    def set_text_hoehe(self, iHoehe):
        oSel = self.doc.CurrentSelection.getByIndex(0) # get the current selection
        oTC = self.text.createTextCursorByRange(oSel) # TextCursor erzeugen      
        oEnum = oTC.Text.createEnumeration()
        while oEnum.hasMoreElements():
            oPar = oEnum.nextElement()
            oPar.CharHeight = iHoehe
        pass
    #-----------------------------------------------------------------------------------------------
    # Seite:
    #-----------------------------------------------------------------------------------------------
    def set_seitenformat(self, sPapierformat, IstQuerformat, iRandLi, iRandRe, iRandOb, iRandUn):
        #pageStyle = self.doc.getStyleFamilies().getByName("PageStyles")
        #page = pageStyle.getByName("Default")
        oViewCursor = self.doc.CurrentController.getViewCursor()
        pageStyle = oViewCursor.PageStyleName
        page = self.doc.StyleFamilies.getByName("PageStyles").getByName(pageStyle)
        # Seitenränder:
        # 500 == 5mm
        page.LeftMargin = iRandLi
        page.RightMargin = iRandRe
        page.TopMargin = iRandOb
        page.BottomMargin = iRandUn 
        # Seitenformat:
        if(sPapierformat == "A4"):
            if(IstQuerformat == False):
                # A4 hoch:
                page.IsLandscape = False
                page.Width = 21000
                page.Height = 29700
            else:
                # A4 quer:
                page.IsLandscape = False
                page.Width = 29700
                page.Height = 21000
        elif(sPapierformat == "A3"):
            if(IstQuerformat == False):
                # A3 hoch:
                page.IsLandscape = False
                page.Width = 29700
                page.Height = 42000
            else:
                # A3 quer:
                page.IsLandscape = True
                page.Width = 42000
                page.Height = 29700   
        elif(sPapierformat == "A6"):
            if(IstQuerformat == False):
                # A6 hoch:
                page.IsLandscape = False
                page.Width = 10500
                page.Height = 14800
            else:
                # A6 quer:
                page.IsLandscape = True
                page.Width = 14800
                page.Height = 10500 
        pass
        # Anwendung: set_setenformat("A3", True, 500, 500, 500 , 500, False, False)
#----------------------------------------------------------------------------------
class slist: # Calc
    def __init__(self):
        self.t = ol_tabelle()
        self.maxistklen = 999  
        # CNC-Pfad des Postprozessors:
        windowsuser = os.getlogin()
        self.cnc_pfad = "C:\\Users\\"
        self.cnc_pfad += windowsuser
        self.cnc_pfad += "\\Documents\\CAM\\von postprozessor\\eigen"
        # Farben bestimmen:
        self.farblos = -1
        self.rot = RGBTo32bitInt(204, 0, 0)
        self.gelb = RGBTo32bitInt(255, 255, 0) 
        self.grau = RGBTo32bitInt(204, 204, 204) 
        self.gruen = RGBTo32bitInt(129, 212, 26) 
        pass
    def tabkopf_anlegen(self):
        self.t.set_zelltext_s("A1", "Bezeichnung")
        self.t.set_zelltext_s("B1", "Anzahl")
        self.t.set_zelltext_s("C1", "Länge")
        self.t.set_zelltext_s("D1", "Breite")
        self.t.set_zelltext_s("E1", "Dicke")
        self.t.set_zelltext_s("F1", "Material")
        self.t.set_zelltext_s("G1", "Kante links     (vo)") 
        self.t.set_zelltext_s("H1", "KaDi")
        self.t.set_zelltext_s("I1", "Kante rechts   (hi)") 
        self.t.set_zelltext_s("J1", "KaDi")
        self.t.set_zelltext_s("K1", "Kante oben     (li)")
        self.t.set_zelltext_s("L1", "KaDi")
        self.t.set_zelltext_s("M1", "Kante unten   (re)")
        self.t.set_zelltext_s("N1", "KaDi")
        self.t.set_zelltext_s("O1", "Bemerkung")
        pass
        # Anwendung: self.tabkopf_anlegen()
    def dicke_aus_artikelnummer_bestimmen(self):
        iSpalteArtNr = 5
        iSpalteDicke = 4
        for i in range(1, self.maxistklen):
            myCellArtNr = self.t.sheet.getCellByPosition(iSpalteArtNr, i)
            myCellDicke = self.t.sheet.getCellByPosition(iSpalteDicke, i)
            if myCellArtNr.getType() == CELLCONTENTTYP_TEXT:
                ArtNr = myCellArtNr.String
                if len(ArtNr) > 0:
                    sDicke = ArtNr[0:2] # Zeichen 0 bis 2
                    sDicke += "."
                    sDicke += ArtNr[2:3] # Zeichen 2 bis 3
                    myCellDicke.Value = sDicke
                else:
                    break
        pass
        # Anwendung: self.dicke_aus_artikelnummer_bestimmen()
    def text_zu_zahl_i(self, iSpalte):
        for i in range(1, self.maxistklen):
            myCell = self.t.sheet.getCellByPosition(iSpalte, i)
            if myCell.getType() == CELLCONTENTTYP_TEXT:
                sWert = myCell.String
                if len(sWert) > 0:
                    myCell.Value = sWert # ersetzt den "." mit einem "," als Dezimaltrenner
                else:
                    break
        pass
    def umwandeln_von_BCtoCSV(self):
        # sortiert Stueckliste um
        # von CSV-Format zum Einlesen ins PIOS nach Std-Stuckliste
        # Tabelle kopieren in neue Registerkarte:
        ret = self.t.tab_kopieren("Stueckliste", 99)
        if ret == 0: # keine Fehler
            self.t.set_tabfokus_s("Stueckliste")
            #----Spalten einfügen um Platz zu schaffen:
            self.t.insert_spalte_re_i(0, 15)
            #----Spalten umsortieren:
            # Bezeichnung:
            self.t.spalte_verschieben_i(19, 0)
            # Anzahl:
            self.t.spalte_verschieben_i(20, 1)
            # Länge:
            self.t.spalte_verschieben_i(21, 2)
            # Breite:
            self.t.spalte_verschieben_i(22, 3)
            # Dicke:
            # >>Die Dicke ist nicht in der Tabelle enthalten und muss aus der Artikelnummer berechnet werden
            # Material:
            self.t.spalte_verschieben_i(16, 5)
            # Kante links:
            self.t.spalte_verschieben_i(24, 6)
            # KaDi li:
            self.t.spalte_verschieben_i(29, 7)
            # Kante rechts:
            self.t.spalte_verschieben_i(25, 8)
            # KaDi re:
            self.t.spalte_verschieben_i(30, 9)
            # Kante oben = vorne:
            self.t.spalte_verschieben_i(26, 10)
            # KaDi ob:
            self.t.spalte_verschieben_i(31, 11)
            # Kante unten = hinten:
            self.t.spalte_verschieben_i(27, 12)
            # KaDi un:
            self.t.spalte_verschieben_i(32, 13)
            # Bemerkung:
            self.t.spalte_verschieben_i(28, 14)
            #----nicht gebrauchte Zellen entfernen:
            self.t.delelte_spalten_re_i(15, 100)
            #----Tabellenkopf beschriften :
            self.tabkopf_anlegen()
            #----optimale Zellbreite festlegen:
            self.t.optimale_spaltenbreiten()
            #----Plattendicke aus Artikelnummer berechnen:
            self.dicke_aus_artikelnummer_bestimmen()
            #----dezimaltrennerkorrektur:
            self.text_zu_zahl_i(1)
            self.text_zu_zahl_i(2)
            self.text_zu_zahl_i(3)
            self.text_zu_zahl_i(7)
            self.text_zu_zahl_i(9)
            self.text_zu_zahl_i(11)
            self.text_zu_zahl_i(13)
        return ret
        # Anwendung: self.umwandeln_von_BCtoCSV()
    def formatieren(self):
        # Alle Zellen sichtbar machen:
        for i in range(0, 15):
            self.t.set_spalte_sichtbar_i(i, True)        
        # Zellgrößen anpassen:
        self.t.set_zeilenhoehen(700)
        self.t.set_spaltenbreite_i(0, 3900) # Bezeichnung
        self.t.set_spaltenbreite_i(1, 1410) # Anzahl
        self.t.set_spaltenbreite_i(2, 1320) # Länge
        self.t.set_spaltenbreite_i(3, 1320) # Breite
        self.t.set_spaltenbreite_i(4, 1220) # Dicke
        self.t.set_spaltenbreite_i(5, 3830) # Matieral
        self.t.set_spaltenbreite_i(6, 4300) # Kante links
        self.t.set_spaltenbreite_i(7, 900) # KaDi links
        self.t.set_spaltenbreite_i(8, 4300) # Kante rechts
        self.t.set_spaltenbreite_i(9, 900) # KaDi re
        self.t.set_spaltenbreite_i(10, 4300) # Kante oben
        self.t.set_spaltenbreite_i(11, 900) # KaDi oben
        self.t.set_spaltenbreite_i(12, 4300) # Kante unten
        self.t.set_spaltenbreite_i(13, 900) # KaDi unten
        self.t.set_spaltenbreite_i(14, 5460) # Bemerkung
        # Tabellenkopf farbig machen:
        for i in range(0,15):
            self.t.set_zellfarbe_i(0, i, self.grau)
            pass
        # Zellausrichtung:
        # self.t.set_zellausrichtungHori_s("B1:B1000", "mi")
        pass
    def entferneKaDiNull(self): # Nullen in den Feldern der KaDi löschen:        
        for i in range(1, self.maxistklen):
            wert = self.t.get_zelltext_i(i, 7) # KdDi links
            if wert == "0":
                self.t.set_zelltext_i(i, 7, "")
            wert = self.t.get_zelltext_i(i, 9) # KdDi rechts
            if wert == "0":
                self.t.set_zelltext_i(i, 9, "")
            wert = self.t.get_zelltext_i(i, 11) # KdDi oben
            if wert == "0":
                self.t.set_zelltext_i(i, 11, "")
            wert = self.t.get_zelltext_i(i, 13) # KdDi unten
            if wert == "0":
                self.t.set_zelltext_i(i, 13, "")
        pass
        # Anwendung: self.entferneKaDiNull()
    def autoformat(self):
        # Prüfen, ob Stückliste bereits im richtigen Format vorliegt
        # und ggf zuerst Format umwandeln:
        a = self.t.get_zelltext_s("A1")
        b = self.t.get_zelltext_s("B1")
        c = self.t.get_zelltext_s("C1")
        if a == "Aufkb" and b == "Plakb" and c == "Elnr": # Tabelle von BarcodeToCSV
            anzFehler = self.umwandeln_von_BCtoCSV()
            if anzFehler != 0:
                return False
            self.formatieren()
            return True
        elif len(a) == 0 and len(b) == 0 and len(c) == 0: # Tabellenkopf fehlt, evtl istes ein eleere Tabelle
            self.tabkopf_anlegen()
            self.formatieren()
            return True
        elif a == "Bezeichnung" and b == "Anzahl" and c == "Länge": # Tabelle ist bereits richtig formartiert
            self.formatieren()
            return True
        return False
        # Anwendung: self.autoformat()
    def formeln_edit(self):
        if self.autoformat() != True:
            return False
        # eventuellen vorherigen Inhalt löschen:
        self.t.delelte_spalten_re_i(15, 100)
        # Tabellenkopf ergänzen:
        self.t.set_zelltext_s("Q1", "Anz Fehler")
        self.t.set_zelltext_s("S1", "PlattenDi")
        self.t.set_zelltext_s("T1", "KD li")
        self.t.set_zelltext_s("U1", "KD re")
        self.t.set_zelltext_s("V1", "KD ob")
        self.t.set_zelltext_s("W1", "KD un")
        self.t.set_zelltext_s("X1", "Anz 0")
        self.t.set_zelltext_s("Y1", "L<70")
        self.t.set_zelltext_s("Z1", "B<70")
        self.t.set_zelltext_s("AA1", "BC zu lang")
        # --
        self.t.set_zelltext_s("AC1", "KD li")
        self.t.set_zelltext_s("AD1", "KD re")
        self.t.set_zelltext_s("AE1", "KD ob")
        self.t.set_zelltext_s("AF1", "KD un")
        # --
        self.t.set_zelltext_s("P1", "Projekt")
        self.t.set_zelltext_s("P2", "ABC01")
        self.t.set_zellfarbe_s("P2", self.gelb)
        # Breiten der ergänzten Spalten anpassen:
        for i in range (15, 35):
            self.t.optimale_spaltenbreite_i(i)
        self.t.set_spaltenbreite_i(17, 700)
        # Nullen in den Feldern der KaDi löschen:
        self.entferneKaDiNull()
        # Formeln einfügen:
        # Es müssen immer die englischen Funktionsnamen für die Calc-Funktionen verwendet werden!
        for i in range (1, 25):
            sZellname = "Q" + str(i+1)
            sFormel = "=IF(SUM(S" + str(i+1) + ":AB" + str(i+1) + ")=0;0;" + "\"Fehler\"" + ")"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- Anzahl der Fehler:
            sZellname = "R" + str(i+1)
            sFormel = "=SUM(S" + str(i+1) + ":AB" + str(i+1) + ")"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- PlattenDi:
            sZellname = "S" + str(i+1)
            sFormel = "=IF(ISBLANK((INDIRECT(" + "\"F\"" + "&ROW())));0;(IF(NUMBERVALUE((CONCATENATE(LEFT((INDIRECT(" + "\"F\"" + "&ROW()));2);" + "\",\"" + ";RIGHT(LEFT((INDIRECT(" + "\"F\"" + "&ROW()));3);1)));" + "\",\"" + ")=(INDIRECT(" + "\"E\"" + "&ROW()));;1)))"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi links:
            sZellname = "T" + str(i+1)
            # Formel für alte Kantennummer:
            #sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += "+" # Jetzt folg nächste Prüfung:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            # Formel für aktuelle Kantennummer:
            sFormel = "=IF(LEN(INDIRECT(" + "\"G\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "0" # Nichts tun
            sFormel += ";" # Sonst
            sFormel += "(IF(INDIRECT(" + "\"H\"" + "&ROW())" # Wenn der Wert
            sFormel += "=((NUMBERVALUE(LEFT(INDIRECT(" + "\"G\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"G\"" + "&ROW());5);2)))/10)" # minus Fügemaß
            sFormel += ";" # Dann
            sFormel += "0" # Kein Fehler
            sFormel += ";" # Sonst
            sFormel += "1)))" # Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi rechts:
            sZellname = "U" + str(i+1)
            # Formel für alte Kantennummer:
            #sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += "+" # Jetzt folg nächste Prüfung:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            # Formel für aktuelle Kantennummer:
            sFormel = "=IF(LEN(INDIRECT(" + "\"I\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "0" # Nichts tun
            sFormel += ";" # Sonst
            sFormel += "(IF(INDIRECT(" + "\"J\"" + "&ROW())" # Wenn der Wert
            sFormel += "=((NUMBERVALUE(LEFT(INDIRECT(" + "\"I\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"I\"" + "&ROW());5);2)))/10)" # minus Fügemaß
            sFormel += ";" # Dann
            sFormel += "0" # Kein Fehler
            sFormel += ";" # Sonst
            sFormel += "1)))" # Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi oben:
            sZellname = "V" + str(i+1)
            # Formel für alte Kantennummer:
            #sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += "+" # Jetzt folg nächste Prüfung:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            # Formel für aktuelle Kantennummer:
            sFormel = "=IF(LEN(INDIRECT(" + "\"K\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "0" # Nichts tun
            sFormel += ";" # Sonst
            sFormel += "(IF(INDIRECT(" + "\"L\"" + "&ROW())" # Wenn der Wert
            sFormel += "=((NUMBERVALUE(LEFT(INDIRECT(" + "\"K\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"K\"" + "&ROW());5);2)))/10)" # minus Fügemaß
            sFormel += ";" # Dann
            sFormel += "0" # Kein Fehler
            sFormel += ";" # Sonst
            sFormel += "1)))" # Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi unten:
            sZellname = "W" + str(i+1)
            # Formel für alte Kantennummer:
            #sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += "+" # Jetzt folg nächste Prüfung:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            #sFormel += ";" # Sonst Wenn:
            #sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            #sFormel += ";" # Dann
            #sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            # Formel für aktuelle Kantennummer:
            sFormel = "=IF(LEN(INDIRECT(" + "\"M\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "0" # Nichts tun
            sFormel += ";" # Sonst
            sFormel += "(IF(INDIRECT(" + "\"N\"" + "&ROW())" # Wenn der Wert
            sFormel += "=((NUMBERVALUE(LEFT(INDIRECT(" + "\"M\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"M\"" + "&ROW());5);2)))/10)" # minus Fügemaß
            sFormel += ";" # Dann
            sFormel += "0" # Kein Fehler
            sFormel += ";" # Sonst
            sFormel += "1)))" # Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- Anz 0:
            sZellname = "X" + str(i+1)
            sFormel = "=IF(ISBLANK(INDIRECT(" + "\"B\"" + "&ROW()));0;(IF(INDIRECT(" + "\"B\"" + "&ROW())=0;1;0)))" # Wenn Anz leer ist oder 0 dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- L < 70:
            sZellname = "Y" + str(i+1)
            sFormel = "=IF(INDIRECT(" + "\"C\"" + "&ROW())<70;IF(ISBLANK(INDIRECT(" + "\"C\"" + "&ROW()));0;1);0)" # Wenn L < 70 dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- B < 70:
            sZellname = "Z" + str(i+1)
            sFormel = "=IF(INDIRECT(" + "\"D\"" + "&ROW())<70;IF(ISBLANK(INDIRECT(" + "\"D\"" + "&ROW()));0;1);0)" # Wenn L < 70 dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- BC zu lang:
            sZellname = "AA" + str(i+1)
            sFormel = "=IF((LEN(P$2)+LEN(INDIRECT(" + "\"A\"" + "&ROW()))+6)>25;1;0)" # Wenn BC > 28 dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi links korrekter Wert:
            sZellname = "AC" + str(i+1)
            sFormel = "=IF(LEN(INDIRECT(" + "\"G\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "\"---\"" # Kein Ergebnis
            sFormel += ";" # Sonst
            sFormel += "((NUMBERVALUE(LEFT(INDIRECT(" + "\"G\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"G\"" + "&ROW());5);2)))/10))" # minus Fügemaß
            self.t.set_zellformel_s(sZellname, sFormel)
            self.t.set_zellausrichtungHori_s(sZellname, "mi")
            # --- KaDi rechts korrekter Wert:
            sZellname = "AD" + str(i+1)
            sFormel = "=IF(LEN(INDIRECT(" + "\"I\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "\"---\"" # Kein Ergebnis
            sFormel += ";" # Sonst
            sFormel += "((NUMBERVALUE(LEFT(INDIRECT(" + "\"I\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"I\"" + "&ROW());5);2)))/10))" # minus Fügemaß
            self.t.set_zellformel_s(sZellname, sFormel)
            self.t.set_zellausrichtungHori_s(sZellname, "mi")
            # --- KaDi oben korrekter Wert:
            sZellname = "AE" + str(i+1)
            sFormel = "=IF(LEN(INDIRECT(" + "\"K\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "\"---\"" # Kein Ergebnis
            sFormel += ";" # Sonst
            sFormel += "((NUMBERVALUE(LEFT(INDIRECT(" + "\"K\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"K\"" + "&ROW());5);2)))/10))" # minus Fügemaß
            self.t.set_zellformel_s(sZellname, sFormel)
            self.t.set_zellausrichtungHori_s(sZellname, "mi")
            # --- KaDi unten korrekter Wert:
            sZellname = "AF" + str(i+1)
            sFormel = "=IF(LEN(INDIRECT(" + "\"M\"" + "&ROW()))<5" #Wenn Kanteninfo aus weniger als 5 Zeichen besteht
            sFormel += ";" # Dann
            sFormel += "\"---\"" # Kein Ergebnis
            sFormel += ";" # Sonst
            sFormel += "((NUMBERVALUE(LEFT(INDIRECT(" + "\"M\"" + "&ROW());3))" # Kantendicke ist
            sFormel += "-NUMBERVALUE(RIGHT(LEFT(INDIRECT(" + "\"M\"" + "&ROW());5);2)))/10))" # minus Fügemaß
            self.t.set_zellformel_s(sZellname, sFormel)
            self.t.set_zellausrichtungHori_s(sZellname, "mi")
        pass
        # Anwendung: self.formeln_edit()
    def formeln_kante(self):
        if self.autoformat() != True:
            return False
        # eventuell noch umbenennen:
        sTabName = self.t.get_tabname()
        if(sTabName == "Stueckliste"):
            self.t.tab_setName("Kantenberechnung")
        # eventuellen vorherigen Inhalt löschen:
        self.t.delelte_spalten_re_i(15, 100)
        maxi = 0        
        # KaDi ausblenden:
        self.t.set_spalte_sichtbar_i(7,False)
        self.t.set_spalte_sichtbar_i(9,False)
        self.t.set_spalte_sichtbar_i(11,False)
        self.t.set_spalte_sichtbar_i(13,False)
        # Tabellenkopf ergänzen:
        self.t.set_zelltext_s("P1", "zu kurz")
        self.t.set_zelltext_s("Q1", "KantenNr")
        self.t.set_zelltext_s("R1", "lfdm")
        self.t.set_zelltext_s("S1", " = ca.")
        self.t.set_spaltenbreite_i(16, 4500) # KantenNr
        self.t.set_spaltenbreite_i(17, 1500) # lfdm
        self.t.set_spaltenbreite_i(18, 1500) # lfdm ca
        self.t.set_spaltenausrichtung_i(18, "mi")
        # Kantensorten ermitteln:
        aKanten = [] # leere Liste
        iSpalteBez = 0
        iSpalteMat = 5
        iSpalteKaLi = 6
        iSpalteKaRe = 8
        iSpalteKaOb = 10
        iSpalteKaUn = 12
        for i in range (1, self.maxistklen):
            myCellBez = self.t.get_zelle_i(i, iSpalteBez)
            myCellMat = self.t.get_zelle_i(i, iSpalteMat)
            myCellKaLi = self.t.get_zelle_i(i, iSpalteKaLi)
            myCellKaRe = self.t.get_zelle_i(i, iSpalteKaRe)
            myCellKaOb = self.t.get_zelle_i(i, iSpalteKaOb)
            myCellKaUn = self.t.get_zelle_i(i, iSpalteKaUn)
            if (len(myCellBez.String) > 0) or (len(myCellMat.String) > 0) or (i < 10):
                sKaLi = myCellKaLi.String
                sKaRe = myCellKaRe.String
                sKaOb = myCellKaOb.String
                sKaUn = myCellKaUn.String
                if (len(sKaLi) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaLi:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        if len(sKaLi) > 0:
                            aKanten.append(sKaLi)
                if (len(sKaRe) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaRe:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        if len(sKaRe) > 0:
                            aKanten.append(sKaRe)
                if (len(sKaOb) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaOb:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        if len(sKaOb) > 0:
                            aKanten.append(sKaOb)
                if (len(sKaUn) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaUn:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        if len(sKaUn) > 0:
                            aKanten.append(sKaUn)
                maxi += 1
            else: 
                break # für For-Schleife i
            pass
        for i in range (0, len(aKanten)):
            # Kantennummer:
            self.t.set_zelltext_i(i+1, 16, aKanten[i])
            # lfdm:
            formel =  "=("
            formel += "SUMPRODUCT((G$2:G$1000=Q" + str(i+2) + ")*IF((C$2:C$1000+50)<320;320;C$2:C$1000+50)*(B$2:B$1000))"
            formel += "+"
            formel += "SUMPRODUCT((I$2:I$1000=Q" + str(i+2) + ")*IF((C$2:C$1000+50)<320;320;C$2:C$1000+50)*(B$2:B$1000))"
            formel += "+"
            formel += "SUMPRODUCT((K$2:K$1000=Q" + str(i+2) + ")*IF((D$2:D$1000+50)<320;320;D$2:D$1000+50)*(B$2:B$1000))"
            formel += "+"
            formel += "SUMPRODUCT((M$2:M$1000=Q" + str(i+2) + ")*IF((D$2:D$1000+50)<320;320;D$2:D$1000+50)*(B$2:B$1000))"
            formel += ")/1000"
            self.t.set_zellformel_i(i+1, 17, formel)
            # lfdm gerundet:
            formel = "=ROUNDUP(R" + str(i+2) + "/5;0)*5"
            self.t.set_zellformel_i(i+1, 18, formel)
            # Kantennummer Ostermann:
            formel =  "=LEFT(RIGHT(Q" + str(i+2) + ";LEN(Q" + str(i+2) + ")-12);3)&\".\"&RIGHT(Q" + str(i+2) + ";LEN(Q"
            formel += str(i+2) + ")-15)&\".\"&RIGHT(LEFT(Q" + str(i+2) + ";9);3)&RIGHT(LEFT(Q" + str(i+2) + ";3);2)"
            self.t.set_zellformel_i(i+1, 19, formel)
            pass
        # Formeln für Kantenfehler einfügen:
        self.t.set_spaltenausrichtung_i(15, "mi")
        for i in range (1, maxi+1):
            formel =  "=IF(C" + str(i+1) + "<240;IF(NOT(G" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<240;IF(NOT(I" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<240;IF(NOT(K" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<240;IF(NOT(M" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<90;IF(NOT(K" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<90;IF(NOT(M" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<90;IF(NOT(G" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<90;IF(NOT(I" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            self.t.set_zellformel_i(i, 15, formel)
            sErgebnis = self.t.get_zelltext_i(i, 15)
            if(sErgebnis != "0") and (len(sErgebnis) >0 ):
                self.t.set_zellfarbe_i(i, 2, self.rot) # Länge
                self.t.set_zellfarbe_i(i, 3, self.rot) # Breite
                self.t.set_zellfarbe_i(i, 15, self.rot) # Formel
            else:
                self.t.set_zellfarbe_i(i, 2, self.farblos) # Länge
                self.t.set_zellfarbe_i(i, 3, self.farblos) # Breite
                self.t.set_zellfarbe_i(i, 15, self.farblos) # Formel
            pass
        # Bemerkungen mit Kanteninfo farbig machen:
        for i in range (1, maxi+1):
            sZelltext = self.t.get_zelltext_i(i, 14)
            iGefunden = 0
            if "K10" in sZelltext:
                iGefunden += 1
            if "K20" in sZelltext:
                iGefunden += 1
            if "K30" in sZelltext:
                iGefunden += 1
            if "K05" in sZelltext:
                iGefunden += 1
            if "K08" in sZelltext:
                iGefunden += 1
            if iGefunden > 0:
                self.t.set_zellfarbe_i(i, 14, self.gelb)
            else:
                self.t.set_zellfarbe_i(i, 14, self.farblos)
            pass
        pass
        # Anwendung: self.formeln_kante()
    def kanteninfo_beraeumen(self):
        badStrings = ["Ger", "Gehr", "Zugabe", "Schräg", "Schmiege", "DA"]
        for i in range(1, self.maxistklen): # Schleife beginnt unter den Tabellenkopf
            for ii in range(0, len(badStrings)): # Alle Argumente von badStrings[] durchlaufen
                sZelltextLi = self.t.get_zelltext_i(i, 6) # Kante links
                sZelltextRe = self.t.get_zelltext_i(i, 8) # Kante rechts
                sZelltextOb = self.t.get_zelltext_i(i, 10) # Kante oben
                sZelltextUn = self.t.get_zelltext_i(i, 12) # Kante unten
                if badStrings[ii] in sZelltextLi:
                    self.t.set_zelltext_i(i, 6, "")
                if badStrings[ii] in sZelltextRe:
                    self.t.set_zelltext_i(i, 8, "")
                if badStrings[ii] in sZelltextOb:
                    self.t.set_zelltext_i(i, 10, "")
                if badStrings[ii] in sZelltextUn:
                    self.t.set_zelltext_i(i, 12, "")
                pass
            pass
        pass
        # Anwendung: self.kanteninfo_beraeumen()
    def teil_drehen(self):
        iZeileStart = self.t.get_selection_zeile_start()
        iZeileEnde  = self.t.get_selection_zeile_ende()
        tmpPosDiff = 5000
        iPosLaenge = 2
        iPosBreite = 3
        iPosKanteLi = 6
        iPosKaDiLi = 7
        iPosKanteRe = 8
        iPosKaDiRe = 9
        iPosKanteOb = 10
        iPosKaDiOb = 11
        iPosKanteUn = 12
        iPosKaDiUn = 13
        for i in range(iZeileStart, iZeileEnde+1):
            # Zellen nach unten verschieben:
            self.t.zelle_verschieben_i(i, iPosLaenge, i+tmpPosDiff, iPosLaenge)
            self.t.zelle_verschieben_i(i, iPosBreite, i+tmpPosDiff, iPosBreite)
            self.t.zelle_verschieben_i(i, iPosKanteLi, i+tmpPosDiff, iPosKanteLi)
            self.t.zelle_verschieben_i(i, iPosKaDiLi, i+tmpPosDiff, iPosKaDiLi)
            self.t.zelle_verschieben_i(i, iPosKanteRe, i+tmpPosDiff, iPosKanteRe)
            self.t.zelle_verschieben_i(i, iPosKaDiRe, i+tmpPosDiff, iPosKaDiRe)
            self.t.zelle_verschieben_i(i, iPosKanteOb, i+tmpPosDiff, iPosKanteOb)
            self.t.zelle_verschieben_i(i, iPosKaDiOb, i+tmpPosDiff, iPosKaDiOb)
            self.t.zelle_verschieben_i(i, iPosKanteUn, i+tmpPosDiff, iPosKanteUn)
            self.t.zelle_verschieben_i(i, iPosKaDiUn, i+tmpPosDiff, iPosKaDiUn)
            # Zellen mit neuer Spaltenzuweisung wieder zurück verschieben
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosLaenge, i, iPosBreite) # L -> B
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosBreite, i, iPosLaenge) # B -> L
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKanteUn, i, iPosKanteLi) # Li -> Un
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKanteRe, i, iPosKanteUn) # Un -> Re
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKanteOb, i, iPosKanteRe) # Re -> Ob
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKanteLi, i, iPosKanteOb) # Ob -> Li
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKaDiUn, i, iPosKaDiLi) # Li -> Un
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKaDiRe, i, iPosKaDiUn) # Un -> Re
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKaDiOb, i, iPosKaDiRe) # Re -> Ob
            self.t.zelle_verschieben_i(i+tmpPosDiff, iPosKaDiLi, i, iPosKaDiOb) # Ob -> Li
            pass
        pass
        # Anwendung: self.teil_drehen()
    def sortieren(self):
        rankingList  = ["Seite_li", "Seite_re", "Seite"]
        rankingList += ["MS_li", "MS_re", "MS"]
        rankingList += ["OB_li", "OB_mi", "OB_re", "OB"]
        rankingList += ["UB_li", "UB_mi", "UB_re", "UB"]
        rankingList += ["KB_ob", "KB_li", "KB_mi", "KB_un", "KB_re", "KB"]
        rankingList += ["Trav_ob", "Trav_un", "Trav_vo", "Trav_hi", "Trav"]
        rankingList += ["Traver_ob", "Traver_un", "Traver_vo", "Traver_hi", "Traver"]
        rankingList += ["Trav_ob", "Trav_un", "Trav_vo", "Trav_hi", "Trav"]
        rankingList += ["EB_ob", "EB_li", "EB_mi", "EB_un", "EB_re", "EB"]
        rankingList += ["RW_ob", "RW_li", "RW_mi", "RW_un", "RW_re", "RW"]
        rankingList += ["Tuer_li", "Tuer_re", "Tuer_A", "Tuer_B", "Tuer_C", "Tuer_D", "Tuer_E", "Tuer"]
        rankingList += ["Front_li", "Front_re", "Front_A", "Front_B", "Front_C", "Front_D", "Front_E", "Front"]
        rankingList += ["Klappe"]
        rankingList += ["SF_A", "SF_B", "SF_C", "SF_D", "SF_E", "SF"]
        rankingList += ["SS_A", "SS_B", "SS_C", "SS_D", "SS_E", "SS"]
        rankingList += ["SV_A", "SV_B", "SV_C", "SV_D", "SV_E", "SV"]
        rankingList += ["SH_A", "SH_B", "SH_C", "SH_D", "SH_E", "SH"]
        rankingList += ["SB_A", "SB_B", "SB_C", "SB_D", "SB_E", "SB"]
        rankingList += ["Sockel_li", "Sockel_mi", "Sockel_re" ,"Sockel"]
        rankingNum = [] # Speichert das Ranking für die jeweilige Zeile
        rankingVonZeile = [] # Speichert die ursprüngliche Zeilennummer
        iZeileStart = self.t.get_selection_zeile_start()
        iZeileEnde  = self.t.get_selection_zeile_ende()
        tmpPosDiff = 5000
        # Ranking ermitteln
        for i in range(iZeileStart, iZeileEnde+1):
            sName = self.t.get_zelltext_i(i, 0)
            rankingVonZeile.append(i) # Ursprungszeile merken
            maxRanking = 99
            iRanking = maxRanking
            for ii in range(0, len(rankingList)):
                if rankingList[ii] in sName:
                    iRanking = ii
                    break # for ii
                pass
            rankingNum.append(iRanking)
            pass
        # Zellen nach unten verschieben:
        source = self.t.sheet.getCellRangeByPosition(0, iZeileStart, 14, iZeileEnde)
        target = self.t.sheet.getCellByPosition(0, iZeileStart+tmpPosDiff)
        self.t.sheet.moveRange(target.CellAddress, source.RangeAddress)
        # Zellen in der richtigen Reihenfolge  wieder nach oben verschieben:
        naechsteZeile = iZeileStart
        for i in range(0, maxRanking+1):# vom kleinen == guten Ranking zum schlechten Ranking
            if i in rankingNum:
                last_ii = 0
                for ii in range(last_ii, len(rankingNum)):
                    if rankingNum[ii] == i:
                        # Zeile nach oben verschieben
                        source = self.t.sheet.getCellRangeByPosition(0, rankingVonZeile[ii]+tmpPosDiff, 14, rankingVonZeile[ii]+tmpPosDiff)
                        target = self.t.sheet.getCellByPosition(0, naechsteZeile)
                        self.t.sheet.moveRange(target.CellAddress, source.RangeAddress)
                        naechsteZeile += 1
                        last_ii = ii
                    pass
            pass
        pass
    def reduzieren(self):
        iZeileStart = self.t.get_selection_zeile_start()
        iZeileEnde  = self.t.get_selection_zeile_ende()
        list_zeiNum = [iZeileStart] # int
        list_bez = [self.t.get_zelltext_i(iZeileStart, 0)]
        list_anz = [self.t.get_zellzahl_i(iZeileStart, 1)] # int
        list_L   = [self.t.get_zelltext_i(iZeileStart, 2)]
        list_B   = [self.t.get_zelltext_i(iZeileStart, 3)]
        list_D   = [self.t.get_zelltext_i(iZeileStart, 4)]
        list_mat = [self.t.get_zelltext_i(iZeileStart, 5)]
        list_KaLi = [self.t.get_zelltext_i(iZeileStart, 6)]
        list_KDLi = [self.t.get_zelltext_i(iZeileStart, 7)]
        list_KaRe = [self.t.get_zelltext_i(iZeileStart, 8)]
        list_KDRe = [self.t.get_zelltext_i(iZeileStart, 9)]
        list_KaOb = [self.t.get_zelltext_i(iZeileStart, 10)]
        list_KDOb = [self.t.get_zelltext_i(iZeileStart, 11)]
        list_KaUn = [self.t.get_zelltext_i(iZeileStart,12)]
        list_KDUn = [self.t.get_zelltext_i(iZeileStart, 13)]
        list_kom = [self.t.get_zelltext_i(iZeileStart, 14)]
        for i in range(iZeileStart+1, iZeileEnde+1):
            bez = self.t.get_zelltext_i(i, 0)
            gefunden = False
            if bez in list_bez:
                for ii in range(0, len(list_bez)):
                    if bez == list_bez[ii]:                        
                        L = self.t.get_zelltext_i(i, 2)
                        B = self.t.get_zelltext_i(i, 3)
                        D = self.t.get_zelltext_i(i, 4)
                        mat = self.t.get_zelltext_i(i, 5)
                        KaLi = self.t.get_zelltext_i(i, 6)
                        KDLi = self.t.get_zelltext_i(i, 7)
                        KaRe = self.t.get_zelltext_i(i, 8)
                        KDRe = self.t.get_zelltext_i(i, 9)
                        KaOb = self.t.get_zelltext_i(i, 10)
                        KDOb = self.t.get_zelltext_i(i, 11)
                        KaUn = self.t.get_zelltext_i(i, 12)
                        KDUn = self.t.get_zelltext_i(i, 13)
                        kom = self.t.get_zelltext_i(i, 14)
                        if( L == list_L[ii] ):
                            if( B == list_B[ii] ):
                                if( D == list_D[ii] ):
                                    if( mat == list_mat[ii] ):
                                        if( KaLi == list_KaLi[ii] ):
                                            if( KDLi == list_KDLi[ii] ):
                                                if( KaRe == list_KaRe[ii] ):
                                                    if( KDRe == list_KDRe[ii] ):
                                                        if( KaOb == list_KaOb[ii] ):
                                                            if( KDOb == list_KDOb[ii] ):
                                                                if( KaUn == list_KaUn[ii] ):
                                                                    if( KDUn == list_KDUn[ii] ):
                                                                        if( kom == list_kom[ii] ):                                                                            
                                                                            gefunden = True
                                                                            # Zeilen zusammenführen:
                                                                            neueAnz = list_anz[ii] + self.t.get_zellzahl_i(i, 1)
                                                                            list_anz[ii] = neueAnz
                                                                            self.t.set_zellzahl_i(list_zeiNum[ii], 1, neueAnz)
                                                                            # Zeileninhalt überschreiben / Dublette löschen:
                                                                            source = self.t.sheet.getCellRangeByPosition(0, 9999, 15, 9999)
                                                                            target = self.t.sheet.getCellByPosition(0, i)
                                                                            self.t.sheet.moveRange(target.CellAddress, source.RangeAddress)
                    pass
            if gefunden == False:
                list_zeiNum.append(i)
                list_bez.append(bez)
                list_anz.append(self.t.get_zellzahl_i(i, 1)) # int
                list_L.append(self.t.get_zelltext_i(i, 2))
                list_B.append(self.t.get_zelltext_i(i, 3))
                list_D.append(self.t.get_zelltext_i(i, 4))
                list_mat.append(self.t.get_zelltext_i(i, 5))
                list_KaLi.append(self.t.get_zelltext_i(i, 6))
                list_KDLi.append(self.t.get_zelltext_i(i, 7))
                list_KaRe.append(self.t.get_zelltext_i(i, 8))
                list_KDRe.append(self.t.get_zelltext_i(i, 9))
                list_KaOb.append(self.t.get_zelltext_i(i, 10))
                list_KDOb.append(self.t.get_zelltext_i(i, 11))
                list_KaUn.append(self.t.get_zelltext_i(i, 12))
                list_KDUn.append(self.t.get_zelltext_i(i, 13))
                list_kom.append(self.t.get_zelltext_i(i, 14))
            pass
        pass
    def std_namen(self):
        iZeileStart = self.t.get_selection_zeile_start()
        iZeileEnde  = self.t.get_selection_zeile_ende()
        for i in range(iZeileStart, iZeileEnde+1):
            sName = self.t.get_zelltext_i(i, 0)     
            sName = sName.replace("Seite Links", "S#_Seite_li")    
            sName = sName.replace("Seite Rechts", "S#_Seite_re")    
            sName = sName.replace("Mittelseite", "S#_MS")
            sName = sName.replace("Boden Oben", "S#_OB")
            sName = sName.replace("Boden Unten", "S#_UB")
            sName = sName.replace("Konstruktionsboden", "S#_KB")
            sName = sName.replace("Fachboden", "S#_EB")
            sName = sName.replace("Rückwand", "S#_RW")
            sName = sName.replace("Tür", "S#_Tuer")
            sName = sName.replace("Klappe", "S#_Klappe")
            sName = sName.replace("Blindfront", "S#_Front")
            sName = sName.replace("Schubkasten Front", "S#_SF")
            sName = sName.replace("Travers Vorne", "S#_Trav_vo")
            sName = sName.replace("Travers Hinten", "S#_Trav_hi")
            sName = sName.replace("Travers Oben Vorne", "S#_Trav_OV")
            sName = sName.replace("Travers Oben Hinten", "S#_Trav_OH")
            #sName = sName.replace("", "S#_")
            self.t.set_zelltext_i(i, 0, sName)
            pass
        pass
    def gehr_masszugabe(self):
        iZeileStart = self.t.get_selection_zeile_start()
        iZeileEnde  = self.t.get_selection_zeile_ende()
        iIndexSpalte_L = 2
        iIndexSpalte_B = 3
        iIndexSpalte_KaLi = 6
        iIndexSpalte_KaRe = 8
        iIndexSpalte_KaOb = 10
        iIndexSpalte_KaUn = 12
        for i in range(iZeileStart, iZeileEnde+1):
            sLaenge = self.t.get_zelltext_i(i, iIndexSpalte_L)
            sBreite = self.t.get_zelltext_i(i, iIndexSpalte_B)
            sKaLi = self.t.get_zelltext_i(i, iIndexSpalte_KaLi)
            sKaRe = self.t.get_zelltext_i(i, iIndexSpalte_KaRe)
            sKaOb = self.t.get_zelltext_i(i, iIndexSpalte_KaOb)
            sKaUn = self.t.get_zelltext_i(i, iIndexSpalte_KaUn)
            iZugabe_L = 0
            iZugabe_B = 0
            if "Gehr" in sKaLi:
                iZugabe_B += 20
            if "Gehr" in sKaRe:
                iZugabe_B += 20
            if "Gehr" in sKaOb:
                iZugabe_L += 20
            if "Gehr" in sKaUn:
                iZugabe_L += 20
            if iZugabe_L > 0:
                sNeue_L = "=" + sLaenge + "+" + str(iZugabe_L)
                sNeue_L = sNeue_L.replace(",", ".")
                self.t.set_zellformel_i(i, iIndexSpalte_L, sNeue_L)
                self.t.set_zellfarbe_i(i, iIndexSpalte_L, self.gelb)
            if iZugabe_B > 0:                
                sNeue_B = "=" + sBreite + "+" + str(iZugabe_B)
                sNeue_B = sNeue_B.replace(",", ".")
                self.t.set_zellformel_i(i, iIndexSpalte_B, sNeue_B)
                self.t.set_zellfarbe_i(i, iIndexSpalte_B, self.gelb)
            pass #for
        pass
    def tap_anlegen_uebersicht(self):
        tabname = "Übersicht"
        if not self.t.tab_existiert(tabname):            
            self.t.tab_anlegen(tabname, 0)
        else:
            self.t.set_tabfokus_s(tabname)
        if self.t.tab_existiert(tabname):
            self.t.set_tabfokus_s(tabname)
            self.t.set_zelltext_s("A1", "Projekt:")
            self.t.set_zelltext_s("B1", "Abc01")
            self.t.set_zelltext_s("C1", "CNC-Pfad:")
            self.t.set_zelltext_s("D1", self.cnc_pfad)
            self.t.set_zelltext_s("A3", "Pos")
            self.t.set_zelltext_s("B3", "Bezeichnung")
            self.t.set_spaltenbreite_i(1, 6000) # Bezeichnung
            self.t.set_zellausrichtungHori_s("A3:A1000", "mi")
        pass
    def tab_anlegen_stklistpos(self):
        akt_tabname = self.t.get_tabname()
        if akt_tabname == "Übersicht":
            projektnummer = self.t.get_zelltext_s("B1")
            tabname = self.t.get_zelltext_akt_auswahl() # Positionsnummer
            if self.t.tab_existiert(tabname):
                self.t.set_tabfokus_s(tabname)
            else:
                self.t.tab_anlegen(tabname, 9999)
                if self.t.tab_existiert(tabname):
                    self.t.set_tabfokus_s(tabname)
                    self.formeln_edit()
                    self.t.set_zelltext_s("P2", projektnummer)
        else:
            titel = "Bediener-Fehler"
            msg   = "Bitte in die Registerkarte \"Grundlagen\" wechseln und die Zelle mit der gewünschten Positionsnummer anklicken um diese Funktion zu nutzen"
            msgbox(msg, titel, 1, 'QUERYBOX')
        pass
    def check_cncdata(self):
        if self.autoformat() == True:
            projekt = self.t.get_zelltext_s("P2")
            pos_nr = self.t.get_tabname()
            grundpfad = self.cnc_pfad
            grundpfad += "\\"
            grundpfad += projekt
            if os.path.isdir(grundpfad):
                grundpfad += "\\"
                grundpfad += self.posnr_formartieren(pos_nr)
                if os.path.isdir(grundpfad):
                    anz_leerzeilen = 0
                    for i in range (1, self.maxistklen):
                        bezeichnung = self.t.get_zelltext_i(i, 0)
                        if len(bezeichnung) == 0:
                            anz_leerzeilen = anz_leerzeilen + 1
                            if anz_leerzeilen > 20:
                                break #for
                        else: # Bezeichnung ist nicht leer
                            baugruppe = self.baugruppe(bezeichnung)
                            wstname = "0"
                            akt_pfad = grundpfad
                            if len(baugruppe) == 0: # es gibt keine Baugruppe/Schranknummer
                                wstname = bezeichnung
                                akt_pfad = grundpfad
                            else: # es gibt eine Baugruppe/Schranknummer
                                wstname = bezeichnung[len(baugruppe)+1:]
                                akt_pfad = grundpfad
                                akt_pfad += "\\"
                                akt_pfad += baugruppe
                            akt_datei  = akt_pfad
                            akt_datei += "\\"
                            akt_datei += wstname
                            akt_datei += ".ppf"
                            if os.path.isfile(akt_datei):
                                self.t.set_zellfarbe_i(i,0, self.gruen)
                                tabindex_la = 2
                                tabindex_br = 3
                                tabindex_di = 4
                                tabindex_la_s = "C"
                                tabindex_br_s = "D"
                                tabindex_di_s = "E"
                                slist_la = float(self.t.get_zelltext_i(i, tabindex_la))
                                slist_br = float(self.t.get_zelltext_i(i, tabindex_br))
                                slist_di = float(self.t.get_zelltext_i(i, tabindex_di))
                                cnc_la = float(self.ppf_wst_laenge(akt_datei))
                                cnc_br = float(self.ppf_wst_breite(akt_datei))
                                cnc_di = float(self.ppf_wst_dicke(akt_datei))
                                rahmenbreite = 70
                                if(slist_la == cnc_la):
                                    self.t.set_Rahmen_s(tabindex_la_s + str(i+1), rahmenbreite, self.gruen)
                                    if(slist_br == cnc_br):
                                        self.t.set_Rahmen_s(tabindex_br_s + str(i+1), rahmenbreite, self.gruen)
                                if(slist_br == cnc_la):
                                    self.t.set_Rahmen_s(tabindex_br_s + str(i+1), rahmenbreite, self.gruen)
                                    if(slist_la == cnc_br):
                                        self.t.set_Rahmen_s(tabindex_la_s + str(i+1), rahmenbreite, self.gruen)
                                if(slist_di == cnc_di):
                                    self.t.set_Rahmen_s(tabindex_di_s + str(i+1), rahmenbreite, self.gruen)                            
                else: # Ordner für Projektpos nicht gefunden
                    titel = "Klasse: slist, Funktion: check_wstmass()"
                    msg   = "Ordner wurde nicht gefunden!"
                    msg  += "\n"
                    msg  += grundpfad
                    msgbox(msg, titel, 1, 'QUERYBOX')
            else: # Projektordner nicht gefunden
                titel = "Klasse: slist, Funktion: check_wstmass()"
                msg   = "Projektordner wurde nicht gefunden!"
                msg  += "\n"
                msg  += grundpfad
                msgbox(msg, titel, 1, 'QUERYBOX')
        else: # ist keine Stückliste
            titel = "Klasse: slist, Funktion: check_wstmass()"
            msg   = "Die Tabelle ist keine Stückliste. Die Funktion wird nicht ausgeführt."
            msgbox(msg, titel, 1, 'QUERYBOX')
        pass
    def posnr_formartieren(self, posnr):
        formartierte_posnr = "0"
        if "," in posnr:
            index = posnr.find(",")
            posnr_ohne_komma = posnr[0:index]
            nachkommastellen = posnr[index+1:]
            if len(posnr_ohne_komma) == 4:
                formartierte_posnr = posnr
            elif len(posnr_ohne_komma) == 3:
                formartierte_posnr = "0"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += ","
                formartierte_posnr += nachkommastellen
            elif len(posnr_ohne_komma) == 2:
                formartierte_posnr = "00"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += ","
                formartierte_posnr += nachkommastellen
            elif len(posnr_ohne_komma) == 1:
                formartierte_posnr = "000"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += ","
                formartierte_posnr += nachkommastellen
        elif "." in posnr:
            index = posnr.find(".")
            posnr_ohne_komma = posnr[0:index]
            nachkommastellen = posnr[index+1:]
            if len(posnr_ohne_komma) == 4:
                formartierte_posnr = posnr
            elif len(posnr_ohne_komma) == 3:
                formartierte_posnr = "0"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += "."
                formartierte_posnr += nachkommastellen
            elif len(posnr_ohne_komma) == 2:
                formartierte_posnr = "00"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += "."
                formartierte_posnr += nachkommastellen
            elif len(posnr_ohne_komma) == 1:
                formartierte_posnr = "000"
                formartierte_posnr += posnr_ohne_komma
                formartierte_posnr += "."
                formartierte_posnr += nachkommastellen
        else:
            if len(posnr) == 4:
                formartierte_posnr = posnr
            elif len(posnr) == 3:
                formartierte_posnr = "0"
                formartierte_posnr += posnr
            elif len(posnr) == 2:
                formartierte_posnr = "00"
                formartierte_posnr += posnr
            elif len(posnr) == 1:
                formartierte_posnr = "000"
                formartierte_posnr += posnr
        return formartierte_posnr
    def baugruppe(self, bezeichnung):
        gruppenbezeichnung = ""
        if "_" in bezeichnung:
            index = bezeichnung.find("_")
            text_links  = bezeichnung[0:index]
            # text_rechts = bezeichnung[index+1:]
            if len(text_links)>1:
                erstes_zeichen = text_links[0:1]
                zweites_zeichen = text_links[1:2]
                if erstes_zeichen == "S":
                    if istZiffer(zweites_zeichen):
                        gruppenbezeichnung = text_links
                        # bauteilname = text_rechts
                elif erstes_zeichen == "#":
                    gruppenbezeichnung = text_links
                elif erstes_zeichen == "@":
                    gruppenbezeichnung = text_links
        return gruppenbezeichnung
    def ppf_wst_laenge(self, datipfad):
        gesuchter_parameter = "0"
        if os.path.isfile(datipfad):
            datei = open(datipfad, "r")
            pkopf = False
            for zeile in datei:                
                if "<<Werkstueck>>" in zeile:
                    pkopf = True
                if "<</Werkstueck>>" in zeile:
                    pkopf = False
                    return gesuchter_parameter
                if pkopf == True:
                    parambez_start = "<Wst_Laenge>"
                    parambez_ende  = "</Wst_Laenge>"
                    if parambez_start in zeile:
                        start_index = zeile.find(parambez_start)
                        start_laenge = len(parambez_start)
                        ende_index = zeile.find(parambez_ende)
                        gesuchter_parameter = zeile[start_index+start_laenge:ende_index]
        return gesuchter_parameter
    def ppf_wst_breite(self, datipfad):
        gesuchter_parameter = "0"
        if os.path.isfile(datipfad):
            datei = open(datipfad, "r")
            pkopf = False
            for zeile in datei:                
                if "<<Werkstueck>>" in zeile:
                    pkopf = True
                if "<</Werkstueck>>" in zeile:
                    pkopf = False
                    return gesuchter_parameter
                if pkopf == True:
                    parambez_start = "<Wst_Breite>"
                    parambez_ende  = "</Wst_Breite>"
                    if parambez_start in zeile:
                        start_index = zeile.find(parambez_start)
                        start_laenge = len(parambez_start)
                        ende_index = zeile.find(parambez_ende)
                        gesuchter_parameter = zeile[start_index+start_laenge:ende_index] 
        return gesuchter_parameter
    def ppf_wst_dicke(self, datipfad):
        gesuchter_parameter = "0"
        if os.path.isfile(datipfad):
            datei = open(datipfad, "r")
            pkopf = False
            for zeile in datei:                
                if "<<Werkstueck>>" in zeile:
                    pkopf = True
                if "<</Werkstueck>>" in zeile:
                    pkopf = False
                    return gesuchter_parameter
                if pkopf == True:
                    parambez_start = "<Wst_Dicke>"
                    parambez_ende  = "</Wst_Dicke>"
                    if parambez_start in zeile:
                        start_index = zeile.find(parambez_start)
                        start_laenge = len(parambez_start)
                        ende_index = zeile.find(parambez_ende)
                        gesuchter_parameter = zeile[start_index+start_laenge:ende_index] 
        return gesuchter_parameter
#----------------------------------------------------------------------------------
class baugrpetk_calc: # Calc
    def __init__(self):
        self.t = ol_tabelle()
        self.maxistklen = 999  
        self.listProjekt = []
        self.listPosNr   = []
        self.listBaugrp  = []   
        self.listMenge   = []   
        pass
    def ermitteln(self):
        iSpalteProjekt = 0 # Spaltennummer mit Projektinformation
        iSpaltePosNr   = 3 # Spaltennummer mit Information über Pos-Nr
        iSpalteBez     = 4 # Spaltennummer mit Teilebezeichnung
        iSpalteMenge   = 5 # Spaltennummer mit Menge
        # Werte Nullen:
        self.listProjekt = []
        self.listPosNr   = []
        self.listBaugrp  = []   
        self.listMenge   = [] 
        for i in range (1, self.maxistklen):
            myCellProj  = self.t.get_zelle_i(i, iSpalteProjekt)
            myCellPosNr = self.t.get_zelle_i(i, iSpaltePosNr)
            myCellBez   = self.t.get_zelle_i(i, iSpalteBez)
            myCellMenge = self.t.get_zelle_i(i, iSpalteMenge)
            sBaugruppe = myCellBez.String
            if "_" in sBaugruppe:
                iPos = sBaugruppe.find("_")
                sBaugruppe = sBaugruppe[0:iPos]
                iGefunden = 0
                if sBaugruppe[0] is "S":
                    if istZiffer(sBaugruppe[1]):
                        iGefunden = 1
                elif sBaugruppe[0] is "#":
                    if istZiffer(sBaugruppe[1]):
                        iGefunden = 1
                elif sBaugruppe[0] is "@":
                    if istZiffer(sBaugruppe[1]):
                        iGefunden = 1
                if iGefunden:
                    sProj  = myCellProj.String
                    sPosNr = myCellPosNr.String
                    sMenge = myCellMenge.String # Es wird angenommen das immer zuerst eine Schrankseite in der 
                                                # Stückliste steht und die Menge dieser Seite der Gesamtmenge entspricht
                    if not sBaugruppe in self.listBaugrp:
                        self.listProjekt += [sProj]
                        self.listPosNr   += [sPosNr]
                        self.listBaugrp  += [sBaugruppe]
                        self.listMenge   += [sMenge]
                        # msgbox(sBaugruppe +'\n' + sPosNr +'\n' + sMenge , 'msgbox1', 1, 'QUERYBOX')
                    else:
                        for ii in range (0, len(self.listBaugrp)):    
                            listIndexex = [] 
                            # alle Indexe finden in der Baugruppenliste für diese Baugruppe:
                            for iii in range (0, len(self.listBaugrp)): 
                                if (self.listBaugrp[iii] == sBaugruppe):
                                    listIndexex += [iii]
                            # Index für Index durchgehen:
                            iBekannt = 0
                            for iii in range (0, len(listIndexex)):
                                if (self.listPosNr[listIndexex[iii]] == sPosNr):
                                    iBekannt = 1
                                    break #for iii
                            if(iBekannt == 0):
                                self.listProjekt += [sProj]
                                self.listPosNr   += [sPosNr]
                                self.listBaugrp  += [sBaugruppe] 
                                self.listMenge   += [sMenge]
                                break #for ii
        #msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        pass
    def auflisten(self):      
        # Registerkarte anlegen:
        self.t.tab_entfernen("labels")
        self.t.tab_anlegen("labels", 1)
        self.t.set_tabfokus_s("labels")
        # Tabellenkopf erstellen:
        self.t.set_zelltext_s("A1", "Projekt")
        self.t.set_zelltext_s("B1", "Opti")
        self.t.set_zelltext_s("C1", "Pos")
        self.t.set_zelltext_s("D1", "Baugruppe")
        self.t.set_zelltext_s("E1", "Menge")
        self.t.set_zelltext_s("F1", "Orte")
        self.t.set_spaltenausrichtung_i(0, "li")
        self.t.set_spaltenausrichtung_i(1, "mi")
        self.t.set_spaltenausrichtung_i(2, "mi")
        self.t.set_spaltenausrichtung_i(3, "mi")
        self.t.set_spaltenausrichtung_i(4, "mi")
        # Tabelle füllen:
        for i in range (0, len(self.listBaugrp)):
            iStartZeile = 1
            self.t.set_zelltext_i(iStartZeile+i, 0, self.listProjekt[i]) # Projekt            
            if i > 0:
                self.t.set_zellformel_i(iStartZeile+i, 1, "=B2") # Opti
            self.t.set_zelltext_i(iStartZeile+i, 2, self.listPosNr[i]) # Pos
            self.t.set_zelltext_i(iStartZeile+i, 3, self.listBaugrp[i]) # Baugruppe
            self.t.set_zelltext_i(iStartZeile+i, 4, self.listMenge[i]) # Menge
        pass
    def speichern(self):
        msg = "" #return-Wert
        iZeileKopf = 0
        iSpalteStart = 0
        iMaxLen = 999
        # Tabellenkof:
        for i in range (iZeileKopf, iMaxLen):
            sProjekt = self.t.get_zelltext_i(i+1, iSpalteStart)
            sOpti    = self.t.get_zelltext_i(i+1, iSpalteStart+1)
            sPosNr   = self.t.get_zelltext_i(i+1, iSpalteStart+2)            
            sBaugrp  = self.t.get_zelltext_i(i+1, iSpalteStart+3)
            iMenge   = self.t.get_zelltext_i(i+1, iSpalteStart+4)
            tmp = "" # Zwischenergebnis für Ausgabe
            # Projekt:
            tmp += "Projekt  : "
            tmp += sProjekt
            tmp += "\n"
            # Projekt:
            tmp += "Opti     : "
            tmp += sOpti
            tmp += "\n"
            # PosNr:
            tmp += "Pos      : "
            tmp += sPosNr
            tmp += "\n"
            tmp += "\n"
            # Baugruppe:
            tmp += "Baugruppe: "
            tmp += sBaugrp
            tmp += "\n"
            tmp += "\n"
            # Ort:
            tmp += "Ort: " 
            #tmp += "\n"
            #tmp += "\n"

            gesund = True
            try:
                val = int(iMenge) # versuche ob der string in einen int umgewandelt werden kann
            except ValueError:
                gesund = False
            if gesund == True:
                for ii in range (0, int(iMenge)):
                    sOrt  = self.t.get_zelltext_i(i+1, iSpalteStart+5+ii)
                    msg += tmp
                    msg += sOrt
                    msg += "\n"
                    msg += "\n"
        
        path = get_userpath()
        path += "\\Desktop\\label"
        path += ".odt"
        if schreibe_in_datei_entferne_bestehende(path, msg) == True:
            msg = "label wurden erfolgreich gespeichert."
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        else:
            msg = "label konnten nicht gespeichert werden."
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        
        pass
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
class raumbuch: #calc
     def __init__(self):
         self.t = ol_tabelle()
         self.grau = RGBTo32bitInt(204, 204, 204) 
         #----------------------------- Verzeichnisse:
         self.quelle = ""
         self.quelle_zelle = "B1"
         self.ziel = ""
         self.ziel_zelle = "B2"
         self.grundrisse = ""
         self.grundrisse_zelle = "B3"
         #----------------------------- WE:
         self.we_info_zeile = "B6"
         self.we_info_spalte_start = "C6"
         self.we_info_spalte_ende = "D6"
         #----
         self.grundrisse_info_zeile = "B7"
         #----------------------------- Dateien:
         self.datei_info_spalte = "B10"
         self.datei_info_zeile_start = "C10"
         self.datei_info_zeile_ende = "D10" 
         #----
         self.pos_info_spalte = "B11"
         self.bez_info_spalte = "B12"
         self.montage_info_spalte = "B13"
         pass
     def spalten_umwandeln(self, buchstabe): # noch ergänzen!!!
         if buchstabe == "A":
             return 1
         elif buchstabe == "B":
             return 2
         elif buchstabe == "C":
             return 3
         elif buchstabe == "D":
             return 4
         elif buchstabe == "E":
             return 5
         elif buchstabe == "F":
             return 6
         elif buchstabe == "G":
             return 7
         elif buchstabe == "H":
             return 8
         elif buchstabe == "I":
             return 9
         elif buchstabe == "J":
             return 10
         elif buchstabe == "K":
             return 11
         elif buchstabe == "L":
             return 12
         elif buchstabe == "M":
             return 13
         elif buchstabe == "N":
             return 14
         elif buchstabe == "O":
             return 15
         elif buchstabe == "P":
             return 16
         elif buchstabe == "Q":
             return 17
         elif buchstabe == "R":
             return 18
         elif buchstabe == "S":
             return 19
         elif buchstabe == "T":
             return 20
         elif buchstabe == "U":
             return 21
         elif buchstabe == "V":
             return 22
         elif buchstabe == "W":
             return 23
         elif buchstabe == "X":
             return 24
         elif buchstabe == "Y":
             return 25
         elif buchstabe == "Z":
             return 26
             #---------------------------
         elif buchstabe == "AA":
             return 27
         elif buchstabe == "AB":
             return 28
         elif buchstabe == "AC":
             return 29
         elif buchstabe == "AD":
             return 30
         elif buchstabe == "AE":
             return 31
         elif buchstabe == "AF":
             return 32
         elif buchstabe == "AG":
             return 33
         elif buchstabe == "AH":
             return 34
         elif buchstabe == "AI":
             return 35
         elif buchstabe == "AJ":
             return 36
         elif buchstabe == "AK":
             return 37
         elif buchstabe == "AL":
             return 38
         elif buchstabe == "AM":
             return 39
         elif buchstabe == "AN":
             return 40
         elif buchstabe == "AO":
             return 41
         elif buchstabe == "AP":
             return 42
         elif buchstabe == "AQ":
             return 43
         elif buchstabe == "AR":
             return 44
         elif buchstabe == "AS":
             return 45
         elif buchstabe == "AT":
             return 46
         elif buchstabe == "AU":
             return 47
         elif buchstabe == "AV":
             return 48
         elif buchstabe == "AW":
             return 49
         elif buchstabe == "AX":
             return 50
         elif buchstabe == "AY":
             return 51
         elif buchstabe == "AZ":
             return 52
             #---------------------------
         elif buchstabe == "BA":
             return 53
         elif buchstabe == "BB":
             return 54
         elif buchstabe == "BC":
             return 55
         elif buchstabe == "BD":
             return 56
         elif buchstabe == "BE":
             return 57
         elif buchstabe == "BF":
             return 58
         elif buchstabe == "BG":
             return 59
         elif buchstabe == "BH":
             return 60
         elif buchstabe == "BI":
             return 61
         elif buchstabe == "BJ":
             return 62
         elif buchstabe == "BK":
             return 63
         elif buchstabe == "BL":
             return 64
         elif buchstabe == "BM":
             return 65
         elif buchstabe == "BN":
             return 66
         elif buchstabe == "BO":
             return 67
         elif buchstabe == "BP":
             return 68
         elif buchstabe == "BQ":
             return 69
         elif buchstabe == "BR":
             return 70
         elif buchstabe == "BS":
             return 71
         elif buchstabe == "BT":
             return 72
         elif buchstabe == "BU":
             return 73
         elif buchstabe == "BV":
             return 74
         elif buchstabe == "BW":
             return 75
         elif buchstabe == "BX":
             return 76
         elif buchstabe == "BY":
             return 77
         elif buchstabe == "BZ":
             return 78
             #---------------------------
         elif buchstabe == "CA":
             return 79
         elif buchstabe == "CB":
             return 80
         elif buchstabe == "CC":
             return 81
         elif buchstabe == "CD":
             return 82
         elif buchstabe == "CE":
             return 83
         elif buchstabe == "CF":
             return 84
         elif buchstabe == "CG":
             return 85
         elif buchstabe == "CH":
             return 86
         elif buchstabe == "CI":
             return 87
         elif buchstabe == "CJ":
             return 88
         elif buchstabe == "CK":
             return 89
         elif buchstabe == "CL":
             return 90
         elif buchstabe == "CM":
             return 91
         elif buchstabe == "CN":
             return 92
         elif buchstabe == "CO":
             return 93
         elif buchstabe == "CP":
             return 94
         elif buchstabe == "CQ":
             return 95
         elif buchstabe == "CR":
             return 96
         elif buchstabe == "CS":
             return 97
         elif buchstabe == "CT":
             return 98
         elif buchstabe == "CU":
             return 99
         elif buchstabe == "CV":
             return 100
         elif buchstabe == "CW":
             return 101
         elif buchstabe == "CX":
             return 102
         elif buchstabe == "CY":
             return 103
         elif buchstabe == "CZ":
             return 104
             #---------------------------
         elif buchstabe == "DA":
             return 105
         elif buchstabe == "DB":
             return 106
         elif buchstabe == "DC":
             return 107
         elif buchstabe == "DD":
             return 108
         elif buchstabe == "DE":
             return 109
         elif buchstabe == "DF":
             return 110
         elif buchstabe == "DG":
             return 111
         elif buchstabe == "DH":
             return 112
         elif buchstabe == "DI":
             return 113
         elif buchstabe == "DJ":
             return 114
         elif buchstabe == "DK":
             return 115
         elif buchstabe == "DL":
             return 116
         elif buchstabe == "DM":
             return 117
         elif buchstabe == "DN":
             return 118
         elif buchstabe == "DO":
             return 119
         elif buchstabe == "DP":
             return 120
         elif buchstabe == "DQ":
             return 121
         elif buchstabe == "DR":
             return 122
         elif buchstabe == "DS":
             return 123
         elif buchstabe == "DT":
             return 124
         elif buchstabe == "DU":
             return 125
         elif buchstabe == "DV":
             return 126
         elif buchstabe == "DW":
             return 127
         elif buchstabe == "DX":
             return 128
         elif buchstabe == "DY":
             return 129
         elif buchstabe == "DZ":
             return 130
             #---------------------------
         elif buchstabe == "EA":
             return 131
         elif buchstabe == "EB":
             return 132
         elif buchstabe == "EC":
             return 133
         elif buchstabe == "ED":
             return 134
         elif buchstabe == "EE":
             return 135
         elif buchstabe == "EF":
             return 136
         elif buchstabe == "EG":
             return 137
         elif buchstabe == "EH":
             return 138
         elif buchstabe == "EI":
             return 139
         elif buchstabe == "EJ":
             return 140
         elif buchstabe == "EK":
             return 141
         elif buchstabe == "EL":
             return 142
         elif buchstabe == "EM":
             return 143
         elif buchstabe == "EN":
             return 144
         elif buchstabe == "EO":
             return 145
         elif buchstabe == "EP":
             return 146
         elif buchstabe == "EQ":
             return 147
         elif buchstabe == "ER":
             return 148
         elif buchstabe == "ES":
             return 149
         elif buchstabe == "ET":
             return 150
         elif buchstabe == "EU":
             return 151
         elif buchstabe == "EV":
             return 152
         elif buchstabe == "EW":
             return 153
         elif buchstabe == "EX":
             return 154
         elif buchstabe == "EY":
             return 155
         elif buchstabe == "EZ":
             return 156
             #---------------------------
         elif buchstabe == "FA":
             return 157
         elif buchstabe == "FB":
             return 158
         elif buchstabe == "FC":
             return 159
         elif buchstabe == "FD":
             return 160
         elif buchstabe == "FE":
             return 161
         elif buchstabe == "FF":
             return 162
         elif buchstabe == "FG":
             return 163
         elif buchstabe == "FH":
             return 164
         elif buchstabe == "FI":
             return 165
         elif buchstabe == "FJ":
             return 166
         elif buchstabe == "FK":
             return 167
         elif buchstabe == "FL":
             return 168
         elif buchstabe == "FM":
             return 169
         elif buchstabe == "FN":
             return 170
         elif buchstabe == "FO":
             return 171
         elif buchstabe == "FP":
             return 172
         elif buchstabe == "FQ":
             return 173
         elif buchstabe == "FR":
             return 174
         elif buchstabe == "FS":
             return 175
         elif buchstabe == "FT":
             return 176
         elif buchstabe == "FU":
             return 177
         elif buchstabe == "FV":
             return 178
         elif buchstabe == "FW":
             return 179
         elif buchstabe == "FX":
             return 180
         elif buchstabe == "FY":
             return 181
         elif buchstabe == "FZ":
             return 182
             #---------------------------
         elif buchstabe == "GA":
             return 183
         elif buchstabe == "GB":
             return 184
         elif buchstabe == "GC":
             return 185
         elif buchstabe == "GD":
             return 186
         elif buchstabe == "GE":
             return 187
         elif buchstabe == "GF":
             return 188
         elif buchstabe == "GG":
             return 189
         elif buchstabe == "GH":
             return 190
         elif buchstabe == "GI":
             return 191
         elif buchstabe == "GJ":
             return 192
         elif buchstabe == "GK":
             return 193
         elif buchstabe == "GL":
             return 194
         elif buchstabe == "GM":
             return 195
         elif buchstabe == "GN":
             return 196
         elif buchstabe == "GO":
             return 197
         elif buchstabe == "GP":
             return 198
         elif buchstabe == "GQ":
             return 199
         elif buchstabe == "GR":
             return 200
         elif buchstabe == "GS":
             return 201
         elif buchstabe == "GT":
             return 202
         elif buchstabe == "GU":
             return 203
         elif buchstabe == "GV":
             return 204
         elif buchstabe == "GW":
             return 205
         elif buchstabe == "GX":
             return 206
         elif buchstabe == "GY":
             return 207
         elif buchstabe == "GZ":
             return 208
             #---------------------------
         elif buchstabe == "HA":
             return 209
         elif buchstabe == "HB":
             return 210
         elif buchstabe == "HC":
             return 211
         elif buchstabe == "HD":
             return 212
         elif buchstabe == "HE":
             return 213
         elif buchstabe == "HF":
             return 214
         elif buchstabe == "HG":
             return 215
         elif buchstabe == "HH":
             return 216
         elif buchstabe == "HI":
             return 217
         elif buchstabe == "HJ":
             return 218
         elif buchstabe == "HK":
             return 219
         elif buchstabe == "HL":
             return 220
         elif buchstabe == "HM":
             return 221
         elif buchstabe == "HN":
             return 222
         elif buchstabe == "HO":
             return 223
         elif buchstabe == "HP":
             return 224
         elif buchstabe == "HQ":
             return 225
         elif buchstabe == "HR":
             return 226
         elif buchstabe == "HS":
             return 227
         elif buchstabe == "HT":
             return 228
         elif buchstabe == "HU":
             return 229
         elif buchstabe == "HV":
             return 230
         elif buchstabe == "HW":
             return 231
         elif buchstabe == "HX":
             return 232
         elif buchstabe == "HY":
             return 233
         elif buchstabe == "HZ":
             return 234
             #---------------------------
         elif buchstabe == "IA":
             return 235
         elif buchstabe == "IB":
             return 236
         elif buchstabe == "IC":
             return 237
         elif buchstabe == "ID":
             return 238
         elif buchstabe == "IE":
             return 239
         elif buchstabe == "IF":
             return 240
         elif buchstabe == "IG":
             return 241
         elif buchstabe == "IH":
             return 242
         elif buchstabe == "II":
             return 243
         elif buchstabe == "IJ":
             return 244
         elif buchstabe == "IK":
             return 245
         elif buchstabe == "IL":
             return 246
         elif buchstabe == "IM":
             return 247
         elif buchstabe == "IN":
             return 248
         elif buchstabe == "IO":
             return 249
         elif buchstabe == "IP":
             return 250
         elif buchstabe == "IQ":
             return 251
         elif buchstabe == "IR":
             return 252
         elif buchstabe == "IS":
             return 253
         elif buchstabe == "IT":
             return 254
         elif buchstabe == "IU":
             return 255
         elif buchstabe == "IV":
             return 256
         elif buchstabe == "IW":
             return 257
         elif buchstabe == "IX":
             return 258
         elif buchstabe == "IY":
             return 259
         elif buchstabe == "IZ":
             return 260
             #---------------------------
         elif buchstabe == "JA":
             return 261
         elif buchstabe == "JB":
             return 262
         elif buchstabe == "JC":
             return 263
         elif buchstabe == "JD":
             return 264
         elif buchstabe == "JE":
             return 265
         elif buchstabe == "JF":
             return 266
         elif buchstabe == "JG":
             return 267
         elif buchstabe == "JH":
             return 268
         elif buchstabe == "JI":
             return 269
         elif buchstabe == "JJ":
             return 270
         elif buchstabe == "JK":
             return 271
         elif buchstabe == "JL":
             return 272
         elif buchstabe == "JM":
             return 273
         elif buchstabe == "JN":
             return 274
         elif buchstabe == "JO":
             return 275
         elif buchstabe == "JP":
             return 276
         elif buchstabe == "JQ":
             return 277
         elif buchstabe == "JR":
             return 278
         elif buchstabe == "JS":
             return 279
         elif buchstabe == "JT":
             return 280
         elif buchstabe == "JU":
             return 281
         elif buchstabe == "JV":
             return 282
         elif buchstabe == "JW":
             return 283
         elif buchstabe == "JX":
             return 284
         elif buchstabe == "JY":
             return 285
         elif buchstabe == "JZ":
             return 286
             #---------------------------
         elif buchstabe == "KA":
             return 287
         elif buchstabe == "KB":
             return 288
         elif buchstabe == "KC":
             return 289
         elif buchstabe == "KD":
             return 290
         elif buchstabe == "KE":
             return 291
         elif buchstabe == "KF":
             return 292
         elif buchstabe == "KG":
             return 293
         elif buchstabe == "KH":
             return 294
         elif buchstabe == "KI":
             return 295
         elif buchstabe == "KJ":
             return 296
         elif buchstabe == "KK":
             return 297
         elif buchstabe == "KL":
             return 298
         elif buchstabe == "KM":
             return 299
         elif buchstabe == "KN":
             return 300
         elif buchstabe == "KO":
             return 301
         elif buchstabe == "KP":
             return 302
         elif buchstabe == "KQ":
             return 303
         elif buchstabe == "KR":
             return 304
         elif buchstabe == "KS":
             return 305
         elif buchstabe == "KT":
             return 306
         elif buchstabe == "KU":
             return 307
         elif buchstabe == "KV":
             return 308
         elif buchstabe == "KW":
             return 309
         elif buchstabe == "KX":
             return 310
         elif buchstabe == "KY":
             return 311
         elif buchstabe == "KZ":
             return 312
             #---------------------------
         elif buchstabe == "LA":
             return 313
         elif buchstabe == "LB":
             return 314
         elif buchstabe == "LC":
             return 315
         elif buchstabe == "LD":
             return 316
         elif buchstabe == "LE":
             return 317
         elif buchstabe == "LF":
             return 318
         elif buchstabe == "LG":
             return 319
         elif buchstabe == "LH":
             return 320
         elif buchstabe == "LI":
             return 321
         elif buchstabe == "LJ":
             return 322
         elif buchstabe == "LK":
             return 323
         elif buchstabe == "LL":
             return 324
         elif buchstabe == "LM":
             return 325
         elif buchstabe == "LN":
             return 326
         elif buchstabe == "LO":
             return 327
         elif buchstabe == "LP":
             return 328
         elif buchstabe == "LQ":
             return 329
         elif buchstabe == "LR":
             return 330
         elif buchstabe == "LS":
             return 331
         elif buchstabe == "LT":
             return 332
         elif buchstabe == "LU":
             return 333
         elif buchstabe == "LV":
             return 334
         elif buchstabe == "LW":
             return 335
         elif buchstabe == "LX":
             return 336
         elif buchstabe == "LY":
             return 337
         elif buchstabe == "LZ":
             return 338
             #---------------------------
         else:
             return 0
         pass
     def spalten_umwandeln_num(self, spaltennummer): # noch ergänzen!!!
         if spaltennummer == 1:
             return "A"
         elif spaltennummer == 2:
             return "B"
         elif spaltennummer == 3:
             return "C"
         elif spaltennummer == 4:
             return "D"
         elif spaltennummer == 5:
             return "E"
         elif spaltennummer == 6:
             return "F"
         elif spaltennummer == 7:
             return "G"
         elif spaltennummer == 8:
             return "H"
         elif spaltennummer == 9:
             return "I"
         elif spaltennummer == 10:
             return "J"
         elif spaltennummer == 11:
             return "K"
         elif spaltennummer == 12:
             return "L"
         elif spaltennummer == 13:
             return "M"
         elif spaltennummer == 14:
             return "N"
         elif spaltennummer == 15:
             return "O"
         elif spaltennummer == 16:
             return "P"
         elif spaltennummer == 17:
             return "Q"
         elif spaltennummer == 18:
             return "R"
         elif spaltennummer == 19:
             return "S"
         elif spaltennummer == 20:
             return "T"
         elif spaltennummer == 21:
             return "U"
         elif spaltennummer == 22:
             return "V"
         elif spaltennummer == 23:
             return "W"
         elif spaltennummer == 24:
             return "X"
         elif spaltennummer == 25:
             return "Y"
         elif spaltennummer == 26:
             return "Z"
             #---------------------------
         elif spaltennummer == 27:
             return "AA"
         elif spaltennummer == 28:
             return "AB"
         elif spaltennummer == 29:
             return "AC"
         elif spaltennummer == 30:
             return "AD"
         elif spaltennummer == 31:
             return "AE"
         elif spaltennummer == 32:
             return "AF"
         elif spaltennummer == 33:
             return "AG"
         elif spaltennummer == 34:
             return "AH"
         elif spaltennummer == 35:
             return "AI"
         elif spaltennummer == 36:
             return "AJ"
         elif spaltennummer == 37:
             return "AK"
         elif spaltennummer == 38:
             return "AL"
         elif spaltennummer == 39:
             return "AM"
         elif spaltennummer == 40:
             return "AN"
         elif spaltennummer == 41:
             return "AO"
         elif spaltennummer == 42:
             return "AP"
         elif spaltennummer == 43:
             return "AQ"
         elif spaltennummer == 44:
             return "AR"
         elif spaltennummer == 45:
             return "AS"
         elif spaltennummer == 46:
             return "AT"
         elif spaltennummer == 47:
             return "AU"
         elif spaltennummer == 48:
             return "AV"
         elif spaltennummer == 49:
             return "AW"
         elif spaltennummer == 50:
             return "AX"
         elif spaltennummer == 51:
             return "AY"
         elif spaltennummer == 52:
             return "AZ"
             #---------------------------
         elif spaltennummer == 53:
             return "BA"
         elif spaltennummer == 54:
             return "BB"
         elif spaltennummer == 55:
             return "BC"
         elif spaltennummer == 56:
             return "BD"
         elif spaltennummer == 57:
             return "BE"
         elif spaltennummer == 58:
             return "BF"
         elif spaltennummer == 59:
             return "BG"
         elif spaltennummer == 60:
             return "BH"
         elif spaltennummer == 61:
             return "BI"
         elif spaltennummer == 62:
             return "BJ"
         elif spaltennummer == 63:
             return "BK"
         elif spaltennummer == 64:
             return "BL"
         elif spaltennummer == 65:
             return "BM"
         elif spaltennummer == 66:
             return "BN"
         elif spaltennummer == 67:
             return "BO"
         elif spaltennummer == 68:
             return "BP"
         elif spaltennummer == 69:
             return "BQ"
         elif spaltennummer == 70:
             return "BR"
         elif spaltennummer == 71:
             return "BS"
         elif spaltennummer == 72:
             return "BT"
         elif spaltennummer == 73:
             return "BU"
         elif spaltennummer == 74:
             return "BV"
         elif spaltennummer == 75:
             return "BW"
         elif spaltennummer == 76:
             return "BX"
         elif spaltennummer == 77:
             return "BY"
         elif spaltennummer == 78:
             return "BZ"
             #---------------------------
         elif spaltennummer == 79:
             return "CA"
         elif spaltennummer == 80:
             return "CB"
         elif spaltennummer == 81:
             return "CC"
         elif spaltennummer == 82:
             return "CD"
         elif spaltennummer == 83:
             return "CE"
         elif spaltennummer == 84:
             return "CF"
         elif spaltennummer == 85:
             return "CG"
         elif spaltennummer == 86:
             return "CH"
         elif spaltennummer == 87:
             return "CI"
         elif spaltennummer == 88:
             return "CJ"
         elif spaltennummer == 89:
             return "CK"
         elif spaltennummer == 90:
             return "CL"
         elif spaltennummer == 91:
             return "CM"
         elif spaltennummer == 92:
             return "CN"
         elif spaltennummer == 93:
             return "CO"
         elif spaltennummer == 94:
             return "CP"
         elif spaltennummer == 95:
             return "CQ"
         elif spaltennummer == 96:
             return "CR"
         elif spaltennummer == 97:
             return "CS"
         elif spaltennummer == 98:
             return "CT"
         elif spaltennummer == 99:
             return "CU"
         elif spaltennummer == 100:
             return "CV"
         elif spaltennummer == 101:
             return "CW"
         elif spaltennummer == 102:
             return "CX"
         elif spaltennummer == 103:
             return "CY"
         elif spaltennummer == 104:
             return "CZ"
             #---------------------------
         elif spaltennummer == 105:
             return "DA"
         elif spaltennummer == 106:
             return "DB"
         elif spaltennummer == 107:
             return "DC"
         elif spaltennummer == 108:
             return "DD"
         elif spaltennummer == 109:
             return "DE"
         elif spaltennummer == 110:
             return "DF"
         elif spaltennummer == 111:
             return "DG"
         elif spaltennummer == 112:
             return "DH"
         elif spaltennummer == 113:
             return "DI"
         elif spaltennummer == 114:
             return "DJ"
         elif spaltennummer == 115:
             return "DK"
         elif spaltennummer == 116:
             return "DL"
         elif spaltennummer == 117:
             return "DM"
         elif spaltennummer == 118:
             return "DN"
         elif spaltennummer == 119:
             return "DO"
         elif spaltennummer == 120:
             return "DP"
         elif spaltennummer == 121:
             return "DQ"
         elif spaltennummer == 122:
             return "DR"
         elif spaltennummer == 123:
             return "DS"
         elif spaltennummer == 124:
             return "DT"
         elif spaltennummer == 125:
             return "DU"
         elif spaltennummer == 126:
             return "DV"
         elif spaltennummer == 127:
             return "DW"
         elif spaltennummer == 128:
             return "DX"
         elif spaltennummer == 129:
             return "DY"
         elif spaltennummer == 130:
             return "DZ"
             #---------------------------
         elif spaltennummer == 131:
             return "EA"
         elif spaltennummer == 132:
             return "EB"
         elif spaltennummer == 133:
             return "EC"
         elif spaltennummer == 134:
             return "ED"
         elif spaltennummer == 135:
             return "EE"
         elif spaltennummer == 136:
             return "EF"
         elif spaltennummer == 137:
             return "EG"
         elif spaltennummer == 138:
             return "EH"
         elif spaltennummer == 139:
             return "EI"
         elif spaltennummer == 140:
             return "EJ"
         elif spaltennummer == 141:
             return "EK"
         elif spaltennummer == 142:
             return "EL"
         elif spaltennummer == 143:
             return "EM"
         elif spaltennummer == 144:
             return "EN"
         elif spaltennummer == 145:
             return "EO"
         elif spaltennummer == 146:
             return "EP"
         elif spaltennummer == 147:
             return "EQ"
         elif spaltennummer == 148:
             return "ER"
         elif spaltennummer == 149:
             return "ES"
         elif spaltennummer == 150:
             return "ET"
         elif spaltennummer == 151:
             return "EU"
         elif spaltennummer == 152:
             return "EV"
         elif spaltennummer == 153:
             return "EW"
         elif spaltennummer == 154:
             return "EX"
         elif spaltennummer == 155:
             return "EY"
         elif spaltennummer == 156:
             return "EZ"
             #---------------------------
         elif spaltennummer == 157:
             return "FA"
         elif spaltennummer == 158:
             return "FB"
         elif spaltennummer == 159:
             return "FC"
         elif spaltennummer == 160:
             return "FD"
         elif spaltennummer == 161:
             return "FE"
         elif spaltennummer == 162:
             return "FF"
         elif spaltennummer == 163:
             return "FG"
         elif spaltennummer == 164:
             return "FH"
         elif spaltennummer == 165:
             return "FI"
         elif spaltennummer == 166:
             return "FJ"
         elif spaltennummer == 167:
             return "FK"
         elif spaltennummer == 168:
             return "FL"
         elif spaltennummer == 169:
             return "FM"
         elif spaltennummer == 170:
             return "FN"
         elif spaltennummer == 171:
             return "FO"
         elif spaltennummer == 172:
             return "FP"
         elif spaltennummer == 173:
             return "FQ"
         elif spaltennummer == 174:
             return "FR"
         elif spaltennummer == 175:
             return "FS"
         elif spaltennummer == 176:
             return "FT"
         elif spaltennummer == 177:
             return "FU"
         elif spaltennummer == 178:
             return "FV"
         elif spaltennummer == 179:
             return "FW"
         elif spaltennummer == 180:
             return "FX"
         elif spaltennummer == 181:
             return "FY"
         elif spaltennummer == 182:
             return "FZ"
             #---------------------------
         elif spaltennummer == 183:
             return "GA"
         elif spaltennummer == 184:
             return "GB"
         elif spaltennummer == 185:
             return "GC"
         elif spaltennummer == 186:
             return "GD"
         elif spaltennummer == 187:
             return "GE"
         elif spaltennummer == 188:
             return "GF"
         elif spaltennummer == 189:
             return "GG"
         elif spaltennummer == 190:
             return "GH"
         elif spaltennummer == 191:
             return "GI"
         elif spaltennummer == 192:
             return "GJ"
         elif spaltennummer == 193:
             return "GK"
         elif spaltennummer == 194:
             return "GL"
         elif spaltennummer == 195:
             return "GM"
         elif spaltennummer == 196:
             return "GN"
         elif spaltennummer == 197:
             return "GO"
         elif spaltennummer == 198:
             return "GP"
         elif spaltennummer == 199:
             return "GQ"
         elif spaltennummer == 200:
             return "GR"
         elif spaltennummer == 201:
             return "GS"
         elif spaltennummer == 202:
             return "GT"
         elif spaltennummer == 203:
             return "GU"
         elif spaltennummer == 204:
             return "GV"
         elif spaltennummer == 205:
             return "GW"
         elif spaltennummer == 206:
             return "GX"
         elif spaltennummer == 207:
             return "GY"
         elif spaltennummer == 208:
             return "GZ"
             #---------------------------
         elif spaltennummer == 209:
             return "HA"
         elif spaltennummer == 210:
             return "HB"
         elif spaltennummer == 211:
             return "HC"
         elif spaltennummer == 212:
             return "HD"
         elif spaltennummer == 213:
             return "HE"
         elif spaltennummer == 214:
             return "HF"
         elif spaltennummer == 215:
             return "HG"
         elif spaltennummer == 216:
             return "HH"
         elif spaltennummer == 217:
             return "HI"
         elif spaltennummer == 218:
             return "HJ"
         elif spaltennummer == 219:
             return "HK"
         elif spaltennummer == 220:
             return "HL"
         elif spaltennummer == 221:
             return "HM"
         elif spaltennummer == 222:
             return "HN"
         elif spaltennummer == 223:
             return "HO"
         elif spaltennummer == 224:
             return "HP"
         elif spaltennummer == 225:
             return "HQ"
         elif spaltennummer == 226:
             return "HR"
         elif spaltennummer == 227:
             return "HS"
         elif spaltennummer == 228:
             return "HT"
         elif spaltennummer == 229:
             return "HU"
         elif spaltennummer == 230:
             return "HV"
         elif spaltennummer == 231:
             return "HW"
         elif spaltennummer == 232:
             return "HX"
         elif spaltennummer == 233:
             return "HY"
         elif spaltennummer == 234:
             return "HZ"
             #---------------------------
         elif spaltennummer == 235:
             return "IA"
         elif spaltennummer == 236:
             return "IB"
         elif spaltennummer == 237:
             return "IC"
         elif spaltennummer == 238:
             return "ID"
         elif spaltennummer == 239:
             return "IE"
         elif spaltennummer == 240:
             return "IF"
         elif spaltennummer == 241:
             return "IG"
         elif spaltennummer == 242:
             return "IH"
         elif spaltennummer == 243:
             return "II"
         elif spaltennummer == 244:
             return "IJ"
         elif spaltennummer == 245:
             return "IK"
         elif spaltennummer == 246:
             return "IL"
         elif spaltennummer == 247:
             return "IM"
         elif spaltennummer == 248:
             return "IN"
         elif spaltennummer == 249:
             return "IO"
         elif spaltennummer == 250:
             return "IP"
         elif spaltennummer == 251:
             return "IQ"
         elif spaltennummer == 252:
             return "IR"
         elif spaltennummer == 253:
             return "IS"
         elif spaltennummer == 254:
             return "IT"
         elif spaltennummer == 255:
             return "IU"
         elif spaltennummer == 256:
             return "IV"
         elif spaltennummer == 257:
             return "IW"
         elif spaltennummer == 258:
             return "IX"
         elif spaltennummer == 259:
             return "IY"
         elif spaltennummer == 260:
             return "IZ"
             #---------------------------
         elif spaltennummer == 261:
             return "JA"
         elif spaltennummer == 262:
             return "JB"
         elif spaltennummer == 263:
             return "JC"
         elif spaltennummer == 264:
             return "JD"
         elif spaltennummer == 265:
             return "JE"
         elif spaltennummer == 266:
             return "JF"
         elif spaltennummer == 267:
             return "JG"
         elif spaltennummer == 268:
             return "JH"
         elif spaltennummer == 269:
             return "JI"
         elif spaltennummer == 270:
             return "JJ"
         elif spaltennummer == 271:
             return "JK"
         elif spaltennummer == 272:
             return "JL"
         elif spaltennummer == 273:
             return "JM"
         elif spaltennummer == 274:
             return "JN"
         elif spaltennummer == 275:
             return "JO"
         elif spaltennummer == 276:
             return "JP"
         elif spaltennummer == 277:
             return "JQ"
         elif spaltennummer == 278:
             return "JR"
         elif spaltennummer == 279:
             return "JS"
         elif spaltennummer == 280:
             return "JT"
         elif spaltennummer == 281:
             return "JU"
         elif spaltennummer == 282:
             return "JV"
         elif spaltennummer == 283:
             return "JW"
         elif spaltennummer == 284:
             return "JX"
         elif spaltennummer == 285:
             return "JY"
         elif spaltennummer == 286:
             return "JZ"
             #---------------------------
         elif spaltennummer == 287:
             return "KA"
         elif spaltennummer == 288:
             return "KB"
         elif spaltennummer == 289:
             return "KC"
         elif spaltennummer == 290:
             return "KD"
         elif spaltennummer == 291:
             return "KE"
         elif spaltennummer == 292:
             return "KF"
         elif spaltennummer == 293:
             return "KG"
         elif spaltennummer == 294:
             return "KH"
         elif spaltennummer == 295:
             return "KI"
         elif spaltennummer == 296:
             return "KJ"
         elif spaltennummer == 297:
             return "KK"
         elif spaltennummer == 298:
             return "KL"
         elif spaltennummer == 299:
             return "KM"
         elif spaltennummer == 300:
             return "KN"
         elif spaltennummer == 301:
             return "KO"
         elif spaltennummer == 302:
             return "KP"
         elif spaltennummer == 303:
             return "KQ"
         elif spaltennummer == 304:
             return "KR"
         elif spaltennummer == 305:
             return "KS"
         elif spaltennummer == 306:
             return "KT"
         elif spaltennummer == 307:
             return "KU"
         elif spaltennummer == 308:
             return "KV"
         elif spaltennummer == 309:
             return "KW"
         elif spaltennummer == 310:
             return "KX"
         elif spaltennummer == 311:
             return "KY"
         elif spaltennummer == 312:
             return "KZ"
             #---------------------------
         elif spaltennummer == 313:
             return "LA"
         elif spaltennummer == 314:
             return "LB"
         elif spaltennummer == 315:
             return "LC"
         elif spaltennummer == 316:
             return "LD"
         elif spaltennummer == 317:
             return "LE"
         elif spaltennummer == 318:
             return "LF"
         elif spaltennummer == 319:
             return "LG"
         elif spaltennummer == 320:
             return "LH"
         elif spaltennummer == 321:
             return "LI"
         elif spaltennummer == 322:
             return "LJ"
         elif spaltennummer == 323:
             return "LK"
         elif spaltennummer == 324:
             return "LL"
         elif spaltennummer == 325:
             return "LM"
         elif spaltennummer == 326:
             return "LN"
         elif spaltennummer == 327:
             return "LO"
         elif spaltennummer == 328:
             return "LP"
         elif spaltennummer == 329:
             return "LQ"
         elif spaltennummer == 330:
             return "LR"
         elif spaltennummer == 331:
             return "LS"
         elif spaltennummer == 332:
             return "LT"
         elif spaltennummer == 333:
             return "LU"
         elif spaltennummer == 334:
             return "LV"
         elif spaltennummer == 335:
             return "LW"
         elif spaltennummer == 336:
             return "LX"
         elif spaltennummer == 337:
             return "LY"
         elif spaltennummer == 338:
             return "LZ"
             #---------------------------
         else:
             return "ZZZ"
         pass
     def RB_Blankoliste(self):
         # ---------
         tabname = "Raumbuch"
         self.t.tab_anlegen(tabname, 99)
         self.t.set_tabfokus_s(tabname)
         # ---------
         self.t.set_Rahmen_komplett_s("D1:I1", 20)
         self.t.set_Rahmen_komplett_s("A2:I11", 20)
         self.t.set_zellfarbe_s("A2:I2", self.grau)
         self.t.set_zellfarbe_s("A2:C11", self.grau)
         self.t.set_zellfarbe_s("I2:I11", self.grau)
         # ---------
         self.t.set_spaltenbreite_i(1, 5000) # B == Bezeichnung
         self.t.set_spaltenbreite_i(2, 6000) # C == Dateiname
         self.t.set_spaltenbreite_i(3, 2000) # WE
         self.t.set_spaltenbreite_i(4, 2000) # WE
         self.t.set_spaltenbreite_i(5, 2000) # WE
         self.t.set_spaltenbreite_i(6, 2000) # WE
         self.t.set_spaltenbreite_i(7, 2000) # WE
         self.t.set_spaltenbreite_i(8, 2000) # Summe
         self.t.set_spaltenbreite_i(10, 2500) # VK-Preis
         self.t.set_spaltenbreite_i(11, 2500) # Montagebudget
         # ---------
         self.t.set_spaltenausrichtung_i(8, "mi")
         # ---------
         self.t.set_zelltext_s("A1", "Projekt xy")
         self.t.set_zelltext_s("A2", "Pos")
         self.t.set_zelltext_s("B2", "Bezeichnung")
         self.t.set_zelltext_s("C2", "Datei")
         self.t.set_zelltext_s("D2", "WE001")
         self.t.set_zelltext_s("E2", "WE002")
         self.t.set_zelltext_s("F2", "WE003")
         self.t.set_zelltext_s("G2", "WE004")
         self.t.set_zelltext_s("H2", "leer")
         self.t.set_zelltext_s("I2", "Summe")
         # ---------
         for i in range (2, 11):
            formel = "=SUM(D" + str(i+1) + ":H" + str(i+1) + ")"
            self.t.set_zellformel_i(i, 8, formel)
            pass
         # ---------Grundrisse:
         self.t.set_zelltext_s("C13", "Grundrissname:")
         self.t.set_Rahmen_komplett_s("C13:H13", 20)
         self.t.set_zellfarbe_s("C13:C13", self.grau)
         # ---------Kalkulatorische Auswertung untere Zellen:
         self.t.set_zelltext_s("C15", "VK-Preis:")
         self.t.set_zelltext_s("C16", "Montagebudget:")
         self.t.set_Rahmen_komplett_s("C15:H16", 20)
         self.t.set_zellfarbe_s("C15:C16", self.grau)
         for spa in range (3, 8):
            spaAplha = self.spalten_umwandeln_num(spa+1)
            #VK_Preis:
            formel  = "=SUMPRODUCT(" + spaAplha + "3:" + spaAplha + "11*$K3:$K11)"
            self.t.set_zellformel_i(15-1, spa, formel)
            #Montagebudget:
            formel  = "=SUMPRODUCT(" + spaAplha + "3:" + spaAplha + "11*$L3:$L11)"
            self.t.set_zellformel_i(16-1, spa, formel)
            pass
         self.t.set_zellformat_s("D15:H16", "#.##0,00 [$€-407];[ROT]-#.##0,00 [$€-407]") # Währung
         # ---------Kalkulatorische Auswertung rechte Zellen:
         self.t.set_zelltext_s("K2", "VK-Preis:")
         self.t.set_zelltext_s("L2", "Montagebudget:")
         self.t.set_Rahmen_komplett_s("K2:L11", 20)
         self.t.set_zellfarbe_s("K2:L2", self.grau)
         self.t.set_zellformat_s("K3:L11", "#.##0,00 [$€-407];[ROT]-#.##0,00 [$€-407]") # Währung
         # ---------
         pass
     def LList_Formblatt (self):
         # ---------
         tabname = "Lieferlisten"
         self.t.tab_anlegen(tabname, 99)
         self.t.set_tabfokus_s(tabname)
         self.t.set_spaltenbreite_i(0, 4500)
         # ---------Verzeichnisse:
         text = "Quell-Verzeichnis"
         pos = "A1"
         self.t.set_zelltext_s(pos, text)
         text = "Bitte hier den Pfad eintragen"
         pos = "B1"
         self.t.set_zelltext_s(pos, text)
         text = "Ziel-Verzeichnis"
         pos = "A2"
         self.t.set_zelltext_s(pos, text)
         text = "Bitte hier den Pfad eintragen"
         pos = "B2"
         self.t.set_zelltext_s(pos, text)
         text = "Grundrisse-Verzeichnis"
         pos = "A3"
         self.t.set_zelltext_s(pos, text)
         text = "Bitte hier den Pfad eintragen"
         pos = "B3"
         self.t.set_zelltext_s(pos, text)
         # ---------Erste Tabelle:
         text = "WE-Bezeichnungen"
         pos = "A6"
         self.t.set_zelltext_s(pos, text)
         text = "Grundrisse"
         pos = "A7"
         self.t.set_zelltext_s(pos, text)
         text = "Zeile"
         pos = "B5"
         self.t.set_zelltext_s(pos, text)
         text = "Spalte-Start"
         pos = "C5"
         self.t.set_zelltext_s(pos, text)
         text = "Spalte-Ende"
         pos = "D5"
         self.t.set_zelltext_s(pos, text)

         self.t.set_zellfarbe_s("A5:D5", self.grau)
         self.t.set_zellfarbe_s("A5:A7", self.grau)
         self.t.set_zellfarbe_s("C7:D7", self.grau)
         self.t.set_Rahmen_komplett_s("A5:D7", 20)
         self.t.set_zellausrichtungHori_s("B6:D7", "mi")
         text = "2"
         pos = "B6"
         self.t.set_zelltext_s(pos, text)
         text = "D"
         pos = "C6"
         self.t.set_zelltext_s(pos, text)
         text = "H"
         pos = "D6"
         self.t.set_zelltext_s(pos, text)
         text = "13"
         pos = "B7"
         self.t.set_zelltext_s(pos, text)
         # ---------Zweite Tabelle:
         text = "Dateinamen"
         pos = "A10"
         self.t.set_zelltext_s(pos, text)
         text = "Pos-Nummern"
         pos = "A11"
         self.t.set_zelltext_s(pos, text)
         text = "Pos-Bez."
         pos = "A12"
         self.t.set_zelltext_s(pos, text)
         text = "Montagebudget"
         pos = "A13"
         self.t.set_zelltext_s(pos, text)
         text = "Spalte"
         pos = "B9"
         self.t.set_zelltext_s(pos, text)
         text = "Zeile-Start"
         pos = "C9"
         self.t.set_zelltext_s(pos, text)
         text = "Zeile-Ende"
         pos = "D9"
         self.t.set_zelltext_s(pos, text)

         self.t.set_zellfarbe_s("A9:D9", self.grau)
         self.t.set_zellfarbe_s("A9:A13", self.grau)
         self.t.set_zellfarbe_s("C11:D13", self.grau)
         self.t.set_Rahmen_komplett_s("A9:D13", 20)
         self.t.set_zellausrichtungHori_s("B10:D13", "mi")
         text = "C"
         pos = "B10"
         self.t.set_zelltext_s(pos, text)
         text = "A"
         pos = "B11"
         self.t.set_zelltext_s(pos, text)
         text = "B"
         pos = "B12"
         self.t.set_zelltext_s(pos, text)
         text = "L"
         pos = "B13"
         self.t.set_zelltext_s(pos, text)
         text = "3"
         pos = "C10"
         self.t.set_zelltext_s(pos, text)
         text = "11"
         pos = "D10"
         self.t.set_zelltext_s(pos, text)
         # ---------Bedienerhinweise:
         text = "Der Platzhalter [WE] kann bei der Benennung der Dateinamen für"
         pos = "A15"
         self.t.set_zelltext_s(pos, text)
         text = "die Zeichnungen und Grundrisse verwendet werden. Das Makro tauscht"
         pos = "A16"
         self.t.set_zelltext_s(pos, text)
         text = "den Platzhalter dann mit der Bezeichnung der jeweiligen Wohnung aus."
         pos = "A17"
         self.t.set_zelltext_s(pos, text)
         # ---------
         pass
     def csvZelle(self, msg):
         ret = "\""
         ret += msg
         ret += "\""
         ret += ";"
         return ret
     def LList_start(self):
         # ------------------------------------------------------ ist die richtige Registerkarte geöffnet? :
         reg = self.t.get_tabname()
         if reg != "Lieferlisten":
             msgbox('Bitte in die Registerkarte \"Lieferlisten\" wechseln', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         # ------------------------------------------------------
         self.quelle = self.t.get_zelltext_s(self.quelle_zelle)
         self.ziel = self.t.get_zelltext_s(self.ziel_zelle)
         self.grundrisse = self.t.get_zelltext_s(self.grundrisse_zelle)
         # ------------------------------------------------------ WE Zeile:
         we_zei = 0
         tmp = self.t.get_zelltext_s(self.we_info_zeile)
         if len(tmp) == 0:
             msgbox('Eingabe von WE-Zeile ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         else:
             we_zei = tmp
         # ------------------------------------------------------ WE Spalte Start:
         we_spaS = 0
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.we_info_spalte_start))
         if tmp != 0:
            we_spaS = tmp
         else:
             msgbox('Eingabe von WE-Spalte Start ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         # ------------------------------------------------------ WE Spalte Ende:
         we_spaE = 0 
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.we_info_spalte_ende))
         if tmp != 0:
            we_spaE = tmp
         else:
             msgbox('Eingabe von WE-Spalte Ende ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         # ------------------------------------------------------ Grundrisse Zeile:
         gru_zei = 0
         tmp = self.t.get_zelltext_s(self.grundrisse_info_zeile)
         if len(tmp) == 0:
             msgbox('Eingabe von Grundrisse ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         else:
             gru_zei  = tmp
         # ------------------------------------------------------
         # ------------------------------------------------------ Dateinahmen Spalte:
         dn_spa = 0
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.datei_info_spalte))
         if tmp != 0:
            dn_spa = tmp
         else:
             msgbox('Eingabe von Dateinamen-Spalte Start ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return  
         # ------------------------------------------------------ Dateinahmen Zeile Start:
         dn_zeiS = 0
         tmp = self.t.get_zelltext_s(self.datei_info_zeile_start)
         if len(tmp) == 0:
             msgbox('Eingabe von WE-Zeile ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         else:
             dn_zeiS = tmp
         # ------------------------------------------------------ Dateinahmen Zeile Ende:
         dn_zeiE = 0
         tmp = self.t.get_zelltext_s(self.datei_info_zeile_ende)
         if len(tmp) == 0:
             msgbox('Eingabe von WE-Zeile ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return
         else:
             dn_zeiE = tmp
         # ------------------------------------------------------ Position Spalte:
         pos_spa = 0
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.pos_info_spalte))
         if tmp != 0:
            pos_spa = tmp
         else:
             msgbox('Eingabe von Pos-Nummer-Spalte ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return  
         # ------------------------------------------------------ Positios-Bezeichnung Spalte:
         bez_spa = 0
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.bez_info_spalte))
         if tmp != 0:
            bez_spa = tmp
         else:
             msgbox('Eingabe von Pos.Bez.-Spalte ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return  
         # ------------------------------------------------------ Montagebudget Spalte:
         montage_spa = 0
         tmp = self.spalten_umwandeln(self.t.get_zelltext_s(self.montage_info_spalte))
         if tmp != 0:
            montage_spa = tmp
         else:
             msgbox('Eingabe von Montage.-Spalte ungültig', 'Makro Lieferlisten', 1, 'QUERYBOX')
             return  
         # ------------------------------------------------------ Prüfen ob alle Dateien existieren:
         # Dateinamen mit Platzhaltern nicht prüfen
         # Platzhalter:
         # [WE] --> Wird durch die jeweiligewohnungsnummer ersetzt
         PLATZHALTER_WE = "[WE]" # Konstante
         tabname_raumbuch = "Raumbuch"
         if self.t.tab_existiert(tabname_raumbuch) == True:
             self.t.set_tabfokus_s(tabname_raumbuch)
             for i in range (int(dn_zeiS), int(dn_zeiE)+1):
                datnam = self.t.get_zelltext_i(int(i)-1, int(dn_spa)-1)
                # msgbox(datnam, 'msgbox', 1, 'QUERYBOX')
                if ((len(datnam) > 4) and (PLATZHALTER_WE not in datnam)):
                    datnam = self.quelle + "\\" + datnam
                    fileObj = Path(datnam)
                    if fileObj.is_file() == False:
                        msg = "Die Datei \""
                        msg += datnam
                        msg += "\" wurde ncht gefunden!"
                        msgbox(msg, 'msgbox', 1, 'QUERYBOX')
                        return
                    pass
             pass
         else:
             msg =  "Die Registerkarte \""
             msg += tabname_raumbuch
             msg += "\" existiert nicht!"
             msgbox(msg, 'msgbox', 1, 'QUERYBOX')
         # ------------------------------------------------------ WE Prüfen:
         WEs = []
         for i in range (we_spaS-1, we_spaE):              
             we = self.t.get_zelltext_i(int(we_zei)-1, i)
             if len(we) == 0:
                 msg  = "Bitte Raumbuch prüfen!\n"
                 msg += "Im angegebenen Bereich fehlen WE-Angaben.\n"
                 msg += "Spalte "
                 msg += str(i+1)
                 msgbox(msg, 'msgbox', 1, 'QUERYBOX')
                 return
             elif we in WEs:
                 msg = "Die WE \""
                 msg += we
                 msg += "\" ist mehrfach vorhanden!"
                 msgbox(msg, 'msgbox', 1, 'QUERYBOX')
                 return
             else:
                WEs += [we] 
             pass
         # ------------------------------------------------------Ordner für WEs erstellen:
         if(os.path.isdir(self.ziel)):
             for i in WEs:
                 dir = self.ziel
                 dir += "\\"
                 dir += i
                 if os.path.isdir(dir) == False:
                    try: 
                        os.makedirs(dir)
                    except OSError:
                        titel = "Ordner erstellen"
                        msg = "Der Ordner \n"
                        msg += dir
                        msg += "\n"
                        msg += "kann nicht erstellt werden! Makro wird abgebrochen."
                        msgbox(msg, titel, 1, 'QUERYBOX')
                        return
         else:
             msg  = "Das angegebene Zielverzeichnis\n"
             msg += self.ziel
             msg += "ist nicht zugreifbar."
             msgbox(msg, 'msgbox', 1, 'QUERYBOX')
             return
         # ------------------------------------------------------Dateien in WE-Ordnern ablegen:
         msg_zaehler = 0
         for spa in range (we_spaS-1, we_spaE): # WE für WE durchgehen
             weName = self.t.get_zelltext_i(int(we_zei)-1, spa)
             sRBtext  = self.csvZelle("Raumbuch ")
             sRBtext += self.csvZelle("")
             sRBtext += self.csvZelle(weName)
             sRBtext += "\n"
             sRBtext += "\n"
             sRBtext += self.csvZelle("Pos")
             sRBtext += self.csvZelle("Bezeichnung")
             sRBtext += self.csvZelle("Menge")
             sRBtext += self.csvZelle("Montagebudget")
             sRBtext += "\n"
             
             for zei in range (int(dn_zeiS), int(dn_zeiE)+1): # Pos für Pos durchgehen
                 zelakt = self.t.get_zelltext_i(zei-1, spa) # Mengenangabe in dieser WE
                 datnam = self.t.get_zelltext_i(int(zei)-1, int(dn_spa)-1) # Dateiname
                 if( len(zelakt) > 0 ):
                    if( len(datnam) > 0 ): # Dateien nur ablegen wenn Dateinamen angegeben worden sind.
                                           # Ohne diese Prüfung kommt es zu Fehlern bei fehlender Dateiangabe.
                        if(PLATZHALTER_WE in datnam):
                            datnam = datnam.replace(PLATZHALTER_WE, weName, 1) 
                            datnam_ohne_daterw = datnam.replace(".pdf", "", 1)
                            quelldateien = findeDateien(datnam_ohne_daterw, self.quelle) 
                            for i in quelldateien:
                                aktZieldatei = self.ziel + "\\"
                                aktZieldatei += self.t.get_zelltext_i(int(we_zei)-1, spa) + "\\"
                                aktZieldatei += os.path.basename(i)
                                if(os.path.isfile(i) == True):
                                    copyfile(i, aktZieldatei)
                        else:
                            zieldatei = self.ziel
                            zieldatei += "\\"
                            zieldatei += self.t.get_zelltext_i(int(we_zei)-1, spa)
                            zieldatei += "\\" 
                            zieldatei += datnam
                            quelldatei  = self.quelle 
                            quelldatei += "\\" 
                            quelldatei += datnam
                            if(os.path.isfile(quelldatei) == True):                            
                                copyfile(quelldatei, zieldatei)
                            else:
                                if(msg_zaehler < 5):
                                    msg = "Die Datei: \n"
                                    msg += quelldatei
                                    msg += "\n wurde nicht gefunden und wird übersprungen."
                                    fenstertitel = "Datei nicht gefunden (" + str(msg_zaehler+1) + ")"
                                    msgbox(msg, fenstertitel, 1, 'QUERYBOX')
                                    msg_zaehler = msg_zaehler + 1
                                elif(msg_zaehler == 5):
                                    msg = "Alle weiteren nicht vorhandenen Dateien werden ohne weitere Meldungen übersprungen."
                                    msgbox(msg, 'Datei nicht gefunden', 1, 'QUERYBOX')
                                    msg_zaehler = msg_zaehler + 1                            
                    # Auch wenn keine Datei angegeben wurde soll das Bauteil im Raumbuch
                    # dieser Wohnung aufgeführt werden:
                    sRBtext += self.csvZelle(self.t.get_zelltext_i(int(zei)-1, pos_spa-1)) #Pos
                    sRBtext += self.csvZelle(self.t.get_zelltext_i(int(zei)-1, bez_spa-1))#Bez
                    sRBtext += self.csvZelle(self.t.get_zelltext_i(int(zei)-1, spa))#Menge
                    sRBtext += self.csvZelle(self.t.get_zelltext_i(int(zei)-1, montage_spa-1))#Montagebudget
                    sRBtext += "\n"
                 pass
             
             rbDatNam = self.ziel
             rbDatNam += "\\\\"
             rbDatNam += weName
             rbDatNam += "\\\\"
             rbDatNam += "Raumbuch "
             rbDatNam += weName
             rbDatNam += ".csv"
             file = open(rbDatNam, "w")
             file.write(sRBtext)
             file.close()
             pass
         # ------------------------------------------------------Grundrisse in WE-Ordnern ablegen:
         for spa in range (we_spaS-1, we_spaE): # WE für WE durchgehen
             weName = self.t.get_zelltext_i(int(we_zei)-1, spa)
             datnam = self.t.get_zelltext_i(int(gru_zei)-1, spa) # Dateiname
             if( len(datnam) > 0 ):
                 if(PLATZHALTER_WE in datnam):
                     datnam = datnam.replace(PLATZHALTER_WE, weName, 1)
                     pass
                 quelldatei  = self.grundrisse
                 quelldatei += "\\" 
                 quelldatei += datnam
                 zieldatei = self.ziel
                 zieldatei += "\\"
                 zieldatei += self.t.get_zelltext_i(int(we_zei)-1, spa)
                 zieldatei += "\\" 
                 zieldatei += datnam
                 if(os.path.isfile(quelldatei) == True):                            
                    copyfile(quelldatei, zieldatei)
                 else:
                     msg = "Die Datei: \n"
                     msg += quelldatei
                     msg += "\n wurde nicht gefunden und wird übersprungen."
                     fenstertitel = "Datei für Grundris nicht gefunden (" + str(msg_zaehler+1) + ")"
                     msgbox(msg, fenstertitel, 1, 'QUERYBOX')
             pass
         # ------------------------------------------------------
         msg = "Makro erfolgreich abgeschlossen."
         msgbox(msg, 'Makro Lieferlisten', 1, 'QUERYBOX')
         pass
        
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
class WoPlan: # Calc
    def __init__(self):
        self.context = XSCRIPTCONTEXT # globale Variable im sOffice-kontext
        self.doc = self.context.getDocument() #aktuelles Document per Methodenaufruf ! mit Klammern !
        self.RahLinDi = 25 # entspricht "0,7pt"
        self.grau = RGBTo32bitInt(204, 204, 204)  
        self.blau = RGBTo32bitInt(0, 102, 204) 
        self.tab = ol_tabelle()
        self.tabGrundlagen = ol_tabelle()
        self.setup_tab_grundlagen()
        self.tabGrundlagen.set_tabname("Grundlagen")   
        self.jahr = 2020 
        self.kw = 1          
        pass
    def setup_tab_grundlagen(self):
        anzFehler = self.tabelle_anlegen("Grundlagen", True)
        if(anzFehler == 0):
            t = ol_tabelle()
            t.set_tabname("Grundlagen")
            #---
            t.set_zelltext_s("A1", "Mitarbeiter")
            t.set_zelltext_s("B1", "Gruppe")
            t.set_zelltext_s("C1", "Tätigkeit")   
            t.set_zelltext_s("D1", "Geburtstag") 
            t.set_SchriftFett_s("A1:D1", True)
            t.set_zellfarbe_s("A1:D1", self.grau)
            t.set_Rahmen_komplett_s("A1:D20", self.RahLinDi)
            t.set_spaltenbreite_i(0, 4100)
            t.set_spaltenbreite_i(1, 2260)
            t.set_spaltenbreite_i(2, 2260)
            t.set_spaltenbreite_i(3, 2260)
            t.set_zellformat_s("D2:D20", "TT.MM.JJJJ")
            #---
            t.set_zelltext_s("F1", "Kalender-Jahr")
            t.set_zelltext_s("G1", "2020")
            t.set_zelltext_s("F2", "KW1 beginnt am")
            t.set_zelltext_datum_s("G2", "2021", "12", "30")
            t.set_zelltext_s("F3", "KW")
            t.set_zelltext_s("G3", "1")
            t.set_SchriftFett_s("F1:F3", True)            
            t.set_zellfarbe_s("F1:F3", self.grau)            
            t.set_Rahmen_komplett_s("F1:G3", self.RahLinDi)
            t.set_spaltenbreite_i(5, 3250) 
            #---
            t.set_zelltext_s("F5", "Gruppen")
            t.set_SchriftFett_s("F5", True)            
            t.set_zellfarbe_s("F5", self.grau)            
            t.set_Rahmen_komplett_s("F5:F15", self.RahLinDi)
            t.set_zelltext_s("F6", "Halle1")
            t.set_zelltext_s("F7", "Halle2")
            t.set_zelltext_s("F8", "Halle3")
            t.set_zelltext_s("F9", "Kraftfahrer")
            t.set_zelltext_s("F10", "Lehrlinge")
            t.set_zelltext_s("F11", "Büro")
            #---
        pass
    def tab_Grundlagen(self):
        self.setup_tab_grundlagen()
        t = ol_tabelle()
        t.set_tabfokus_s("Grundlagen")
        pass
    def wochenplan_erstellen(self):
        anzFehler = self.tabelle_anlegen(self.get_kw())
        if(anzFehler == 0):
            self.set_fokus_tab_kw()
            self.set_spaltenbreiten()
            self.set_tabellenkopf()
            self.set_tabellenrumpf()
            self.setup_for_printing()
        self.set_fokus_tab_kw()
        pass
    def tabelle_anlegen(self, sTabname, bIgnoreError = False):
        tabNames = self.doc.Sheets.ElementNames
        bereits_vorhanden = False
        anzFehler = 0
        for i in range(0, len(tabNames)):
            if(sTabname == tabNames[i]):
                bereits_vorhanden = True
                break # for i
            pass
        if(bereits_vorhanden == True):
            if(bIgnoreError == False):
                msg = "Die Registerkarte \"" + str(sTabname) + "\" existiert bereits!"
                msgbox(msg, 'msgbox', 1, 'QUERYBOX')
            anzFehler += 1
        else:
            tabIndex = 99
            self.tab.tab_anlegen(sTabname, tabIndex)
        return anzFehler
    def set_fokus_tab_kw(self):
        self.tab.set_tabfokus_s(self.get_kw())
    def get_kw(self):
        self.kw = self.tabGrundlagen.get_zelltext_s("G3")
        return self.kw
    def get_jahr(self):
        self.jahr = self.tabGrundlagen.get_zelltext_s("G1")
        return self.jahr
    def set_spaltenbreiten(self):
        t = ol_tabelle()
        t.set_tabname(self.get_kw())
        t.set_spaltenbreite_i(0, 4500)
        t.set_spaltenbreite_i(1, 2400)
        t.set_spaltenbreite_i(2, 2250)
        t.set_spaltenbreite_i(3, 5200)
        t.set_spaltenbreite_i(4, 5200)
        t.set_spaltenbreite_i(5, 5200)
        t.set_spaltenbreite_i(6, 5200)
        t.set_spaltenbreite_i(7, 5200)
        t.set_spaltenbreite_i(8, 5200)
        pass
    def set_tabellenkopf(self):
        t = ol_tabelle()
        t.set_tabname(self.get_kw())
        # Zeile 1:
        t.set_zelltext_s("A1", "Wochenplan")
        t.set_zelltext_s("D1", str(self.get_kw()))
        t.set_zellausrichtungHori_s("D1", "re")
        t.set_zelltext_s("E1", ".KW")
        t.set_zelltext_s("F1", str(self.get_jahr()))
        t.set_zeilenhoehe_i(0, 1100)
        t.set_SchriftGroesse_s("A1:I1", 26)
        t.set_SchriftFett_s("A1:I1", True)
        # Zeile 2:
        t.set_zeilenhoehe_i(1, 300)
        # Zeile 3:
        t.set_zellfarbe_s("A3:I3", self.grau)
        t.set_SchriftFett_s("A3:I3", True)
        t.set_Rahmen_komplett_s("A3:I3", self.RahLinDi)
        t.set_zellausrichtungHori_s("A3:I3", "mi")
        t.set_zelltext_s("B3", "Tätigkeit")
        t.set_zelltext_s("C3", "KFZ")
        # Beschriftung Montag:
        startdatum = "$Grundlagen.G2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)" # startdatum + ( (KW-1) * 7 )
        t.set_zellformel_s("D3", formel)
        t.set_zellformat_s("D3", "TT.MM.JJJJ")
        tmp = "Montag "
        tmp += t.get_zelltext_s("D3")
        t.set_zelltext_s("D3", tmp)
        # Beschriftung Dienstag:
        startdatum = "$Grundlagen.G2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+1" # startdatum + ( (KW-1) * 7 ) + 1 Tag
        t.set_zellformel_s("E3", formel)
        t.set_zellformat_s("E3", "TT.MM.JJJJ")
        tmp = "Dienstag "
        tmp += t.get_zelltext_s("E3")
        t.set_zelltext_s("E3", tmp)
        # Beschriftung Mittwoch:
        startdatum = "$Grundlagen.G2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+2" # startdatum + ( (KW-1) * 7 ) + 2 Tage
        t.set_zellformel_s("F3", formel)
        t.set_zellformat_s("F3", "TT.MM.JJJJ")
        tmp = "Mittwoch "
        tmp += t.get_zelltext_s("F3")
        t.set_zelltext_s("F3", tmp)
        # Beschriftung Donnerstag:
        startdatum = "$Grundlagen.G2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+3" # startdatum + ( (KW-1) * 7 ) + 3 Tage
        t.set_zellformel_s("G3", formel)
        t.set_zellformat_s("G3", "TT.MM.JJJJ")
        tmp = "Donnerstag "
        tmp += t.get_zelltext_s("G3")
        t.set_zelltext_s("G3", tmp)
        # Beschriftung Freitag:
        startdatum = "$Grundlagen.G2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+4" # startdatum + ( (KW-1) * 7 ) + 4 Tage
        t.set_zellformel_s("H3", formel)
        t.set_zellformat_s("H3", "TT.MM.JJJJ")
        tmp = "Freitag "
        tmp += t.get_zelltext_s("H3")
        t.set_zelltext_s("H3", tmp)
        pass
    def get_gruppennamen(self):
        t = ol_tabelle()
        t.set_tabname("Grundlagen")
        # Namen und Anzahl der Gruppen ermitteln:
        gruppenNamen = []
        idZeile = 5
        idSpalte = 5
        for i in range(0, 10):
            tmp = t.get_zelltext_i(idZeile+i, idSpalte)
            if(len(tmp) > 0):
                tmp2 = [tmp] # Kapselung nötig da sonst jeder einzelne Buchstabe als Einzelwert gedeutet wird
                gruppenNamen += tmp2
            pass
        return gruppenNamen
    def set_tabellenrumpf(self):
        t = ol_tabelle()
        t.set_tabname("Grundlagen")
        # Namen und Anzahl der Gruppen ermitteln:
        gruppenNamen = self.get_gruppennamen()
        # Mitarbeiter, Gruppe und Tätigkeit ermitteln:
        mitarb = []
        gruppe = []
        taetig = []
        gebtag = []
        idZeile = 1
        idSpalte = 0
        for i in range(0, 50):
            tmp_mitarb = t.get_zelltext_i(idZeile+i, idSpalte)
            tmp_gruppe = t.get_zelltext_i(idZeile+i, idSpalte+1)
            tmp_taetig = t.get_zelltext_i(idZeile+i, idSpalte+2)
            tmp_gebtag = t.get_zelltext_i(idZeile+i, idSpalte+3)
            if(len(tmp_gruppe) > 0):
                mitarb += [tmp_mitarb] # Kapselung nötig da sonst jeder einzelne Buchstabe als Einzelwert gedeutet wird
                gruppe += [tmp_gruppe]
                taetig += [tmp_taetig]
                gebtag += [tmp_gebtag]
            pass
        #msgbox(gebtag, 'msgbox', 1, 'QUERYBOX')
        # Tabellenrumpf füllen:
        t.set_tabname(self.get_kw()) # ab jetzt tab der KW ansprechen
        aktZeile = 3
        for i in range(0, len(gruppenNamen)):
            if(gruppenNamen[i] != "Büro"):
                # Gruppenname:
                t.set_zelltext_i(aktZeile, 0, gruppenNamen[i])            
                zellname = "A" + str(aktZeile+1)
                t.set_zellausrichtungHori_s(zellname, "mi")
                t.set_SchriftFett_s(zellname, True)
                zellname += ":I" + str(aktZeile+1)
                t.set_zellfarbe_s(zellname, self.grau)
                t.set_Rahmen_komplett_s(zellname, self.RahLinDi)
                t.set_SchriftGroesse_s(zellname, 8)
                aktZeile += 1
                # Gruppenmitglieder:
                for ii in range(0, len(gruppe)):
                    zellname = "A" + str(aktZeile+1) + ":I" + str(aktZeile+1)
                    t.set_Rahmen_komplett_s(zellname, self.RahLinDi)
                    t.set_SchriftGroesse_s(zellname, 8)
                    if(gruppe[ii] == gruppenNamen[i]):
                        t.set_zelltext_i(aktZeile, 0, mitarb[ii])
                        t.set_zelltext_i(aktZeile, 1, taetig[ii])
                        # Prüfen ob Mitarbeiter Geburtstag hat:
                        if(len(gebtag[ii])>0):
                            zeitformat = "%d.%m.%Y"
                            tmp_s = t.get_zelltext_s("D3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            montag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("E3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            dienstag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("F3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            mitwoch = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("G3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            donnerstag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("H3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            freitag = time.strptime(tmp_s, zeitformat)
                            geburtstag = time.strptime(gebtag[ii], zeitformat)
                            alter = int(self.jahr) - int(geburtstag.tm_year)                            
                            gebtagtext = mitarb[ii] + " hat " + str(alter) + ". Geburtstag"
                            if geburtstag.tm_yday == montag.tm_yday:
                                t.set_zelltext_i(aktZeile, 3, gebtagtext)
                            if geburtstag.tm_yday == dienstag.tm_yday:
                                t.set_zelltext_i(aktZeile, 4, gebtagtext)
                            if geburtstag.tm_yday == mitwoch.tm_yday:
                                t.set_zelltext_i(aktZeile, 5, gebtagtext)
                            if geburtstag.tm_yday == donnerstag.tm_yday:
                                t.set_zelltext_i(aktZeile, 6, gebtagtext)
                            if geburtstag.tm_yday == freitag.tm_yday:
                                t.set_zelltext_i(aktZeile, 7, gebtagtext)
                            pass
                        aktZeile += 1
                    pass
                zellname = "A" + str(aktZeile+1) + ":I" + str(aktZeile+1)
                t.set_Rahmen_komplett_s(zellname, self.RahLinDi)
                t.set_SchriftGroesse_s(zellname, 8)
                aktZeile += 1
            pass
        zellname = "A" + str(aktZeile+1) + ":I" + str(aktZeile+1)
        t.set_Rahmen_komplett_s(zellname, 0)
        t.set_zeilenhoehe_i(aktZeile, 260)
        aktZeile += 1
        # Wochenziele
        zeilenmengeWochenziele = 6
        zellname = "A" + str(aktZeile+1) + ":I" + str(aktZeile+zeilenmengeWochenziele)
        t.set_SchriftGroesse_s(zellname, 12)
        t.set_SchriftFett_s(zellname, True)
        t.set_zelltext_i(aktZeile, 3, "Wochenziele:")
        for i in range(1, zeilenmengeWochenziele+1):
            zellname = "E" + str(aktZeile+i)
            t.set_SchriftFarbe_s(zellname, self.blau)
            # t.set_zelltext_s(zellname, "alles schaffen :-)")
            pass
        aktZeile += zeilenmengeWochenziele
        # Arbeitszeiten:
        t.set_zelltext_i(aktZeile, 0, "     Werte Kollegen  es gelten bis auf weiteres folgende Arbeitszeiten:                           Frühschicht von 6.00 bis 15.00 Uhr")
        zellname = "A" + str(aktZeile+1)
        t.set_SchriftGroesse_s(zellname, 18)
        t.set_SchriftFett_s(zellname, True)
        aktZeile += 1
        # Bereich Büro:
        aktZeile += 1 # Leerzeile
        for i in range(0, len(gruppenNamen)):
            if(gruppenNamen[i] == "Büro"):
                # Gruppenname:
                t.set_zelltext_i(aktZeile, 0, gruppenNamen[i])            
                zellname = "A" + str(aktZeile+1)
                t.set_zellausrichtungHori_s(zellname, "mi")
                t.set_SchriftFett_s(zellname, True)
                zellname += ":I" + str(aktZeile+1)
                t.set_zellfarbe_s(zellname, self.grau)
                t.set_Rahmen_komplett_s(zellname, self.RahLinDi)
                t.set_SchriftGroesse_s(zellname, 8)
                aktZeile += 1
                # Gruppenmitglieder:
                for ii in range(0, len(gruppe)):
                    zellname = "A" + str(aktZeile+1) + ":I" + str(aktZeile+1)
                    t.set_Rahmen_komplett_s(zellname, self.RahLinDi)
                    t.set_SchriftGroesse_s(zellname, 8)
                    if(gruppe[ii] == gruppenNamen[i]):
                        t.set_zelltext_i(aktZeile, 0, mitarb[ii])
                        t.set_zelltext_i(aktZeile, 1, taetig[ii])
                        # Prüfen ob Mitarbeiter Geburtstag hat:
                        if(len(gebtag[ii])>0):
                            zeitformat = "%d.%m.%Y"
                            tmp_s = t.get_zelltext_s("D3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            montag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("E3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            dienstag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("F3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            mitwoch = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("G3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            donnerstag = time.strptime(tmp_s, zeitformat)
                            tmp_s = t.get_zelltext_s("H3")
                            tmp_s = tmp_s[len(tmp_s)-10: ]
                            freitag = time.strptime(tmp_s, zeitformat)
                            geburtstag = time.strptime(gebtag[ii], zeitformat)
                            alter = int(self.jahr) - int(geburtstag.tm_year)                            
                            gebtagtext = mitarb[ii] + " hat " + str(alter) + ". Geburtstag"
                            if geburtstag.tm_yday == montag.tm_yday:
                                t.set_zelltext_i(aktZeile, 3, gebtagtext)
                            if geburtstag.tm_yday == dienstag.tm_yday:
                                t.set_zelltext_i(aktZeile, 4, gebtagtext)
                            if geburtstag.tm_yday == mitwoch.tm_yday:
                                t.set_zelltext_i(aktZeile, 5, gebtagtext)
                            if geburtstag.tm_yday == donnerstag.tm_yday:
                                t.set_zelltext_i(aktZeile, 6, gebtagtext)
                            if geburtstag.tm_yday == freitag.tm_yday:
                                t.set_zelltext_i(aktZeile, 7, gebtagtext)
                            pass
                        aktZeile += 1
                    pass
                aktZeile += 1
            pass
        pass  
    def setup_for_printing(self):
        tab = ol_tabelle()
        tab.set_seitenformat("A3", True, 3000, 500, 500 , 500, False, False)
        tab.set_pageScaling(82)
        pass
    def ist_Urlaub(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Urlaub")
                farbe = RGBTo32bitInt(153, 204, 0) # grün
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Zeitausgleich(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Zeitausgleich")
                farbe = RGBTo32bitInt(153, 204, 0) # grün
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Lieferung(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Lieferung")
                farbe = RGBTo32bitInt(153, 204, 0) # grün
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Kurzarbeit(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Kurzarbeit")
                farbe = RGBTo32bitInt(153, 204, 0) # grün
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Montage(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Montage")
                farbe = RGBTo32bitInt(255, 102, 0) # orange
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_krank(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "krank")
                farbe = RGBTo32bitInt(255, 0, 0) # rot
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Berufsschule(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Berufsschule")
                farbe = RGBTo32bitInt(153, 204, 255) # blau
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def ist_Lehrgang(self):
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):
            for s in range(iSpalteStart, iSpalteEnde+1):
                tab.set_zelltext_i(z, s, "Lehrgang")
                farbe = RGBTo32bitInt(153, 204, 255) # blau
                tab.set_zellfarbe_i(z, s, farbe)
                pass
            pass
        pass
    def get_tagesplan(self):
        tab = ol_tabelle()
        kw = tab.get_tabname()
        gesund = True
        try:
            val = int(kw) # versuche ob der string kw in einen int umgewandelt werden kann
        except ValueError:
            msg = "Bitte zuerst in einen konkreten Wochenplan wechseln!"
            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
            gesund = False

        if gesund == True:
            path = get_userpath()
            path += "\\Desktop\\Tagesplan_"
            path += self.get_jahr()
            path += "_KW"
            # path += self.get_kw()
            path += kw
            path += ".odt"
            # erstelle_datei(path)
            # schreibe_in_datei(path, "Andre Klapper\n123")
            # Inhalt erfassen:
            gruppenNamen = self.get_gruppennamen()
            mitarb = []
            motag = []
            ditag = []
            mitwo = []
            dotag = []
            frtag = []
            satag = []
            for i in range(3, 99):
                ma = tab.get_zelltext_i(i, 0)
                mo = tab.get_zelltext_i(i, 3)
                di = tab.get_zelltext_i(i, 4)
                mi = tab.get_zelltext_i(i, 5)
                do = tab.get_zelltext_i(i, 6)
                fr = tab.get_zelltext_i(i, 7)
                sa = tab.get_zelltext_i(i, 8)
                if ma in gruppenNamen:
                    continue # for
                if len(ma) == 0:
                    continue # for
                if len(ma) > 50: #      Werte Kollegen  es gelten bis auf weiteres folgende Arbeitszeiten:                           Frühschicht von 6.00 bis 15.00 Uhr
                    break # for
                mitarb += [ma] # Kapselung notwendig
                motag += [mo]
                ditag += [di]
                mitwo += [mi]
                dotag += [do]
                frtag += [fr]
                satag += [sa]
                pass
            # msgbox(mitarb, 'msgbox', 1, 'QUERYBOX')
            # Inhalt zusammenstellen:
            taplan = ""
            for i in range(0, len(mitarb)):
                # Kopfzeile:
                lentren = 55
                lentren -= len(mitarb[i])
                lentren -= 2
                lentren = lentren/2
                trenner = ""
                for ii in range(0, int(lentren)):
                    trenner += "-"
                    pass
                taplan += "KW "
                taplan += tab.get_tabname()
                taplan += "/"
                taplan += self.get_jahr()
                taplan += " "
                taplan += trenner
                taplan += " "
                taplan += mitarb[i]
                taplan += " "
                taplan += trenner
                taplan += "\n"
                taplan += "\n"
                # 
                bemerkung = "    - "
                # Montag:
                # label = tab.get_zelltext_s("D3")
                label = "Montag    "
                taplan += label
                taplan += ": "
                taplan += motag[i]
                taplan += "\n"
                taplan += bemerkung
                taplan += "\n"
                # Dienstag:
                # label = tab.get_zelltext_s("E3")
                label = "Dienstag  "
                taplan += label
                taplan += ": "
                taplan += ditag[i]
                taplan += "\n"
                taplan += bemerkung
                taplan += "\n"
                # Mittwoch:
                # label = tab.get_zelltext_s("F3")
                label = "Mittwoch  "
                taplan += label
                taplan += ": "
                taplan += mitwo[i]
                taplan += "\n"
                taplan += bemerkung
                taplan += "\n"
                # Donnerstag:
                # label = tab.get_zelltext_s("G3")
                label = "Donnerstag"
                taplan += label
                taplan += ": "
                taplan += dotag[i]
                taplan += "\n"
                taplan += bemerkung
                taplan += "\n"
                # Freitag:
                # label = tab.get_zelltext_s("H3")
                label = "Freitag   "
                taplan += label
                taplan += ": "
                taplan += frtag[i]
                taplan += "\n"
                taplan += "    - "
                taplan += "\n"
                # Samstag:
                if len(satag[i]) > 0:
                    label = tab.get_zelltext_s("I3")
                    if len(label) > 0:
                        taplan += label
                        taplan += ": "
                    taplan += satag[i]
                    taplan += "\n"
                    taplan += bemerkung
                    taplan += "\n"
                # Freizeilen:
                taplan += "\n"
                taplan += "\n"
                taplan += "\n"
                pass
            
            if schreibe_in_datei(path, taplan) == True:
                msg = "Tagespläne wurden erfolgrich gespeichert."
                msgbox(msg, 'msgbox', 1, 'QUERYBOX')
            else:
                msg = "Tagespläne konnten nicht gespeichert werden."
                msgbox(msg, 'msgbox', 1, 'QUERYBOX')
        pass
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
class kalkulation: #Calc
    def __init__(self):
        self.context = XSCRIPTCONTEXT # globale Variable im sOffice-kontext
        self.doc = self.context.getDocument() #aktuelles Document per Methodenaufruf ! mit Klammern !
        self.tab = ol_tabelle()        
        pass
    def get_zelltext(self):
        iZeile = 0
        iSpalte = 0
        tab = ol_tabelle()
        iZeileStart = tab.get_selection_zeile_start()
        iZeileEnde  = tab.get_selection_zeile_ende()
        iSpalteStart = tab.get_selection_spalte_start()
        iSpalteEnde = tab.get_selection_spalte_ende()
        for z in range(iZeileStart, iZeileEnde+1):# wird gebraucht zur Typenumwandlung
            for s in range(iSpalteStart, iSpalteEnde+1):
                iZeile = z
                iSpalte = s
                break # nur 1 durchlauf erwünscht
            break # nur 1 durchlauf erwünscht
        return self.tab.get_zelltext_i(iZeile, iSpalte)
    def set_fokus_tab(self, tabname):
        if self.tab.tab_existiert(tabname):
            self.tab.set_tabfokus_s(tabname)
        pass
    def erstelle_tab(self):
        sPosNr = self.get_zelltext()
        tab = ol_tabelle()
        gesund = tab.tab_kopieren2("leer", sPosNr, 9999)
        if self.tab.tab_existiert(sPosNr):
            tab.set_tabfokus_s(sPosNr)
            if gesund == 0:
                tab.set_zelltext_s("B3", sPosNr)
        pass
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
class TaPlan: # Writer
    def __init__(self):
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.text = self.doc.getText()
        self.desktop = XSCRIPTCONTEXT.getDesktop()
        self.model = self.desktop.getCurrentComponent()        
        pass
    def formartieren(self):
        self.set_text_hoehe(12)
        self.set_text_fett()
        pass
    def set_text_hoehe(self, iHoehe):
        oSel = self.doc.CurrentSelection.getByIndex(0) # get the current selection
        oTC = self.text.createTextCursorByRange(oSel) # TextCursor erzeugen
        oEnum = oTC.Text.createEnumeration()
        # oTC = oText.createTextCursorByRange(oSel)
        while oEnum.hasMoreElements():
            oPar = oEnum.nextElement()
            oPar.CharHeight = iHoehe
        pass
    def set_text_fett(self):
        fettMachen = []
        fettMachen += ["Montag    :"]
        fettMachen += ["Dienstag  :"]
        fettMachen += ["Mittwoch  :"]
        fettMachen += ["Donnerstag:"]
        fettMachen += ["Freitag   :"]
        for i in range(0, len(fettMachen)):
            suche = self.doc.createSearchDescriptor()
            # suche.SearchString = "Montag"
            suche.SearchString = fettMachen[i]
            suche.SearchWords = True # nur ganze Wörter suchen
            suche.SearchCaseSensitive = True # Groß/Klein-Schreibung beachten
            funde = self.doc.findAll(suche)
            for ii in range(0, funde.getCount()):
                fund = funde.getByIndex(ii)
                fund.CharWeight = FONT_BOLD
                fund.CharUnderline = FONT_UNDERLINED_SINGLE
                # fund.setString("neuer text") # Suchergebnis ersetzen durch
                pass
            pass
        pass
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
class baugrpetk_writer: # Writer
    def __init__(self):      
        self.t = ol_textdatei()
        pass
    def formartieren(self):
        self.t.set_text_hoehe(24)
        self.set_text_fett()
        self.t.set_seitenformat("A6", True, 2000,1000,1000,1000)
        pass
    def set_text_fett(self):
        fettMachen = []
        fettMachen += ["Projekt  :"]
        fettMachen += ["Opti     :"]
        fettMachen += ["Pos      :"]
        fettMachen += ["Baugruppe:"]
        fettMachen += ["Ort:"]
        for i in range(0, len(fettMachen)):
            suche = self.t.doc.createSearchDescriptor()
            # suche.SearchString = "Montag"
            suche.SearchString = fettMachen[i]
            suche.SearchWords = True # nur ganze Wörter suchen
            suche.SearchCaseSensitive = True # Groß/Klein-Schreibung beachten
            funde = self.t.doc.findAll(suche)
            for ii in range(0, funde.getCount()):
                fund = funde.getByIndex(ii)
                fund.CharWeight = FONT_BOLD
                fund.CharUnderline = FONT_UNDERLINED_SINGLE
                # fund.setString("neuer text") # Suchergebnis ersetzen durch
                pass
            pass
        pass

#----------------------------------------------------------------------------------

def test_123():
    # sli = slist()
    # sli.reduzieren()
    # wplan = WoPlan()
    # wplan.wochenplan_erstellen()
    # create_file("C:\\Users\\AV6\\Desktop\\Unbekannt123.odt")
    # path = get_userpath()
    # path += "\\Desktop\\Unbekannt123.odt"
    # create_file(path)
    # os.system('notepad.exe')
    # os.system("swriter.exe")
    # os.system("C:\\Users\\AV6\\Desktop\\Unbekannt123.odt")
    # t = ol_textdatei()
    # wpl = WoPlan()
    # wpl.get_tagesplan()
    # t = TaPlan()
    # t.set_text_hoehe(12)
    # t.set_text_fett()
    #sli = slist()
    #sTabname = sli.t.get_tabname()
    #sli.t.tab_kopieren("hans", 99)
    msg = "Die Testfunktion ist derzeit nicht in Nutzung."
    msgbox(msg, 'msgbox', 1, 'QUERYBOX')
    pass

#----------------------------------------------------------------------------------
# Starter für die Bedienung im Menü:
def tabelle_set_tabfokus_uebersicht():
    tab = ol_tabelle()
    tabname = "Übersicht"
    if tab.tab_existiert(tabname):
        tab.set_tabfokus_s(tabname)
    pass
#---------
def SList_autoformat():
    sli = slist()
    sli.autoformat()
    pass
def SList_Formeln_edit():
    sli = slist()
    sli.formeln_edit()
    pass
def SList_Formeln_Kante():
    sli = slist()
    sli.formeln_kante()
    pass
def SList_Kanteninfo_beraeumen():
    sli = slist()
    sli.kanteninfo_beraeumen()
    pass
def SList_Teil_drehen():
    sli = slist()
    sli.teil_drehen()
    pass
def SList_sortieren():
    sli = slist()
    sli.sortieren()
    pass
def SList_reduzieren():
    sli = slist()
    sli.reduzieren()
    pass
def SList_sortieren_reduzieren():
    sli = slist()
    sli.std_namen()
    sli.reduzieren()
    sli.sortieren()
    pass
def SList_gehr_masszugabe():
    sli = slist()
    sli.gehr_masszugabe()
    pass
def SList_tap_anlegen_uebersicht():
    sli = slist()
    sli.tap_anlegen_uebersicht()
    pass
def SList_tab_anlegen_stklistpos():
    sli = slist()
    sli.tab_anlegen_stklistpos()
    pass
def SList_check_cncdata():
    sli = slist()
    sli.check_cncdata()
    pass
#---------
def RB_Blancoliste():
    l = raumbuch()
    l.RB_Blankoliste()
    pass
def RB_LList_Formblatt():
    l = raumbuch()
    l.LList_Formblatt()
    pass
def RB_LList_start():
    l = raumbuch()
    l.LList_start()
    pass
#---------
def baugrpetk_calc_ermitteln():
    sli = baugrpetk_calc()
    sli.ermitteln()
    sli.auflisten()
    pass
def baugrpetk_calc_speichern():
    sli = baugrpetk_calc()
    sli.speichern()
    pass
#---------
def baugrpetk_writer_formartieren():
    obj = baugrpetk_writer()
    obj.formartieren()
    pass
#---------
def WoPlan_tab_Grundlagen():
    wpl = WoPlan()
    wpl.tab_Grundlagen()
    pass
def WoPlan_tab_KW():
    wpl = WoPlan()
    wpl.wochenplan_erstellen()
    pass
def WoPlan_ist_Urlaub():
    wpl = WoPlan()
    wpl.ist_Urlaub()
    pass
def WoPlan_ist_Zeitausgleich():
    wpl = WoPlan()
    wpl.ist_Zeitausgleich()
    pass
def WoPlan_ist_Lieferung():
    wpl = WoPlan()
    wpl.ist_Lieferung()
    pass
def WoPlan_ist_Kurzarbeit():
    wpl = WoPlan()
    wpl.ist_Kurzarbeit()
    pass
def WoPlan_ist_Montage():
    wpl = WoPlan()
    wpl.ist_Montage()
    pass
def WoPlan_ist_krank():
    wpl = WoPlan()
    wpl.ist_krank()
    pass
def WoPlan_ist_Berufsschule():
    wpl = WoPlan()
    wpl.ist_Berufsschule()
    pass
def WoPlan_ist_Lehrgang():
    wpl = WoPlan()
    wpl.ist_Lehrgang()
    pass
def WoPlan_Tagesplan():
    wpl = WoPlan()
    wpl.get_tagesplan()
    pass
#---------
def Kalkulation_set_tab_fokus():
    kal = kalkulation()
    kal.set_fokus_tab(kal.get_zelltext())    
    pass
def Kalkulation_pos_erstellen():
    kal = kalkulation()
    kal.erstelle_tab()    
    pass
#---------
def TaPlan_formartieren(): 
    tpl = TaPlan()
    tpl.formartieren()
    pass
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
# Starter für die Bedienung in der Symbolleiste:
def tabelle_set_tabfokus_uebersicht_BTN(self):
    tab = ol_tabelle()
    tabname = "Übersicht"
    if tab.tab_existiert(tabname):
        tab.set_tabfokus_s(tabname)
    pass
#---------
def SList_autoformat_BTN(self):
    sli = slist()
    sli.autoformat()
    pass
def SList_Formeln_edit_BTN(self):
    sli = slist()
    sli.formeln_edit()
    pass
def SList_Formeln_Kante_BTN(self):
    sli = slist()
    sli.formeln_kante()
    pass
def SList_Kanteninfo_beraeumen_BTN(self):
    sli = slist()
    sli.kanteninfo_beraeumen()
    pass
def SList_Teil_drehen_BTN(self):
    sli = slist()
    sli.teil_drehen()
    pass
def SList_sortieren_BTN(self):
    sli = slist()
    sli.sortieren()
    pass
def SList_reduzieren_BTN(self):
    sli = slist()
    sli.reduzieren()
    pass
def SList_sortieren_reduzieren_BTN(self):
    sli = slist()
    sli.std_namen()
    sli.reduzieren()
    sli.sortieren()
    pass
def SList_gehr_masszugabe_BTN(self):
    sli = slist()
    sli.gehr_masszugabe()
    pass
def SList_tap_anlegen_uebersicht_BTN(self):
    sli = slist()
    sli.tap_anlegen_uebersicht()
    pass
def SList_tab_anlegen_stklistpos_BTN(self):
    sli = slist()
    sli.tab_anlegen_stklistpos()
    pass
def SList_check_cncdata_BTN(self):
    sli = slist()
    sli.check_cncdata()
    pass
#---------
def RB_Blancoliste_BTN(self):
    l = raumbuch()
    l.RB_Blankoliste()
    pass
def RB_LList_Formblatt_BTN(self):
    l = raumbuch()
    l.LList_Formblatt()
    pass
def RB_LList_start_BTN(self):
    l = raumbuch()
    l.LList_start()
    pass
#---------
def WoPlan_tab_Grundlagen_BTN(self):
    wpl = WoPlan()
    wpl.tab_Grundlagen()
    pass
def WoPlan_tab_KW_BTN(self):
    wpl = WoPlan()
    wpl.wochenplan_erstellen()
    pass
def WoPlan_ist_Urlaub_BTN(self):
    wpl = WoPlan()
    wpl.ist_Urlaub()
    pass
def WoPlan_ist_Zeitausgleich_BTN(self):
    wpl = WoPlan()
    wpl.ist_Zeitausgleich()
    pass
def WoPlan_ist_Lieferung_BTN(self):
    wpl = WoPlan()
    wpl.ist_Lieferung()
    pass
def WoPlan_ist_Kurzarbeit_BTN(self):
    wpl = WoPlan()
    wpl.ist_Kurzarbeit()
    pass
def WoPlan_ist_Montage_BTN(self):
    wpl = WoPlan()
    wpl.ist_Montage()
    pass
def WoPlan_ist_krank_BTN(self):
    wpl = WoPlan()
    wpl.ist_krank()
    pass
def WoPlan_ist_Berufsschule_BTN(self):
    wpl = WoPlan()
    wpl.ist_Berufsschule()
    pass
def WoPlan_ist_Lehrgang_BTN(self):
    wpl = WoPlan()
    wpl.ist_Lehrgang()
    pass
def WoPlan_Tagesplan_BTN(self):
    wpl = WoPlan()
    wpl.get_tagesplan()
    pass
#---------
def Kalkulation_set_tab_fokus_BTN(self):
    kal = kalkulation()
    kal.set_fokus_tab(kal.get_zelltext())    
    pass
def Kalkulation_pos_erstellen_BTN(self):
    kal = kalkulation()
    kal.erstelle_tab()    
    pass
#---------
def TaPlan_formartieren_BTN(self): 
    tpl = TaPlan()
    tpl.formartieren()
    pass
#----------------------------------------------------------------------------------

# Notizen:

#----------------------------------------------------------------------------------
#def getCurrentRegion(oRange):
#    """Get current region around given range."""
#    oCursor = oRange.getSpreadsheet().createCursorByRange(oRange)
#    #oCursor.collapseToCurrentRegion()
#    return oCursor

#def getCurrentColumnsAddress(oRange):
#    """Get address of intersection between range and current region's columns"""
#    oCurrent = getCurrentRegion(oRange)
#    oAddr = oRange.getRangeAddress()
#    oCurrAddr = oCurrent.getRangeAddress()
#    oAddr.StartColumn = oCurrAddr.StartColumn
#    oAddr.EndColumn = oCurrAddr.EndColumn
#    return oAddr
#----------------------------------------------------------------------------------