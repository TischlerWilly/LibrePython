from __future__ import unicode_literals
import os
from os.path import expanduser
from pathlib import Path
import uno
import datetime
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
        msgbox(msg, 'msgbox', 1, 'QUERYBOX')
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
def get_userpath():
    return expanduser("~")
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
    def tab_anlegen(self, sTabname, iTabIndex):
        self.doc.Sheets.insertNewByName(sTabname, iTabIndex)
        pass
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
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.text = self.doc.getText()
        self.text.setString('Hello World in Python in Writer')

#----------------------------------------------------------------------------------
class slist: # Calc
    def __init__(self):
        self.t = ol_tabelle()
        self.maxistklen = 999   
        # Farben bestimmen:
        self.farblos = -1
        self.rot = RGBTo32bitInt(204, 0, 0)
        self.gelb = RGBTo32bitInt(255, 255, 0) 
        self.grau = RGBTo32bitInt(204, 204, 204) 
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
        pass
        # Anwendung: self.umwandeln_von_BCtoCSV()
    def formatieren(self):
        # Alle Zellen sichtbar machen:
        for i in range(0, 15):
            self.t.set_spalte_sichtbar_i(i, True)        
        # Zellgrößen anpassen:
        self.t.set_zeilenhoehen(700)
        self.t.set_spaltenbreite_i(0, 4330) # Bezeichnung
        self.t.set_spaltenbreite_i(1, 1410) # Anzahl
        self.t.set_spaltenbreite_i(2, 1320) # Länge
        self.t.set_spaltenbreite_i(3, 1320) # Breite
        self.t.set_spaltenbreite_i(4, 1220) # Dicke
        self.t.set_spaltenbreite_i(5, 3830) # Matieral
        self.t.set_spaltenbreite_i(6, 3000) # Kante links
        self.t.set_spaltenbreite_i(7, 900) # KaDi links
        self.t.set_spaltenbreite_i(8, 3000) # Kante rechts
        self.t.set_spaltenbreite_i(9, 900) # KaDi re
        self.t.set_spaltenbreite_i(10, 3000) # Kante oben
        self.t.set_spaltenbreite_i(11, 900) # KaDi oben
        self.t.set_spaltenbreite_i(12, 3000) # Kante unten
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
            self.umwandeln_von_BCtoCSV()
            self.formatieren()
            return True
        elif len(a) == 0 and len(b) == 0 and len(c) == 0: # Tabellenkopf fehlt, evtl istes ein eleere Tabelle
            self.tabkopf_anlegen()
            self.formatieren()
            return True
        elif a == "Bezeichnung" and b == "Anzahl" and c == "Länge": # Tabell ist bereits richtig formartiert
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
        self.t.set_zelltext_s("P1", "Projekt")
        self.t.set_zelltext_s("P2", "ABC01")
        self.t.set_zellfarbe_s("P2", self.gelb)
        # Breiten der ergänzten Spalten anpassen:
        for i in range (15, 30):
            self.t.optimale_spaltenbreite_i(i)
        self.t.set_spaltenbreite_i(17, 700)
        # Nullen in den Feldern der KaDi löschen:
        self.entferneKaDiNull()
        # Formeln einfügen:
        # Es müssen immer die englischen Funktionsnamen für die Calc-Funktionen verwendet werden!
        for i in range (1, 25):
            sZellname = "Q" + str(i+1)
            sFormel = "=IF(SUM(S" + str(i+1) + ":ZZ" + str(i+1) + ")=0;0;" + "\"Fehler\"" + ")"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- Anzahl der Fehler:
            sZellname = "R" + str(i+1)
            sFormel = "=SUM(S" + str(i+1) + ":ZZ" + str(i+1) + ")"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- PlattenDi:
            sZellname = "S" + str(i+1)
            sFormel = "=IF(ISBLANK((INDIRECT(" + "\"F\"" + "&ROW())));0;(IF(NUMBERVALUE((CONCATENATE(LEFT((INDIRECT(" + "\"F\"" + "&ROW()));2);" + "\",\"" + ";RIGHT(LEFT((INDIRECT(" + "\"F\"" + "&ROW()));3);1)));" + "\",\"" + ")=(INDIRECT(" + "\"E\"" + "&ROW()));;1)))"
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi links:
            sZellname = "T" + str(i+1)
            sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += "+" # Jetzt folg nächste Prüfung:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"G\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"H\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi rechts:
            sZellname = "U" + str(i+1)
            sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += "+" # Jetzt folg nächste Prüfung:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"I\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"J\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi oben:
            sZellname = "V" + str(i+1)
            sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += "+" # Jetzt folg nächste Prüfung:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"K\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"L\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
            # --- KaDi unten:
            sZellname = "W" + str(i+1)
            sFormel = "=IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));3);1))=" +"\"N\"" # Wenn Kante ein "N" als 3. Zeichen Enthällt (z.B. 10N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));4);1))=" + "\"N\"" # Wenn Kante ein "N" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += "+" # Jetzt folg nächste Prüfung:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));3);1))=" + "\"X\"" # Wenn Kante ein "X" als 3. Zeichen Enthällt (z.B. 10X410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0))" # wenn Feld für KaDi leer dann Fehler
            sFormel += ";" # Sonst Wenn:
            sFormel += "IF((RIGHT(LEFT((INDIRECT(" + "\"M\"" + "&ROW()));4);1))=" + "\"X\"" # Wenn Kante ein "X" als 4. Zeichen Enthällt (z.B. 100N410040_23)
            sFormel += ";" # Dann
            sFormel += "(IF(ISBLANK((INDIRECT(" + "\"N\"" + "&ROW())));1;0));0))" # wenn Feld für KaDi leer dann Fehler
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
            sFormel = "=IF((LEN(P$2)+LEN(INDIRECT(" + "\"A\"" + "&ROW()))+6)>28;1;0)" # Wenn BC > 28 dann Fehler
            self.t.set_zellformel_s(sZellname, sFormel)
        pass
        # Anwendung: self.formeln_edit()
    def formeln_kante(self):
        if self.autoformat() != True:
            return False
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
        self.t.set_spaltenbreite_i(16, 2700) # KantenNr
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
            pass
        # Formeln für Kantenfehler einfügen:
        self.t.set_spaltenausrichtung_i(15, "mi")
        for i in range (1, maxi+1):
            formel =  "=IF(C" + str(i+1) + "<240;IF(NOT(G" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<240;IF(NOT(I" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<240;IF(NOT(K" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<240;IF(NOT(M" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<80;IF(NOT(K" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(C" + str(i+1) + "<80;IF(NOT(M" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<80;IF(NOT(G" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
            formel += "+IF(D" + str(i+1) + "<80;IF(NOT(I" + str(i+1) + "=" + "\"\"" + ");1;0);0)"
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
        rankingList += ["EB_ob", "EB_li", "EB_mi", "EB_un", "EB_re", "EB"]
        rankingList += ["RW_ob", "RW_li", "RW_mi", "RW_un", "RW_re", "RW"]
        rankingList += ["Tuer_li", "Tuer_re", "Tuer_A", "Tuer_B", "Tuer_C", "Tuer_D", "Tuer_E", "Tuer"]
        rankingList += ["Front_li", "Front_re", "Front_A", "Front_B", "Front_C", "Front_D", "Front_E", "Front"]
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
            sName = sName.replace("Schubkasten Front", "S#_SF")
            sName = sName.replace("Travers Vorne", "S#_Traver_vo")
            sName = sName.replace("Travers Hinten", "S#_Traver_hi")
            #sName = sName.replace("", "S#_")
            self.t.set_zelltext_i(i, 0, sName)
            pass
        pass
        
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
            t.set_SchriftFett_s("A1:C1", True)
            t.set_zellfarbe_s("A1:C1", self.grau)
            t.set_Rahmen_komplett_s("A1:C20", self.RahLinDi)
            t.set_spaltenbreite_i(0, 4100)
            t.set_spaltenbreite_i(1, 2260)
            t.set_spaltenbreite_i(2, 2260)
            #---
            t.set_zelltext_s("E1", "Kalender-Jahr")
            t.set_zelltext_s("F1", "2020")
            t.set_zelltext_s("E2", "KW1 beginnt am")
            t.set_zelltext_datum_s("F2", "2019", "12", "30")
            t.set_zelltext_s("E3", "KW")
            t.set_zelltext_s("F3", "1")
            t.set_SchriftFett_s("E1:E3", True)            
            t.set_zellfarbe_s("E1:E3", self.grau)            
            t.set_Rahmen_komplett_s("E1:F3", self.RahLinDi)
            t.set_spaltenbreite_i(4, 3250) 
            #---
            t.set_zelltext_s("E5", "Gruppen")
            t.set_SchriftFett_s("E5", True)            
            t.set_zellfarbe_s("E5", self.grau)            
            t.set_Rahmen_komplett_s("E5:E15", self.RahLinDi)
            t.set_zelltext_s("E6", "Halle1")
            t.set_zelltext_s("E7", "Halle2")
            t.set_zelltext_s("E8", "Halle3")
            t.set_zelltext_s("E9", "Kraftfahrer")
            t.set_zelltext_s("E10", "Lehrlinge")
            t.set_zelltext_s("E11", "Büro")
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
        return self.tabGrundlagen.get_zelltext_s("F3")
    def get_jahr(self):
        return self.tabGrundlagen.get_zelltext_s("F1")
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
        startdatum = "$Grundlagen.F2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)" # startdatum + ( (KW-1) * 7 )
        t.set_zellformel_s("D3", formel)
        t.set_zellformat_s("D3", "TT.MM.JJJJ")
        tmp = "Montag "
        tmp += t.get_zelltext_s("D3")
        t.set_zelltext_s("D3", tmp)
        # Beschriftung Dienstag:
        startdatum = "$Grundlagen.F2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+1" # startdatum + ( (KW-1) * 7 ) + 1 Tag
        t.set_zellformel_s("E3", formel)
        t.set_zellformat_s("E3", "TT.MM.JJJJ")
        tmp = "Dienstag "
        tmp += t.get_zelltext_s("E3")
        t.set_zelltext_s("E3", tmp)
        # Beschriftung Mittwoch:
        startdatum = "$Grundlagen.F2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+2" # startdatum + ( (KW-1) * 7 ) + 2 Tage
        t.set_zellformel_s("F3", formel)
        t.set_zellformat_s("F3", "TT.MM.JJJJ")
        tmp = "Mittwoch "
        tmp += t.get_zelltext_s("F3")
        t.set_zelltext_s("F3", tmp)
        # Beschriftung Donnerstag:
        startdatum = "$Grundlagen.F2"
        formel = "=" + startdatum + "+((D1-1)" + "*7)+3" # startdatum + ( (KW-1) * 7 ) + 3 Tage
        t.set_zellformel_s("G3", formel)
        t.set_zellformat_s("G3", "TT.MM.JJJJ")
        tmp = "Donnerstag "
        tmp += t.get_zelltext_s("G3")
        t.set_zelltext_s("G3", tmp)
        # Beschriftung Freitag:
        startdatum = "$Grundlagen.F2"
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
        idSpalte = 4
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
        idZeile = 1
        idSpalte = 0
        for i in range(0, 50):
            tmp_mitarb = t.get_zelltext_i(idZeile+i, idSpalte)
            tmp_gruppe = t.get_zelltext_i(idZeile+i, idSpalte+1)
            tmp_taetig = t.get_zelltext_i(idZeile+i, idSpalte+2)
            if(len(tmp_gruppe) > 0):
                mitarb += [tmp_mitarb] # Kapselung nötig da sonst jeder einzelne Buchstabe als Einzelwert gedeutet wird
                gruppe += [tmp_gruppe]
                taetig += [tmp_taetig]
            pass
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
    msg = "Die Testfunktion ist derzeit nicht in Nutzung."
    msgbox(msg, 'msgbox', 1, 'QUERYBOX')
    pass

#----------------------------------------------------------------------------------
# Starter für die Bedienung im Menü:
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
def WoPlan_Tagesplan():
    wpl = WoPlan()
    wpl.get_tagesplan()
    pass
#---------
def TaPlan_formartieren(): 
    tpl = TaPlan()
    tpl.formartieren()
    pass
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
# Starter für die Bedienung in der Symbolleiste:
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
def WoPlan_Tagesplan_BTN(self):
    wpl = WoPlan()
    wpl.get_tagesplan()
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