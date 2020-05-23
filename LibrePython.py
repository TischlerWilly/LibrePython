from __future__ import unicode_literals
import uno
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.sheet.CellInsertMode import RIGHT as INSERT_RE
from com.sun.star.sheet.CellInsertMode import DOWN as INSERT_UN
from com.sun.star.table.CellHoriJustify import LEFT as AUSRICHTUNG_HORI_Li
from com.sun.star.table.CellHoriJustify import CENTER as AUSRICHTUNG_HORI_MI
from com.sun.star.table.CellHoriJustify import RIGHT as AUSRICHTUNG_HORI_RE
from com.sun.star.sheet.CellDeleteMode import LEFT as DEL_LI
from com.sun.star.table.CellContentType import TEXT as CELLCONTENTTYP_TEXT

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
    """ Anwendung:
    msgbox('Hallo Oliver', 'msgbox', 1, 'QUERYBOX')
    """
#----------------------------------------------------------------------------------
def RGBTo32bitInt(r, g, b):
  return int('%02x%02x%02x' % (r, g, b), 16)
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
    #-----------------------------------------------------------------------------------------------
    # Zellen:
    #-----------------------------------------------------------------------------------------------
    def get_zelle_i(self, zeile, spalte):
        return self.sheet.getCellByPosition(spalte, zeile)
        # Anwendung: text = t.get_zelle_i(1,1)
    def zelle_verschieben_i(self, iZeileVon, iSpalteVon, iZeileNach, iSpalteNach):
        source = self.sheet.getCellRangeByPosition(iSpalteVon, iZeileVon, iSpalteVon, iZeileVon)
        target = self.sheet.getCellByPosition(iSpalteNach, iZeileNach)
        self.sheet.moveRange(target.CellAddress, source.RangeAddress)
        pass
    def set_zelltext_s(self, zellname, text): # self muss immer als erster Parameter übergeben werden
        self.sheet.getCellRangeByName(zellname).String = text
        pass
        # Anwendung: t.set_zelltext_s('A1', 'Hallo 1')
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
    def set_zellfarbe_s(self, zellname, farbe): # farbe ist ein long-wert
        self.sheet.getCellRangeByName(zellname).CellBackColor = farbe
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
class slist:
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
        self.t.set_zelltext_s("G1", "Kante links")
        self.t.set_zelltext_s("H1", "KaDi")
        self.t.set_zelltext_s("I1", "Kante rechts")
        self.t.set_zelltext_s("J1", "KaDi")
        self.t.set_zelltext_s("K1", "Kante oben")
        self.t.set_zelltext_s("L1", "KaDi")
        self.t.set_zelltext_s("M1", "Kante unten")
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
                        aKanten.append(sKaLi)
                if (len(sKaRe) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaRe:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        aKanten.append(sKaRe)
                if (len(sKaOb) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaOb:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        aKanten.append(sKaOb)
                if (len(sKaUn) > 0):
                    bBekannt = False
                    for ii in range (0, len(aKanten)):
                        if aKanten[ii] == sKaOb:
                            bBekannt = True
                            break # für For-Schleife ii
                    if bBekannt == False:
                        aKanten.append(sKaOb)
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
        rankingList = ["Seite_li", "Seite_re", "Seite", "MS_li", "MS_re", "MS", "OB", "UB", "KB_ob", "KB_mi", "KB_un", "KB", "EB", "RW", "Tuer_li", "Tuer_re", "Tuer", "SF_A", "SF_B", "SF_C", "SF_D", "SF_E", "SF", "Sockel"]
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
        source = self.t.sheet.getCellRangeByPosition(0, iZeileStart, 15, iZeileEnde)
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
                        source = self.t.sheet.getCellRangeByPosition(0, rankingVonZeile[ii]+tmpPosDiff, 15, rankingVonZeile[ii]+tmpPosDiff)
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
                msg = "Zeile " + str(i) + " | Bez in liste"
                msgbox(msg, 'msgbox', 1, 'QUERYBOX')
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
                                                                            msg = "Zeile " + str(i) + "zusammenführen"
                                                                            msgbox(msg, 'msgbox', 1, 'QUERYBOX')
                                                                            gefunden = True
                                                                            # Zeilen zusammenführen:
                                                                            neueAnz = list_anz[ii] + self.t.get_zellzahl_i(i, 1)
                                                                            self.t.set_zellzahl_i(list_zeiNum[ii], 1, neueAnz)
                                                                            # Zeile löschen:
                                                                            # ....
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
        
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
def test_123():
    sli = slist()
    # sli.tabkopf_anlegen()
    # sli.dicke_aus_artikelnummer_bestimmen()
    sli.reduzieren()
    pass






#----------------------------------------------------------------------------------
# Starter für die Bedienung im Calc-Menü:
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
#----------------------------------------------------------------------------------
#----------------------------------------------------------------------------------
# Starter für die Bedienung in der Calc-Symbolleiste:
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