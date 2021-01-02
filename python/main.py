# coding: utf-8
from __future__ import unicode_literals
import uno
import os
from sys import platform


from com.sun.star.beans import PropertyValue
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS

doc = XSCRIPTCONTEXT.getDocument()
CTX = uno.getComponentContext()
SM = CTX.getServiceManager()

def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance

def call_dispatch(doc, url, args=()):
        frame = doc.getCurrentController().getFrame()
        dispatch = create_instance('com.sun.star.frame.DispatchHelper')
        dispatch.executeDispatch(frame, url, '', 0, args)
        return

def dict_to_property(values, uno_any=False):
    ps = tuple([PropertyValue(Name=n, Value=v) for n, v in values.items()])
    if uno_any:
        ps = uno.Any('[]com.sun.star.beans.PropertyValue', ps)
    return ps

def getLanguage():
    Access   = "com.sun.star.configuration.ConfigurationAccess"
    ConfigProvider = create_instance("com.sun.star.configuration.ConfigurationProvider")
    Prop = {"nodepath":"org.openoffice.Office.Linguistic/General"}
    properties = dict_to_property(Prop)
    key = "UILocale"
    Set = ConfigProvider.createInstanceWithArguments(Access, properties)
    if Set and (Set.hasByName(key)):
        Lang = Set.getPropertyValue(key)

    if not (Lang and not Lang.isspace()):
        Prop = {"nodepath":"/org.openoffice.Setup/L10N"}
        properties = dict_to_property(Prop)
        key = "ooLocale"
        Set = ConfigProvider.createInstanceWithArguments(Access, properties)
        if Set and (Set.hasByName(key)):
            Lang = Set.getPropertyValue(key)
    return Lang

def initLn():

    language = getLanguage()

    if language == "it":
        ln = {
        "TitleFirstPage" : "Imposta Prima Pagina",
        "OptionLabel1" : "Imposta la pagina 1 con lo stile \"Prima Pagina\" e inserisce una interruzione con lo \"Stile Prefefinito\"",
        "OptionLabel2" : "Crea tre pagine, separate da interruzioni, con i seguenti stili: 1) Prima pagina, 2) Stile predefinito, 3) Pagina destra",
        "OptionLabel3" : "Imposta la pagina 1 con lo stile \"Prima Pagina\", inserisce una interruzione con lo \"Stile Prefefinito\". Le successive utilizzeranno gli stili \"Destra/Sinistra\"",
        "OptionLabel4" : "Imposta la pagina 1 con lo stile \"Prima Pagina\", la seconda utilizzerà lo \"Stile Prefefinito\". Dalla terza si alterneranno gli stili \"Destra/Sinistra\"",
        "OptionLabel5" : "Imposta la pagina 1 come \"Prima Pagina\" e le successive \"Sinistra\" e \"Destra\"",
        "CancelLabel" : "Annulla",
        "TitleInformation" : "Informazioni",
        "labelInformation" : "Per qualsiasi informazione o segnalazione potete contattare:",
        "heightPage" : "Altezza",
        "widthPage" : "Larghezza",
        "infoGroup" : "Informazioni sullo stile applicato a questa pagina:",
        "pageStyle" : "Stile pagina:",
        "followStyle" : "Stile successivo:",
        "nrPage" : "Numero pagina:"
        }

    elif language == "en":
        ln = {
        "TitleFirstPage" : "Set Front Page",
        "OptionLabel1" : "Set page 1 with the \"First Page\" style and insert a break with the \"Default Style\"",
        "OptionLabel2" : "Create three pages, separated by breaks, with the following styles: 1) First page, 2) Default style, 3) Right page",
        "OptionLabel3" : "Set page 1 with the \"First Page\" style, insert a break with the \"Default Style\". The next ones will use the \"Right/Left\" styles",
        "OptionLabel4" : "Set page 1 with the \"First Page\" style, the second will use the \"Default Style\". From the third, \"Right/Left\" styles will alternate",
        "OptionLabel5" : "Set page 1 as \"First Page\" and the following ones \"Left\" and \"Right\"",
        "CancelLabel" : "Cancel",
        "TitleInformation" : "Information",
        "labelInformation" : "You can contact for any information or report:",
        "heightPage" : "Page height",
        "widthPage" : "Page width",
        "infoGroup" : "Style information applied to this page:",
        "pageStyle" : "Page style:",
        "followStyle" : "Next style:",
        "nrPage" : "Page number:"}

    elif language == "de":
        ln = {
        "TitleFirstPage" : "Erste Seite einstellen",
        "OptionLabel1" : "Setzt Seite 1 mit Vorlage \"Erste Seite\" und fügt einen Umbruch mit der Vorlage \"Standard\" ein.",
        "OptionLabel2" : "Erstellt drei Seiten, die mit Seitenumbrüchen getrennt sind, mit den folgenden Vorlagen: 1) Erste Seite, 2) Standard, 3) Rechte Seite.",
        "OptionLabel3" : "Setzt Seite 1 mit Vorlage \"Erste Seite\" und fügt einen Umbruch mit der Vorlage \"Standard\" ein. Nachfolgende Seiten verwenden die Vorlagen \"Rechts/Links\".",
        "OptionLabel4" : "Setzt Seite 1 mit der Seitenvorlage \"Erste Seite\", die zweite Seite wird die Seitenvorlage \"Standard\" verwenden. Ab der dritten Seite wird zwischen den Seitenvorlagen \"Rechts/Links\" gewechselt.",
        "OptionLabel5" : "Setz Seite 1 als \"Erste Seite\" und die folgenden Seiten als \"Links\" und \"Rechts\".",
        "CancelLabel" : "Abbrechen",
        "TitleInformation" : "Über uns",
        "labelInformation" : "Für weitere Informationen oder Hinweise kontaktieren Sie uns bitte:",
        "heightPage" : "Seitenhöhe",
        "widthPage" : "Seitenbreite",
        "infoGroup" : "Angewandte Seitenvorlageninformationen:",
        "pageStyle" : "Seitenvorlage:",
        "followStyle" : "Folgevorlage:",
        "nrPage" : "Seitennummer:"}

    elif language == "fr":
        ln = {
        "TitleFirstPage" : "Définir la page d'accueil",
        "OptionLabel1" : "Définissez la page 1 avec le style \"Première page\" et insérez un saut avec le \"Style par défaut\"",
        "OptionLabel2" : "Créez trois pages, séparées par des sauts, avec les styles suivants: 1) Première page, 2) Style par défaut, 3) Page de droite",
        "OptionLabel3" : "Définissez la page 1 avec le style \"Première page\", insérez un saut avec le \"Style par défaut\". Ce qui suit utilisera les styles \"Droite/Gauche\"",
        "OptionLabel4" : "Définissez la page 1 avec le style \"Première page\", la seconde utilisera le \"Style par défaut\". À partir du troisième, les styles \"Droite/Gauche\" alterneront",
        "OptionLabel5" : "Définissez la page 1 comme \"Première page\" et les suivantes \"Gauche\" et \"Droite\"",
        "CancelLabel" : "Annuler",
        "TitleInformation" : "Information",
        "labelInformation" : "Vous pouvez contacter pour toute information ou rapport:",
        "heightPage" : "Page Hauteur",
        "widthPage" : "Page Largeur",
        "infoGroup" : "Des informations de style appliqué à cette page:",
        "pageStyle" : "Style de page:",
        "followStyle" : "Style suivant:",
        "nrPage" : "Numéro de page:"}

    elif language == "es":
        ln = {
        "TitleFirstPage" : "Establecer portada",
        "OptionLabel1" : "Establecer la página 1 con el estilo \"Primera página\" e insertar un salto con el \"Estilo predeterminado\"",
        "OptionLabel2" : "Cree tres páginas, separadas por saltos, con los siguientes estilos: 1) Primera página, 2) Estilo predeterminado, 3) Página derecha",
        "OptionLabel3" : "Establezca la página 1 con el estilo \"Primera página\", inserte un salto con el \"Estilo predeterminado\". Lo siguiente utilizará los estilos \"Derecha/Izquierda\"",
        "OptionLabel4" : "Configure la página 1 con el estilo \"Primera página\", la segunda utilizará el \"Estilo predeterminado\". A partir del tercero, se alternarán los estilos \"Derecha/Izquierda\"",
        "OptionLabel5" : "Establecer la página 1 como \"Primera página\" y las páginas siguientes \"Izquierda\" y \"Derecha\"",
        "CancelLabel" : "Cancelar",
        "TitleInformation" : "Información",
        "labelInformation" : "Puede ponerse en contacto para cualquier información o informe:",
        "heightPage" : "Altura de la página",
        "widthPage" : "Ancho de página",
        "infoGroup" : "Información de estilo aplicado a esta página:",
        "pageStyle" : "Estilo de página:",
        "followStyle" : "Siguiente estilo:",
        "nrPage" : "Número de página:"}

    elif language == "pt":
        ln = {
        "TitleFirstPage" : "Definir primeira página",
        "OptionLabel1" : "Defina a página 1 com o estilo \"Primeira página\" e insira uma quebra com o \"Estilo padrão\"",
        "OptionLabel2" : "Crie três páginas, separadas por quebras, com os seguintes estilos: 1) Primeira página, 2) Estilo padrão, 3) Página direita",
        "OptionLabel3" : "Defina a página 1 com o estilo \"Primeira página\", insira uma quebra com o \"Estilo padrão\". O seguinte usará os estilos \"Direita/Esquerda\"",
        "OptionLabel4" : "Defina a página 1 com o estilo \"Primeira página\", a segunda usará o \"Estilo padrão\". A partir do terceiro, os estilos \"Direita/Esquerda\" serão alternados",
        "OptionLabel5" : "Defina a página 1 como \"Primeira página\" e as seguintes \"Esquerda\" e \"Direita\"",
        "CancelLabel" : "Cancelar",
        "TitleInformation" : "Informação",
        "labelInformation" : "Você pode entrar em contato para qualquer informação ou relatório:",
        "heightPage" : "Altura da página",
        "widthPage" : "Largura da página",
        "infoGroup" : "Informações de estilo aplicado a esta página:",
        "pageStyle" : "Estilo de página:",
        "followStyle" : "Próximo estilo:",
        "nrPage" : "Número de página:"}

    else:
        ln = {
        "TitleFirstPage" : "Set Front Page",
        "OptionLabel1" : "Set page 1 with the \"First Page\" style and insert a break with the \"Default Style\"",
        "OptionLabel2" : "Create three pages, separated by breaks, with the following styles: 1) First page, 2) Default style, 3) Right page",
        "OptionLabel3" : "Set page 1 with the \"First Page\" style, insert a break with the \"Default Style\". The next ones will use the \"Right/Left\" styles",
        "OptionLabel4" : "Set page 1 with the \"First Page\" style, the second will use the \"Default Style\". From the third, \"Right/Left\" styles will alternate",
        "OptionLabel5" : "Set page 1 as \"First Page\" and the following ones \"Left\" and \"Right\"",
        "CancelLabel" : "Cancel",
        "TitleInformation" : "Information",
        "labelInformation" : "You can contact for any information or report:",
        "heightPage" : "Height page",
        "widthPage" : "Width page",
        "infoGroup" : "Style information applied to this page:",
        "pageStyle" : "Page style:",
        "followStyle" : "Next style:",
        "nrPage" : "Page number:"}

    ln["labelEmail"] = "antonio.faccioli@libreoffice.org"

    return ln

def informationdlg(*param):

    ln = initLn()

    dp = SM.createInstanceWithContext("com.sun.star.awt.DialogProvider", CTX)
    dialog = dp.createDialog( "vnd.sun.star.script:PortraitOrLandscape.Information?location=application")

    dlg = dialog.Model

    dlg.Title = ln["TitleInformation"]

    infoGroupTitle = dlg.getByName("infoGroup")
    infoGroupTitle.Label = ln["infoGroup"]

    infoNrPage = dlg.getByName("nrPage")
    infoNrPage.Label = ln["nrPage"] + " " + str(pageNumber())

    infoStyle = dlg.getByName("infoStyle")
    infoStyle.Label = ln["pageStyle"] + " " + nameCurrentStyle()

    infoFollowStyle = dlg.getByName("followStyle")
    infoFollowStyle.Label = ln["followStyle"] + " " + styleFollow(indexCurrentStyle())

    infoSize = dlg.getByName("infoSize")
    infoSize.Label = ln["widthPage"] + ": " + "{:.2f}".format(width()/1000) + " cm - " + ln["heightPage"] + ": " + "{:.2f}".format(height()/1000) + " cm"

    infoLabel = dlg.getByName("labelInformation")
    infoLabel.Label = ln["labelInformation"]

    emailLabel = dlg.getByName("labelEmail")
    emailLabel.Label = ln["labelEmail"]

    dialog.execute()

def firstPageDialog(*param):

    ln = initLn()

    dp = SM.createInstanceWithContext("com.sun.star.awt.DialogProvider", CTX)
    dialog = dp.createDialog( "vnd.sun.star.script:PortraitOrLandscape.FirstPage?location=application")

    dlg = dialog.Model
    dlg.Title = ln["TitleFirstPage"]

    OptionButton1 = dlg.getByName("OptionButton1")
    OptionButton2 = dlg.getByName("OptionButton2")
    OptionButton3 = dlg.getByName("OptionButton3")
    OptionButton4 = dlg.getByName("OptionButton4")
    OptionButton5 = dlg.getByName("OptionButton5")

    OptionButton1.Label = ln["OptionLabel1"]
    OptionButton2.Label = ln["OptionLabel2"]
    OptionButton3.Label = ln["OptionLabel3"]
    OptionButton4.Label = ln["OptionLabel4"]
    OptionButton5.Label = ln["OptionLabel5"]

    OptionButton1.State = True

    CancelButton = dlg.getByName("Cancel")
    CancelButton.Label = ln["CancelLabel"]

    if dialog.execute() == 1:
        if OptionButton1.State == True:
            firstPageBreak()
        elif OptionButton2.State == True:
            firstPageBreakRight()
        elif OptionButton3.State == True:
            firstPageBreakSetRight()
        elif OptionButton4.State == True:
            firstPageSetDefRight()
        elif OptionButton5.State == True:
            firstPageSetLeft()
    else:
        dialog.endExecute()

def insertPageBreak(IndexStyle):

    styleName = doc.StyleFamilies.getByName("PageStyles").getByIndex(IndexStyle).DisplayName

    args = { 'Kind':3,
             'TemplateName':styleName,
             'PageNumber':0}

    args = dict_to_property(args)

    call_dispatch(doc, '.uno:InsertBreak', args)


def pageNumber():

    cursor = doc.CurrentController.getViewCursor()

    pageNumber = cursor.getPage()

    return pageNumber


def styleApply(IndexStyle):

    styleName = doc.StyleFamilies.getByName("PageStyles").getByIndex(IndexStyle).DisplayName

    args = {"Template":styleName,
            "Family":8}

    args = dict_to_property(args)

    call_dispatch(doc, '.uno:StyleApply', args)

def setFollowStyle(IndexStyle, IndexFollow):

    styleFollow = doc.StyleFamilies.getByName("PageStyles").getByIndex(IndexFollow).DisplayName

    pageStyle = doc.StyleFamilies.getByName("PageStyles").getByIndex(IndexStyle)
    pageStyle.FollowStyle = styleFollow

def styleFollow(IndexStyle):

    pageStyle = doc.StyleFamilies.getByName("PageStyles").getByIndex(IndexStyle)
    followStyle = doc.StyleFamilies.getByName("PageStyles").getByName(pageStyle.FollowStyle).DisplayName

    return followStyle

def changeSize(intWidth, intHeight):

    cursor = doc.CurrentController.getViewCursor()

    pageStyle = doc.StyleFamilies.getByName("PageStyles").getByName(cursor.PageStyleName)
    pageStyle.Width = intWidth
    pageStyle.Height = intHeight

def nameCurrentStyle():

    cursor = doc.CurrentController.getViewCursor()

    nameCurrentStyle = doc.StyleFamilies.getByName("PageStyles").getByName(cursor.PageStyleName).DisplayName

    return nameCurrentStyle

def nameStyle(indexStyle):

    nameStyle = doc.StyleFamilies.getByName("PageStyles").getByIndex(indexStyle).DisplayName

    return nameStyle

def changeOrientation(*param):

    if nameCurrentStyle() == nameStyle(0):
        styleApply(9)
    else:
        styleApply(0)

def indexCurrentStyle():

    if nameCurrentStyle() == nameStyle(0):

        indexCurrentStyle = 0

    elif nameCurrentStyle() ==  nameStyle(1):

        indexCurrentStyle = 1

    elif nameCurrentStyle() ==  nameStyle(2):

        indexCurrentStyle = 2

    elif nameCurrentStyle() ==  nameStyle(3):

        indexCurrentStyle = 3

    elif nameCurrentStyle() ==  nameStyle(9):

        indexCurrentStyle = 9

    return indexCurrentStyle

def indexFollowStyle(IndexStyle):

    if styleFollow(IndexStyle) == nameStyle(0):

        indexFollowStyle = 0

    elif styleFollow(IndexStyle) ==  nameStyle(2):

        indexFollowStyle = 2

    elif styleFollow(IndexStyle) ==  nameStyle(3):

        indexFollowStyle = 3

    elif styleFollow(IndexStyle) ==  nameStyle(9):

        indexFollowStyle = 9

    return indexFollowStyle

def A5(*param):

    if indexCurrentStyle() == 9:
        changeSize(21000, 14800)
    else:
        changeSize(14800, 21000)

def A4(*param):

    if indexCurrentStyle() == 9:
        changeSize(29700, 21000)
    else:
        changeSize(21000, 29700)

def A3(*param):

    if indexCurrentStyle() == 9:
        changeSize(42000, 29700)
    else:
        changeSize(29700, 42000)

def pathUser():


    path = os.path.dirname(__file__)
    path = path.split("///")
    sepPath = os.path.sep

    if platform == "linux" or platform == "linux2":
        url = sepPath+path[1]+sepPath
    elif platform == "darwin":
        urlTemp = path[1].replace('%20',' ')
        url = sepPath+urlTemp+sepPath
    else:
        url = path[1]+sepPath

    return url

def insLandscape(*param):

    if indexCurrentStyle() != 9:
        prevStyle = indexCurrentStyle()

        try:
            path = pathUser()
            ini = open(path+"info.ini", "w")
            ini.write(str(prevStyle))
            ini.close
        except:
            pass

    insertPageBreak(9)

def insPortrait(*param):

    try:
        path = pathUser()
        ini = open(path+"info.ini", "r")
        style = ini.readline()
        if style == 'S':
            pass
        else:
            style = int(style)
        ini.close

        if style == 2 or style == 3:
            if pageNumber() % 2 == 0:
                insertPageBreak(3)
            else:
                insertPageBreak(2)
        elif style == 1:
            if indexFollowStyle(1) == 0:
                insertPageBreak(0)
            elif indexFollowStyle(1) == 2 or indexFollowStyle(1) == 3:
                if pageNumber() % 2 == 0:
                    insertPageBreak(3)
                else:
                    insertPageBreak(2)
        elif style == 0:
            if indexFollowStyle(0) == 0:
                insertPageBreak(0)
            elif indexFollowStyle(0) == 2 or indexFollowStyle(0) == 3:
                if pageNumber() % 2 == 0:
                    insertPageBreak(3)
                else:
                    insertPageBreak(2)
        else:
            iStyle = indexCurrentStyle()
            fStyle = indexFollowStyle(iStyle)
            
            if fStyle == 2 or fStyle == 3:
                if pageNumber() % 2 == 0:
                    insertPageBreak(3)
                else:
                    insertPageBreak(2)
            else:
                insertPageBreak(fStyle)
        
        ini = open(path+"info.ini", "w")
        ini.write('S')
        ini.close

    except:
        insertPageBreak(0)

def width():

    cursor = doc.CurrentController.getViewCursor()

    pageStyle = doc.StyleFamilies.getByName("PageStyles").getByName(cursor.PageStyleName)
    width = pageStyle.Width

    return width

def height():

    cursor = doc.CurrentController.getViewCursor()

    pageStyle = doc.StyleFamilies.getByName("PageStyles").getByName(cursor.PageStyleName)
    height = pageStyle.Height

    return height

def firstPage(*param):

    if nameCurrentStyle() == nameStyle(0):
        styleApply(1)
    else:
        styleApply(0)

def firstPageBreak():

    styleApply(1)
    insertPageBreak(0)

def firstPageBreakRight():

    styleApply(1)
    insertPageBreak(0)
    insertPageBreak(3)

def firstPageBreakSetRight():

    styleApply(1)
    insertPageBreak(0)
    setFollowStyle(0,3)

def firstPageSetDefRight():

    styleApply(1)
    setFollowStyle(1,0)
    setFollowStyle(0,3)

def firstPageSetLeft():

    styleApply(1)
    setFollowStyle(1,2)
