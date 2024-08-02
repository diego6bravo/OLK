/* Copyright (c) Business Objects 2006. All rights reserved. */

var L_bobj_crv_MainReport = "Hoofdrapport";
// Viewer Toolbar tooltips
var L_bobj_crv_FirstPage = "Ga naar eerste pagina";
var L_bobj_crv_PrevPage = "Ga naar vorige pagina";
var L_bobj_crv_NextPage = "Ga naar volgende pagina";
var L_bobj_crv_LastPage = "Ga naar laatste pagina";
var L_bobj_crv_ParamPanel = "Parametervenster";
var L_bobj_crv_Parameters = "Parameters";
var L_bobj_crv_GroupTree = "Groepsstructuur";
var L_bobj_crv_DrillUp = "Diepteanalyse";
var L_bobj_crv_Refresh = "Rapport vernieuwen";
var L_bobj_crv_Zoom = "In-/uitzoomen";
var L_bobj_crv_SelectPage = "Ga naar pagina";
var L_bobj_crv_SearchText = "Zoeken naar tekst";
var L_bobj_crv_Export = "Dit rapport exporteren";
var L_bobj_crv_Print = "Dit rapport afdrukken";
var L_bobj_crv_TabList = "Tablijst";
var L_bobj_crv_Close = "Sluiten";
var L_bobj_crv_Logo=  "Business Objects Logo"

var L_bobj_crv_Show = "Weergeven";
var L_bobj_crv_Hide = "Verbergen";

var L_bobj_crv_submitBtnLbl = "Exporteren";
var L_bobj_crv_ActiveXPrintDialogTitle = "Afdrukken";
var L_bobj_crv_PDFPrintDialogTitle = "Afdrukken naar PDF";
var L_bobj_crv_PrintRangeLbl = "Paginabereik:";
var L_bobj_crv_PrintAllLbl = "Alle pagina\'s";
var L_bobj_crv_PrintPagesLbl = "Pagina\'s selecteren";
var L_bobj_crv_PrintFromLbl = "Van:";
var L_bobj_crv_PrintToLbl = "Tot:";
var L_bobj_crv_PrintInfoTitle = "Afdrukken naar PDF:";
var L_bobj_crv_PrintInfo1 = 'De viewer moet naar PDF exporteren om te kunnen afdrukken. Kies de optie Afdrukken in de PDF reader wanneer het document is geopend.';
var L_bobj_crv_PrintInfo2 = 'Opmerking: u moet een PDF reader (bijvoorbeeld Adobe Reader) hebben ge\u00EFnstalleerd om te kunnen afdrukken.';
var L_bobj_crv_PrintPageRangeError = "Voer een geldig paginabereik in.";

var L_bobj_crv_ExportBtnLbl = "Exporteren";
var L_bobj_crv_ExportDialogTitle = "Exporteren";
var L_bobj_crv_ExportFormatLbl = "Bestandsindeling:";
var L_bobj_crv_ExportInfoTitle = "U kunt als volgt exporteren:";

var L_bobj_crv_ParamsApply = "Toepassen";
var L_bobj_crv_ParamsAdvDlg = "Parameterwaarde bewerken";
var L_bobj_crv_ParamsDelete = "Verwijderen";
var L_bobj_crv_ParamsDeleteTooltip = "Parameterwaarde verwijderen";
var L_bobj_crv_ParamsAddValue = "Klik om toe te voegen...";
var L_bobj_crv_ParamsApplyTip = "Parameterwaarden toevoegen";
var L_bobj_crv_ParamsDlgTitle = "Waarden invoeren";
var L_bobj_crv_ParamsCalBtn = "Een agenda weergeven";
var L_bobj_crv_Undo= "Ongedaan maken";
var L_bobj_crv_UndoTip = "Alle wijzigingen herstellen";
var L_bobj_crv_ParamsDirtyTip = "Klik op de knop Toepassen om wijzigingen toe te passen";
var L_bobj_crv_ParamsDataTip = "Dit is een parameter voor het  ophalen van gegevens";
var L_bobj_crv_ParamsMaxNumDefaultValues = "Klik hier voor meer items...";

var L_bobj_crv_ParamsInvalidTitle = "De parameterwaarde is niet geldig";
var L_bobj_crv_ParamsTooLong = "Parameterwaarde mag niet langer zijn dan %1 tekens";
var L_bobj_crv_ParamsTooShort = "Parameterwaarde moet minstens %1 tekens lang zijn";
var L_bobj_crv_ParamsBadNumber = "Deze parameter is van het type \"Getal\" en mag alleen een minteken, cijfers (\"0-9\"), cijfergroepeersymbolen of een decimaalteken bevatten.";
var L_bobj_crv_ParamsBadCurrency = "Deze parameter is van het type \"Valuta\" en mag alleen een minteken, cijfers (\"0-9\"), cijfergroepeersymbolen of een decimaalteken bevatten.";
var L_bobj_crv_ParamsBadDate = "Deze parameter is van het type \"Datum\" en de juiste notatie is \"%1\" waarbij \"jjjj\" is het viercijferige jaartal, \"mm\" de maand (bijvoorbeeld januari = 1) en \"dd\" de dag van de maand zijn.";
var L_bobj_crv_ParamsBadTime = "Deze parameter is van het type \"Tijd\" en de juiste notatie is \"uuh:mm:ss\" waarbij \"uu\" het uur in een 24-uurs klok, \"mm\" de minuten en \"ss\" de seconden zijn.";
var L_bobj_crv_ParamsBadDateTime = "Deze parameter is van het type \"DatumTijd\" en de juiste notatie is \"%1 uu:mm:ss\". \"jjjj\" is het viercijferige jaartal, \"mm\" is de maand (bijvoorbeeld januari = 1), \"dd\" is de dag van de maand, \"uu\" is het uur in een 24-uurs klok, \"mm\" zijn de minuten en \"ss\" de seconden.";
var L_bobj_crv_ParamsMinTooltip = "Geef een %1-waarde op die groter dan of gelijk aan %2 is.";
var L_bobj_crv_ParamsMaxTooltip = "Geef een %1-waarde op die kleiner dan of gelijk aan %2 is.";
var L_bobj_crv_ParamsMinAndMaxTooltip = "Geef een %1-waarde tussen %2 en %3 op.";
var L_bobj_crv_ParamsStringMinOrMaxTooltip = "De %1 lengte voor dit veld is %2.";
var L_bobj_crv_ParamsStringMinAndMaxTooltip = "De waarde moet tussen %1 en %2 tekens lang zijn.";
var L_bobj_crv_ParamsYearToken = "jjjj";
var L_bobj_crv_ParamsMonthToken = "mm";
var L_bobj_crv_ParamsDayToken = "dd";
var L_bobj_crv_ParamsReadOnly = "Deze parameter is van het type \"Alleen-lezen\".";
var L_bobj_crv_ParamsNoValue = "Geen waarde";
var L_bobj_crv_ParamsDuplicateValue = "Dubbele waarden zijn niet toegestaan.";
var L_bobj_crv_Error = "FOUT";
var L_bobj_crv_OK = "OK";
var L_bobj_crv_Cancel = "Annuleren";
var L_bobj_crv_showDetails = "Details weergeven";
var L_bobj_crv_hideDetails = "Details verbergen";
var L_bobj_crv_RequestError = "Uw verzoek kan niet worden verwerkt";
var L_bobj_crv_ServletMissing = "De viewer kan geen verbinding maken met de CrystalReportViewerServlet die asynchrone verzoeken afhandelt.\nControleer of de Servlet en Servlet-Mapping correct zijn gedeclareerd in het bestand web.xml van de toepassing.";

var L_bobj_crv_Tree_Drilldown_Node = "Knooppunt %1 analyseren";

var L_bobj_crv_ReportProcessingMessage = "Wacht totdat het document is verwerkt.";

var L_bobj_crv_SundayShort = "Z";
var L_bobj_crv_MondayShort = "M";
var L_bobj_crv_TuesdayShort = "D";
var L_bobj_crv_WednesdayShort = "W";
var L_bobj_crv_ThursdayShort = "D";
var L_bobj_crv_FridayShort = "V";
var L_bobj_crv_SaturdayShort = "Z";

var L_bobj_crv_Minimum = "minimum";
var L_bobj_crv_Maximum = "maximum";

var L_bobj_crv_Date = "Datum";
var L_bobj_crv_Time = "Tijd";
var L_bobj_crv_DateTime = "Datum en tijd";
var L_bobj_crv_Boolean = "Boolean-waarde";
var L_bobj_crv_Number = "Getal";

var L_bobj_crv_InteractiveParam_NoAjax = "De webbrowser die u gebruikt, is niet geconfigureerd om het parametervenster weer te geven.";
