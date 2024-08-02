/* Copyright (c) Business Objects 2006. All rights reserved. */

var L_bobj_crv_MainReport = "Hovedrapport";
// Viewer Toolbar tooltips
var L_bobj_crv_FirstPage = "G\u00E5 til f\u00F8rste side";
var L_bobj_crv_PrevPage = "G\u00E5 til forrige side";
var L_bobj_crv_NextPage = "G\u00E5 til neste side";
var L_bobj_crv_LastPage = "G\u00E5 til siste side";
var L_bobj_crv_ParamPanel = "Parameterpanel";
var L_bobj_crv_Parameters = "Parametere";
var L_bobj_crv_GroupTree = "Gruppetre";
var L_bobj_crv_DrillUp = "Analyser opp";
var L_bobj_crv_Refresh = "Oppdater rapport";
var L_bobj_crv_Zoom = "Zoom";
var L_bobj_crv_SelectPage = "G\u00E5 til side";
var L_bobj_crv_SearchText = "S\u00F8k etter tekst";
var L_bobj_crv_Export = "Eksporter denne rapporten";
var L_bobj_crv_Print = "Skriv ut denne rapporten";
var L_bobj_crv_TabList = "Kategoriliste";
var L_bobj_crv_Close = "Lukk";
var L_bobj_crv_Logo=  "Business Objects Logo"

var L_bobj_crv_Show = "Vis";
var L_bobj_crv_Hide = "Skjul";

var L_bobj_crv_submitBtnLbl = "Eksporter";
var L_bobj_crv_ActiveXPrintDialogTitle = "Skriv ut";
var L_bobj_crv_PDFPrintDialogTitle = "Skriv ut til PDF";
var L_bobj_crv_PrintRangeLbl = "Sideomr\u00E5de:";
var L_bobj_crv_PrintAllLbl = "Alle sider";
var L_bobj_crv_PrintPagesLbl = "Velg sider";
var L_bobj_crv_PrintFromLbl = "Fra:";
var L_bobj_crv_PrintToLbl = "Til:";
var L_bobj_crv_PrintInfoTitle = "Skriv ut til PDF:";
var L_bobj_crv_PrintInfo1 = 'Visningen m\u00E5 eksporteres til PDF for \u00E5 skrives ut. Velg utskriftsalternativet i PDF-leserapplikasjonen n\u00E5r dokumentet er \u00E5pnet.';
var L_bobj_crv_PrintInfo2 = 'Merk: Du m\u00E5 ha en installert PDF-leser for \u00E5 skrive ut (f.eks. Adobe Reader).';
var L_bobj_crv_PrintPageRangeError = "Angi et gyldig sideomr\u00E5de.";

var L_bobj_crv_ExportBtnLbl = "Eksporter";
var L_bobj_crv_ExportDialogTitle = "Eksporter";
var L_bobj_crv_ExportFormatLbl = "Filformat:";
var L_bobj_crv_ExportInfoTitle = "For eksport:";

var L_bobj_crv_ParamsApply = "Bruk";
var L_bobj_crv_ParamsAdvDlg = "Rediger parameterverdi";
var L_bobj_crv_ParamsDelete = "Slett";
var L_bobj_crv_ParamsDeleteTooltip = "Slett parameterverdi";
var L_bobj_crv_ParamsAddValue = "Klikk for \u00E5 legge til...";
var L_bobj_crv_ParamsApplyTip = "Bruk parameterverdier";
var L_bobj_crv_ParamsDlgTitle = "Angi verdier";
var L_bobj_crv_ParamsCalBtn = "Vis en kalender";
var L_bobj_crv_Undo= "Angre";
var L_bobj_crv_UndoTip = "Tilbakestill alle endringer";
var L_bobj_crv_ParamsDirtyTip = "Klikk p\u00E5 knappen Bruk for \u00E5 aktivere endringer";
var L_bobj_crv_ParamsDataTip = "Dette er en datahentingsparameter";
var L_bobj_crv_ParamsMaxNumDefaultValues = "Klikk her for flere elementer...";

var L_bobj_crv_ParamsInvalidTitle = "Parameterverdien er ikke gyldig";
var L_bobj_crv_ParamsTooLong = "Parameterverdien kan ikke v\u00E6re lengre enn %1 tegn";
var L_bobj_crv_ParamsTooShort = "Parameterverdien m\u00E5 best\u00E5 av minst %1 tegn";
var L_bobj_crv_ParamsBadNumber = "Denne parameteren er av typen \Tall\ og kan bare inneholde et negativt tegnsymbol, sifre (\0-9\), siffergrupperingssymboler eller et desimalsymbol.";
var L_bobj_crv_ParamsBadCurrency = "Denne parameteren er av typen \Valuta\ og kan bare inneholde et negativt tegnsymbol, sifre (\0-9\), siffergrupperingssymboler eller et desimalsymbol.";
var L_bobj_crv_ParamsBadDate = "Denne parameteren er av typen \Dato\, og det riktige formatet er \%1\ der \\u00E5\u00E5\u00E5\u00E5\ er \u00E5rstallet skrevet med fire tall, \mm\ er m\u00E5neden (f.eks. januar = 1) og \dd\ er dagen i m\u00E5neden.";
var L_bobj_crv_ParamsBadTime = "Denne parameteren er av typen \Klokkeslett\, og det riktige formatet er \tt:mm:ss\ der \tt\ er timer i 24-timers klokkeformat, \mm\ er minutter og \ss\ er sekunder.";
var L_bobj_crv_ParamsBadDateTime = "Denne parameteren er av typen \Dato/klokkeslett\, og det riktige formatet er \%1 tt:mm:ss\. \\u00E5\u00E5\u00E5\u00E5\ er \u00E5rstallet skrevet med fire tall, \mm\ er m\u00E5neden (f.eks. januar = 1), \dd\ er dagen i m\u00E5neden, \tt\ er timer i 24-timers klokkeformat, \mm\ er minutter og \ss\ er sekunder.";
var L_bobj_crv_ParamsMinTooltip = "Angi en verdi for %1 som er st\u00F8rre enn eller lik %2.";
var L_bobj_crv_ParamsMaxTooltip = "Angi en verdi for %1 som er mindre enn eller lik %2.";
var L_bobj_crv_ParamsMinAndMaxTooltip = "Angi en verdi for %1 som er mellom %2 og %3.";
var L_bobj_crv_ParamsStringMinOrMaxTooltip = "Lengden p\u00E5 %1 for dette feltet er %2.";
var L_bobj_crv_ParamsStringMinAndMaxTooltip = "Verdien m\u00E5 best\u00E5 av mellom %1 og %2 tegn.";
var L_bobj_crv_ParamsYearToken = "\u00E5\u00E5\u00E5\u00E5";
var L_bobj_crv_ParamsMonthToken = "mm";
var L_bobj_crv_ParamsDayToken = "dd";
var L_bobj_crv_ParamsReadOnly = "Denne parameteren er av typen \Skrivebeskyttet\.";
var L_bobj_crv_ParamsNoValue = "Ingen verdi";
var L_bobj_crv_ParamsDuplicateValue = "Dupliserte verdier er ikke tillatt.";
var L_bobj_crv_Error = "Feil";
var L_bobj_crv_OK = "OK";
var L_bobj_crv_Cancel = "Avbryt";
var L_bobj_crv_showDetails = "Vis detaljer";
var L_bobj_crv_hideDetails = "Skjul detaljer";
var L_bobj_crv_RequestError = "Kan ikke behandle foresp\u00F8rselen";
var L_bobj_crv_ServletMissing = "Viewer-en kan ikke koble til CrystalReportViewer-servleten som behandler asynkrone foresp\u00F8rsler.\nKontroller at servleten og servlettilordningen er riktig angitt i web.xml-filen til applikasjonen.";

var L_bobj_crv_Tree_Drilldown_Node = "Analyser ned i knutepunkt %1";

var L_bobj_crv_ReportProcessingMessage = "Vent mens dokumentet behandles.";

var L_bobj_crv_SundayShort = "S";
var L_bobj_crv_MondayShort = "M";
var L_bobj_crv_TuesdayShort = "T";
var L_bobj_crv_WednesdayShort = "O";
var L_bobj_crv_ThursdayShort = "T";
var L_bobj_crv_FridayShort = "F";
var L_bobj_crv_SaturdayShort = "L";

var L_bobj_crv_Minimum = "minimum";
var L_bobj_crv_Maximum = "maksimum";

var L_bobj_crv_Date = "Dato";
var L_bobj_crv_Time = "Klokkeslett";
var L_bobj_crv_DateTime = "Dato/klokkeslett";
var L_bobj_crv_Boolean = "Boolsk";
var L_bobj_crv_Number = "Tall";

var L_bobj_crv_InteractiveParam_NoAjax = "Nettleseren du bruker, er ikke konfigurert for visning av parameterpanelet.";
