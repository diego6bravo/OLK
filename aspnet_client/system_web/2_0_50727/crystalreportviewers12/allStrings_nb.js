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
// <script>
/*
=============================================================
WebIntelligence(r) Report Panel
Copyright(c) 2001-2003 Business Objects S.A.
All rights reserved

Use and support of this software is governed by the terms
and conditions of the software license agreement and support
policy of Business Objects S.A. and/or its subsidiaries. 
The Business Objects products and technology are protected
by the US patent number 5,555,403 and 6,247,008

File: labels.js


=============================================================
*/

_default="Standard"
_black="Svart"
_brown="Brun"
_oliveGreen="Olivengrønn"
_darkGreen="Mørkegrønn"
_darkTeal="Mørk blågrønn"
_navyBlue="Marineblå"
_indigo="Indigo"
_darkGray="Mørk grå"
_darkRed="Mørkerød"
_orange="Oransje"
_darkYellow="Mørkegul"
_green="Grønn"
_teal="Blågrønn"
_blue="Blå"
_blueGray="Blågrå"
_mediumGray="Mellomgrå"
_red="Rød"
_lightOrange="Lys oransje"
_lime="Sitrusgrønn"
_seaGreen="Sjøgrønn"
_aqua="Akvamarin"
_lightBlue="Lyseblå"
_violet="Fiolett"
_gray="Grå"
_magenta="Magenta"
_gold="Gull"
_yellow="Gul"
_brightGreen="Skarp grønn"
_cyan="Cyan"
_skyBlue="Himmelblå"
_plum="Plomme"
_lightGray="Lys grå"
_pink="Rosa"
_tan="Lysebrun"
_lightYellow="Lys gul"
_lightGreen="Lysegrønn"
_lightTurquoise="Lys turkis"
_paleBlue="Blekblå"
_lavender="Lavendel"
_white="Hvit"
_lastUsed="Sist brukte:"
_moreColors="Flere farger..."

_month=new Array

_month[0]="JANUAR"
_month[1]="FEBRUAR"
_month[2]="MARS"
_month[3]="APRIL"
_month[4]="MAI"
_month[5]="JUNI"
_month[6]="JULI"
_month[7]="AUGUST"
_month[8]="SEPTEMBER"
_month[9]="OKTOBER"
_month[10]="NOVEMBER"
_month[11]="DESEMBER"

_day=new Array
_day[0]="S"
_day[1]="M"
_day[2]="T"
_day[3]="O"
_day[4]="T"
_day[5]="F"
_day[6]="L"

_today="I dag"

_AM="AM"
_PM="PM"

_closeDialog="Lukk vindu"

_lstMoveUpLab="Flytt opp"
_lstMoveDownLab="Flytt ned"
_lstMoveLeftLab="Flytt  mot venstre" 
_lstMoveRightLab="Flytt mot høyre"
_lstNewNodeLab="Legg til nøstet filter"
_lstAndLabel="AND"
_lstOrLabel="OR"
_lstSelectedLabel="Valgt"
_lstQuickFilterLab="Legg til hurtigfilter"

_openMenuPart1="Klikk her for å få tilgang"
_openMenuPart2=" alternativer"
_openCalendarLab="Åpne kalender"

_scroll_first_tab="Rull til første kategori"
_scroll_previous_tab="Rull til forrige kategori"
_scroll_next_tab="Rull til neste kategori"
_scroll_last_tab="Rull til siste kategori"

_expandedLab="Utvidet"
_collapsedLab="Sammenslått"
_selectedLab="Valgt"

_expandNode="Utvid knutepunkt %1"
_collapseNode="Slå sammen knutepunkt %1"

_checkedPromptLab="Angitt"
_nocheckedPromptLab="Ikke angitt"
_selectionPromptLab="verdier lik"
_noselectionPromptLab="ingen verdier"

_lovTextFieldLab="Skriv inn verdier her"
_lovCalendarLab="Skriv inn data her"
_lovPrevChunkLab="Gå til forrige skive"
_lovNextChunkLab="Gå til neste skive"
_lovComboChunkLab="Skive"
_lovRefreshLab="Oppdater"
_lovSearchFieldLab="Skriv inn tekst for søk her"
_lovSearchLab="Søk"
_lovNormalLab="Normal"
_lovMatchCase="Skill mellom store og små bokstaver"
_lovRefreshValuesLab="Oppdater verdier"

_calendarNextMonthLab="Gå til neste måned"
_calendarPrevMonthLab="Gå til forrige måned"
_calendarNextYearLab="Gå til neste år"
_calendarPrevYearLab="Gå til forrige år"
_calendarSelectionLab="Valgt dag "

_menuCheckLab="Merket"
_menuDisableLab="Deaktivert"

_RGBTxtBegin= "RGB ("
_RGBTxtEnd= ")"

_helpLab="Hjelp"

_waitTitleLab="Vent litt"
_cancelButtonLab="Avbryt"

_modifiers= new Array
_modifiers[0]="Ctrl+"
_modifiers[1]="Shift+"
_modifiers[2]="Alt+"

_bordersMoreColorsLabel="Flere kantlinjer..."
_bordersTooltip=new Array
_bordersTooltip[0]="Ingen kantlinje"
_bordersTooltip[1]="Venstre kantlinje"
_bordersTooltip[2]="Høyre kantlinje"
_bordersTooltip[3]="Nedre kantlinje"
_bordersTooltip[4]="Middels nedre kantlinje"
_bordersTooltip[5]="Bred nedre kantlinje"
_bordersTooltip[6]="Øvre og nedre kantlinje"
_bordersTooltip[7]="Øvre og middels nedre kantlinje"
_bordersTooltip[8]="Øvre og bred nedre kantlinje"
_bordersTooltip[9]="Alle kantlinjer"
_bordersTooltip[10]="Alle kantlinjer middels"
_bordersTooltip[11]="Alle kantlinjer brede"/* Copyright (c) Business Objects 2006. All rights reserved. */

// LOCALIZATION STRING

// Strings for calendar.js and calendar_param.js
var L_Today     = "I dag";
var L_January   = "Januar";
var L_February  = "Februar";
var L_March     = "Mars";
var L_April     = "April";
var L_May       = "Mai";
var L_June      = "Juni";
var L_July      = "Juli";
var L_August    = "August";
var L_September = "September";
var L_October   = "Oktober";
var L_November  = "November";
var L_December  = "Desember";
var L_Su        = "S\u00F8";
var L_Mo        = "Ma";
var L_Tu        = "Ti";
var L_We        = "On";
var L_Th        = "To";
var L_Fr        = "Fr";
var L_Sa        = "L\u00F8";

// Strings for prompts.js and prompts_param.js
var L_YYYY          = "\u00E5\u00E5\u00E5\u00E5";
var L_MM            = "mm";
var L_DD            = "dd";
var L_BadNumber     = "Denne parameteren er av typen Nummer og kan bare inneholde et negativt tegnsymbol, sifre (0-9), grupperingssymboler for sifre eller et desimalsymbol. Korriger den angitte parameterverdien.";
var L_BadCurrency   = "Denne parameteren er av typen Valuta og kan bare inneholde et symbol for negativt tegn, sifre (0-9), grupperingssymboler for sifre eller et desimalsymbol. Korriger den angitte parameterverdien.";
var L_BadDate       = "Denne parameteren er av typen Dato og m\u00E5 ha formatet %1, der \u00E5\u00E5\u00E5\u00E5 er \u00E5rstallet med fire sifre, mm er m\u00E5neden (f.eks. januar = 1) og dd er dagen i m\u00E5neden.";
var L_BadDateTime   = "Denne parameteren er av typen Dato/klokkeslett og m\u00E5 ha formatet %1 tt:mm:ss, der \u00E5\u00E5\u00E5\u00E5 er \u00E5rstallet med fire sifre, mm er m\u00E5neden (f.eks. januar = 1), dd er dagen i m\u00E5neden, tt er timer i en klokke med 24-timers format, mm er minutter og ss er sekunder.";
var L_BadTime       = "Denne parameteren er av typen Klokkeslett og m\u00E5 ha formatet tt:mm:ss, der tt er timer i en klokke med 24-timers format, mm er minutter og ss er sekunder.";
var L_NoValue       = "Ingen verdi";
var L_BadValue      = "Hvis du vil angi Ingen verdi, m\u00E5 du sette b\u00E5de Fra og Til til Ingen verdi.";
var L_BadBound      = "Du kan ikke angi Ingen nedre grense sammen med Ingen \u00F8vre grense.";
var L_NoValueAlready = "Denne parameteren er allerede satt til Ingen verdi. Fjern Ingen verdi f\u00F8r du legger til andre verdier.";
var L_RangeError    = "Starten p\u00E5 omr\u00E5det kan ikke v\u00E6re st\u00F8rre enn slutten p\u00E5 omr\u00E5det.";
var L_NoDateEntered = "Du m\u00E5 angi en dato.";

// Strings for filter dialog
var L_closeDialog="Lukk vindu";

var L_SetFilter = "Angi filter";
var L_OK        = "OK";
var L_Cancel    = "Avbryt";

 /* Crystal Decisions Confidential Proprietary Information */
