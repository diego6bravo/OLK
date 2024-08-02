<%
'****************************************************************************************
'**  Copyright Notice
'**
'**  Web Wiz Guide - Web Wiz Rich Text Editor
'**  http://www.richtexteditor.org
'**
'**  Copyright 2002-2005 Bruce Corkhill All Rights Reserved.
'**
'**  This program is free software; you can modify (at your own risk) any part of it
'**  under the terms of the License that accompanies this software and use it both
'**  privately and commercially.
'**
'**  All copyright notices must remain in tacked in the scripts and the
'**  outputted HTML.
'**
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even
'**  if it is modified or reverse engineered in whole or in part without express
'**  permission from the author.
'**
'**  You may not pass the whole or any part of this application off as your own work.
'**
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**
'**  You should have received a copy of the License along with this program;
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**
'**
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**
'**  Support questions are NOT answered by email ever!
'**
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.info
'**
'**  or at: -
'**
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'**
'****************************************************************************************

Const strTxtTextFormat = "Formato de Texto"
Const strTxtMode = "Modo"
Const strTxtPrompt = "Aviso"
Const strTxtBasic = "Básico"
Const strTxtAddEmailLink = "Añadir vinculo hacia Email"
Const strTxtList = "Lista"
Const strTxtCentre = "Centro"

Const strTxtEnterBoldText = "Escriba el texto que desea en Negrita"
Const strTxtEnterItalicText = "Escriba el texto que desea en Itálica"
Const strTxtEnterUnderlineText = "Escriba el texto que desea Subrayado"
Const strTxtEnterCentredText = "Escriba el texto que desea Centrado"

Const strTxtEnterHyperlinkText = "Escriba el texto que aparece en pantalla para el Hipervínculo"
Const strTxtEnterHeperlinkURL = "Escriba la dirección URL para el Hipervínculo"
Const strTxtEnterEmailText = "Escriba el texto que aparece en pantalla para la dirección Email"
Const strTxtEnterEmailMailto = "Escriba la dirección Email a vincular"
Const strTxtEnterImageURL = "Escriba la dirección Web de la imagen"
Const strTxtEnterTypeOfList = "Tipo de lista"
Const strTxtEnterEnter = "Escriba"
Const strTxtEnterNumOrBlankList = "para enumerado o deje en blanco para viñetas"
Const strTxtEnterListError = "ERROR! Favor Escribir"
Const strEnterLeaveBlankForEndList = "Dejar en blanco para terminar la lista"
Const strTxtErrorInsertingObject = "Error al insertar objeto en esta ubicación"


Const strTxtFontStyle = "Formato"
Const strTxtFontTypes = "Fuente"
Const strTxtFontSizes ="Tamaño"
Const strTxtEmoticons = "Emociones"
Const strTxtFontSize = "Tamaño de Fuente"


Const strTxtFontColours ="Colores de Fuente"
Const strTxtBlack = "Negro"
Const strTxtWhite = "Blanco"
Const strTxtBlue = "Azul"
Const strTxtRed = "Rojo"
Const strTxtGreen = "Verde"
Const strTxtYellow = "Amarillo"
Const strTxtOrange = "Naranja"
Const strTxtBrown = "Chocolate"
Const strTxtMagenta = "Magenta"
Const strTxtCyan = "Celeste"
Const strTxtLimeGreen = "Verde Lima"



Const strTxtCut = "Cortar"
Const strTxtCopy = "Copiar"
Const strTxtPaste = "Pegar"
Const strTxtBold = "Negrita"
Const strTxtItalic = "Itálica"
Const strTxtUnderline = "Subrayado"
Const strTxtLeftJustify = "Alineación Izquierda"
Const strTxtCentrejustify = "Alineación Centrada"
Const strTxtRightJustify = "Alineación Derecha"
Const strTxtJustify = "Justificar"
Const strTxtUnorderedList = "Viñetas"
Const strTxtOutdent = "Disminuir Sangría"
Const strTxtIndent = "Aumentar Sangría"
Const strTxtAddHyperlink = "Insertar Hipervínculo"
Const strTxtAddImage = "Insertar Imagen"
Const strTxtJavaScriptEnabled = "JavaScript debe estar habilitado en su navegador para poder utilizar Rich Text Editor!"
Const strTxtFontColour = "Color"
Const strTxtstrTxtOrderedList = "Numeración"
Const strTxtTextColour = "Color de Texto"
Const strTxtBackgroundColour = "Color de Fondo"
Const strTxtUndo = "Deshacer"
Const strTxtRedo = "Rehacer"
Const strTxtstrSpellCheck = "Corrector Ortográfico"
Const strTxtToggleHTMLView = "Habilitar/Deshabilitar Vista HTML"
Const strTxtAboutRichTextEditor = "Acerca de Rich Text Editor"
Const strTxtInsertTable = "Insertar Tabla"
Const strTxtSpecialCharacters = "Caracteres Especiales"
Const strTxtPrint = "Imprimir"
Const strTxtImage = "Imagen"
Const strTxtStrikeThrough = "Tachado"
Const strTxtSubscript = "Subíndice"
Const strTxtSuperscript = "Superíndice"
Const strTxtHorizontalRule = "Línea Horizontal"


Const strTxtIeSpellNotDetected = "Necesita instalar el corrector ortográfico \'ieSpell\' para utilizar esta función.  Haga clic en Aceptar para descargar \'ieSpell\'."
Const strTxtSpellBoundNotDetected = " Necesita instalar el corrector ortográfico \'SpellBound 0.7.0+\' para utilizar esta función.  Haga clic en Aceptar para descargar \'SpellBound 0.7.0+\'."

Const strTxtOK = "Aceptar"
Const strTxtCancel = "Cancelar"


Const strTxtImageUpload = "Subir Imagen"
Const strTxtFileUpload = "Subir Archivo"
Const strTxtUpload = "Subir"
Const strTxtPath = "Ruta"
Const strTxtFileName = "Nombre del Archivo"
Const strTxtFileURL = "Archivo URL"

Const strTxtParentDirectory = "Directorio Padre"

Const strTxtImagesMustBeOfTheType = "Imágenes deben ser de tipo"
Const strTxtAndHaveMaximumFileSizeOf = "y tener un tamaño máximo de archivo de"
Const strTxtImageOfTheWrongFileType = "La imagen subida es de tipo incorrecto"
Const strTxtImageFileSizeToLarge = "El tamaño del archivo de imagen es muy grande en"
Const strTxtMaximumFileSizeMustBe = "El tamaño máximo de archivo debe ser"
Const strTxtErrorUploadingImage = "Error subiendo imagen!!"
Const strTxtNoImageToUpload = "Favor utilizar el botón \'Browse...\' para seleccionar la imagen a subir."

Const strTxtFile = "Archivo"
Const strTxtFilesMustBeOfTheType = "Archivos deben ser de tipo"
Const strTxtFileOfTheWrongFileType = "El archivo subido es del tipo incorrecto"
Const strTxtFileSizeToLarge = " El tamaño del archivo es muy grande en"
Const strTxtErrorUploadingFile = "Error subiendo archivo!!"
Const strTxtNoFileToUpload = "Favor utilizar el botón \'Browse...\' para seleccionar el archivo a subir."


Const strTxtPleaseWaitWhileFileIsUploaded = "Favor esperar mientras el archivo se transfiere hacia el servidor."
Const strTxtPleaseWaitWhileImageIsUploaded = " Favor esperar mientras la imagen se transfiere hacia el servidor."


Const strTxtCloseWindow = "Cerrar Ventana"


Const strTxtPreview = "Vista Preliminar"
Const strTxtThereIsNothingToPreview = "No hay nada que Ver Preliminarmente."

Const strResetFormConfirm = "Esta seguro que desea reiniciar el formulario?"
Const strResetWarningFormConfirm = "ADVERTENCIA: Toda la información del formulario se perderá!!"
Const strResetWarningEditorConfirm = "ADVERTENCIA: Toda la información del editor se perderá!!"


Const strTxtSubmitForm = "Aceptar Formulario"
Const strTxtResetForm = "Reiniciar Formulario"

Const strTxtDisplayMessage = "Desplegar Mensaje"
Const strTxtThereIsNothingToShow = "No hay mensaje que desplegar"


Const strTxtTableProperties = "Propiedades de la Tabla"

Const strTxtImageProperties = "Propiedades de la Imagen"

Const strTxtImageURL = "Imagen&nbsp;URL"
Const strTxtAlternativeText = "Texto Alternativo"
Const strTxtLayout = "Layout"
Const strTxtAlignment = "Alineación"
Const strTxtBorder = "Borde"
Const strTxtSpacing = "Espacio"
Const strTxtHorizontal = "Horizontal"
Const strTxtVertical = "Vertical"

Const strTxtRows = "Filas"
Const strTxtColumns = "Columnas"
Const strTxtWidth = "Ancho"
Const strTxtpixels = "píxeles"
Const strTxtCellPad = "Colchón de Celda"
Const strTxtCellSpace = "Espacio de Celda"

Const strTxtHeight = "Alto"


Const strTxtSelectTextToTurnIntoHyperlink = "Favor seleccionar texto a convertir a Hipervínculo"

Const strTxtYourBrowserSettingsDoNotPermit = "La configuración de su navegador no le permite al editor invocar"
Const strTxtPleaseUseKeybordsShortcut = "operaciones. \nFavor utilizar el comando con el teclado."
Const strTxtWindowsUsers = "Usuarios de Windows:"
Const strTxtMacUsers = "Usuarios de Mac:"


Const strTxtHyperlinkProperties = "Propiedades del Hipervínculo"
Const strTxtNoPreviewAvailableForLink = "No hay vista premilitar disponible"
Const strTxtAddress = "Dirección"
Const strTxtLinkType = "Tipo de hipervínculo"
Const strTxtTitle = "Titulo"
Const strTxtWindow = "Ventana"
Const strTxtEmail = "Email"
Const strTxtSubject = "Asunto"
Const strTxtPleaseWaitWhilePreviewLoaded = "Favor esperar mientras la vista preliminar carga..."
Const strTxtErrorLoadingPreview = "Error al cargar vista preliminar.\nFavor revisar que la ruta y nombre sean correctos."


Const strTxAttachFileProperties = "Propiedades del archivo adjunto"

Const strTxtNewBlankDoc = "Nuevo Documento en Blanco"
Const strTxtOpen = "Abrir"
Const strTxtSave = "Guardar"

Const strTxtFileAlreadyExistsRenamedAs = "Un archivo con el mismo nombre ya existe o hay un problema con el nombre de archivo escrito.\nEl Archivo ha sido guardado como"
Const strTxtTheFile = "El Archivo"
Const strTxtHasBeenSaved = "ha sido guardado"


Const strTxtPasteFromWord = "Pegar desde Word"
Const strTxtPasteFromWordDialog = "Esta acción limpiara los documentos pegados desde Word. Favor pegar dentro del recuadro siguiente utilizando el atajo de teclado (Usuarios de Windows: ctrl. + ‘v’, Usuarios MAC: Apple + ’v’) y presione ‘Aceptar’."

Const strUpload = "Subir"
Const strAlignDef = "Predeterminado"
Const strAlignLeft = "Izquierda"
Const strAlignRight = "Derecha"
Const strAlignCenter = "Centro"
Const strAlignAbsTop = "Arriba del texto"
Const strAlignAbsMdl = "Centro Absoluto"
Const strAlignAbsBottom = "Fondo Absoluto"
Const strAlignBottom = "Abajo"
Const strAlignMiddle = "Centro"
Const strAlignTop = "Arriba"

%>