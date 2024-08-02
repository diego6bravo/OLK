<!--

/*
Configure menu styles below
NOTE: To edit the link colors, go to the STYLE tags and edit the ssm2Items colors
*/
YOffset=150; // no quotes!!
XOffset=0;
staticYOffset=30; // no quotes!!
slideSpeed=20 // no quotes!!
waitTime=100; // no quotes!! this sets the time the menu stays out for after the mouse goes off it.
menuBGColor="black";
menuIsStatic="yes"; //this sets whether menu should stay static on the screen
menuWidth=160; // Must be a multiple of 10! no quotes!!
menuCols=2;
hdrFontFamily="verdana";
hdrFontSize="2";
hdrFontColor="white";
hdrBGColor="#003399";
hdrAlign="left";
hdrVAlign="center";
hdrHeight="15";
linkFontFamily="Verdana";
linkFontSize="2";
linkBGColor="#FEFFEC";
linkOverBGColor="#DDE6FB";
linkTarget="_top";
linkAlign="Left";
barBGColor="#1A4EC8";
barFontFamily="Verdana";
barFontSize="2";
barFontColor="white";
barVAlign="center";
barWidth=20; // no quotes!!
barText="MENU"; // <IMG> tag supported. Put exact html for an image to show.

///////////////////////////

// ssmItems[...]=[name, link, target, colspan, endrow?] - leave 'link' and 'target' blank to make a header
ssmItems[0]=["Menu"] //create header
ssmItems[1]=["Alertas", "adminalert.asp", "_self"]
ssmItems[2]=["Generales", "adminnew.asp",""]
ssmItems[3]=["Contraseña Admin", "adminuser.asp", "_self"]
ssmItems[4]=["Contraseña Clientes", "adminuserc.asp", "_self"]
ssmItems[5]=["Contraseña Ventas", "adminuserv.asp", "_self"]
ssmItems[6]=["Inventario/Disponible", "admininv.asp", "_self"]
ssmItems[7]=["Detalles de articulos", "admininvopt.asp", "_self"]
ssmItems[8]=["Lista de precios(C)", "adminplc.asp", "_self"]
ssmItems[9]=["Lista de precios(V)", "adminplv.asp", "_self"]
ssmItems[10]=["Portal", "adminportal.asp", "_self"]
ssmItems[11]=["Noticias", "adminnews.asp", "_self"]
buildMenu();

//-->