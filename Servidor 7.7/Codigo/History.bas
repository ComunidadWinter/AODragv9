Attribute VB_Name = "History"
Option Explicit
'History Log by Morgolock

'13-2-2003
'---------
'1) Modifiqué todas las llamadas a las funciones Mid, Left y
'Right por Mid$, Left$ y Right$ para que devuelvan strings
'en vez de variants. Se deberia ganar considerable velocidad.
'2) Quite el comando /GRABAR ya que generaba problemas con
'las mascotas y no era demasiado útil ya que los usuarios
'consiguen el mismo efecto saliendo y volviendo a entrar
'en el juego.
'3) Agregué el MOTD, el servidor levanta el mensaje del archivo
'motd.ini del directorio dat del servidor, les envia el motd
'a los usuarios cuando entran al juego.

'12-2-2003
'---------
'1) Limité a tres la máxima cantidad de mascotas
'2) A los newbies se les caen los objetos no newbies


