# O365---Gestiones

Script hecho por Pau Juan Nieto
Fecha 04/10/2018
Versión 1.3

Changelog:
************************************************
08/07/2019 - 1.3.1
- Se desactiva por defecto la opción de automapping a la hora de activar permisos.
02/01/2019 - 1.3
- Se soluciona bug por el cual no se agregaban permisos del tipo "send on behalf of": "Add-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount"
15/11/2018 - 1.2
- Se añade la funcionalidad de desactivar el automapping de las cuentas de correo
06/10/2018 - 1.1
- Se extrae la función "RemovePSSession" del bucle do principal para evitar que se vuelvan a solicitar credenciales
- Se extrae la función "Import-PSSession" del bucle do principal para evitar que se vuelva a importar la configuración de MSO
04/10/2018 - 1.0
- Primera Versión
************************************************ 
