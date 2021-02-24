<# 
Script hecho por Pau Juan Nieto
Fecha 04/10/2018
Versión 1.4

Changelog:
************************************************
22/02/2021 - 1.4
-Se añade la funcionalidad de comprobar las licencias asignadas a una cuenta.
-Se retira el desactivado automático del automapping
-Se añade la opción de agregar el permiso "ReadPermission"

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
#>

# Guardamos las credenciales de inicio de sesión de manera encriptada (Usuario Administrador de Office365)

$creds = Get-Credential

# Iniciamos la conexión contra Azure y la nube

Connect-AzureAD -Credential $creds
Connect-MsolService -Credential $creds

# Establecemos una sesión remota de PS contra O365

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $creds -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

do {
    
    # Menu

    Clear-Host
    $validall = 'n'

    Write-Output ''
    Write-Output 'Herramientas de Office 365'
    Write-Output '*******************************************'
    Write-Output '1. Revisar usuarios en la nube'
    Write-Output '2. Ver permisos de correo sobre una cuenta'
    Write-Output '3. Modificar permisos de correo sobre una cuenta'
    Write-Output '4. Ver licencia(s) asignadas sobre una cuenta'
    Write-Output '5. Desactivar AutoMapping'
    Write-Output '*******************************************'

    $accion = Read-Host 'Opcion'

    switch ($accion) {

        1 {
            # Leemos el String y se realiza la busqueda

            $recurrent = "s"

            while ($recurrent -eq "s") {

                Clear-Host

                $azure_string = Read-Host -Prompt 'Introduce string'
                $result = Get-MsolUser -SearchString "$azure_string" | Select-Object UserPrincipalName, DisplayName, isLicensed, ProxyAddresses | Sort-Object UserPrincipalName | Format-Table

                # Error por si la variable $result tiene un valor nulo

                if ($result) { $result }else { Write-Host -ForegroundColor Red "No se encontraron resultados para $azure_string" }

                # Submenu

                $recurrent = Read-Host -Prompt '¿Realizar otra busqueda? (s/n)'
                if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }

            }
        }
        2 {
            # Iniciamos la consulta sobre la cuenta de correo

            $recurrent = "s"

            while ($recurrent -eq "s") {

                Clear-Host
    
                $account = Read-Host -Prompt 'Introducir cuenta de correo'
                $result = get-mailboxpermission $account | Select-Object User, AccessRights | Sort-Object User | Format-Table

                # Error por si la variable $result tiene un valor nulo

                if ($result) { $result }else { Write-Host -ForegroundColor Red "No se encontraron resultados para $account" }

                # Submenu

                $recurrent = Read-Host -Prompt '¿Realizar otra consulta? (s/n)'
                if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }

            }

        }
        3 {
            # Menu
            
            Clear-Host

            Write-Output ''
            Write-Output 'Modificar permisos de correo en Office 365'
            Write-Output '*******************************************'
            Write-Output '1. Agregar permisos'
            Write-Output '2. Quitar permisos'
            Write-Output '*******************************************'

            $operation = Read-Host 'Opción'

            switch ($operation) {

                1 {
                    $recurrent = "s"
                    while ($recurrent -eq "s") {
                    
                        # Menu
                    
                        Clear-Host
    
                        Write-Output ''
                        Write-Output 'Agregar Permisos'
                        Write-Output '*******************************************'
                        Write-Output '1. Agregar permisos a una cuenta (full access)'
                        Write-Output '2. Agregar permisos a una cuenta (read access)'
                        Write-Output '3. Usar archivo TXT con listado de usuarios (full access)'
                        Write-Output '*******************************************'

                        $opsadd = Read-Host 'Opcion'

                        switch ($opsadd) {

                            # Se leen los datos de la operación y se agregan los permisos sobre las cuentas especificadas

                            1 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $useraccount = Read-Host 'Introduce la cuenta que obtendra los permisos'
                                Add-MailboxPermission -Identity $account -User $useraccount -AccessRights FullAccess -InheritanceType All
                                Add-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                            }

                            2 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $useraccount = Read-Host 'Introduce la cuenta que obtendra los permisos'
                                Add-MailboxPermission -Identity $account -User $useraccount -AccessRights ReadPermission -InheritanceType All
                                Add-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                            }

                            3 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $path = Read-Host 'Introduce la ruta del fichero TXT con los usuarios'
                                foreach ($useraccount in [System.IO.File]::ReadLines("$path")) {
                                    Write-Output "Procesando $useraccount"
                                    Add-MailboxPermission -Identity $account -User $useraccount -AccessRights FullAccess -InheritanceType All
                                    Add-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                                }
                            }
                        }
                        $recurrent = Read-Host -Prompt '¿Seguir modificando permisos? (s/n)'
                        if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }
                    }
                }

                2 {

                    # Menu

                    $recurrent = "s"
                    while ($recurrent -eq "s") {
                        Clear-Host
    
                        Write-Output ''
                        Write-Output 'Quitar Permisos'
                        Write-Output '*******************************************'
                        Write-Output '1. Quitar permisos a una cuenta (full access)'
                        Write-Output '2. Quitar permisos a una cuenta (read access)'
                        Write-Output '3. Usar archivo TXT con listado de usuarios (full access)'
                        Write-Output '*******************************************'

                        $opsadd = Read-Host 'Opcion'

                        switch ($opsadd) {
                        
                            # Se leen los datos de la operación y se quitan los permisos sobre las cuentas especificadas

                            1 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $useraccount = Read-Host 'Introduce la cuenta que perdera los permisos'
                                Remove-MailboxPermission -Identity $account -User $useraccount -AccessRights FullAccess
                                Remove-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                            }
                        
                            2 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $useraccount = Read-Host 'Introduce la cuenta que perdera los permisos'
                                Remove-MailboxPermission -Identity $account -User $useraccount -AccessRights ReadPermission
                                Remove-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                            }

                            3 {
                                $account = Read-Host 'Introduce la cuenta sobre la cual se modificaran los permisos'
                                $path = Read-Host 'Introduce la ruta del fichero TXT con los usuarios'
                                foreach ($useraccount in [System.IO.File]::ReadLines("$path")) {
                                    Write-Output "Procesando $useraccount"
                                    Remove-MailboxPermission -Identity $account -User $useraccount -AccessRights FullAccess
                                    Remove-RecipientPermission $account -AccessRights SendAs -Trustee $useraccount
                                }
                            }
                        }
                        $recurrent = Read-Host -Prompt '¿Seguir modificando permisos? (s/n)'
                        if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }
                    }
                }
            }
        }
        4 {
            # Leemos string y se realiza la búsqueda
            $recurrent = "s"

            while ($recurrent -eq "s") {

                Clear-Host

                $azure_string = Read-Host -Prompt 'Introduce la cuenta sobre la cual se quieren revisar los permisos'

                $result = Get-MsolUser -SearchString "$azure_string" | Select-Object UserPrincipalName, DisplayName, isLicensed, licenses | Sort-Object UserPrincipalName | Format-Table

                # Error por si la variable $result tiene un valor nulo

                if ($result) { $result }else { Write-Host -ForegroundColor Red "No se encontraron resultados para $azure_string" }

                # Submenu

                $recurrent = Read-Host -Prompt '¿Realizar otra busqueda? (s/n)'
                if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }

            }
        }
        5 {
            # Leemos el String y se realiza la operación de desactivar el automapeo

            $recurrent = "s"
            while ($recurrent -eq "s") {
                Clear-Host

                Write-Output ''
                Write-Output 'Selecciona una opción'
                Write-Output '*******************************************'
                Write-Output '1. Desactivar automapping para un usuario de un buzón compartido'
                Write-Output '2. Desactivar automapping para todos los usuarios de un buzón compartido '
                Write-Output '*******************************************'
                Write-Output ''
            
                $opsadd = Read-Host 'Opcion'

                switch ($opsadd) {

                    1 {
                        $account = Read-Host -Prompt 'Introduce la cuenta de correo compartida'
                        $usermapping = Read-Host -Prompt 'Introduce la cuenta de correo del usuario'
                        Remove-MailboxPermission -Identity $account -User $usermapping -AccessRights FullAccess
                        Add-MailboxPermission -Identity $account  -User  $usermapping -AccessRights FullAccess
                    }

                    2 {
                        $account = Read-Host -Prompt 'Introduce la cuenta de correo'
                        Remove-MailboxPermission $account  –ClearAutoMapping
                    }

                }
                $recurrent = Read-Host -Prompt '¿Desactivar automapping de otra cuenta? (s/n)'
                if ($recurrent -eq 'n') { $validall = Read-Host -Prompt '¿Realizar otra operación? (s/n)' }

            }
        }
        default {Write-Host "Opción no válida" -BackgroundColor Red}
    }

} while ($validall -eq 's') 

# Se cierra la sesión remota contra O365

Disconnect-AzureAD
Remove-PSSession $Session
