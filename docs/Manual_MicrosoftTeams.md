



# MicrosoftTeams
  
Este módulo permite conectarse a la API de Microsoft Teams crear y administrar equipos, grupos y canales  

*Read this in other languages: [English](Manual_MicrosoftTeams.md), [Português](Manual_MicrosoftTeams.pr.md), [Español](Manual_MicrosoftTeams.es.md)*
  
![banner](imgs/Banner_MicrosoftTeams.png o jpg)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Como usar este modulo

Antes de usar este modulo, es necesario registrar tu aplicación en el portal de Azure App Registrations. 

1. Inicie sesión en Azure Portal (Registración de aplicaciones: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade ).
2. Seleccione "Nuevo registro".
3. En “Tipos de cuenta compatibles” soportados elija:
    - "Cuentas en cualquier directorio organizativo (cualquier directorio de Azure AD: multiinquilino) y cuentas de Microsoft personales (como Skype o Xbox)" para este caso utilizar  ID Inquilino = **common**.
    - "Solo cuentas de este directorio organizativo (solo esta cuenta: inquilino único) para este caso utilizar **ID Inquilino** especifico de la aplicación.
    - "Solo cuentas personales de Microsoft " for this case use use Tenant ID = **consumers**.
4. Establezca la uri de redirección (Web) como: https://localhost:5001/ y haga click en "Registrar".
5. Copie el ID de la aplicación (cliente). Necesitará este valor.

6. Dentro de "Certificados y secretos", genere un nuevo secreto de cliente. Establezca la caducidad (preferiblemente 24 meses). Copie el VALOR del secreto de cliente creado (NO el ID de Secreto). El mismo se ocultará al cabo de unos minutos.
7. Dentro de "Permisos de API", haga click en "Agregar un permiso", seleccione "Microsoft Graph", luego "Permisos delegados", busque y seleccione "Directory.ReadWrite.All", "Group.ReadWrite.All", "TeamMember.ReadWrite.All", "ChannelSettings.Read.All", "ChannelSettings.ReadWrite.All", "Directory.Read.All", "Group.Read.All", "ChannelMessage.Read.All" y por ultimo "Agregar permisos".
8. Codigo de acceso, generar codigo ingresando al siguiente link:
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={**client_id**}&response_type=code&redirect_uri={**redirect_uri**}&response_mode=query&scope=offline_access%20files.readwrite.all&state=12345
Reemplazar dentro del link {tennat}, {client_id} y {redirect_uri}, por los datos 
correspondientes a la applicación creada.
9. Si la operación tuvo exito, la URL del navedador cambiara por: http://localhost:5001/?code={**CODE**}&state=12345#!/ 
El valor que figurara en {CODE}, copiarlo y utilizarlo en el comando de Rocketbot en el campo "code" para realizar la conexión.

Nota: El navegador NO cargara ninguna pagina.

## Descripción de los comandos

### Establecer credenciales
  
Establece las credenciales para tener disponible la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|client_id|ID de cliente obtenido en la creación de la aplicación|Your client_id|
|client_secret|Secreto del cliente obtenido en la creación de la aplicación|Your client_secret|
|redirect_uri|URL de redireccionamiento de la aplicación|http://localhost:5000|
|code|Dato obtenido al colocar la URL de autenticación. Revisa la documentación para más información|code|
|tenant|Identificador del tenant al que se desea conectar|tenant|
|Resultado|Variable para guardar resultado. Si la conexion es exitosa retornara True, caso contrario sera False|connection|
|session|Variable para guardar el identificador de sesión. Utilizar en caso de que desee conectarse a más de una cuenta de forma simultánea|session|

### Crear Equipo
  
Crea un nuevo equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre|Nombre del equipo|Rocketbot|
|Descripcion|Descripcion del equipo (opcional)|Team Rocketbot|
|Visibilidad|Visibilidad Public o Private|Public|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contrario sera False|res|
|session|Identificador de sesión|session|

### Listar Equipos
  
Listar equipos a los que pertenece
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Variable donde se guardará el resultado de la consulta|res|
|session|Identificador de sesión|session|

### Obtener detalles de un equipo
  
Obtener detalles de un equipo específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Team ID|Identificador de equipo|Team ID|
|Resultado|Nombre de variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Eliminar equipo
  
Elimina un equipo específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Team ID|Identificador de equipo|ID Team|
|Resultado|Nombre de variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Listar miembros
  
Lista miembros de un equipo específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Team ID|Identificador de equipo|ID Team|
|Resultado|Nombre de variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Añadir miembro
  
Añadir miembro a un equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Team|Identificador de equipo|ID Team|
|User ID|Identificador de usuario|ID|
|Correo del usuario|Correo del usuario|test@test.com|
|Resultado|Nombre de variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Eliminar miembro
  
Eliminar miembro de un equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Team|Identificador de equipo|ID Team|
|User ID|Identificador de usuario|User ID|
|Resultado|Nombre de variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Listar canales
  
Listar canales dentro de un equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo|ID del equipo para listar canales|Team ID|
|Ordenar por|Parámetros para ordenar los resultados de la consulta realizada|name desc|
|Filter by|Filtro a aplicar para realizar la consulta|name eq 'file.txt'|
|Cantidad|Cantidad de items a obtener. Devolvera el top de items de la consulta.|10|
|Resultado|Nombre de la variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Obtener detalles de un canal
  
Obtener detalles de un canal específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo|ID del equipo|Team ID|
|ID del canal|ID del canal.|Channel ID|
|Resultado|Nombre de la variable donde se guardará el resultado|res|
|session|Identificador de sesión|session|

### Crear Canal
  
Crear un nuevo canal en un equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del aquipo|ID del aquipo|23XWM5ASR67M67S6KYNCV66KFMQFOTOPDL|
|Nombre||Name|
|Descripcion (Opcional)||Description (Optional)|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contrario sera False|res|
|session|Identificador de sesión|session|

### Subir archivo
  
Eliminar un canal de un equipo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo|ID del equipo.|Team ID|
|ID del canal|ID del canal.|channel id|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contrario sera False|res|
|session|Identificador de sesión|session|

### Listar mensajes
  
Listar mensajes en un canal
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo||Team ID|
|ID del canal||15ZLM4OKQTAC3M7UDDR5DBUKPA4U8ULNXW|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contraria sera False|res|
|session||session|

### Obtener detalles de un mensaje
  
Obtener detalles de un mensaje específico en un canal
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo|ID del equipo|id|
|ID del canal|ID del canal|id|
|ID del Mensaje|ID del Mensaje|id|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contrario sera False|res|
|session|Identificador de sesión|session|

### Enviar un mensaje
  
Enviar un mensaje a un canal
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del equipo|ID del equipo|id|
|ID del Canal|ID del Canal|id|
|Cuerpo del mensaje|Cuerpo del mensaje|content|
|Asunto|Asunto|subject|
|Resultado|Variable para guardar resultado. Si la operacion es exitosa retornara True, caso contrario sera False|res|
|session|Identificador de sesión|session|
