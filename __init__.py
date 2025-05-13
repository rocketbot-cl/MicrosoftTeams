# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import os
import sys
import json


base_path = tmp_global_obj["basepath"]
cur_path = base_path + "modules" + os.sep + "MicrosoftTeams" + os.sep + "libs" + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)
"""
    Obtengo el modulo que fue invocado
"""
from MicrosoftTeams import MicrosoftTeams

module = GetParams("module") 

global mod_Teams_session

session = GetParams("session")
if not session:
    session = 'default'

try:
    if not mod_Teams_session : #type:ignore
        mod_Teams_session = {}
except NameError:
    mod_Teams_session = {}


if module == "setCredentials":
    client_secret = GetParams("client_secret")
    client_id = GetParams("client_id")
    redirect_uri = GetParams("redirect_uri")
    code = GetParams("code")
    tenant = GetParams("tenant")
    res = GetParams("res")

    if session == '': 
         credentials_filename = "teams_credentials.json"
    else:
        credentials_filename = "teams_credentials_{s}.json".format(s=session)


    path_credentials = base_path + "modules" + os.sep + "MicrosoftTeams" + os.sep + credentials_filename

    mod_Teams_session[session] = MicrosoftTeams(client_id=client_id, client_secret=client_secret, tenant=tenant, redirect_uri=redirect_uri,
                                                path_credentials=path_credentials)

    try:
        try:
            with open(path_credentials) as json_file:
                data = json.load(json_file)
            auth_param = {'refresh_token': data['refresh_token']}
            grant_type = 'refresh_token'
            response = mod_Teams_session[session].get_token(auth_param, grant_type)
        except IOError:
            grant_type = 'authorization_code'
            auth_param = {'code': code}
            response = mod_Teams_session[session].get_token(auth_param, grant_type)
        is_connected = mod_Teams_session[session].create_tokens_file(response)

        SetVar(res,is_connected)

    except Exception as e:
        SetVar(res, response)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

try:
    mod_Teams_session[session]
except:
    raise Exception("Must be connected before executing any command...")

# Crear un nuevo equipo
if module == "createTeam":
    display_name = GetParams("display_name")
    description = GetParams("description") or None
    visibility = GetParams("visibility") or "Public" # "Public", "Private"
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].create_team(display_name, description=description, visibility=visibility)
        # Note: Team creation is async, response might contain a location header
        SetVar(res, response)

    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción creando equipo: {e}'})
        raise e
    
#obtiene los equipos a los que pertenece    
if module == "listJoinedTeams":
    res = GetParams("res")

    try:
        response =  mod_Teams_session[session].list_joined_teams()
        if 'error' not in response:
            teams_list = [{'id': team.get('id'), 'displayName': team.get('displayName')} for team in response.get('value', [])]
            SetVar(res, teams_list)
        else:            
            SetVar(res, response) 
            print("\x1B[" + "31;40mError listando equipos unidos.\x1B[" + "0m")
    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción listando equipos: {e}'})
        raise e

# Obtener detalles de un equipo específico
if module == "getTeamDetails":
    team_id = GetParams("team_id")
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].get_team_details(team_id)
        SetVar(res, response)

    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción obteniendo detalles del equipo: {e}'})
        raise e

# Eliminar un equipo
if module == "deleteTeam":
    team_id = GetParams("team_id")
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].delete_team(team_id)
        SetVar(res, response)
    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción eliminando equipo: {e}'})
        raise e

# Listar miembros de un equipo
if module == "listMembers":
    team_id = GetParams("team_id")
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].list_members(team_id)
        SetVar(res, members_list)

    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción listando miembros: {e}'})
        raise e

# Añadir miembro a un equipo
if module == "addMember":
    team_id = GetParams("team_id")
    user_id = GetParams("user_id") or None # User's object ID
    user_principal_name = GetParams("user_principal_name") or None # User's email address
    res = GetParams("res")

    try:
        # Pass either user_id or user_principal_name to the library function
        response = mod_Teams_session[session].add_member(team_id, user_id=user_id, user_principal_name=user_principal_name)
        SetVar(res, response)
    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción añadiendo miembro: {e}'})
        raise e

# Eliminar miembro de un equipo
if module == "removeMember":
    team_id = GetParams("team_id")
    member_id = GetParams("member_id")
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].remove_member(team_id, member_id)
        SetVar(res, response)
    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción eliminando miembro: {e}'})
        raise e

# Listar canales dentro de un equipo
if module == "listChannels":
    team_id = GetParams("team_id") 
    filter_by = GetParams("filter_by") or None 
    order_by = GetParams("order_by") or None 
    top = GetParams("top") or None 
    res = GetParams("res") 

    try:
        response = mod_Teams_session[session].list_channels(team_id, filter_by=filter_by, order_by=order_by, top=top)
        if 'error' not in response:
            channels_list = [{'id': channel.get('id'), 'displayName': channel.get('displayName'), 'description': channel.get('description')} for channel in response.get('value', [])]
            SetVar(res, channels_list)

    except Exception as e:
            PrintException(e)
            SetVar(res, {'error': f'Excepción listando canales: {e}'})
            raise e

# Obtener detalles de un canal específico
if module == "getChannelDetails":
    team_id = GetParams("team_id") 
    channel_id = GetParams("channel_id") 
    res = GetParams("res") 

    try:
        response = mod_Teams_session[session].get_channel_details(team_id, channel_id)
        SetVar(res, response)
    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción obteniendo detalles del canal: {e}'})
        raise e

# Crear un nuevo canal en un equipo
if module == "createChannel":
    team_id = GetParams("team_id")
    display_name = GetParams("display_name")
    description = GetParams("description") or None 
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].create_channel(team_id, display_name, description=description)
        SetVar(res, response)
    except Exception as e:
            PrintException(e)
            SetVar(res, {'error': f'Excepción creando canal: {e}'})
            raise e

# Eliminar un canal de un equipo
if module == "deleteChannel":
    team_id = GetParams("team_id") 
    channel_id = GetParams("channel_id") 
    res = GetParams("res") 
    try:
        response = mod_Teams_session[session].delete_channel(team_id, channel_id)
        SetVar(res, response)

    except Exception as e:
        PrintException(e)
        SetVar(res, {'error': f'Excepción eliminando canal: {e}'})
        raise e


# Listar mensajes en un canal
if module == "listMessages":
    team_id = GetParams("team_id") 
    channel_id = GetParams("channel_id")
    res = GetParams("res")

    try:
        response = mod_Teams_session[session].list_messages(team_id, channel_id)
        messages_list = [{'id': msg.get('id'),
                        'sentDateTime': msg.get('sentDateTime'),
                        'from': msg.get('from', {}).get('user', {}).get('displayName', 'Usuario Desconocido'),
                        'bodyPreview': msg.get('body', {}).get('content', '')[:100] + '...'
                        } for msg in response.get('value', [])]
        SetVar(res, messages_list)
    except Exception as e:
            PrintException(e)
            SetVar(res, {'error': f'Excepción listando mensajes: {e}'})
            raise e

# Obtener detalles de un mensaje específico en un canal
if module == "getMessageDetails":
    team_id = GetParams("team_id")
    channel_id = GetParams("channel_id")
    message_id = GetParams("message_id")
    res = GetParams("res")
    try:
        response = mod_Teams_session[session].get_message_details(team_id, channel_id, message_id)
        SetVar(res, response) 
    except Exception as e:
            PrintException(e)
            SetVar(res, {'error': f'Excepción obteniendo mensaje: {e}'})
            raise e

# Enviar un mensaje a un canal
if module == "sendChannelMessage":
    team_id = GetParams("team_id") 
    channel_id = GetParams("channel_id")
    content = GetParams("content") 
    subject = GetParams("subject") or None 
    res = GetParams("res") 

    try:
            response = mod_Teams_session[session].send_channel_message(team_id, channel_id, content, subject=subject)
            SetVar(res, response) 
    except Exception as e:
            PrintException(e)
            SetVar(res, {'error': f'Excepción enviando mensaje: {e}'})
            raise e



