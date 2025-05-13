import requests
import json
import os

class MicrosoftTeams:
    def __init__(self, *, client_id, client_secret, tenant, redirect_uri, path_credentials):
        self.client_id = client_id
        self.client_secret = client_secret
        # Default scope for common Teams operations. Adjust as needed based on required permissions.
        # https://learn.microsoft.com/en-us/graph/permissions-reference#teams-permissions
        self.scope = 'Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All ChannelMessage.Send Chat.Read Chat.ReadWrite offline_access'
        self.redirect_uri = redirect_uri
        self.tenant = tenant
        self.access_token = None
        self.refresh_token = None
        self.path_credentials = path_credentials

    def get_token(self, auth_code_or_refresh_token, grant_type):
        """ Get the access_token or refresh_token.

        Obtains token based on authorization code (first time) or refresh token.

        Parameters
        ----------
        auth_code_or_refresh_token : dict
            Contains 'code' or 'refresh_token'.
        grant_type : str
            Type of grant_type, 'authorization_code' or 'refresh_token'.

        Returns
        -------
        dict
            a json with the credentials, or an error response.
        """
        url, params = self.build_request(auth_code_or_refresh_token, grant_type)
        try:
            response = requests.post(url, data=params)
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            json_response = response.json()
            self.access_token = json_response.get('access_token')
            self.refresh_token = json_response.get('refresh_token', self.refresh_token) # Keep old refresh token if new one not provided
            return json_response
        except requests.exceptions.RequestException as e:
            print(f"Error getting token: {e}")
            return {'error': f'Failed to get token: {e}', 'response_text': response.text if response else None}

    def build_request(self, auth_code_or_refresh_token, grant_type):
        """ Build the token request parameters.

        Parameters
        ----------
        auth_code_or_refresh_token : dict
            Contains 'code' or 'refresh_token'.
        grant_type : str
            Type of grant_type, 'authorization_code' or 'refresh_token'.

        Returns
        -------
        string, dict
            a formed url and a dict with parameters
        """
        params = {
            'client_id': self.client_id,
            'scope': self.scope,
            'redirect_uri': self.redirect_uri,
            'grant_type': grant_type,
            'client_secret': self.client_secret,
        }
        params.update(auth_code_or_refresh_token)
        url = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'.format(tenant=self.tenant)
        return url, params

    def create_tokens_file(self, credentials):
        """ Create or update a json file with credentials.

        Parameters
        ----------
        credentials : dict
            Contains the credentials (access_token, refresh_token, etc.)

        Returns
        -------
        bool
            True if file was created successfully, False otherwise.
        """
        try:
            with open(self.path_credentials, 'w') as credfile:
                json.dump(credentials, credfile)
            return True
        except Exception as e:
            print(f"Error creating credentials file: {e}")
            return False

    def _graph_request(self, method, url_suffix, json_data=None, params=None):
        """ Generic helper for making Graph API requests. """
        if not self.access_token:
            return {'error': 'No access token available. Authenticate first.'}

        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            'Content-Type': 'application/json'
        }
        url = f"https://graph.microsoft.com/v1.0/{url_suffix}"

        try:
            response = requests.request(method, url, headers=headers, json=json_data, params=params)
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            # Some requests (like DELETE, or async ops) might not return JSON
            if response.status_code == 204: # No Content
                return {'success': 'Item deleted'}
            if response.status_code == 202: # Accepted (async operations)
                 # Return the location header for monitoring async operation status
                 return {'status': 'Accepted', 'location': response.headers.get('Location')}
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Graph API request failed ({method} {url}): {e}")
            try:
                error_details = response.json()
            except json.JSONDecodeError:
                 error_details = {'response_text': response.text}

            return {'error': f'Graph API request failed: {e}', 'details': error_details}


    # --- Teams Specific Methods ---
    def create_team(self, display_name, description=None, visibility="Public"):
        """ Creates a new Microsoft Team.
            https://learn.microsoft.com/en-us/graph/api/team-post?view=graph-rest-v1.0&tabs=http
        """
        # Note: Creating a team is an asynchronous operation.
        # The response will contain a 'Location' header pointing to the operation status URL.
        # You might need to poll that URL to know when the team creation is complete.
        payload = {
            "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')", # 'standard' or 'education'
            "displayName": display_name,
            "description": description or "",
            "visibility": visibility # e.g., "Public", "Private"
            }
        return self._graph_request('POST', 'teams', json_data=payload)
    
    def list_joined_teams(self):
        """ Lists all Microsoft Teams that the user is a member of. """
        return self._graph_request('GET', 'me/joinedTeams')

    def get_team_details(self, team_id):
        """ Gets details for a specific Microsoft Team. """
        return self._graph_request('GET', f'teams/{team_id}')
    
    def delete_team(self, team_id):
        """ Deletes a Microsoft Team.
            https://learn.microsoft.com/en-us/graph/api/team-delete?view=graph-rest-v1.0&tabs=http
        """
        # Deleting a team also deletes the underlying Microsoft 365 group.
        return self._graph_request('DELETE', f'teams/{team_id}')
    
    def list_members(self, team_id):
        """ Lists members of a specific team.
            https://learn.microsoft.com/en-us/graph/api/team-list-members?view=graph-rest-v1.0&tabs=http
        """
        url_suffix = f'groups/{team_id}/members'
        return self._graph_request('GET', url_suffix)


    def add_member(self, team_id, user_id=None, user_principal_name=None):
        """ Adds a member to a team (adds a member to the underlying Microsoft 365 Group).
            Requires the user ID or User Principal Name (email).
            https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-v1.0&tabs=http
        """
        if not user_id and not user_principal_name:
            return {'error': 'Either user_id or user_principal_name must be provided.'}

        if user_id:
            payload = {
                "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"
            }
            return self._graph_request('POST', f'groups/{team_id}/members/$ref', json_data=payload)
        elif user_principal_name:
             # Option 1: Search for the user by UPN first to get their ID
             search_response = self._graph_request('GET', f"users?$filter=userPrincipalName eq '{user_principal_name}'&$select=id")
             if 'error' in search_response:
                  return {'error': f'Failed to find user with UPN {user_principal_name}', 'details': search_response['details']}
             users = search_response.get('value', [])
             if not users:
                  return {'error': f'No user found with UPN {user_principal_name}'}
             user_id_to_add = users[0].get('id')
             if not user_id_to_add:
                  return {'error': f'Could not get ID for user with UPN {user_principal_name}'}

             # Now add the user by ID
             payload = {
                "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id_to_add}"
             }
             return self._graph_request('POST', f'groups/{team_id}/members/$ref', json_data=payload)


    def remove_member(self, team_id, member_id):
        """ Removes a member from a team (removes a member from the underlying Microsoft 365 Group).
            Requires the member's ID (the user object ID).
            https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-v1.0&tabs=http
        """
        # Note: This removes the user from the group, which effectively removes them from the team.
        # The member_id is the user object ID.
        return self._graph_request('DELETE', f'groups/{team_id}/members/{member_id}/$ref')

    def list_channels(self, team_id, filter_by=None, order_by=None, top=None):
        """ Lists channels within a specific team. Handles pagination. """
        url_suffix = f'teams/{team_id}/channels'
        params = {}
        if filter_by:
            params['$filter'] = filter_by
        if order_by:
            params['$orderby'] = order_by
        if top:
             params['$top'] = top # $top is not always supported by Graph API list methods

        all_channels = []
        response = self._graph_request('GET', url_suffix, params=params)

        if 'error' in response:
            return response # Return error immediately

        all_channels.extend(response.get('value', []))

        # Handle pagination
        while '@odata.nextLink' in response:
            next_link = response['@odata.nextLink']
            # Need to make a full URL request for nextLink
            try:
                next_response = requests.get(next_link, headers={'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json'})
                next_response.raise_for_status()
                response = next_response.json()
                all_channels.extend(response.get('value', []))
            except requests.exceptions.RequestException as e:
                print(f"Error fetching next page: {e}")
                return {'error': f'Failed to fetch subsequent pages: {e}', 'partial_results': all_channels}


        return {'value': all_channels} # Return in a similar structure to original response


    def get_channel_details(self, team_id, channel_id):
        """ Gets details for a specific channel within a team. """
        return self._graph_request('GET', f'teams/{team_id}/channels/{channel_id}')

    def create_channel(self, team_id, display_name, description=None):
        """ Creates a new channel within a team. """
        payload = {
            "displayName": display_name,
            "description": description or "",
            "membershipType": "standard" # Or "private" etc.
        }
        # Note: Creating private channels is more complex and requires member definition
        # https://learn.microsoft.com/en-us/graph/api/channel-post?view=graph-rest-v1.0&tabs=http#example-3-create-a-private-channel
        return self._graph_request('POST', f'teams/{team_id}/channels', json_data=payload)

    def delete_channel(self, team_id, channel_id):
        """ Deletes a channel within a team. """
        return self._graph_request('DELETE', f'teams/{team_id}/channels/{channel_id}')

    def list_messages(self, team_id, channel_id):
        """ Lists messages in a specific channel. Handles pagination. """
        url_suffix = f'teams/{team_id}/channels/{channel_id}/messages'
        # Can add $top, $filter, $orderby params here if needed, similar to list_channels
        all_messages = []
        response = self._graph_request('GET', url_suffix)

        if 'error' in response:
            return response # Return error immediately

        all_messages.extend(response.get('value', []))

        # Handle pagination
        while '@odata.nextLink' in response:
             next_link = response['@odata.nextLink']
             try:
                 next_response = requests.get(next_link, headers={'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json'})
                 next_response.raise_for_status()
                 response = next_response.json()
                 all_messages.extend(response.get('value', []))
             except requests.exceptions.RequestException as e:
                 print(f"Error fetching next page for messages: {e}")
                 return {'error': f'Failed to fetch subsequent message pages: {e}', 'partial_results': all_messages}

        return {'value': all_messages} # Return in a similar structure

    def get_message_details(self, team_id, channel_id, message_id):
        """ Gets details for a specific message in a channel. """
        return self._graph_request('GET', f'teams/{team_id}/channels/{channel_id}/messages/{message_id}')

    def send_channel_message(self, team_id, channel_id, content, subject=None, attachments=None, mentions=None):
        """ Sends a message to a channel.
            https://learn.microsoft.com/en-us/graph/api/channelmessage-post?view=graph-rest-v1.0&tabs=http
        """
        payload = {
            "body": {
                "content": content
            }
        }
        if subject:
            payload['subject'] = subject
        if attachments:
             # Attachments require specific formatting, see Graph API docs
             payload['attachments'] = attachments
        if mentions:
             # Mentions require specific formatting, see Graph API docs
             payload['mentions'] = mentions


        return self._graph_request('POST', f'teams/{team_id}/channels/{channel_id}/messages', json_data=payload)

