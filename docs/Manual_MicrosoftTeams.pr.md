



# MicrosoftTeams
  
 Este módulo permite conectar-se à API do Microsoft Teams para criar e gerenciar equipes, grupos e canais  

*Read this in other languages: [English](Manual_MicrosoftTeams.md), [Português](Manual_MicrosoftTeams.pr.md), [Español](Manual_MicrosoftTeams.es.md)*
  
![banner](imgs/Banner_MicrosoftTeams.png o jpg)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Como usar este módulo

Antes de usar este módulo, você precisa registrar seu aplicativo no portal de Registros de Aplicativo do Azure.

1. Entre no portal do Azure (Registro de Aplicativos: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade ).
2. Selecione "Novo registro".
3. Em "Tipos de conta suportados", escolha:
    - "Contas em qualquer diretório organizacional (qualquer diretório do Azure AD: multilocatário) e contas pessoais da Microsoft (como Skype ou Xbox)" para este caso, use ID de locatário = **common**.
    - "Somente contas deste diretório organizacional (somente esta conta: locatário único) para este caso usam **ID de locatário específico** do aplicativo.
    - "Somente contas pessoais da Microsoft" para este caso, use ID do locatário = **consumers**.
4. Defina o redirecionamento uri (Web) como: https://localhost:5001/ e clique em "Registrar".
5. Copie o ID do aplicativo (cliente). Você vai precisar desse valor.
6. Em "Certificados e 
segredos", gere um novo segredo do cliente. Defina a validade (de preferência 24 meses). Copie o VALUE do segredo do cliente criado (NÃO o ID do segredo). Ele vai esconder depois de alguns minutos.
7. Em "Permissões de API", clique em "Adicionar uma permissão", selecione "Microsoft Graph", depois "Permissões delegadas", localize e selecione "Directory.ReadWrite.All", "Group.ReadWrite.All", "TeamMember.ReadWrite.All", "ChannelSettings.Read.All", "ChannelSettings.ReadWrite.All", "Directory.Read.All", "Group.Read.All", "ChannelMessage.Read.All" e, finalmente, "Adicionar permissões".
8. Acesse o código, gere o código entrando no seguinte link:
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={**client_id**}&response_type=code&redirect_uri={**redirect_uri**}&response_mode=query&scope=offline_access%20files.readwrite.all&state=12345
Substitua no link {tennat}, {client_id} e {redirect_uri}, pelos dados correspondentes ao aplicativo criado.
9. Se a operação for bem-
sucedida, a URL do navegador será alterada para: http://localhost:5001/?code={**CODE**}&state=12345#!/
O valor que aparece em {CODE}, copie-o e use-o no comando Rocketbot no campo "code" para fazer a conexão.

Nota: O navegador NÃO carregará nenhuma página.
## Descrição do comando

### Definir credenciais
  
Defina as credenciais para ter a API disponível
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|client_id|ID do cliente obtido na criação do aplicativo|Your client_id|
|client_secret|Segredo do cliente obtido na criação do aplicativo|Your client_secret|
|redirect_uri|URL de redirecionamento do aplicativo|http://localhost:5000|
|code|Dados obtidos colocando a URL de autenticação. Verifique a documentação para mais informações|code|
|tenant|Identificador do tenant ao qual você deseja se conectar|tenant|
|Resultado|Variável para armazenar resultado. Se a conexão for bem sucedida retornará True, caso contrário será False|connection|
|session|Variável para armazenar o identificador da sessão. Use caso você queira se conectar a mais de uma conta ao mesmo tempo|session|

### Criar Time
  
Criar um novo time
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome|Nome da equepe|Rocketbot|
|Descrição|Descrição da Equipe (opcional)|Team Rocketbot|
|Visibilidade|Visibilidade Pública ou Privada|Public|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session|ID da sessão|session|

### Listar Equipes
  
Listar equipes a que pertence
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado|Variável onde o resultado da consulta será salvo|res|
|session|ID da sessão|session|

### Obter detalhes de uma equipe
  
Obter detalhes de uma equipe específica
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da Equipe|ID da Equipe|Team ID|
|Resultado|Nome da variável onde o resultado será guardado|res|
|session|ID da sessão|session|

### Excluir um time
  
Excluir um time específico
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da Equipe|ID da Equipe|ID Team|
|Resultado|Nome da variável onde o resultado será guardado|res|
|session|ID da sessão|session|

### Listar Membros
  
Listar Membros de uma equipe específica
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da Equipe|ID da Equipe|ID Team|
|Resultado|Nome da variável onde o resultado será guardado|res|
|session|ID da sessão|session|

### Adicionar Membro
  
Adicionar Membro de uma equipe específica
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Team|ID da Equipe|ID Team|
|ID do Usuário|ID do Usuário|ID|
|e-mail do usuário|Endereço de e-mail do usuário|test@test.com|
|Resultado|Nome da variável onde o resultado será guardado|res|
|session|ID da sessão|session|

### Remover membro
  
Remover membro de um time específico
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Team|ID da Equipe|ID Team|
|ID do usuário|ID do usuário|User ID|
|Resultado|Nome da variável onde o resultado será guardado|res|
|session|ID da sessão|session|

### Listar canais
  
Listar canais dentro de uma equipe
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe| ID da Equipe para listar canais|Team ID|
|Ordenar por|Parâmetros para ordenar os resultados da consulta realizada|name desc|
|Filtrar por|Filtro a ser aplicado para realizar a consulta|name eq 'file.txt'|
|Quantia|Número de itens a serem obtidos. Ele retornará os principais itens da consulta|10|
|Resultado|Nome da variável para salvar o resultado|res|
|session|ID da sessão|session|

### Obter detalhes de um canal
  
Obter detalhes de um canal específico
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe|ID da equipe|Team ID|
|ID do canal|ID do canal.|Channel ID|
|Resultado|Nome da variável para salvar o resultado|res|
|session|ID da sessão|session|

### Criar um canal
  
Criar um novo canal em uma equipe
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe|ID da equipe|23XWM5ASR67M67S6KYNCV66KFMQFOTOPDL|
|Nome||Name|
|Descrição (Opcional)||Description (Optional)|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session|ID da sessão|session|

### Excluir um canal
  
Excluir um canal em uma equipe
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe|ID da equipe.|Team ID|
|ID do canal|ID do canal.|channel id|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session|ID da sessão|session|

### Listar mensagens
  
Listar mensagens em um canal
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe||Team ID|
|ID do canal||15ZLM4OKQTAC3M7UDDR5DBUKPA4U8ULNXW|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session||session|

### Obter detalhes de uma mensagem
  
Obter detalhes de uma mensagem específica em um canal
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe|ID da equipe|id|
|ID do canal|ID do canal|id|
|ID da Mensagem|ID da Mensagem|id|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session|ID da sessão|session|

### Enviar uma mensagem
  
Enviar uma mensagem para um canal
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID da equipe|ID da equipe|id|
|ID do canal|ID do canal|id|
|Corpo da mensagem|Corpo da mensagem|content|
|Assunto|Assunto|subject|
|Resultado|Variável para armazenar resultado. Se a tarefa for bem sucedida, retornará True, caso contrário, retornará False|res|
|session|ID da sessão|session|
