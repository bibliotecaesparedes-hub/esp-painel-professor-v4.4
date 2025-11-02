# REFERÊNCIAS OFICIAIS — Microsoft Learn

## Upload/Replace de ficheiros via Microsoft Graph (PUT …:/content)
- **Descrição**: Carregar ou substituir conteúdo de ficheiros até 250 MB.
- **Endpoint**:
  ```
  PUT /sites/{site-id}/drive/root:{path-to-file}:/content
  ```
- **Doc**: https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0

## Listagem de itens numa pasta (…:/children)
- **Endpoint**:
  ```
  GET /sites/{site-id}/drive/root:{folder-path}:/children
  ```
- **Doc**: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0

## MSAL para SPA — acquireTokenSilent ➜ acquireTokenRedirect
- **Docs**:
  - https://learn.microsoft.com/en-us/entra/identity-platform/scenario-spa-acquire-token
  - https://learn.microsoft.com/en-us/entra/msal/javascript/browser/acquire-token
