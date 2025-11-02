# ESP.EE · Painel Professor v4.4.1

SPA HTML/CSS/JS (sem build) com autenticação **MSAL** e **Microsoft Graph**.

## Publicação (GitHub Pages)
- Settings ▸ Pages ▸ Source: `Deploy from a branch` → Branch: `main` → Folder: `/ (root)`
- `index.html` inclui `<base href="/esp-painel-professor-v4.4/">`.

## App Registration (Entra ID)
- Redirect URI: `https://bibliotecaesparedes-hub.github.io/esp-painel-professor-v4.4/`
- Permissões delegadas: `Files.ReadWrite.All`, `User.Read` (+ `openid profile offline_access`).

## Endpoints Graph usados
- Ler JSON: `GET /sites/{siteId}/drive/root:{path}:/content`
- Gravar JSON: `PUT /sites/{siteId}/drive/root:{path}:/content`
- Listar backups: `GET /sites/{siteId}/drive/root:{folderPath}:/children`

## Estrutura
- `config_especial.json` — configuração (professores, alunos, disciplinas, grupos, calendario)
- `2registos_alunos.json` — registos (v1)
