Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://cdn.jsdelivr.net/gh/fylcr/easy-ai/OmniVoice/setup.ps1 | Invoke-Expression
