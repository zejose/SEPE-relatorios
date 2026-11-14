# SEPE — Gerador de Relatórios de Vistoria

Gera os relatórios DOCX de vistoria direto das submissões do ODK Central
(com as fotos), a partir de um modelo `formulario.docx`.

## Publicar no Streamlit Community Cloud

1. Em <https://share.streamlit.io> → **Create app** → escolha o repositório
   `zejose/SEPE-relatorios`, branch `main`, arquivo `app.py`.
2. Em **Advanced settings → Secrets** (opcional), cole o conteúdo de
   `.streamlit/secrets.toml.example` preenchido. O que estiver ali só
   preenche a tela de conexão.
3. Se colocar a senha do ODK nos Secrets, deixe o app **privado**
   (Share → "Only specific people"): senão qualquer pessoa com o link
   consegue baixar os dados.

Cada `git push` na `main` atualiza o app sozinho.

## Uso

- **ODK Central**: informe e-mail/senha e clique em *Buscar dados*. Só o
  CSV é baixado nesse momento; as fotos são baixadas na hora de gerar, e só
  as dos registros escolhidos.
- **Arquivo exportado**: envie o ZIP exportado pelo ODK Central (CSV +
  pasta `media`) — gera com as fotos sem precisar de senha. Só o CSV também
  funciona, mas os relatórios saem sem foto.
- Escolha os registros (todos, marcando na tabela ou digitando `1, 3, 5-10`),
  envie o modelo DOCX e clique em *Gerar*.

O número do relatório é a linha da planilha (o 1º registro é o relatório
2), igual à versão anterior.

## Campos do modelo

`relatorio`, `data`, `id_proj`, `processo_sei`, `id_tipo_rel`,
`possui_placa`, `plano_trabalho`, `cidade`, `lat`, `long`, `tipo_proj`,
`observacao`, `meta`, `responsavel`, `imagem_1` a `imagem_5`.

Cada campo vem de uma posição fixa do CSV do ODK — ver as constantes
`COL_*` no topo de `sepe_core.py`, que é o único lugar a mudar se o
formulário mudar de ordem.

## Rodar local

```
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\streamlit run app.py
```
