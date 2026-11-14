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
- Escolha os registros (todos, marcando na tabela ou digitando `1, 3, 5-10`)
  e clique em *Gerar*. Cada arquivo sai como `<#>_ID<id>.docx` (o # é o da
  coluna da tabela).

## Modelo

O app já vem com o modelo do Relatório de Visita — Programa Adote uma
Praça (`modelos/relatorio_visita_adote_uma_praca.docx`, formulário
"Vistoria - Adote uma Praça - SEPE", projeto 4). Para mudar o layout, baixe
o modelo pelo botão do app, edite no Word mantendo os campos `{{ ... }}`, e
envie no app (vale só para aquela geração) ou substitua o arquivo em
`modelos/`.

Campos disponíveis: `id_proj`, `data`, `processo_sei`, `id_tipo_rel`,
`responsavel`, `cidade`, `lat`, `long`, `observacao`, `adotante`,
`possui_placa`, `plano_trabalho`, `obs_plano_trabalho`, `tipo_proj`, `meta`,
`relatorio` (o # da tabela), `imagem_1` (croqui/planta de localização) e
`imagem_2` a `imagem_5` (Foto 1 a 4).

Cada campo vem de uma coluna do CSV do ODK pelo nome (dicionário `CAMPOS`
no topo de `sepe_core.py`) — se o formulário mudar o nome de um campo, é só
ali.

## Rodar local

```
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\streamlit run app.py
```
