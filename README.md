# Automatizador de Visitas

Ferramenta simples para pegar uma planilha Excel de visitas e gerar uma nova planilha com as visitas separadas por gerente.

O script lê a planilha base, identifica o gerente na coluna D e cria um arquivo final com:

- uma aba `Base Total`, com todos os registros que possuem gerente;
- uma aba para cada gerente encontrado;
- as mesmas colunas principais da planilha original.

## Pastas do projeto

```text
automatizador_visitas/
├── entrada/                 # coloque aqui a planilha original .xlsx
├── saida/                   # o arquivo gerado será salvo aqui
├── separar_gerentes.py      # script principal
├── separar_gerentes.ps1     # atalho para rodar no PowerShell
└── README.md                # esta documentação
```

## Como usar

### 1. Coloque a planilha na pasta de entrada

Copie o arquivo Excel original para:

```text
entrada/
```

A planilha precisa estar no formato `.xlsx`.

Se houver mais de uma planilha na pasta `entrada`, o sistema usa a mais recente.

### 2. Rode o automatizador

No terminal, dentro da pasta do projeto, execute:

```bash
python separar_gerentes.py
```

No PowerShell do Windows, também pode usar:

```powershell
.\separar_gerentes.ps1
```

### 3. Veja o resultado

O arquivo final será criado em:

```text
saida/visitas_por_gerente.xlsx
```

## Comandos opcionais

Se quiser indicar manualmente o arquivo de entrada e o arquivo de saída:

```bash
python separar_gerentes.py -i entrada/minha_planilha.xlsx -o saida/resultado.xlsx
```

Se a coluna do gerente mudar, informe o número da coluna. Por padrão, o gerente fica na coluna D, que é a coluna 4.

Exemplo usando a coluna E:

```bash
python separar_gerentes.py --manager-column 5
```

## Observações importantes

- Feche o Excel antes de rodar, principalmente se o arquivo de saída já existir.
- Arquivos temporários do Excel que começam com `~$` são ignorados.
- Arquivos gerados com nome `visitas_por_gerente` também são ignorados como entrada.
- O script não altera a planilha original; ele cria uma nova planilha na pasta `saida`.

## Exemplo rápido

```bash
# 1. entrar na pasta do projeto
cd "C:\Users\marcos_barros\Downloads\automatizador_visitas"

# 2. rodar
python separar_gerentes.py

# 3. abrir o resultado
# saida/visitas_por_gerente.xlsx
```
