# 🧾 Notificações Extrajudiciais Automatizadas

Este projeto gera notificações extrajudiciais personalizadas em `.docx`, a partir de um modelo Word e uma planilha Excel com os dados dos clientes e suas operações.

---

## ✅ Como usar

1. **Crie e ative o ambiente virtual (opcional, recomendado):**

```bash
python3 -m venv venv
source venv/bin/activate
```

2. **Instale as dependências:**

```bash
pip install -r requirements.txt
```

3. **Configure o ambiente (opcional):**

Você pode definir variáveis no terminal ou via `.env`:

```bash
export SUA_CIDADE="Macapá-AP"
export AGENCIA="0001-X"
export ENDERECO_AGENCIA="Rua da Agência, 123"
export CNPJ="00.000.000/0001-00"
```
OU

Você pode alterar os dados das constantes para os dados que precisa dentro do código.


4. **Organize as pastas esperadas:**

```
.
├── main.py
├── modelo/
│   └── Modelo_notificacao.docx
├── planilhas/
│   └── criar_notificacoes.xlsx
├── notificacoes/
└── requirements.txt
```

---

## 📁 Requisitos dos Arquivos

### 📄 1. Modelo Word (`modelo/Modelo_notificacao.docx`)

- Use o marcador `«OPERACOES_BLOCO»` no ponto onde deseja que a tabela de operações apareça.
- Os demais campos obrigatórios devem estar como `«NOME_DO_CAMPO»`, por exemplo:
  - `«CLIENTE»`, `«ENDERECO_CLIENTE»`, `«CIDADE_CLIENTE»`, `«UF_CLIENTE»`, etc.

### 📊 2. Planilha Excel (`planilhas/criar_notificacoes.xlsx`)

Cada linha representa **uma operação** de um cliente.
mesmo que o cliente se repita com várias operações, será criada apenas uma notificação

A planilha deve conter as seguintes colunas:

| Coluna              | Descrição                            |
|---------------------|----------------------------------------|
| cod_cliente         | Identificador único do cliente         |
| nome_cliente        | Nome completo do cliente               |
| endereco_cliente    | Endereço do cliente                    |
| cidade_cliente      | Cidade                                 |
| uf_cliente          | UF                                     |
| cep_cliente         | CEP                                    |
| produto             | Nome do produto/contrato               |
| operacao            | Código ou número da operação           |
| vencimento          | Data de vencimento da operação         |

---

## ✅ Saída

Os arquivos gerados serão salvos na pasta `notificacoes/` com o nome:

```
notificacao_<primeiro_nome>_<cod_cliente>_<data>.docx
```

Exemplo:

```
notificacao_Joao_123_2025-05-20.docx
```

---

## 📌 Exemplo de execução

Ponha os dados na planilha `planilhas/criar_notificacoes.xlsx` 

lembre de criar as pastas para os arquivos.

Rode via terminal
```bash
python main.py
```


---
