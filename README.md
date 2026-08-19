# 🔍 Auditoria de Segregação de Funções em Git Log (RPA + Python)

Script em Python que automatiza uma checagem clássica de auditoria de TI: **garantir que quem desenvolve (Pull Request) não seja a mesma pessoa que publica a versão (Release)** — e vice-versa.

Em vez de revisar centenas de linhas de log manualmente, o script lê o arquivo `.log`, classifica cada pessoa em um grupo (`Pull Request` ou `Release`) com base no seu comportamento histórico, e sinaliza automaticamente qualquer linha em que essa regra de segregação de funções (SoD) tenha sido quebrada — tudo consolidado em uma planilha Excel pronta para servir de evidência.

> 📖 Este projeto nasceu de um case real aplicado em Auditoria Interna de TI. Leia a história completa no artigo do LinkedIn: **[cole aqui o link do seu post]**

---

## 🎯 A regra de negócio

| Grupo | Pode fazer | Não pode fazer |
|---|---|---|
| **Pull Request** (desenvolvimento) | Abrir e mesclar PRs | Fechar/tagear uma release |
| **Release** (publicação) | Fechar/tagear uma release | Abrir PRs |

Sempre que uma pessoa aparece atuando fora do seu grupo predominante, o script marca a ocorrência como **conflito de segregação**.

---

## 📂 Estrutura do repositório

```
.
├── analisar_git_log.py              # Script principal
├── requirements.txt                 # Dependências do projeto
├── logs/
│   ├── git-log-jan-jun-2025.log     # Log de exemplo — 1º semestre (sem conflitos)
│   └── git-log-jul-dez-2025.log     # Log de exemplo — 2º semestre (com conflitos)
└── exemplos_saida/
    ├── relatorio_1_semestre_2025.xlsx
    └── relatorio_2_semestre_2025.xlsx
```

---

## ⚙️ Como o script funciona

1. **Leitura e parsing do log** (`re`): identifica cada `commit` (Pull Request) e cada `tag` (Release), extraindo autor, data, PR relacionado e mensagem.
2. **Classificação de grupo por pessoa**: cada pessoa é classificada como `Pull Request` ou `Release` conforme o que ela faz na **maioria** das vezes ao longo do log.
3. **Checagem linha a linha**: todo evento (commit ou release) é comparado ao grupo esperado de quem o executou. Divergências são marcadas como conflito, com o motivo já descrito.
4. **Geração do relatório** (`pandas` + `openpyxl`): consolida tudo em um `.xlsx` formatado, com cabeçalho estilizado, colunas ajustadas automaticamente e as linhas de conflito destacadas em vermelho.

### Saída gerada (`.xlsx`)

| Aba | Conteúdo |
|---|---|
| **Resumo** | Visão macro: totais, pessoas em ambos os grupos e o veredito final da segregação |
| **Eventos** | Todos os commits e releases numa única tabela cronológica, com grupo do autor e flag de conflito |
| **Filhos_por_Pai** | Quantos Pull Requests (filhos) cada responsável por Release (pai) efetivamente fechou |

---

## 🚀 Como rodar

### 1. Clone o repositório
```bash
git clone https://github.com/SEU-USUARIO/NOME-DO-REPOSITORIO.git
cd NOME-DO-REPOSITORIO
```

### 2. Crie e ative um ambiente virtual
```bash
python3 -m venv venv

# Linux / Mac
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Rode a análise
```bash
python analisar_git_log.py logs/git-log-jan-jun-2025.log relatorio_1_semestre.xlsx
python analisar_git_log.py logs/git-log-jul-dez-2025.log relatorio_2_semestre.xlsx
```

O segundo argumento (nome do arquivo de saída) é opcional — se omitido, gera `relatorio_git_log.xlsx` na pasta atual.

---

## 🧪 Resultados dos testes incluídos

| Período | Commits/PRs | Releases | Conflitos encontrados | Resultado |
|---|---|---|---|---|
| 1º semestre de 2025 | 657 | 43 | 0 | ✅ Segregação respeitada |
| 2º semestre de 2025 | 609 | 41 | 3 | ❌ Segregação violada |

No 2º semestre, o script identificou:
- 2 pessoas do grupo de desenvolvimento que fecharam release em algum momento;
- 1 pessoa do grupo de release que abriu um pull request.

Cada ocorrência vem com data, versão/PR envolvido e o nome do responsável — pronta para virar evidência de auditoria.

---

## 💡 Por que isso importa

Esse case ilustra como rotinas de auditoria de TI podem sair do modelo 100% manual quando o objeto de análise é estruturado (logs, planilhas, exportações de sistema). Os ganhos:

- **Elimina o erro humano** de comparar centenas de nomes manualmente;
- **Reduz drasticamente o tempo de execução** — de horas de revisão para segundos de processamento;
- **Gera evidência documentada e replicável**, pronta para o papel de trabalho da auditoria;
- **É reaplicável** em qualquer novo ciclo de desenvolvimento, virando um controle contínuo.

---

## 📄 Licença

Este projeto é disponibilizado para fins educacionais e de referência. Os dados de log incluídos são **fictícios**, gerados para fins de demonstração.
