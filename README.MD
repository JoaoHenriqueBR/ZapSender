# ZapSender - Automação Assistida de Mensagens e Mídia para WhatsApp Web

![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)

> **Status do Projeto:** Ativo (v1.0.0)  
> **Aviso Legal:** Este projeto é uma ferramenta independente e **não possui afiliação, associação ou autorização** da Meta Platforms, Inc. ou do WhatsApp.

O **ZapSender** é uma aplicação CLI em Python que auxilia um operador no envio de mensagens personalizadas e anexos pelo **WhatsApp Web**, utilizando **Selenium**.

O fluxo atual é **assistido pelo usuário**: a pessoa executa o script, escolhe o modo de envio, informa os arquivos necessários, abre a sessão no navegador, escaneia o QR Code do WhatsApp Web e acompanha a execução.

O objetivo do projeto é atender fluxos legítimos de comunicação com contatos próprios, como escolas, cursos, equipes internas ou organizações que já possuem relação prévia com os destinatários.

---

## 🚀 Funcionalidades Principais

* **Envio assistido de mídia:** suporte a imagens, áudios e documentos com mensagens associadas.
* **Personalização de conteúdo:** uso de placeholders como nome do contato, saudações e variações de texto.
* **Leitura de planilhas:** carregamento de contatos a partir de arquivos Excel.
* **Rastreamento de processamento:** registro local de números já processados para evitar reenvios acidentais em execuções futuras.
* **Execução interativa:** configuração do fluxo diretamente pelo terminal.

---

## ✅ O que este projeto faz

- Abre uma sessão do WhatsApp Web em um navegador controlado por Selenium
- Lê dados de uma planilha Excel
- Monta mensagens personalizadas por contato
- Envia imagens, áudios ou documentos conforme o modo escolhido
- Mantém um controle local de números já processados

## 🚫 O que este projeto não faz

- Não é uma integração oficial com o WhatsApp
- Não concede permissão para contatar pessoas sem consentimento
- Não substitui APIs oficiais quando há necessidade de escala institucional, auditoria ou compliance formal
- Não elimina a responsabilidade do operador pelo uso adequado da ferramenta

Se o seu caso de uso exige operação institucional em larga escala, rastreabilidade centralizada, múltiplos operadores ou garantias formais de conformidade, o mais adequado é utilizar um provedor oficial ou uma API apropriada para esse cenário.

---

## 🛠️ Modelo de Operação Atual

Atualmente, o ZapSender funciona como uma automação **interativa** via Selenium.

O operador:

1. inicia o programa pelo terminal
2. escolhe o tipo de envio
3. informa a planilha e o arquivo de mídia
4. configura as mensagens
5. aguarda a abertura do navegador
6. escaneia o QR Code do WhatsApp Web
7. acompanha a execução

Ou seja, o projeto **não opera como serviço em segundo plano**, **não utiliza API oficial** e **depende de uma sessão autenticada do WhatsApp Web no navegador**.

---

## 🧩 Modos de Envio

O hub principal (`main.py`) permite escolher entre:

- **Imagens + Mensagens**
- **Audios + Mensagens**
- **Documentos + Mensagens**

Cada fluxo carrega contatos da planilha, prepara a mensagem e envia o conteúdo pelo navegador.

---

## 📝 Personalização de Mensagens

As mensagens configuradas no terminal podem usar placeholders como:

- `{nome}`
- `{escolha_saudacao}`
- `{escolha_pergunta}`
- `{escolha_emoji}`
- `{escolha_emoji2}`
- `{escolha_emoji3}`

Isso permite adaptar o conteúdo para cada destinatário dentro de um fluxo supervisionado pelo operador.

---

## 📊 Uso de Planilhas

Os contatos são lidos de um arquivo Excel, com colunas configuráveis para:

- número de celular
- nome do destinatário

O projeto também gera um arquivo local de controle com os números já processados, usando como base o nome da planilha.

---

## 🛠️ Especificações Técnicas (Setup)

O ZapSender opera como uma aplicação de console (CLI) interativa via Selenium.

### Requisitos do Sistema
* **Navegador:** Google Chrome instalado.
* **Dependências de SO:**
    * **Windows (Nativo):** Utiliza `pywin32` para manipulação de área de transferência.
    * **Linux (Experimental):** Requer o pacote `xclip` para suporte a imagens (`sudo apt install xclip`).
* **WhatsApp Web:** Requer login via escaneamento de QR Code na instância do navegador aberta pelo script.

### Como Executar
1. Clone o repositório.
2. Instale as dependências: `pip install -r requirements.txt`
3. Execute o hub: `python main.py`

### Dependências Principais
* `selenium` + `webdriver-manager` — Automação de navegador.
* `pandas` + `openpyxl` — Processamento de dados.
* `pillow` — Processamento de imagens.
* `python-dotenv` — Gestão de variáveis de ambiente.
* `requests` — Integrações e requisições HTTP.

---

## ⚖️ Uso Responsável

Antes de utilizar a ferramenta, certifique-se de que:

- existe relação legítima com os destinatários
- há consentimento quando exigido pela legislação ou pela política da sua instituição
- o conteúdo enviado é pertinente e esperado pelo destinatário
- o uso está em conformidade com os Termos de Serviço da plataforma e com normas de privacidade, como a LGPD

O uso desta ferramenta é de responsabilidade exclusiva do operador.

---


## 💰 Investimento e Suporte

Se gostou do meu trabalho e deseja vê-lo evoluir, poderá sempre apoiar através de uma doação, que será usado para investimentos e sustentabilidade do projeto:

**Chave PIX (E-mail):** `joaohenrique.jh103@protonmail.com`


**Atenção:** Estou trabalhando em benefícios exclusivo, como a criação do programa com uma **interface gráfica intuitiva (GUI)**, usando a **API oficial do WhatsApp** para apoiadores e contribuidores, mais detalhes em breve!


## 🤝 Contribuição e Ética

Contribuições no código são muito bem-vindas e geram benefícios diretos, dependendo da relevância da contribuição:

* **Acesso VIP:** Benefícios exclusivos da futura versão de apoiador.
* **Reconhecimento:** Nome listado nos Créditos e no Hall de Contribuidores.
* **Networking:** Recomendação técnica no LinkedIn e acesso ao nosso Discord.
* **Experiência:** Desenvolvimento prático em automações de alto nível.

Para contribuir com sua mudança:
1. Faça um Fork do projeto.
2. Assine o **CLA** (nosso bot solicitará no seu primeiro Pull Request).
3. Siga as diretrizes de estilo PEP 8.

*O **CLA (Contributor License Agreement)** é o que garante a segurança jurídica para que o projeto continue existindo e crescendo, protegendo os direitos do fundador e garantindo benefícios aos investidores*

## 🛡️ Compromisso com o Código Aberto (Open Source Guarantee)

A existência de uma versão comercial para Apoiador serve exclusivamente para financiar o desenvolvimento contínuo, suporte e infraestrutura do ecossistema **ZapSender.**

Eu, **João Henrique Alves Ferreira,** comprometo-me publicamente com os seguintes pontos:

1. **Núcleo Gratuito:** O motor principal desenvolvido de forma comunitária e as funcionalidades de envio de mensagens sempre terão uma versão gratuita e funcional sob a licença **AGPLv3.**

2. **Transparência:** Nenhuma funcionalidade essencial que já exista na versão _Community_ será removida para se tornar exclusiva de uma versão paga.

3. **Continuidade:** Caso o projeto comercial seja descontinuado, o código-fonte da versão _Community_ permanecerá disponível para a comunidade manter e evoluir.

---

**Nota sobre LGPD:** O uso desta ferramenta para coleta de dados ou spam sem consentimento é de responsabilidade exclusiva do usuário. O desenvolvedor não se responsabiliza por suspensões de contas ou penalidades aplicadas pela Meta Platforms, Inc.

---

## 👤 Fundador
* **João Henrique Alves Ferreira** - *Software Developer*

---

## 🌟 Apoiadores e Contribuidores 
*(Espaço reservado para aqueles que contribuirem ao projeto)*
