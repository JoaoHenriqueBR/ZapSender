# ZapSender - Automação de Mensagens e Mídia para WhatsApp

![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)

> **Status do Projeto:** Ativo (v1.0.0)  
> **Aviso Legal:** Este projeto é uma ferramenta independente e **não possui afiliação, associação ou autorização** da Meta Platforms, Inc. ou do WhatsApp. O uso deve respeitar os Termos de Serviço da plataforma.

O **ZapSender** é um conjunto de scripts em Python projetado para automação inteligente de mensagens, permitindo o envio em massa com foco em personalização, tratamento de dados e segurança (anti-spam).

---

## 🚀 Funcionalidades Principais

* **Multimídia:** Envio automático de textos, áudios e imagens.
* **Inteligência de Dados:** Leitura e tratamento de planilhas de dados.
* **Validação:** Verificação de números ativos antes do envio.
* **Anti-Spam Nativo:** Sistema de *cooldown* (tempo de espera) totalmente customizável.
* **Personalização Dinâmica:** Saudações, emojis aleatórios e múltiplas versões da mesma mensagem para evitar detecção.
* **Feedback Responsivo:** Acompanhamento em tempo real do processo de envio.

---

## 🛠️ Especificações Técnicas (Setup)

O ZapSender opera como uma aplicação de console (CLI) interativa via Selenium.

### Requisitos do Sistema
* **Navegador:** Google Chrome ou Chromium instalado.
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

## 💰 Investimento e Suporte

Se gostou do meu trabalho e deseja vê-lo evoluir, poderá sempre apoiar através de uma doação, que será usado para investimentos e sustentabilidade do projeto:

**Chave PIX (E-mail):** `joaohenrique.jh103@protonmail.com`


**Atenção:** Estou trabalhando em benefícios exclusivos e uma interface gráfica (GUI) para apoiadores e contribuidores, mais detalhes em breve!


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

A existência de uma versão comercial para Apoiador(GUI/Premium) serve exclusivamente para financiar o desenvolvimento contínuo, suporte e infraestrutura do ecossistema **ZapSender.**

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
