🛠️ OS Manager Pro
Sistema de Gerenciamento de Ordens de Serviço
📌 Visão Geral

O OS Manager Pro é uma aplicação web desenvolvida em Python com Flask, criada para otimizar o gerenciamento de Ordens de Serviço (OS).
O sistema permite a importação de dados via planilhas Excel, centralizando as informações em um banco de dados SQLite, garantindo organização, padronização e facilidade de consulta.

Este projeto foi desenvolvido com foco em boas práticas de backend, persistência de dados e aplicações funcionais de uso real, sendo ideal tanto para fins acadêmicos quanto para portfólio profissional.

🎯 Objetivos do Projeto

Automatizar o cadastro de Ordens de Serviço

Reduzir erros manuais através da importação de planilhas

Centralizar informações operacionais em banco de dados

Aplicar conceitos sólidos de desenvolvimento web backend

Demonstrar habilidades técnicas em um projeto completo e funcional

🚀 Funcionalidades

📄 Gerenciamento completo de Ordens de Serviço

📥 Upload e leitura de arquivos Excel (.xlsx)

🗄️ Persistência de dados utilizando SQLite

🌐 Interface web simples e responsiva

⚙️ Execução local via Flask

🧰 Tecnologias Utilizadas
Tecnologia	Descrição
Python	Linguagem principal
Flask	Framework web backend
SQLite	Banco de dados relacional
HTML5	Estrutura da interface
Tailwind CSS	Estilização responsiva
Excel (XLSX)	Importação de dados
📂 Estrutura do Projeto
📦 os-manager-pro
 ┣ 📜 app_desktop.py              # Aplicação principal Flask
 ┣ 📜 ordens_servico_completo.db  # Banco de dados SQLite
 ┣ 📂 uploads                     # Arquivos Excel enviados
 ┃ ┗ 📜 modelo_planilha.xlsx
 ┣ 📄 Tutorial_de_Uso_OS_Manager_Pro.docx
 ┣ 📜 README.md

⚙️ Como Executar o Projeto
🔧 Pré-requisitos

Python 3.8+

Pip

▶️ Passo a passo
# Clone o repositório
git clone https://github.com/seu-usuario/os-manager-pro.git

# Acesse a pasta
cd os-manager-pro

# Instale as dependências
pip install flask

# Execute a aplicação
python app_desktop.py


A aplicação ficará disponível em:

http://localhost:5000

🧠 Conceitos Técnicos Aplicados

Desenvolvimento de aplicações web com Flask

Manipulação e validação de arquivos

Leitura e importação de dados Excel

Persistência de dados com SQLite

Organização de código backend

Estruturação de projetos para portfólio

📈 Evoluções Planejadas

🔐 Autenticação e controle de usuários

🧱 Arquitetura MVC

🧪 Testes automatizados

☁️ Deploy em ambiente de produção

🗃️ Migração para banco de dados escalável (PostgreSQL)

👤 Autor

Eduardo Felype Liberal Santos
🎓 Engenharia de Software
💻 Desenvolvedor em formação
