# 📊 GitHub VBA-Excel Framework

Sistema desenvolvido em **VBA para Excel** que implementa um ambiente completo com interface gráfica, painel administrativo, segurança de usuários, geração automática de documentação e ferramentas para desenvolvedores VBA.

Este projeto demonstra como o **VBA pode ser utilizado para construir aplicações estruturadas dentro do Excel**, com recursos normalmente vistos em softwares desktop.

---

# 🚀 Principais Recursos

### 🔐 Segurança

* Sistema de **login de usuários**
* **Hash de senha**
* **Criptografia de dados**
* **Verificação de integridade do arquivo**
* Usuário **administrador interno**

### 🖥 Interface

* **Splash Screen animada**
* **Painel administrativo**
* **Tela de cadastro de usuários**
* **Tela de informações do desenvolvedor**
* **Animações de interface**

### 🧰 Ferramentas para Desenvolvedores

* Gerador automático de **documentação do projeto VBA**
* Exportação de **dados Excel → TXT**
* Importação de **dados TextBox → Planilha**
* Sistema de **Easter Egg com códigos bônus**

---

# 🧩 Estrutura do Projeto

## Forms

* `form_splash_screen`
* `form_projeto`
* `form_painel_de_controle`
* `form_cadastro`
* `form_conceito`
* `form_gerador_vba`
* `form_informacao_desenvolvedor`

---

## Módulos

* `gerador_hash`
* `gerador_criptografia`
* `gerar_animation`
* `gerar_bonus`
* `gerar_documentacao`

---

## Planilhas

* `workbook_dashboard`
* `plan_dashboard`

---

# ⚙ Funcionalidades do Sistema

## 🔑 Sistema de Login

O sistema possui autenticação baseada em:

* Usuário administrador interno
* Usuários armazenados em arquivo criptografado

Arquivo utilizado:

```
usuarios.dat
```

Cada registro contém:

```
usuario | hash_da_senha
```

---

## 🔐 Criptografia e Integridade

O sistema utiliza:

* **criptografia simples para armazenamento**
* **assinatura de integridade do arquivo**

Se o arquivo for alterado manualmente, o sistema bloqueia o acesso.

---

## 📄 Gerador de Documentação VBA

O projeto possui um módulo que analisa automaticamente o código VBA e gera um relatório contendo:

* módulos
* forms
* subs
* functions

Exemplo de saída:

```
DOCUMENTAÇÃO AUTOMÁTICA DO PROJETO VBA
Arquivo: Projeto GitHub VBA-Excel v1.xlsm
Gerado em: 12/03/2026
```

---

# 🎁 Easter Egg

Ao visitar os perfis do desenvolvedor:

* LinkedIn
* GitHub

um **frame oculto é desbloqueado** contendo códigos VBA bônus.

---

# 📦 Como usar

1️⃣ Baixe o arquivo `.xlsm`

2️⃣ Habilite macros no Excel

3️⃣ Execute o dashboard inicial

4️⃣ Utilize o painel administrativo para acessar os recursos

---

# 🛠 Tecnologias utilizadas

* VBA (Visual Basic for Applications)
* Microsoft Excel
* Windows API (para animações)

---

# 👨‍💻 Desenvolvedor

**Guilherme Mesquita Machado**

📧 Email
[gmmgmmgoias@gmail.com](mailto:gmmgmmgoias@gmail.com)

🔗 LinkedIn
https://linkedin.com/in/guilherme-mesquita-machado-07664b2b9

💻 GitHub
https://github.com/gmmsystem

---

# 📜 Licença

Este projeto está disponível para fins de estudo e demonstração de técnicas em VBA.
