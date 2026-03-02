# excel-vba-turno-automacao
Sistema de automação em Excel VBA para geração automatizada de turnos diários com registro de auditoria e controle estrutural via ocultação programática de planilhas.

 _Sistema de Criação de Turnos em Excel (VBA)_

 
 👨‍💻 Autor:
 Carlos Gomes |
_Projeto desenvolvido para fins de automação corporativa e portfólio técnico._


**Descrição**

Projeto desenvolvido em VBA para Excel com o objetivo de automatizar a criação diária de planilhas de turno e registrar automaticamente o histórico de criação (log de auditoria).

**O sistema permite:**

-> Criar automaticamente uma nova planilha com base em um modelo padrão

-> Nomear a planilha com a data atual do Windows

-> Registrar usuário, data e hora de criação

-> Ocultar planilhas estruturais (modelo e log) para proteção do sistema

-> Controle administrativo via VBA

**Arquitetura do Sistema**

O projeto é composto por três planilhas principais:

Menu → Interface do usuário

MODELO_TURNO → Template base (VeryHidden)

LOG → Histórico de criação (VeryHidden)

**Fluxo do processo:**
Usuário clica no botão
↓
Sistema verifica se já existe turno do dia
↓
Copia modelo
↓
Renomeia com data atual
↓
Registra log (usuário, data, hora)

**Recursos de Segurança**

-> Uso de "xlSheetVeryHidden"

-> Modularização do código

-> Separação de responsabilidades (criação e log)

-> Controle de visibilidade temporária para cópia segura

**Conceitos Aplicados:**

-> Manipulação de objetos Worksheet

-> Controle de erros

-> Funções privadas

-> Modularização

-> Registro de auditoria (Audit Trail)

-> Automação de processos internos

**Aplicabilidade Empresarial**

_Este sistema pode ser aplicado em:_

-> Indústrias

-> Empresas com controle de turno operacional

-> Centros logísticos

-> Hospitais

-> Empresas com gestão diária de equipe
