# Repositório de Apresentações Executivas

Este repositório tem **uma única finalidade**: criar apresentações PowerPoint executivas de alta qualidade.

## Comportamento automático

**Qualquer mensagem com um tema, assunto ou pedido neste repositório é automaticamente um pedido de criação de apresentação PowerPoint.**

Não espere o usuário digitar `/ppt`. Assim que receber um tema, execute o processo completo:

1. Leia integralmente `.claude/commands/ppt.md` antes de começar
2. Interprete o tema como o assunto da apresentação
3. Gere o script Python, execute-o e entregue o arquivo `.pptx` pronto
4. Execute o checklist de qualidade antes de entregar

## Exemplos de ativação

| Mensagem do usuário | O que fazer |
|---------------------|-------------|
| "Cash In e Run Rate" | Criar apresentação explicando Cash In e Run Rate |
| "Resultados Q1 2026 da área de crédito" | Criar apresentação de resultados |
| "Onboarding do novo time de Change" | Criar apresentação de onboarding |
| "Estratégia de crédito para o semestre" | Criar apresentação estratégica |
| Qualquer outro tema | Criar apresentação sobre esse tema |

## O que NÃO fazer

- Não pergunte "Você quer que eu crie uma apresentação?" — neste repositório, **sempre** é uma apresentação
- Não aguarde confirmação de estrutura — vá direto ao `.pptx` (a menos que o usuário peça revisão prévia)
- Não use conteúdo, nomes de seções ou dados de apresentações anteriores como exemplo

## Referência técnica

Todas as diretrizes de design, código Python, layouts e checklist de qualidade estão em:

```
.claude/commands/ppt.md
```
