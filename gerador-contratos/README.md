# ContratoFácil — Gerador de Contratos

Aplicativo web (roda 100% no navegador, sem servidor) para gerar contratos em **Word (.docx)** e **PDF** a partir de modelos do Word, mantendo a formatação original.

## Como usar

- **Online:** publique a pasta `gerador-contratos/` em qualquer hospedagem estática (GitHub Pages, Netlify, Vercel). Com https, o app funciona offline e pode ser instalado como aplicativo.
- **Arquivo único:** `ContratoFacil.html` tem o sistema inteiro embutido — basta copiar para o computador e abrir com dois cliques (Chrome/Edge). É gerado automaticamente pelo `npm run build`.
- **Local:** abra `index.html` no Chrome/Edge (ou rode `python3 -m http.server` dentro da pasta).

## O que já vem pronto

- Modelo **Locação de Bem Móvel sem Operação (Rent Loc)** com 53 campos configurados: valores por extenso, término calculado pelo prazo, valor global = soma dos equipamentos × prazo, número automático do contrato (0063/2026…).
- Tabela de equipamentos com **várias linhas** (a linha se repete para cada equipamento).
- Busca de **CNPJ** (razão social, endereço e sócio-administrador) e **CEP**.
- Edição de **cláusulas** direto no documento: editar texto (negrito/itálico/sublinhado), inserir, duplicar, mover, excluir, mudar o tipo — a numeração (1.1, §1º, a) …) se refaz sozinha.
- Edição do **modelo**: importar qualquer .docx, **detectar campos automaticamente** (CPF, CNPJ, CEP, valores, datas, nomes das partes) ou selecionar um texto e “Transformar em campo”.
- Contratos salvos, status (rascunho/finalizado/assinado…), **vencimentos** no painel inicial, duplicar para renovação, histórico, desfazer/refazer, salvamento automático.
- **Cadastros** reutilizáveis (clientes, equipamentos) e **Biblioteca de cláusulas**.
- Backup/restauração em arquivo, tema claro/escuro.

## Desenvolvimento

```bash
npm install
npm run build    # gera app.js, app.css e sw.js
npm run watch    # recompila ao salvar
```

Código-fonte em `src/`:

| Pasta | Conteúdo |
|---|---|
| `src/docx/` | Motor DOCX: leitura do pacote, estilos, numeração, renderização fiel em páginas, paginação, edição de parágrafos, exportação |
| `src/fields/` | Tipos de campo, fórmulas, extensos, datas calculadas |
| `src/br/` | Extenso em português, máscaras/validação CPF/CNPJ/CEP, consultas BrasilAPI/ViaCEP |
| `src/ui/` | Interface (Preact + htm): páginas, editor, diálogos |
| `src/seed/` | Modelo inicial Rent Loc (.docx + definição dos campos) |
| `src/store/` | Armazenamento local (IndexedDB) e backup |

Os dados ficam somente no navegador do usuário (IndexedDB). Recomende backups periódicos em **Ajustes › Backup**.
