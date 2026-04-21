# Code Review: Tarefas Sugeridas

## 1) Corrigir erro de digitação (typo)
**Problema encontrado:** O `README.md` orienta instalar dependências com `pip install -r requirements.txt`, mas o repositório possui `requirements/prod.txt` e `requirements/dev.txt`.

**Tarefa sugerida:** Atualizar o `README.md` para usar `requirements/prod.txt` (e opcionalmente documentar `requirements/dev.txt` para ambiente de desenvolvimento).

**Critério de aceite:** As instruções de instalação do README funcionam sem ajuste manual de caminho.

---

## 2) Corrigir bug funcional
**Problema encontrado:** O método `MRPGUI._load_table` assume por padrão o arquivo `itens_criticos.xlsx`, mas `_execute_analysis` gera arquivo com timestamp (`itens_criticos_YYYYMMDD_HHMMSS.xlsx`). Isso pode quebrar o botão de recarregar/fluxo padrão quando nenhum caminho explícito é passado.

**Tarefa sugerida:** Persistir `last_analysis_file` em `AppState` após análise bem-sucedida e usar esse caminho como fonte prioritária no `_load_table` (com fallback seguro).

**Critério de aceite:** Após rodar uma análise, a ação de recarregar tabela funciona consistentemente sem depender de nome fixo de arquivo.

---

## 3) Ajustar comentário de código / discrepância de documentação
**Problema encontrado:** A docstring de `_run_analysis` diz que a análise roda em “separate thread”, mas a implementação usa `root.after(...)`, que apenas agenda execução no loop principal do Tkinter.

**Tarefa sugerida:** Ou (a) atualizar docstring/comentários para refletir comportamento real (agendamento no event loop), ou (b) mover execução para thread de fato e sincronizar atualização de UI com segurança.

**Critério de aceite:** Documentação e implementação ficam consistentes entre si.

---

## 4) Melhorar teste automatizado
**Problema encontrado:** A suíte atual (`tests/unit/test_mrp_analyzer.py`) cobre apenas inicialização/configuração e não valida as regras de negócio (cálculo de estoque disponível, filtro de itens críticos e quantidade a solicitar).

**Tarefa sugerida:** Criar testes unitários parametrizados para `MRPAnalyzer` cobrindo:
- cálculo de `ESTOQUE DISPONÍVEL` (`ESTQ10 + ESTQ20/3`),
- regra de filtro de criticidade,
- arredondamento/clamp em `QUANTIDADE A SOLICITAR`,
- cenários com valores-limite (zero, borda, ausência de críticos).

**Critério de aceite:** Novos testes falham ao introduzir regressão nas fórmulas/regras e passam no comportamento esperado.
