#########Analise das variaveis categoricas######################################

# Setup ------------------------------------------------------------------------
library(tidyverse)
require(gtsummary)           # Tabelas automáticas
require(gt)                  # Tabelas automáticas
require(rstatix)             # Coeficiente de Cramer

df <- openxlsx::read.xlsx("Avaliacao_6/data/Nutricao.xlsx") %>%  select(1:20, "Doces")
df_quali <- df %>% select(1:15, "Doces")

# formatação em português (vírgula pra decimais e ponto para milhares)
theme_gtsummary_language("pt", big.mark = ".", decimal.mark = ",")

# Tabelas de contingencia -----------------------------------------------------
tbl_summary(data = df_quali)

tbl_summary(data = df_quali,
            by = Doces)

tbl_summary(data = df_quali,
            by = Doces ,
            percent = "row")
# Verificação da Matriz Esperada ----------------------------------------------
chisq.test(df_quali$Genero,df_quali$Doces)$expected

chisq.test(df_quali$Raca_cor,df_quali$Doces)$expected # casela menor que 5

chisq.test(df_quali$Regiao,df_quali$Doces)$expected

chisq.test(df_quali$Isolamento,df_quali$Doces)$expected

chisq.test(df_quali$Trabalha,df_quali$Doces)$expected

chisq.test(df_quali$Profissional_Saude,df_quali$Doces)$expected

chisq.test(df_quali$Renda_familiar,df_quali$Doces)$expected

chisq.test(df_quali$Escolaridade,df_quali$Doces)$expected # casela menor que 5

chisq.test(df_quali$Covid,df_quali$Doces)$expected

chisq.test(df_quali$Consulta_Nutricionista,df_quali$Doces)$expected

chisq.test(df_quali$Dificuldade_Financeira,df_quali$Doces)$expected

chisq.test(df_quali$Acesso_Alimento,df_quali$Doces)$expected

chisq.test(df_quali$Tempo_Preparo_Refeicao,df_quali$Doces)$expected

chisq.test(df_quali$cigarro,df_quali$Doces)$expected

chisq.test(df_quali$atividade_fisica_pandemia,df_quali$Doces)$expected

# Tabela com o pvalor ----------------------------------------------------------
tbl_summary(
  data = df_quali,
  by = Doces,
  percent = "row"
) %>%
  add_p(
    test = list(
      Escolaridade ~ "fisher.test",
      Raca_cor     ~ "fisher.test",
      everything() ~ "chisq.test"
    )
  ) %>% 
  bold_p(t = 0.05)

# Residuos ---------------------------------------------------------------------
# acima de 1,96 ou abaixo de -1,96 na normal padrao sao os influentes
chisq.test(df_quali$Genero,df_quali$Doces)$stdres

chisq.test(df_quali$Regiao,df_quali$Doces)$stdres

chisq.test(df_quali$Isolamento,df_quali$Doces)$stdres

chisq.test(df_quali$Trabalha,df_quali$Doces)$stdres

chisq.test(df_quali$Profissional_Saude,df_quali$Doces)$stdres

chisq.test(df_quali$Renda_familiar,df_quali$Doces)$stdres

chisq.test(df_quali$Escolaridade,df_quali$Doces)$stdres

chisq.test(df_quali$Consulta_Nutricionista,df_quali$Doces)$stdres

chisq.test(df_quali$Dificuldade_Financeira,df_quali$Doces)$stdres

chisq.test(df_quali$Acesso_Alimento,df_quali$Doces)$stdres

chisq.test(df_quali$Tempo_Preparo_Refeicao,df_quali$Doces)$stdres

chisq.test(df_quali$cigarro,df_quali$Doces)$stdres

chisq.test(df_quali$atividade_fisica_pandemia,df_quali$Doces)$stdres 

# Coeficiente de Cramer --------------------------------------------------------
cramer_fun <- function(data, variable, by, ...) {
  tab <- table(data[[variable]], data[[by]])
  v <- cramer_v(tab)
  tibble::tibble(`**Cramér**` = round(v, 3))
}

# Criando tabela --------------------------------------------------------------

## inicio tabela ----

tabela <- tbl_summary(
  data = df_quali,
  by = Doces,
  percent = "row",
  label = list(
    Genero                    ~ "Gênero<sup>Q</sup>",
    Raca_cor                  ~ "Raça/Cor<sup>F</sup>",
    Regiao                    ~ "Região<sup>Q</sup>",
    Isolamento                ~ "Isolamento<sup>Q</sup>",
    Trabalha                  ~ "Trabalha<sup>Q</sup>",
    Profissional_Saude        ~ "Profissional de Saúde<sup>Q</sup>",
    Renda_familiar            ~ "Renda Familiar<sup>Q</sup>",
    Escolaridade              ~ "Escolaridade<sup>F</sup>",
    Consulta_Nutricionista    ~ "Consulta ao Nutricionista<sup>Q</sup>",
    Dificuldade_Financeira    ~ "Dificuldade Financeira<sup>Q</sup>",
    Acesso_Alimento           ~ "Acesso ao Alimento<sup>Q</sup>",
    Tempo_Preparo_Refeicao    ~ "Tempo de Preparo da Refeição<sup>Q</sup>",
    cigarro                   ~ "Cigarro<sup>Q</sup>",
    atividade_fisica_pandemia ~ "Atividade Física na Pandemia<sup>Q</sup>"
  )
) %>%
  add_p(
    test = list(
      Escolaridade ~ "fisher.test",
      Raca_cor     ~ "fisher.test",
      everything() ~ "chisq.test"
    )
  ) %>% 
  add_stat(fns = everything() ~ cramer_fun)%>%
  modify_spanning_header(all_stat_cols() ~ "**Mudança no consumo de doces**") %>%
  modify_header(label ~ "**Variáveis**") %>%
  bold_labels() %>%
  modify_header(all_stat_cols() ~ "**{level}**<br>{n} ({style_percent(p)}%)")%>%
  bold_p(t = 0.05) %>% 
  modify_footnote(everything() ~ NA) %>%
  
  ## negrito nas caselas ----
  # stat_1 = Aumentou
  modify_table_styling(columns = stat_1, text_format = "bold",
                       rows = (variable == "Genero"                 & label %in% c("Feminino", "Masculino"))             |
                         (variable == "Regiao"                 & label %in% c("Centro-oeste", "Norte", "Sudeste"))   |
                         (variable == "Isolamento"             & label %in% c("Não", "Sim"))                         |
                         (variable == "Profissional_Saude"     & label %in% c("Não", "Sim"))                         |
                         (variable == "Renda_familiar"         & label == "entre R$ 1.255 - R$ 8.640")               |
                         (variable == "Escolaridade"           & label %in% c("Ensino Médio Completo", "Pós-graduação")) |
                         (variable == "Dificuldade_Financeira" & label %in% c("Não", "Sim"))                         |
                         (variable == "Tempo_Preparo_Refeicao" & label %in% c("Diminuiu", "Não alterou"))            |
                         (variable == "atividade_fisica_pandemia" & label %in% c("Aumentou", "Diminuiu", "Não alterou"))
  ) %>%
  
  # stat_2 = Diminuiu
  modify_table_styling(columns = stat_2, text_format = "bold",
                       rows = (variable == "Trabalha"               & label %in% c("Não", "Sim"))                         |
                         (variable == "Profissional_Saude"     & label %in% c("Não", "Sim"))                         |
                         (variable == "Renda_familiar"         & label == "até R$ 1254,00")                           |
                         (variable == "Escolaridade"           & label %in% c("Ensino Médio Completo", "Pós-graduação")) |
                         (variable == "Dificuldade_Financeira" & label %in% c("Não", "Sim"))                         |
                         (variable == "Acesso_Alimento"        & label %in% c("Não", "Sim"))                         |
                         (variable == "atividade_fisica_pandemia" & label %in% c("Aumentou", "Diminuiu"))
  ) %>%
  
  # stat_3 = Não alterou
  modify_table_styling(columns = stat_3, text_format = "bold",
                       rows = (variable == "Genero"                 & label %in% c("Feminino", "Masculino"))              |
                         (variable == "Regiao"                 & label == "Centro-oeste")                             |
                         (variable == "Isolamento"             & label %in% c("Não", "Sim"))                         |
                         (variable == "Trabalha"               & label %in% c("Não", "Sim"))                         |
                         (variable == "Renda_familiar"         & label == "mais de R$ 8.640")                         |
                         (variable == "Escolaridade"           & label %in% c("Ensino Médio Completo", "Pós-graduação")) |
                         (variable == "Consulta_Nutricionista" & label %in% c("Não", "Sim"))                         |
                         (variable == "Dificuldade_Financeira" & label %in% c("Não", "Sim"))                         |
                         (variable == "Acesso_Alimento"        & label %in% c("Não", "Sim"))                         |
                         (variable == "Tempo_Preparo_Refeicao" & label %in% c("Diminuiu", "Não alterou"))            |
                         (variable == "atividade_fisica_pandemia" & label %in% c("Diminuiu", "Não alterou"))
  ) %>%
  
  # stat_4 = Não consumo
  modify_table_styling(columns = stat_4, text_format = "bold",
                       rows = (variable == "Genero"                 & label %in% c("Feminino", "Masculino"))              |
                         (variable == "Regiao"                 & label == "Centro-oeste")                             |
                         (variable == "Consulta_Nutricionista" & label %in% c("Não", "Sim"))                         |
                         (variable == "cigarro"                & label %in% c("Não", "Sim"))
  )

## rmd ----

tabela <- tabela %>%
  as_gt() %>%
  tab_style(
    style = cell_fill(color = "#fef9c3"),
    locations = cells_body(
      rows = p.value < 0.05 & variable != "cigarro"
    )
  ) %>%
  tab_style(
    style = cell_fill(color = "#fefce8"),
    locations = cells_body(
      rows = p.value < 0.05 & variable == "cigarro"
    )
  ) %>%
  tab_style(
    style = gt::cell_borders(
      sides = "bottom",
      color = "#7a7a7a",
      weight = gt::px(1.5) 
    ),
    locations = gt::cells_body(
      rows = !duplicated(variable, fromLast = TRUE)
    )
  ) %>% 
  gt::fmt_markdown(
    columns = "label",
    rows = !grepl("R\\$", label)
  )

## nota roda pé ----

tabela_categorica <- tabela %>%
  tab_source_note(
  source_note = md("**Q**: Teste do qui-quadrado de Pearson; <br>
                     **F**: Teste exato de Fisher; <br>
                     Valores em negrito na coluna *Valor-p* indicam significância estatística (p < 0,05); <br>
                     Valores em negrito nas tabelas de contingência indicam células com resíduos padronizados elevados.")
) %>% 
  tab_options(
    source_notes.font.size = "13px"
  ) %>%
  tab_style(
    style = cell_text(color = "#666666"),
    locations = cells_source_notes()
  )
tabela_categorica

## salvar ----

# saveRDS(tabela_categorica, file = "Avaliacao_6/data/tabela_categorica.rds")

## Residuos coloridos ---------------------------------------------------------

cor_pos <- "#1e3a5f"  # resíduo positivo
cor_neg <- "#7c2d12"  # resíduo negativo

tabela_categorica_colorida <- tabela_categorica %>%
  # ── stat_1 = Aumentou ──────────────────────────────────────────────────────
  
  tab_style(
    style = cell_text(color = cor_pos, weight = "bold"),
    locations = cells_body(columns = stat_1, rows =
                             (variable == "Genero"                    & label == "Feminino")                   |
                             (variable == "Regiao"                    & label == "Sudeste")                    |
                             (variable == "Isolamento"                & label == "Sim")                        |
                             (variable == "Profissional_Saude"        & label == "Sim")                        |
                             (variable == "Renda_familiar"            & label == "entre R$ 1.255 - R$ 8.640") |
                             (variable == "Escolaridade"              & label == "Ensino Médio Completo")      |
                             (variable == "Dificuldade_Financeira"    & label == "Não")                        |
                             (variable == "Tempo_Preparo_Refeicao"    & label == "Diminuiu")                  |
                             (variable == "atividade_fisica_pandemia" & label == "Diminuiu")
    )
  ) %>%
  
  tab_style(
    style = cell_text(color = cor_neg, weight = "bold"),
    locations = cells_body(columns = stat_1, rows =
                             (variable == "Genero"                    & label == "Masculino")                      |
                             (variable == "Regiao"                    & label %in% c("Centro-oeste", "Norte"))     |
                             (variable == "Isolamento"                & label == "Não")                            |
                             (variable == "Profissional_Saude"        & label == "Não")                            |
                             (variable == "Escolaridade"              & label == "Pós-graduação")                  |
                             (variable == "Dificuldade_Financeira"    & label == "Sim")                            |
                             (variable == "Tempo_Preparo_Refeicao"    & label == "Não alterou")                    |
                             (variable == "atividade_fisica_pandemia" & label %in% c("Aumentou", "Não alterou"))
    )
  ) %>%
  
  # ── stat_2 = Diminuiu ──────────────────────────────────────────────────────
  
  tab_style(
    style = cell_text(color = cor_pos, weight = "bold"),
    locations = cells_body(columns = stat_2, rows =
                             (variable == "Trabalha"                  & label == "Não")                        |
                             (variable == "Profissional_Saude"        & label == "Não")                        |
                             (variable == "Renda_familiar"            & label == "até R$ 1254,00")             |
                             (variable == "Escolaridade"              & label == "Ensino Médio Completo")      |
                             (variable == "Dificuldade_Financeira"    & label == "Sim")                        |
                             (variable == "Acesso_Alimento"           & label == "Sim")                        |
                             (variable == "atividade_fisica_pandemia" & label == "Aumentou")
    )
  ) %>%
  
  tab_style(
    style = cell_text(color = cor_neg, weight = "bold"),
    locations = cells_body(columns = stat_2, rows =
                             (variable == "Trabalha"                  & label == "Sim")                        |
                             (variable == "Profissional_Saude"        & label == "Sim")                        |
                             (variable == "Escolaridade"              & label == "Pós-graduação")              |
                             (variable == "Dificuldade_Financeira"    & label == "Não")                        |
                             (variable == "Acesso_Alimento"           & label == "Não")                        |
                             (variable == "atividade_fisica_pandemia" & label == "Diminuiu")
    )
  ) %>%
  
  # ── stat_3 = Não alterou ───────────────────────────────────────────────────
  
  tab_style(
    style = cell_text(color = cor_pos, weight = "bold"),
    locations = cells_body(columns = stat_3, rows =
                             (variable == "Genero"                    & label == "Masculino")                  |
                             (variable == "Regiao"                    & label == "Centro-oeste")               |
                             (variable == "Isolamento"                & label == "Não")                        |
                             (variable == "Trabalha"                  & label == "Sim")                        |
                             (variable == "Renda_familiar"            & label == "mais de R$ 8.640")           |
                             (variable == "Escolaridade"              & label == "Pós-graduação")              |
                             (variable == "Consulta_Nutricionista"    & label == "Não")                        |
                             (variable == "Dificuldade_Financeira"    & label == "Não")                        |
                             (variable == "Acesso_Alimento"           & label == "Não")                        |
                             (variable == "Tempo_Preparo_Refeicao"    & label == "Não alterou")               |
                             (variable == "atividade_fisica_pandemia" & label == "Não alterou")
    )
  ) %>%
  
  tab_style(
    style = cell_text(color = cor_neg, weight = "bold"),
    locations = cells_body(columns = stat_3, rows =
                             (variable == "Genero"                    & label == "Feminino")                   |
                             (variable == "Isolamento"                & label == "Sim")                        |
                             (variable == "Trabalha"                  & label == "Não")                        |
                             (variable == "Escolaridade"              & label == "Ensino Médio Completo")      |
                             (variable == "Consulta_Nutricionista"    & label == "Sim")                        |
                             (variable == "Dificuldade_Financeira"    & label == "Sim")                        |
                             (variable == "Acesso_Alimento"           & label == "Sim")                        |
                             (variable == "Tempo_Preparo_Refeicao"    & label == "Diminuiu")                  |
                             (variable == "atividade_fisica_pandemia" & label == "Diminuiu")
    )
  ) %>%
  
  # ── stat_4 = Não consumo ───────────────────────────────────────────────────
  
  tab_style(
    style = cell_text(color = cor_pos, weight = "bold"),
    locations = cells_body(columns = stat_4, rows =
                             (variable == "Genero"                 & label == "Masculino")   |
                             (variable == "Regiao"                 & label == "Centro-oeste")|
                             (variable == "Consulta_Nutricionista" & label == "Sim")         |
                             (variable == "cigarro"                & label == "Sim")
    )
  ) %>%
  
  tab_style(
    style = cell_text(color = cor_neg, weight = "bold"),
    locations = cells_body(columns = stat_4, rows =
                             (variable == "Genero"                 & label == "Feminino")    |
                             (variable == "Consulta_Nutricionista" & label == "Não")         |
                             (variable == "cigarro"                & label == "Não")
    )
  )
tabela_categorica_colorida

### salvar ----

# saveRDS(tabela_categorica_colorida, file = "Avaliacao_6/data/tabela_categorica_colorida.rds")
