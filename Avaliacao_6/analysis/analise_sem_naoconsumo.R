# Setup ------------------------------------------------------------------------
library(tidyverse)
require(gtsummary)           # Tabelas automáticas
require(gt)                  # Tabelas automáticas
require(rstatix)             # Coeficiente de Cramer
require(qqplotr)             # Gráficos qqplot
require(DescTools)           # Teste de Levene

df <- openxlsx::read.xlsx("Avaliacao_6/data/Nutricao.xlsx") %>% 
  select(1:20, "Doces") %>% 
  filter(Doces != "Não consumo")

# formatação em português (vírgula pra decimais e ponto para milhares)
theme_gtsummary_language("pt", big.mark = ".", decimal.mark = ",")

# Doces ------------------------------------------------------------------------
library(ggplot2)
library(scales)

plot_porp_doce_1 <- ggplot(df, aes(x = Doces)) +
  geom_bar(aes(y = after_stat(prop), group = 1),
           fill = "#c07068") +
  scale_y_continuous(labels = percent_format()) +
  labs(title = "Mudança no consumo de doces (%)",
       x = NULL,
       y = "Proporção") +
  theme_classic()+
  theme(plot.title = element_text(hjust = 0.5))

# saveRDS(plot_porp_doce_1, file = "Avaliacao_6/data/plot_porp_doce_1.rds")

plot_porp_doce_2 <- ggplot(df %>% filter(Doces != "Não consumo"), aes(x = Doces)) +
  geom_bar(aes(y = after_stat(prop), group = 1),
           fill = "#c07068") +
  scale_y_continuous(labels = percent_format()) +
  labs(title = "Mudança no consumo de doces (%)",
       x = NULL,
       y = "Proporção") +
  theme_classic()+
  theme(plot.title = element_text(hjust = 0.5))

# saveRDS(plot_porp_doce_2, file = "Avaliacao_6/data/plot_porp_doce_2.rds")

# Qualitativa ------------------------------------------------------------------

df_quali <- df %>% select(1:15, "Doces")

## Tabelas de contingencia -----------------------------------------------------
tbl_summary(data = df_quali)

tbl_summary(data = df_quali,
            by = Doces)

tbl_summary(data = df_quali,
            by = Doces ,
            percent = "row")
## Verificação da Matriz Esperada ----------------------------------------------
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

### manipulações ----

glimpse(df_quali)

#### Raça

df_quali_sem_ind <- df_quali %>%
  mutate(Raca_cor = recode(Raca_cor,
                           "Indígena" = "Outro"))

chisq.test(df_quali_sem_ind$Raca_cor,df_quali_sem_ind$Doces)$expected # casela menor que 5

df_quali <- df_quali %>%
  mutate(Raca_cor = recode(Raca_cor,
                           "Indígena" = "Outro"))

#### Escolaridade

df_quali_sem_ana <- df_quali %>%
  filter(Escolaridade != "Analfabeto")

chisq.test(df_quali_sem_ana$Escolaridade,df_quali_sem_ana$Doces)$expected # casela menor que 5

df_quali <- df_quali %>%
  filter(Escolaridade != "Analfabeto")

## Tabela com o pvalor ----------------------------------------------------------
tbl_summary(
  data = df_quali,
  by = Doces,
  percent = "row"
) %>%
  add_p(
    test = list(
      Raca_cor     ~ "fisher.test",
      everything() ~ "chisq.test"
    )
  ) %>% 
  bold_p(t = 0.05)

## Residuos ---------------------------------------------------------------------
# acima de 1,96 ou abaixo de -1,96 na normal padrao sao os influentes
chisq.test(df_quali$Genero,df_quali$Doces)$stdres

chisq.test(df_quali$Raca_cor,df_quali$Doces)$stdres

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

## Coeficiente de Cramer --------------------------------------------------------
cramer_fun <- function(data, variable, by, ...) {
  tab <- table(data[[variable]], data[[by]])
  v <- cramer_v(tab)
  tibble::tibble(`**Cramér**` = round(v, 3))
}

## Criando tabela --------------------------------------------------------------

### inicio tabela ----

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
  
  ### negrito nas caselas ----
# stat_1 = Aumentou
modify_table_styling(columns = stat_1, text_format = "bold",
                     rows =
                       (variable == "Genero"                  & label %in% c("Feminino", "Masculino"))                          |
                       (variable == "Raca_cor"                & label %in% c("Branca", "Parda"))                                |
                       (variable == "Regiao"                  & label == "Norte")                                               |
                       (variable == "Isolamento"              & label %in% c("Não", "Sim"))                                     |
                       (variable == "Profissional_Saude"      & label %in% c("Não", "Sim"))                                     |
                       (variable == "Renda_familiar"          & label %in% c("até R$ 1254,00", "entre R$ 1.255 - R$ 8.640"))   |
                       (variable == "Escolaridade"            & label == "Ensino Médio Completo")                               |
                       (variable == "Dificuldade_Financeira"  & label %in% c("Não", "Sim"))                                     |
                       (variable == "Tempo_Preparo_Refeicao"  & label %in% c("Diminuiu", "Não alterou"))                       |
                       (variable == "atividade_fisica_pandemia" & label %in% c("Aumentou", "Diminuiu", "Não alterou"))
) %>%
  
  # stat_2 = Diminuiu
  modify_table_styling(columns = stat_2, text_format = "bold",
                       rows =
                         (variable == "Raca_cor"                & label %in% c("Branca", "Parda", "Preta"))                       |
                         (variable == "Regiao"                  & label == "Norte")                                               |
                         (variable == "Trabalha"                & label %in% c("Não", "Sim"))                                     |
                         (variable == "Profissional_Saude"      & label %in% c("Não", "Sim"))                                     |
                         (variable == "Renda_familiar"          & label == "até R$ 1254,00")                                      |
                         (variable == "Escolaridade"            & label %in% c("Ensino Médio Completo", "Pós-graduação"))         |
                         (variable == "Dificuldade_Financeira"  & label %in% c("Não", "Sim"))                                     |
                         (variable == "Acesso_Alimento"         & label %in% c("Não", "Sim"))                                     |
                         (variable == "atividade_fisica_pandemia" & label %in% c("Aumentou", "Diminuiu"))
  ) %>%
  
  # stat_3 = Não alterou
  modify_table_styling(columns = stat_3, text_format = "bold",
                       rows =
                         (variable == "Genero"                  & label %in% c("Feminino", "Masculino"))                          |
                         (variable == "Regiao"                  & label == "Centro-oeste")                                        |
                         (variable == "Isolamento"              & label %in% c("Não", "Sim"))                                     |
                         (variable == "Trabalha"                & label %in% c("Não", "Sim"))                                     |
                         (variable == "Renda_familiar"          & label == "mais de R$ 8.640")                                    |
                         (variable == "Escolaridade"            & label %in% c("Ensino Médio Completo", "Pós-graduação"))         |
                         (variable == "Consulta_Nutricionista"  & label %in% c("Não", "Sim"))                                     |
                         (variable == "Dificuldade_Financeira"  & label %in% c("Não", "Sim"))                                     |
                         (variable == "Acesso_Alimento"         & label %in% c("Não", "Sim"))                                     |
                         (variable == "Tempo_Preparo_Refeicao"  & label %in% c("Aumentou", "Diminuiu", "Não alterou"))           |
                         (variable == "atividade_fisica_pandemia" & label %in% c("Diminuiu", "Não alterou"))
  )

### rmd ----

tabela <- tabela %>%
  as_gt() %>%
  tab_style(
    style = cell_fill(color = "#fef9c3"),
    locations = cells_body(
      rows = p.value < 0.05
    )
  ) %>%
  gt::fmt_markdown(
    columns = "label",
    rows = !grepl("R\\$", label)
  )


### nota roda pé ----

tabela <- tabela %>%
  tab_source_note(
    source_note = md("**Q**: Teste do qui-quadrado de Pearson; <br>
                     **F**: Teste exato de Fisher; <br>
                     Valores em negrito na coluna *Valor-p* indicam significância estatística (p < 0,05); <br>
                     Valores em negrito nas tabelas de contingência indicam células com resíduos padronizados elevados.")
  )

### salvar ----

# tabela_categorica_sem_naoconsumo <- tabela
# saveRDS(tabela_categorica_sem_naoconsumo, file = "Avaliacao_6/data/tabela_categorica_sem_naoconsumo.rds")

# Quantitativa -----------------------------------------------------------------

df_quant <- df %>% select(16:19, "Doces") %>% filter(Doces != "Não consumo")

## Teste t -----------------------------------------------------------------------

## ANOVA -----------------------------------------------------------------------

### Normalidade ----
# Amostra grande; Shapiro rejeita
# Amostra grande; TCL garante Normalidade

#### Idade ----
# Não Normais
ggplot(df_quant, aes(x = Idade)) +
  geom_density(fill = "blue", alpha = 0.6) +
  facet_wrap(~ Doces) +
  xlab("Idade") +
  ylab("Densidade") +
  theme_bw() +
  theme(text = element_text(size = 14))
ggplot(df_quant, aes(sample = Idade)) +
  stat_qq_band(fill="lightblue") +
  stat_qq_line() +
  stat_qq_point() +
  facet_wrap(~Doces,scales = "free")+
  xlab("Quantis Teoricos")+
  ylab("Quantis Amostrais")+
  labs(title = "Idade")+
  theme_bw() +
  theme(text = element_text(size = 14))
by(df_quant$Idade, df_quant$Doces, shapiro.test)

#### Altura ----
# Normais
ggplot(df_quant, aes(x = Altura)) +
  geom_density(fill = "blue", alpha = 0.6) +
  facet_wrap(~ Doces) +
  xlab("Altura") +
  ylab("Densidade") +
  theme_bw() +
  theme(text = element_text(size = 14))
ggplot(df_quant, aes(sample = Altura)) +
  stat_qq_band(fill="lightblue") +
  stat_qq_line() +
  stat_qq_point() +
  facet_wrap(~Doces,scales = "free")+
  xlab("Quantis Teoricos")+
  ylab("Quantis Amostrais")+
  labs(title = "Altura")+
  theme_bw() +
  theme(text = element_text(size = 14))
by(df_quant$Altura, df_quant$Doces, shapiro.test)

#### Peso ----
# Não Normais
ggplot(df_quant, aes(x = Peso)) +
  geom_density(fill = "blue", alpha = 0.6) +
  facet_wrap(~ Doces) +
  xlab("Peso") +
  ylab("Densidade") +
  theme_bw() +
  theme(text = element_text(size = 14))
ggplot(df_quant, aes(sample = Peso)) +
  stat_qq_band(fill="lightblue") +
  stat_qq_line() +
  stat_qq_point() +
  facet_wrap(~Doces,scales = "free")+
  xlab("Quantis Teoricos")+
  ylab("Quantis Amostrais")+
  labs(title = "Peso")+
  theme_bw() +
  theme(text = element_text(size = 14))
by(df_quant$Peso, df_quant$Doces, shapiro.test)

#### IMC ----
# Não Normais
ggplot(df_quant, aes(x = IMC)) +
  geom_density(fill = "blue", alpha = 0.6) +
  facet_wrap(~ Doces) +
  xlab("IMC") +
  ylab("Densidade") +
  theme_bw() +
  theme(text = element_text(size = 14))
ggplot(df_quant, aes(sample = IMC)) +
  stat_qq_band(fill="lightblue") +
  stat_qq_line() +
  stat_qq_point() +
  facet_wrap(~Doces,scales = "free")+
  xlab("Quantis Teoricos")+
  ylab("Quantis Amostrais")+
  labs(title = "IMC")+
  theme_bw() +
  theme(text = element_text(size = 14))
by(df_quant$IMC, df_quant$Doces, shapiro.test)

### Homocedasticidade ----
# Homocedasticidade: Peso, IMC
# Heterocedasticidade: Idade, Altura

#### Idade ----
# bartlett.test(formula = Idade~Doces,data = df_quant)
LeveneTest(formula = Idade~Doces,data = df_quant)

#### Altura ----
bartlett.test(formula = Altura~Doces,data = df_quant)
# LeveneTest(formula = Altura~Doces,data = df_quant)

#### Peso ----
# bartlett.test(formula = Peso~Doces,data = df_quant)
LeveneTest(formula = Peso~Doces,data = df_quant)

#### IMC ----
# bartlett.test(formula = IMC~Doces,data = df_quant)
LeveneTest(formula = IMC~Doces,data = df_quant)

### Tabela com pvalor ----
tbl_summary(
  data = df_quant,
  by = Doces,
  statistic = all_continuous() ~ "{mean} ({sd})"
) |>
  add_p(
    test = everything() ~ "oneway.test",
    test.args = list(
      c(Idade, Altura) ~ list(var.equal = F),   
      c(Peso,IMC) ~ list(var.equal = T) 
    )
  ) %>% 
  bold_p(t = 0.05)

### games_howell ----
games_howell_test(Idade~Doces, data=df_quant)

# Aumentou:    a
# Diminuiu:    b
# Não alterou: b,c
# Não consumo: c

games_howell_test(Altura~Doces, data=df_quant)

# Aumentou:    a
# Diminuiu:    a,b
# Não alterou: b
# Não consumo: a,b

### Tabela ---------------------------------------------------------------------

comparacoes_letras <- function(data, variable, by, ...) {
  letras <- data.frame(
    variavel = c("Idade", "Altura", "Peso", "IMC"),
    comparacao = c(
      # Idade
      "Aumentou = a<br>Diminuiu = b<br>Não alterou = bc<br>Não consumo = c",
      
      # Altura
      "Aumentou = a<br>Diminuiu = ab<br>Não alterou = b<br>Não consumo = ab",
      
      # Peso (ajuste depois se calcular)
      "Aumentou = a<br>Diminuiu = a<br>Não alterou = a<br>Não consumo = a",
      
      # IMC (ajuste depois se calcular)
      "Aumentou = a<br>Diminuiu = a<br>Não alterou = a<br>Não consumo = a"
    ),
    stringsAsFactors = FALSE
  )
  
  result <- letras$comparacao[letras$variavel == variable]
  
  tibble::tibble(`**Comparações Múltiplas**` = result)
}

tabela <- tbl_summary(
  data = df_quant,
  by = Doces,
  statistic = all_continuous() ~ "{mean} ({sd})",
  label = list(
    Idade  ~ "Idade<sup>W</sup>",
    Altura ~ "Altura<sup>W</sup>",
    Peso   ~ "Peso<sup>A</sup>",
    IMC    ~ "IMC<sup>A</sup>"
  )
) %>%
  add_p(
    test = all_continuous() ~ "oneway.test",
    test.args = list(
      c(Idade, Altura) ~ list(var.equal = F),   
      c(Peso,IMC) ~ list(var.equal = T) 
    ),
    pvalue_fun = label_style_pvalue(digits = 3)
  ) %>%
  bold_p(t = 0.05) %>%
  
  modify_spanning_header(all_stat_cols() ~ "**Mudança no consumo de doces**") %>%
  modify_header(label ~ "**Variáveis**") %>%
  bold_labels() %>%
  
  modify_header(all_stat_cols() ~ "**{level}**<br>{n} ({style_percent(p)}%)") %>%
  modify_footnote(everything() ~ NA) %>%
  
  add_stat(
    fns = everything() ~ comparacoes_letras,
    location = everything() ~ "label"
  ) %>%
  
  as_gt() %>%
  
  tab_style(
    style     = cell_fill(color = "#fef9c3"),
    locations = cells_body(
      columns =  everything() , # p.value
      rows    = p.value < 0.05
    )
  ) %>%
  
  gt::fmt_markdown(columns = c(label, `**Comparações Múltiplas**`)) %>%
  
  tab_options(
    table.font.size = "20px",    
    heading.title.font.size = "26px",
    column_labels.font.size = "22px"
  )

#### nota roda pé ----

tabela <- tabela %>%
  tab_source_note(
    source_note = md("**A**: ANOVA One-Way; <br>
                     **W**: ANOVA de Welch; <br>
                     Valores em negrito indicam significância estatística (p < 0,05); <br>
                     Comparação múltipla aplicada: Games-Howell"
    )
  ) %>% 
  tab_options(
    source_notes.font.size = "15px"
  )

#### salvar ----

# saveRDS(tabela_numerica, file = "Avaliacao_6/data/tabela_numerica.rds")

## Plot ------------------------------------------------------------------------
df_long <- df_quant %>%
  pivot_longer(cols = c(Idade, Altura, Peso, IMC),
               names_to = "Variavel", values_to = "Valor") %>%
  mutate(
    Doces   = factor(Doces, levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )

df_summary <- df_long %>%
  group_by(Variavel, Doces) %>%
  summarise(
    media = mean(Valor, na.rm = TRUE),
    se    = sd(Valor, na.rm = TRUE) / sqrt(n()),
    .groups = "drop"
  )

ggplot(df_summary, aes(x = Doces, y = media)) +
  geom_point(size = 2.6) +
  geom_errorbar(aes(ymin = media - 1.96 * se,
                    ymax = media + 1.96 * se),
                width = 0.2) +
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  scale_x_discrete(guide = guide_axis(n.dodge = 2)) +
  labs(
    x = NULL,
    y = "Média (IC 95%)"
  )+
  theme_classic()

### Com elipses 1 ----
library(ggforce)

# Agrupamentos por letra (baseado nos resultados Games-Howell)
grupos <- tribble(
  ~Variavel, ~Doces,           ~grupo,
  # Idade: a | b | bc | c
  "Idade",   "Aumentou",       "a",
  "Idade",   "Diminuiu",       "b",
  "Idade",   "Não alterou",    "b",
  "Idade",   "Não alterou",    "c",
  "Idade",   "Não consumo",    "c",
  # Altura: a | ab | b | ab
  "Altura",  "Aumentou",       "a",
  "Altura",  "Diminuiu",       "a",
  "Altura",  "Diminuiu",       "b",
  "Altura",  "Não alterou",    "b",
  "Altura",  "Não consumo",    "a",
  "Altura",  "Não consumo",    "b",
  # Peso: todos "a"
  "Peso",    "Aumentou",       "a",
  "Peso",    "Diminuiu",       "a",
  "Peso",    "Não alterou",    "a",
  "Peso",    "Não consumo",    "a",
  # IMC: todos "a"
  "IMC",     "Aumentou",       "a",
  "IMC",     "Diminuiu",       "a",
  "IMC",     "Não alterou",    "a",
  "IMC",     "Não consumo",    "a"
) %>%
  mutate(
    Doces    = factor(Doces,    levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )

df_elipse <- grupos %>%
  left_join(df_summary, by = c("Variavel", "Doces"))

# Plot
ggplot(df_summary, aes(x = Doces, y = media)) +
  geom_mark_ellipse(
    data  = df_elipse,
    aes(x = Doces, y = media, group = grupo),
    color = "gray40", fill = NA,
    expand     = unit(3, "mm"),
    linetype   = "dashed",
    linewidth  = 0.4
  ) +
  geom_point(size = 2.6) +
  geom_errorbar(aes(ymin = media - 1.96 * se,
                    ymax = media + 1.96 * se),
                width = 0.2) +
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  scale_x_discrete(guide = guide_axis(n.dodge = 2)) +
  labs(x = NULL, y = "Média (IC 95%)") +
  theme_classic()

### Com elipses 2 ----

library(ggforce)

# Range de y por variável (para definir b mínimo proporcional)
y_ranges <- df_summary %>%
  group_by(Variavel) %>%
  summarise(y_range = diff(range(media)), .groups = "drop")

# Parâmetros das elipses
df_elipse_params <- df_elipse %>%
  group_by(Variavel, grupo) %>%
  summarise(
    x0    = mean(as.numeric(Doces)),
    y0    = mean(media),
    a_raw = diff(range(as.numeric(Doces))) / 2,
    b_raw = diff(range(media)) / 2,
    .groups = "drop"
  ) %>%
  left_join(y_ranges, by = "Variavel") %>%
  mutate(
    a = pmax(a_raw, 0.4) + 0.35,        # raio horizontal + folga
    b = pmax(b_raw, y_range * 0.06)     # raio vertical mínimo de 6% do range
  )

# Plot com eixo x numérico
ggplot(df_summary, aes(x = as.numeric(Doces), y = media)) +
  geom_ellipse(
    data        = df_elipse_params,
    aes(x0 = x0, y0 = y0, a = a, b = b, angle = 0, color = grupo),  # ← cor por grupo
    inherit.aes = FALSE,
    fill        = NA,
    linetype    = "dashed",
    linewidth   = 0.6
  ) +
  scale_color_manual(
    name   = "Grupo",
    values = c(
      "a" = "#D85A30",
      "b" = "#378ADD",
      "c" = "#1D9E75"
    )
  ) +
  geom_point(size = 2.6) +
  geom_errorbar(aes(ymin = media - 1.96 * se,
                    ymax = media + 1.96 * se),
                width = 0.2) +
  scale_x_continuous(
    breaks = 1:4,
    labels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo"),
    guide  = guide_axis(n.dodge = 2)
  ) +
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  labs(x = NULL, y = "Média (IC 95%)") +
  theme_classic() +
  theme(legend.position = "bottom")

### Com retangulo ----
grupos <- tribble(
  ~Variavel, ~Doces,           ~grupo,
  # Idade: a | b | bc | c
  "Idade",   "Aumentou",       "a",
  "Idade",   "Diminuiu",       "b",
  "Idade",   "Não alterou",    "b",
  "Idade",   "Não alterou",    "c",
  "Idade",   "Não consumo",    "c",
  # Altura: a | ab | b | ab
  "Altura",  "Aumentou",       "a",
  "Altura",  "Diminuiu",       "a",
  "Altura",  "Diminuiu",       "b",
  "Altura",  "Não alterou",    "b",
  "Altura",  "Não consumo",    "a",
  "Altura",  "Não consumo",    "b",
  # Peso: todos "a"
  "Peso",    "Aumentou",       "a",
  "Peso",    "Diminuiu",       "a",
  "Peso",    "Não alterou",    "a",
  "Peso",    "Não consumo",    "a",
  # IMC: todos "a"
  "IMC",     "Aumentou",       "a",
  "IMC",     "Diminuiu",       "a",
  "IMC",     "Não alterou",    "a",
  "IMC",     "Não consumo",    "a"
) %>%
  mutate(
    Doces    = factor(Doces,    levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )

df_rect_params <- grupos %>%
  left_join(df_summary, by = c("Variavel", "Doces")) %>%
  group_by(Variavel, grupo) %>%
  summarise(
    x0    = mean(as.numeric(Doces)),
    y0    = mean(media),
    a_raw = diff(range(as.numeric(Doces))) / 2,
    b_raw = diff(range(media)) / 2,
    .groups = "drop"
  ) %>%
  left_join(y_ranges, by = "Variavel") %>%
  mutate(
    a    = pmax(a_raw, 0.4) + 0.35,
    b    = pmax(b_raw, y_range * 0.06),
    xmin = 0.65,
    xmax = 4.35,
    ymin = y0 - b,
    ymax = y0 + b,
    ymax = case_when(
      Variavel == "Altura" & grupo == "a" ~ 1.654,
      Variavel == "Idade"  & grupo == "b" ~ 35.8,
      TRUE ~ ymax
    ),
    ymin = case_when(
      Variavel == "Altura" & grupo == "b" ~ 1.66,
      TRUE ~ ymin
    )
  )

plot_medias_rect <- ggplot(df_summary, aes(x = as.numeric(Doces), y = media)) +
  geom_rect(
    data        = df_rect_params,
    aes(xmin = xmin, xmax = xmax, ymin = ymin, ymax = ymax, color = grupo),
    inherit.aes = FALSE,
    fill        = NA,
    linetype    = "dashed",
    linewidth   = 0.6
  ) +
  scale_color_manual(
    name   = "Grupo",
    values = c("a" = "#D85A30", "b" = "#378ADD", "c" = "#1D9E75")
  ) +
  geom_point(size = 2.6) +
  geom_errorbar(aes(ymin = media - 1.96 * se,
                    ymax = media + 1.96 * se),
                width = 0.2) +
  scale_x_continuous(
    breaks = 1:4,
    labels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo"),
    guide  = guide_axis(n.dodge = 2)
  ) +
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  labs(x = NULL, y = "Média (IC 95%)") +
  theme_classic() +
  theme(legend.position = "bottom")

plot_medias_rect

# saveRDS(plot_medias_rect, file = "Avaliacao_6/data/plot_medias_rect.rds")

