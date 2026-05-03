#########Analise das variaveis numericas########################################

# Setup ------------------------------------------------------------------------
library(tidyverse)
require(gtsummary)           # Tabelas automáticas
require(gt)                  # Tabelas automáticas
require(qqplotr)             # Gráficos qqplot
require(DescTools)           # Teste de Levene
library(scales)

df <- openxlsx::read.xlsx("Avaliacao_6/data/Nutricao.xlsx") %>%  select(1:20, "Doces")
df_quant <- df %>% select(16:19, "Doces")

# formatação em português (vírgula pra decimais e ponto para milhares)
theme_gtsummary_language("pt", big.mark = ".", decimal.mark = ",")

# Teste t -----------------------------------------------------------------------

# ANOVA -----------------------------------------------------------------------

## Normalidade ----
# Amostra grande; Shapiro rejeita
# Amostra grande; TCL garante Normalidade
# Normais: Altura

### Idade ----
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

library(nortest)
by(df_quant$Idade, df_quant$Doces, ad.test)
by(df_quant$Idade, df_quant$Doces, lillie.test)
by(df_quant$Idade, df_quant$Doces, function(x) {
  ks.test(x, "pnorm", mean = mean(x), sd = sd(x))
})

### Altura ----
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

### Peso ----
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

### IMC ----
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

## Homocedasticidade ----
# Homocedasticidade: Peso, IMC
# Heterocedasticidade: Idade, Altura

### Idade ----
# bartlett.test(formula = Idade~Doces,data = df_quant)
LeveneTest(formula = Idade~Doces,data = df_quant)

### Altura ----
bartlett.test(formula = Altura~Doces,data = df_quant)
# LeveneTest(formula = Altura~Doces,data = df_quant)

### Peso ----
# bartlett.test(formula = Peso~Doces,data = df_quant)
LeveneTest(formula = Peso~Doces,data = df_quant)

### IMC ----
# bartlett.test(formula = IMC~Doces,data = df_quant)
LeveneTest(formula = IMC~Doces,data = df_quant)

## Tabela com pvalor ----
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

## Games_Howell ----
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

## Tabela ---------------------------------------------------------------------

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

### nota roda pé ----

tabela <- tabela %>%
  tab_source_note(
    source_note = md("**A**: ANOVA One-Way; <br>
                     **W**: ANOVA de Welch; <br>
                     Valores em negrito indicam significância estatística (p < 0,05); <br>
                     Comparação múltipla aplicada: Games-Howell"
    )
  ) %>% 
  tab_options(
    source_notes.font.size = "14px"
  ) %>%
  tab_style(
    style = cell_text(color = "#666666"),
    locations = cells_source_notes()
  )

### salvar ----

tabela_numerica <- tabela

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

### Com retangulo 1 ----
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
y_ranges <- df_summary %>%
  group_by(Variavel) %>%
  summarise(y_range = diff(range(media)), .groups = "drop")
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
    data = df_rect_params,
    aes(
      xmin = xmin, xmax = xmax,
      ymin = ymin, ymax = ymax,
      fill = grupo,
      color = grupo
    ),
    inherit.aes = FALSE,
    alpha = 0.10,
    linetype = "solid",
    linewidth = 0.7
  ) +
  
  scale_fill_manual(
    name   = "Grupo",
    values = c(
      "a" = "#c07068",
      "b" = "#5c607a",
      "c" = "#4c9a92"
    )
  ) +
  
  scale_color_manual(
    name   = "Grupo",
    values = c(
      "a" = "#c07068",
      "b" = "#5c607a",
      "c" = "#4c9a92"
    )
  ) +
  
  geom_point(
    size = 2.8,
    color = "#2d2f45"
  ) +
  
  geom_errorbar(
    aes(
      ymin = media - 1.96 * se,
      ymax = media + 1.96 * se
    ),
    width = 0.15,
    color = "#2d2f45",
    linewidth = 0.6
  ) +
  
  scale_x_continuous(
    breaks = 1:4,
    labels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo"),
    guide  = guide_axis(n.dodge = 2)
  ) +
  
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  
  labs(
    x = NULL,
    y = "Média (IC 95%)"
  ) +
  
  theme_minimal(base_family = "Inter") +
  
  theme(
    plot.title = element_text(
      hjust = 0.5,
      size = 18,
      face = "bold",
      color = "#2d2f45"
    ),
    axis.text = element_text(color = "#2d2f45"),
    axis.title.y = element_text(color = "#5c607a"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    panel.grid.major.y = element_line(
      color = "#e5e7eb",
      linewidth = 0.6
    ),
    axis.line = element_line(color = "#2d2f45"),
    strip.text = element_text(
      face = "bold",
      color = "#2d2f45",
      size = 12
    ),
    legend.position = "right",
    legend.title = element_text(color = "#5c607a"),
    legend.text = element_text(color = "#2d2f45")
  )


plot_medias_rect

# saveRDS(plot_medias_rect, file = "Avaliacao_6/data/plot_medias_rect.rds")

### Com retangulo 2 ----

#### Grupos
grupos <- tribble(
  ~Variavel, ~Doces,           ~grupo,
  "Idade",   "Aumentou",       "a",
  "Idade",   "Diminuiu",       "b",
  "Idade",   "Não alterou",    "b",
  "Idade",   "Não alterou",    "c",
  "Idade",   "Não consumo",    "c",
  "Altura",  "Aumentou",       "a",
  "Altura",  "Diminuiu",       "a",
  "Altura",  "Diminuiu",       "b",
  "Altura",  "Não alterou",    "b",
  "Altura",  "Não consumo",    "a",
  "Altura",  "Não consumo",    "b",
  "Peso",    "Aumentou",       "a",
  "Peso",    "Diminuiu",       "a",
  "Peso",    "Não alterou",    "a",
  "Peso",    "Não consumo",    "a",
  "IMC",     "Aumentou",       "a",
  "IMC",     "Diminuiu",       "a",
  "IMC",     "Não alterou",    "a",
  "IMC",     "Não consumo",    "a"
) %>%
  mutate(
    Doces    = factor(Doces, levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )

#### Dados longos
df_long <- df_quant %>%
  pivot_longer(
    cols = c(Idade, Altura, Peso, IMC),
    names_to = "Variavel",
    values_to = "Valor"
  ) %>%
  mutate(
    Doces    = factor(Doces, levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )

#### Resumo
df_summary <- df_long %>%
  group_by(Variavel, Doces) %>%
  summarise(
    media = mean(Valor, na.rm = TRUE),
    se    = sd(Valor, na.rm = TRUE) / sqrt(n()),
    .groups = "drop"
  ) %>%
  mutate(
    ic_inf = media - 1.96 * se,
    ic_sup = media + 1.96 * se
  )

#### Parâmetros dos retângulos (interseção dos ICs)
df_rect_params <- grupos %>%
  left_join(df_summary, by = c("Variavel", "Doces")) %>%
  group_by(Variavel, grupo) %>%
  summarise(
    xmin = min(as.numeric(Doces)) - 0.35,
    xmax = max(as.numeric(Doces)) + 0.35,
    
    ymax = min(ic_sup, na.rm = TRUE),  # menor limite superior
    ymin = max(ic_inf, na.rm = TRUE),  # maior limite inferior
    
    .groups = "drop"
  ) %>%
  # (opcional) evitar retângulos invertidos se não houver interseção
  mutate(
    ymin = ifelse(ymin > ymax, NA, ymin),
    ymax = ifelse(ymin > ymax, NA, ymax)
  )

#### Plot
plot_medias_rect_2 <- ggplot(df_summary, aes(x = as.numeric(Doces), y = media)) +
  
  geom_rect(
    data = df_rect_params,
    aes(
      xmin = xmin, xmax = xmax,
      ymin = ymin, ymax = ymax,
      fill = grupo,
      color = grupo
    ),
    inherit.aes = FALSE,
    alpha = 0.10,
    linewidth = 0.7
  ) +
  
  scale_fill_manual(
    name = "Grupo",
    values = c(
      "a" = "#c07068",
      "b" = "#5c607a",
      "c" = "#4c9a92"
    )
  ) +
  
  scale_color_manual(
    name = "Grupo",
    values = c(
      "a" = "#c07068",
      "b" = "#5c607a",
      "c" = "#4c9a92"
    )
  ) +
  
  geom_point(
    size = 2.8,
    color = "#2d2f45"
  ) +
  
  geom_errorbar(
    aes(
      ymin = ic_inf,
      ymax = ic_sup
    ),
    width = 0.15,
    color = "#2d2f45",
    linewidth = 0.6
  ) +
  
  scale_x_continuous(
    breaks = 1:4,
    labels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo"),
    guide = guide_axis(n.dodge = 2)
  ) +
  
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  
  labs(
    x = NULL,
    y = "Média (IC 95%)",
    caption = "Agrupamento por Comparação Múltipla de Games-Howell."
  ) +
  
  theme_minimal(base_family = "Inter") +
  
  theme(
    axis.text = element_text(color = "#2d2f45"),
    axis.title.y = element_text(color = "#5c607a"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    panel.grid.major.y = element_line(
      color = "#e5e7eb",
      linewidth = 0.6
    ),
    axis.line = element_line(color = "#2d2f45"),
    strip.text = element_text(
      face = "bold",
      color = "#2d2f45",
      size = 12
    ),
    legend.position = "right",
    legend.title = element_text(color = "#5c607a"),
    legend.text = element_text(color = "#2d2f45"),
    plot.caption = element_text(hjust = 0,
                                color = "#666666",
                                size = 7.2)
  )

plot_medias_rect_2

# saveRDS(plot_medias_rect_2, file = "Avaliacao_6/data/plot_medias_rect_2.rds")

### Com retangulo 3 ----
library(tidyverse)
library(ggtext)

grupos <- tribble(
  ~Variavel, ~Doces,           ~grupo,
  "Idade",   "Aumentou",       "a",
  "Idade",   "Diminuiu",       "b",
  "Idade",   "Não alterou",    "b",
  "Idade",   "Não alterou",    "c",
  "Idade",   "Não consumo",    "c",
  "Altura",  "Aumentou",       "a",
  "Altura",  "Diminuiu",       "a",
  "Altura",  "Diminuiu",       "b",
  "Altura",  "Não alterou",    "b",
  "Altura",  "Não consumo",    "a",
  "Altura",  "Não consumo",    "b",
  "Peso",    "Aumentou",       "a",
  "Peso",    "Diminuiu",       "a",
  "Peso",    "Não alterou",    "a",
  "Peso",    "Não consumo",    "a",
  "IMC",     "Aumentou",       "a",
  "IMC",     "Diminuiu",       "a",
  "IMC",     "Não alterou",    "a",
  "IMC",     "Não consumo",    "a"
) %>%
  mutate(
    Doces    = factor(Doces, levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )


cor_letras <- c(
  "a" = "#c07068",
  "b" = "#5c607a",
  "c" = "#4c9a92"
)


grupos_label <- grupos %>%
  group_by(Variavel, Doces) %>%
  summarise(
    label_html = {
      letras <- sort(unique(grupo))
      paste(
        sprintf("<b style='color:%s'>%s</b>", cor_letras[letras], letras),
        collapse = ""
      )
    },
    .groups = "drop"
  )


df_long <- df_quant %>%
  pivot_longer(
    cols      = c(Idade, Altura, Peso, IMC),
    names_to  = "Variavel",
    values_to = "Valor"
  ) %>%
  mutate(
    Doces    = factor(Doces, levels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo")),
    Variavel = factor(Variavel, levels = c("Idade", "Altura", "Peso", "IMC"))
  )
df_summary <- df_long %>%
  group_by(Variavel, Doces) %>%
  summarise(
    media = mean(Valor, na.rm = TRUE),
    se    = sd(Valor, na.rm = TRUE) / sqrt(n()),
    .groups = "drop"
  ) %>%
  mutate(
    ic_inf = media - 1.96 * se,
    ic_sup = media + 1.96 * se
  )

df_summary_label <- df_summary %>%
  left_join(grupos_label, by = c("Variavel", "Doces"))
plot_medias_letras <- ggplot(
  df_summary_label,
  aes(x = as.numeric(Doces), y = media)
) +
  
  geom_errorbar(
    aes(ymin = ic_inf, ymax = ic_sup),
    width     = 0.15,
    color     = "#2d2f45",
    linewidth = 0.6
  ) +
  
  geom_point(
    size  = 2.8,
    color = "#2d2f45"
  ) +
  
  # Letras coloridas na lateral direita do ponto
  geom_richtext(
    aes(x = as.numeric(Doces) + 0.12, y = media, label = label_html),
    hjust         = 0,
    vjust         = 0.5,
    size          = 3.8,
    label.colour  = NA,       # sem borda no box
    fill          = NA,       # sem fundo
    family        = "Inter"
  ) +
  
  scale_x_continuous(
    breaks = 1:4,
    labels = c("Aumentou", "Diminuiu", "Não alterou", "Não consumo"),
    guide  = guide_axis(n.dodge = 2),
    expand = expansion(add = c(0.4, 0.7))  # espaço à direita para as letras
  ) +
  
  facet_wrap(~ Variavel, scales = "free_y", ncol = 2) +
  
  labs(
    x       = NULL,
    y       = "Média (IC 95%)",
    caption = "Letras iguais indicam ausência de diferença significativa (Games-Howell, α = 0,05).\nIC 95% individual exibido apenas para referência descritiva."
  ) +
  
  theme_minimal(base_family = "Inter") +
  
  theme(
    axis.text          = element_text(color = "#2d2f45"),
    axis.title.y       = element_text(color = "#5c607a"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor   = element_blank(),
    panel.grid.major.y = element_line(color = "#e5e7eb", linewidth = 0.6),
    axis.line          = element_line(color = "#2d2f45"),
    strip.text         = element_text(face = "bold", color = "#2d2f45", size = 12),
    legend.position    = "none",
    plot.caption       = element_text(hjust = 0, color = "#666666", size = 7.2)
  )

plot_medias_letras

# saveRDS(plot_medias_letras, file = "Avaliacao_6/data/plot_medias_letras.rds")
