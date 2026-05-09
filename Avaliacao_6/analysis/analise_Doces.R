############Analise da variavel Doces###########################################

# Setup ------------------------------------------------------------------------
library(tidyverse)

df <- openxlsx::read.xlsx("Avaliacao_6/data/Nutricao.xlsx") %>% 
  select(1:20, "Doces") %>%
  mutate(Doces = factor(Doces, levels = c("Aumentou", "Não alterou", "Diminuiu", "Não consumo")))

# Plot -------------------------------------------------------------------------

plot_prop_doce_1 <- ggplot(df, aes(x = Doces)) +
  geom_bar(fill = "#c07068", color = "#2d2f45", width = 0.7) +
  
  # percentual dentro da barra
  geom_text(
    stat = "count",
    aes(label = percent(after_stat(count / sum(count)))),
    position = position_stack(vjust = 0.5),
    color = "white",
    size = 5,
    fontface = "bold"
  ) +
  
  # frequência no topo (caixinha estilizada)
  geom_label(
    stat = "count",
    aes(
      y = after_stat(count),
      label = after_stat(count)
    ),
    vjust = 0.3,
    size = 4,
    fill = "white",
    color = "#2d2f45",
    label.size = 0.6,
    label.r = unit(0.2, "lines")  # cantos levemente arredondados
  ) +
  
  labs(
    title = "Mudança no consumo de doces",
    x = NULL,
    y = "Frequência"
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
    
    legend.position = "none"
  )
plot_prop_doce_1

# saveRDS(plot_prop_doce_1, file = "Avaliacao_6/data/plot_prop_doce_1.rds")