#########Analise da variavel aberta#############################################

# Setup ------------------------------------------------------------------------

require(openxlsx)    # Leitura de arquivos xlsx
require(stringr)     # Manipulação de textos
require(dplyr)       # Manipulação de dados
require(ggwordcloud) # Nuvem de palavras (ggplot2)
library(wordcloud2)  # Nuvem de palavras (html)
# devtools::install_github("lchiffon/wordcloud2")
require(tidyr)       # Limpeza
require(tidytext)    # Tokenização profissional
require(pals)        # Paleta de Cores
require(png)         # Carrega shapes de fora do R

df <- openxlsx::read.xlsx("Avaliacao_6/data/Nutricao.xlsx") %>%  select(1:20, "Doces")
df_abert <- df %>% select(20, "Doces")

# Limpeza ----------------------------------------------------------------------
frequencia_palavras <- df_abert %>%
  mutate(texto_limpo = as.character(sentimentos_pandemia)) %>%
  mutate(texto_limpo = str_replace_all(texto_limpo, "[\\.,!?;:()\\[\\]{}“”\"'-]", "")) %>%
  mutate(texto_limpo = str_to_lower(texto_limpo)) %>%
  unnest_tokens(output = palavra, input = texto_limpo) %>%
  filter(palavra != "") %>%
  filter(str_length(palavra) > 3) %>%
  count(palavra, sort = TRUE, name = "n") %>%
  mutate(porc = round((n / sum(n)) * 100, 2))

frequencia_palavras_Doces <- df_abert %>%
  mutate(texto_limpo = as.character(sentimentos_pandemia)) %>%
  mutate(texto_limpo = str_replace_all(texto_limpo, "[\\.,!?;:()\\[\\]{}“”\"']", "")) %>%
  mutate(texto_limpo = str_to_lower(texto_limpo)) %>%
  unnest_tokens(output = palavra, input = texto_limpo) %>%
  filter(palavra != "") %>%
  filter(str_length(palavra) > 3) %>%
  group_by(Doces, palavra) %>%   
  summarise(n = n()) %>%   
  mutate(porc = round((n / sum(n)) * 100, 2))

# Adequação --------------------------------------------------------------------
frequencia_palavras$angle <- sample(c(0,10,40,60), nrow(frequencia_palavras), replace = TRUE)
frequencia_palavras_Doces$angle <- sample(c(0,10,40,60), nrow(frequencia_palavras_Doces), replace = TRUE)
frequencia_palavras_Doces <- frequencia_palavras_Doces %>%
  group_by(Doces) %>% 
  arrange(desc(porc), .by_group = TRUE) %>%
  mutate(
    ranking = dense_rank(desc(porc))
  )%>%
  ungroup

# Barras -----------------------------------------------------------------------

library(tidyverse)
library(ggtext)

## Processamento ---------------------------------------------------------------
df_plot <- frequencia_palavras_Doces %>%
  group_by(Doces) %>%
  slice_max(order_by = porc, n = 10) %>%
  mutate(rank = rank(-porc, ties.method = "first")) %>%
  ungroup()

df_plot <- df_plot %>%
  group_by(Doces) %>%
  arrange(desc(porc), .by_group = TRUE) %>%
  mutate(rank_plot = row_number()) %>%
  ungroup() %>%
  mutate(rank_plot = case_when(
    Doces == "Aumentou" & palavra == "medo"     ~ 4L,
    Doces == "Aumentou" & palavra == "preguiça" ~ 5L,
    TRUE ~ rank_plot
  )) %>%
  mutate(palavra_ord = reorder_within(palavra, -rank_plot, Doces))

shared_rank <- df_plot %>%
  group_by(palavra, rank) %>%
  summarise(n_grupos = n_distinct(Doces), .groups = "drop") %>%
  filter(n_grupos == 4) %>%
  mutate(chave = paste(palavra, rank, sep = "__"))

df_plot <- df_plot %>%
  mutate(
    chave    = paste(palavra, rank, sep = "__"),
    destaque = chave %in% shared_rank$chave
  )

## Plot ------------------------------------------------------------------------
cores_doces <- c(
  "Aumentou"    = "#c07068",
  "Diminuiu"    = "#5c607a",
  "Não alterou" = "#4c9a92",
  "Não consumo" = "#a07850"
)
cor_destaque <- "#f0c040"
plot_palavras_ranking <- ggplot(df_plot, aes(x = palavra_ord, y = porc, fill = Doces)) +
  
  geom_col(
    aes(color = destaque, linewidth = destaque),
    show.legend = FALSE
  ) +
  
  geom_text(
    aes(label = n),
    position = position_stack(vjust = 0.5),
    size = 3, color = "white", fontface = "bold", family = "Inter"
  ) +
  
  geom_text(
    aes(label = paste0(porc, "%")),
    hjust = -0.1, size = 3, color = "#2d2f45", family = "Inter"
  ) +
  
  geom_point(
    data = filter(df_plot, destaque),
    aes(y = 0),
    shape = 18, size = 3, color = cor_destaque,
    show.legend = FALSE
  ) +
  
  coord_flip() +
  scale_x_reordered() +
  scale_y_continuous(expand = expansion(mult = c(0.05, 0.18))) +
  
  scale_fill_manual(values = cores_doces) +
  
  scale_color_manual(
    values = c("TRUE" = cor_destaque, "FALSE" = NA)
  ) +
  
  scale_linewidth_manual(
    values = c("TRUE" = 1.0, "FALSE" = 0),
    guide  = "none"
  ) +
  
  facet_wrap(~ Doces, scales = "free_y", ncol = 2) +
  
  labs(
    title    = "Top 10 palavras mais frequentes por grupo",
    subtitle = "Palavra aparece na mesma posição de ranking em todos os grupos",
    x        = NULL,
    y        = "Percentual (%)"
  ) +
  
  theme_minimal(base_family = "Inter") +
  
  theme(
    axis.text          = element_text(color = "#2d2f45"),
    axis.title.x       = element_text(color = "#5c607a"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor   = element_blank(),
    panel.grid.major.x = element_line(color = "#e5e7eb", linewidth = 0.6),
    axis.line          = element_line(color = "#2d2f45"),
    strip.text         = element_text(face = "bold", color = "#2d2f45", size = 12),
    plot.title         = element_text(hjust = 0.5, color = "#2d2f45", face = "bold", size = 13),
    plot.subtitle      = element_text(hjust = 0.5, color = cor_destaque, size = 9, face = "bold"),
    legend.position    = "none"
  )
plot_palavras_ranking
# saveRDS(plot_palavras_ranking, "Avaliacao_6/data/plot_palavras_ranking.rds")

# Nuvens -----------------------------------------------------------------------

## ggwordcloud -----------------------------------------------------------------
frequencia_palavras_Doces %>% 
  ggplot(aes(label = palavra, 
             size = log(porc+1.2),
             angle = angle,
             color =  factor(ranking)))+   
  scale_size_area(max_size = 16)+
  scale_color_manual(values =colorRampPalette(c("red", "yellow", "blue"))(15)) +  
  facet_wrap(~Doces,ncol=2)+
  geom_text_wordcloud(seed=123,
                      shape = "square",
                      rm_outside=T)+
  theme_bw()

### Bala Unica ----
mask_image <- png::readPNG("Avaliacao_6/media/candy.png")

frequencia_palavras %>% 
  ggplot(aes(label = palavra, 
             size = log(n+1),
             angle = angle,
             color = as.character(n)))+   
  scale_color_manual(values =glasbey()) +
  geom_text_wordcloud(seed=123,
                      mask = mask_image,
                      rm_outside=T)+
  scale_size_area(max_size = 5)+
  theme_bw()

### Bala Facetada ----
mask_image <- png::readPNG("Avaliacao_6/media/candy.png")

frequencia_palavras_Doces %>% 
  ggplot(aes(label = palavra, 
             size = log(porc+1.2),
             angle = angle,
             color =  factor(ranking)))+   
  scale_size_area(max_size = 16)+
  scale_color_manual(values =colorRampPalette(c("red", "yellow", "blue"))(15)) +  
  facet_wrap(~Doces,ncol=2)+
  geom_text_wordcloud(seed=123,
                      mask = mask_image,
                      rm_outside=T)+
  scale_size_area(max_size = 5)+
  theme_bw()

#### floreio ----
mask_image <- png::readPNG("Avaliacao_6/media/candy.png")

plot_nuvem_facet <- frequencia_palavras_Doces %>% 
  ggplot(aes(label = palavra, 
             size = log(porc+1.2),
             angle = angle,
             color = factor(ranking))) +  
  scale_color_manual(values = colorRampPalette(c("#c07068", "#ffffff"))(15)) +  
  facet_wrap(~Doces, ncol = 2) +
  geom_text_wordcloud(seed = 123,
                      mask = mask_image,
                      rm_outside = TRUE,
                      fontface = "bold") +
  scale_size_area(max_size = 2) +
  theme(
    plot.background  = element_rect(fill = "#2d2f45", color = NA),
    panel.background = element_rect(fill = "#2d2f45", color = NA),
    plot.margin      = margin(0, 0, 0, 0),
    strip.background = element_rect(fill = "#4c9a92", color = NA),
    strip.text       = element_text(color = "white",
                                    face  = "bold",
                                    size  = 12),
    panel.border     = element_blank()
  )
plot_nuvem_facet
# saveRDS(plot_nuvem_facet, "Avaliacao_6/data/plot_nuvem_facet.rds")

## wordcloud2 ------------------------------------------------------------------

### Unica ----
frequencia_palavras_2 = frequencia_palavras%>%
  mutate(n=log(n+1))
wordcloud2(
  data = frequencia_palavras_2,
  size = 0.6,          
  fontWeight = 'bold',  
  color='random-light', 
  backgroundColor="dark",
  shape="circle"
)

### Bala ----
frequencia_palavras_2 = frequencia_palavras%>%
  mutate(n=log(n+1))
wordcloud2(
  data = frequencia_palavras_2,
  size = 0.8,          
  fontWeight = 'bold',  
  color='random-light', 
  backgroundColor="dark",
  figPath = "Avaliacao_6/media/candy.png"
)