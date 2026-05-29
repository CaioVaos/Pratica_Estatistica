#########Analise das variaveis numericas########################################

# Setup ------------------------------------------------------------------------
library(tidyverse)      # Manipulação de bases de dados
library(ggplot2)    # Gráficos
library(geobr)      # Shapefiles do Brasil
# library(ggsflabel)  # Criação de Labels que se repelem
library(ggspatial)  # Rosa dos ventos e escala
library(sf)         # Leitura de shapefiles fora do geobr
library(patchwork) # Juntar plots

# Sys.unsetenv("GITHUB_PAT")
remotes::install_github("ipeaGIT/geobr", subdir = "r-package")

devtools::install_github("yutannihilation/ggsflabel")

# Estrutura --------------------------------------------------------------------
Dados_1 <- read_health_region(year = 2025,
                              geometry_level = "micro",
                              simplified = T)

Dados_2 <- read_indigenous_land(year = 2024,
                                simplified = T)

class(Dados_1)
class(Dados_2)
unique(sf::st_geometry_type(Dados_1))
unique(sf::st_geometry_type(Dados_2))

# read_health_region -----------------------------------------------------------
args(read_health_region)

## Brasil ----

### Dados ----
Dados_muni <- read_health_region(
  year = 2025,
  geometry_level = "municipality"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Municípios")
Dados_micro <- read_health_region(
  year = 2025,
  geometry_level = "micro"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Microrregiões")
Dados_macro <- read_health_region(
  year = 2025,
  geometry_level = "macro"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Macrorregiões")

Dados_plot <- bind_rows(
  Dados_muni,
  Dados_micro,
  Dados_macro
)

### Plot ----
mapa_saude_brasil <- ggplot(Dados_plot) +
  
  geom_sf(
    fill = "#5b8db8",
    color = "white",
    linewidth = 0.03
  ) +
  
  facet_wrap(
    ~nivel,
    ncol = 3
  ) +
  annotation_north_arrow(
    location = "bl",
    which_north = "true",
    pad_x = unit(0.2, "cm"),
    pad_y = unit(0.2, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col = "#0a2535",
      line_col = "#0a2535",
      fill = c("#5b8db8", "white")
    ),
    height = unit(1, "cm"),
    width = unit(1, "cm")
  ) +
  annotation_scale(
    location = "br",
    width_hint = 0.25,
    text_cex = 0.6,
    line_width = 0.7
  ) +
  
  labs(
    title = "Níveis espaciais das Regiões de Saúde no Brasil",
    subtitle = "Pacote geobr • read_health_region() • Ano 2025",
    caption = "Fonte: DataSUS / geobr"
  ) +
  
  theme_minimal()+
  
  theme(
    
    plot.title.position = "plot",
    
    plot.title = element_text(
      hjust = 0.5,
      face = "bold",
      color = "#0a2535"
    ),
    
    plot.subtitle = element_text(
      hjust = 0.5,
      color = "#4a6a7a"
    ),
    
    plot.caption = element_text(
      hjust = 0,
      size = 9,
      color = "gray40"
    ),
    
    panel.grid = element_blank(),
    
    axis.text = element_blank(),
    axis.title = element_blank(),
    axis.ticks = element_blank(),
    
    strip.text = element_text(
      face = "bold",
      size = 9,
      color = "#0a2535"
    ),
    
    panel.border = element_rect(
      color = "#d9d9d9",
      fill = NA,
      linewidth = 0.5
    ),
    
    plot.background = element_rect(
      fill = "#d4ebff",
      color = NA
    ),
    panel.background = element_rect(
      fill = "white",
      color = NA
    )
    
  )
mapa_saude_brasil
ggsave(
  filename = "Avaliacao_9/media/mapa_saude_brasil.png",
  plot = mapa_saude_brasil,
  width = 10,
  height = 4,
  dpi = 150,
  bg = "#d4ebff"
)

## RJ ----

### Dados ----
Dados_muni <- read_health_region(
  year = 2025,
  code_state = 33,
  geometry_level = "municipality"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Municípios")
Dados_micro <- read_health_region(
  year = 2025,
  code_state = 33,
  geometry_level = "micro"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Microrregiões")

Dados_macro <- read_health_region(
  year = 2025,
  code_state = 33,
  geometry_level = "macro"
) |>
  filter(!st_is_empty(geometry)) |>
  mutate(nivel = "Macrorregiões")

Dados_plot <- bind_rows(
  Dados_muni,
  Dados_micro,
  Dados_macro
)

### Plot ----
mapa_saude_rj <- ggplot(Dados_plot) +
  
  geom_sf(
    fill = "#5b8db8",
    color = "white",
    linewidth = 0.25
  ) +
  
  facet_wrap(
    ~nivel,
    ncol = 3
  ) +
  annotation_north_arrow(
    location = "tl",
    which_north = "true",
    pad_x = unit(0.2, "cm"),
    pad_y = unit(0.2, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col = "#0a2535",
      line_col = "#0a2535",
      fill = c("#5b8db8", "white")
    ),
    height = unit(1, "cm"),
    width = unit(1, "cm")
  ) +
  annotation_scale(
    location = "br",
    width_hint = 0.25,
    text_cex = 0.6,
    line_width = 0.7
  ) +
  
  labs(
    title = "Níveis espaciais das Regiões de Saúde no RJ",
    subtitle = "Pacote geobr • read_health_region() • Ano 2025",
    caption = "Fonte: DataSUS / geobr"
  ) +
  
  theme_minimal()+
  
  theme(
    
    plot.title.position = "plot",
    
    plot.title = element_text(
      hjust = 0.5,
      face = "bold",
      color = "#0a2535"
    ),
    
    plot.subtitle = element_text(
      hjust = 0.5,
      color = "#4a6a7a"
    ),
    
    plot.caption = element_text(
      hjust = 0,
      size = 9,
      color = "gray40"
    ),
    
    panel.grid = element_blank(),
    
    axis.text = element_blank(),
    axis.title = element_blank(),
    axis.ticks = element_blank(),
    
    strip.text = element_text(
      face = "bold",
      size = 9,
      color = "#0a2535"
    ),
    
    panel.border = element_rect(
      color = "#d9d9d9",
      fill = NA,
      linewidth = 0.5
    ),

    plot.background = element_rect(
      fill = "#d4ebff",
      color = NA
    ),
    panel.background = element_rect(
      fill = "white",
      color = NA
    ),
    strip.background = element_rect(
      fill = "#d4ebff",
      color = NA
    )
  )
mapa_saude_rj

ggsave(
  filename = "Avaliacao_9/media/mapa_saude_rj.png",
  plot = mapa_saude_rj,
  width = 10,
  height = 4,
  dpi = 150,
  bg = "#d4ebff"
)

# read_indigenous_land -----------------------------------------------------------
args(read_indigenous_land)

## Geral ----

Dados_brasil <- read_indigenous_land(year = 2025)
Dados_am <- read_indigenous_land(year = 2025,
                                 code_state = 13)

contorno_brasil <- read_country(year = 2025, simplified = TRUE)
contorno_am     <- read_state(code_state = 13, year = 2025, simplified = TRUE)

plot_brasil <- ggplot() +
  geom_sf(data = contorno_brasil,
          fill = NA, color = "#0a2535", linewidth = 0.2) +
  geom_sf(data = Dados_brasil,
          fill = "#b5622a", color = "white", linewidth = 0.03) +
  annotation_north_arrow(
    location = "bl",
    which_north = "true",
    pad_x = unit(0.2, "cm"),
    pad_y = unit(0.2, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col = "#0a2535",
      line_col = "#0a2535",
      fill = c("#b5622a", "white")
    ),
    height = unit(1, "cm"),
    width  = unit(1, "cm")
  ) +
  annotation_scale(
    location   = "br",
    width_hint = 0.25,
    text_cex   = 0.6,
    line_width = 0.7
  ) +
  labs(title = "Brasil") +
  theme_minimal() +
  theme(
    plot.title       = element_text(
      face = "bold",
      size = 9,
      color = "#0a2535",
      hjust = 0.5),
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.title       = element_blank(),
    axis.ticks       = element_blank(),
    plot.background  = element_rect(fill = "#e7cfbf", color = NA),
    panel.background = element_rect(fill = "white", color = NA),
    panel.border = element_rect(
      color     = "#d9d9d9",
      fill      = NA,
      linewidth = 0.5
    )
  )
plot_brasil

plot_am <- ggplot() +
  geom_sf(data = contorno_am,
          fill = NA, color = "#0a2535", linewidth = 0.3) +
  geom_sf(data = Dados_am,
          fill = "#b5622a", color = "white", linewidth = 0.25) +
  annotation_north_arrow(
    location = "tl",
    which_north = "true",
    pad_x = unit(0.2, "cm"),
    pad_y = unit(0.2, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col = "#0a2535",
      line_col = "#0a2535",
      fill = c("#b5622a", "white")
    ),
    height = unit(1, "cm"),
    width  = unit(1, "cm")
  ) +
  annotation_scale(
    location   = "br",
    width_hint = 0.25,
    text_cex   = 0.6,
    line_width = 0.7
  ) +
  labs(title = "Amazonas") +
  theme_minimal() +
  theme(
    plot.title       = element_text(
      face = "bold",
      size = 9,
      color = "#0a2535",
      hjust = 0.5),
    panel.grid       = element_blank(),
    axis.text        = element_blank(),
    axis.title       = element_blank(),
    axis.ticks       = element_blank(),
    plot.background  = element_rect(fill = "#e7cfbf", color = NA),
    panel.background = element_rect(fill = "white", color = NA),
    panel.border = element_rect(
      color     = "#d9d9d9",
      fill      = NA,
      linewidth = 0.5
    )
  )
plot_am

mapa_indigena <- (plot_brasil + plot_am) &
  theme(
    plot.background  = element_rect(fill = "#e7cfbf", color = NA),
    plot.margin      = margin(4, 6, 4, 6)
  )

mapa_indigena <- mapa_indigena +
  plot_annotation(
    title    = "Terras Indígenas no Brasil e no Amazonas",
    subtitle = "Pacote geobr • read_indigenous_land() • Ano 2025",
    caption  = "Fonte: FUNAI / geobr",
    theme = theme(
      plot.title    = element_text(hjust = 0.5, face = "bold",
                                   size = 13, color = "#0a2535"),
      plot.subtitle = element_text(hjust = 0.5, color = "#4a6a7a"),
      plot.caption  = element_text(hjust = 0, size = 9, color = "gray40"),
      plot.background = element_rect(fill = "#e7cfbf", color = NA),
      plot.margin     = margin(8, 8, 8, 8)
    )
  )

mapa_indigena

ggsave(
  filename = "Avaliacao_9/media/mapa_indigena.png",
  plot     = mapa_indigena,
  width    = 10,
  height   = 4.5,
  dpi      = 150,
  bg       = "#e7cfbf"
)

