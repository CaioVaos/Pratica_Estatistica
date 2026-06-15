# Setup ----------------------------------------------------
library(tidyverse)
library(plotly)
# Sys.unsetenv("GITHUB_PAT")
# remotes::install_github("ipeaGIT/geobr", subdir = "r-package")
library(geobr)
# devtools::install_github("yutannihilation/ggsflabel")
library(ggsflabel)
library(ggiraph)
library(sf)
library(ggspatial)
library(patchwork)

df <- readxl::read_xlsx("data/MortalidadeCancer.xlsx") |> 
  filter(Cancer == "Câncer de Estomago")
df <- df |>
  mutate(codigo = case_when(
    local == "Centro-Oeste" ~ 5,
    local == "Sudeste"      ~ 3,
    local == "Sul"          ~ 4,
    .default = codigo
  ))

# Serie temporal BR -------------------------------------------

plot <- df |>
filter(local == "Brasil", faixa == "todas as idades", sexo == "Ambos") |>
ggplot(aes(x = ano, y = n)) +
geom_line(aes(color = "Observado"), linewidth = 1.2) +
geom_point(aes(color = "Observado"), size = 2, alpha = 0.8) +
geom_smooth(method = "loess", se = FALSE, aes(color = "Tendência (LOESS)"),
            linetype = "dashed", linewidth = 0.8) +
annotate("rect", xmin = 2020, xmax = 2022, ymin = -Inf, ymax = Inf,
          alpha = 0.08, fill = "gray70") +
annotate("text", x = 2021, y = 11000, label = "COVID-19", 
          size = 3, color = "gray40", fontface = "italic") +
scale_x_continuous(breaks = seq(2001, 2024, by = 3), expand = c(0, 0.5)) +
scale_y_continuous(labels = scales::comma) +
scale_color_manual(name = NULL,
                    values = c("Observado" = "#2c3e50", 
                              "Tendência (LOESS)" = "#f39c12")) +
labs(
  title = "Total de Óbitos no Brasil",
  subtitle = "Evolução anual (2001–2024) | Todas as idades, ambos os sexos",
  x = NULL, y = NULL
) +
theme_minimal(base_size = 12) +
theme(
  plot.title = element_text(face = "bold", size = 15, hjust = 0.5),
  plot.subtitle = element_text(color = "gray40", hjust = 0.5),
  plot.caption = element_text(color = "gray50", size = 8, hjust = 1),
  panel.grid.major.y = element_line(color = "gray90", linewidth = 0.4),
  panel.grid.major.x = element_blank(),
  panel.grid.minor = element_blank(),
  axis.ticks.x = element_line(color = "gray70"),
  axis.text = element_text(color = "gray30"),
  plot.margin = margin(10, 15, 10, 10),
  legend.position = c(0, 1),
  legend.justification = c(0, 1)
)
plot

# Barras -------------------------------------------

estados <- c(
  "Acre",
  "Alagoas",
  "Amapá",
  "Amazonas",
  "Bahia",
  "Ceará",
  "Distrito Federal",
  "Espírito Santo",
  "Goiás",
  "Maranhão",
  "Mato Grosso",
  "Mato Grosso do Sul",
  "Minas Gerais",
  "Pará",
  "Paraíba",
  "Paraná",
  "Pernambuco",
  "Piauí",
  "Rio de Janeiro (Estado)",
  "Rio Grande do Norte",
  "Rio Grande do Sul",
  "Rondônia",
  "Roraima",
  "Santa Catarina",
  "São Paulo",
  "Sergipe",
  "Tocantins"
)

regioes <- c(
  "Norte",
  "Nordeste",
  "Centro-Oeste",
  "Sudeste",
  "Sul"
)
brasil <- "Brasil"

territorio <- regioes



taxa_max <- max(df_plot$taxa)

ggplot(df_plot, aes(x = local, y = taxa)) +
  geom_col(
    fill  = "#b91c1c",
    width = 0.7
  ) +
  geom_text(
    aes(label = round(taxa, 1)),
    hjust    = 1.15,
    y        = df_plot$taxa - max(df_plot$taxa) * 0.02,
    color    = "white",
    fontface = "bold",
    size     = 3.2
  ) +
  scale_y_continuous(
    expand = expansion(mult = c(0, 0.08)),
    labels = scales::number_format(accuracy = 0.1)
  ) +
  coord_flip(clip = "off") +
  labs(
    title = glue::glue("{input$territorio} · {input$ano}"),
    x     = NULL,
    y     = "Taxa por 100 mil habitantes"
  ) +
  theme_minimal(base_size = 11) +
  theme(
    plot.title          = element_text(face = "bold", size = 11,
                                       color = "#1e293b", hjust = 0.5,
                                       margin = margin(b = 8)),
    plot.title.position = "plot",
    plot.margin         = margin(8, 16, 8, 8),
    panel.grid.major.y  = element_blank(),
    panel.grid.major.x  = element_line(color = "#e2e8f0", linewidth = 0.5),
    panel.grid.minor    = element_blank(),
    axis.text.y         = element_text(size = 9, color = "#334155", face = "bold"),
    axis.text.x         = element_text(size = 8, color = "#64748b"),
    axis.title.x        = element_text(color = "#64748b"),
    axis.ticks          = element_blank()
  )

# Gapminder ----------------------------------

ESTADOS <- c(
  "Acre", "Alagoas", "Amapá", "Amazonas", "Bahia", "Ceará",
  "Distrito Federal", "Espírito Santo", "Goiás", "Maranhão",
  "Mato Grosso", "Mato Grosso do Sul", "Minas Gerais", "Pará",
  "Paraíba", "Paraná", "Pernambuco", "Piauí", "Rio de Janeiro (Estado)",
  "Rio Grande do Norte", "Rio Grande do Sul", "Rondônia", "Roraima",
  "Santa Catarina", "São Paulo", "Sergipe", "Tocantins"
)

regiao_map <- c(
  "Acre" = "Norte", "Amapá" = "Norte", "Amazonas" = "Norte",
  "Pará" = "Norte", "Rondônia" = "Norte", "Roraima" = "Norte",
  "Tocantins" = "Norte", "Alagoas" = "Nordeste", "Bahia" = "Nordeste",
  "Ceará" = "Nordeste", "Maranhão" = "Nordeste", "Paraíba" = "Nordeste",
  "Pernambuco" = "Nordeste", "Piauí" = "Nordeste",
  "Rio Grande do Norte" = "Nordeste", "Sergipe" = "Nordeste",
  "Distrito Federal" = "Centro-Oeste", "Goiás" = "Centro-Oeste",
  "Mato Grosso" = "Centro-Oeste", "Mato Grosso do Sul" = "Centro-Oeste",
  "Espírito Santo" = "Sudeste", "Minas Gerais" = "Sudeste",
  "Rio de Janeiro (Estado)" = "Sudeste", "São Paulo" = "Sudeste",
  "Paraná" = "Sul", "Rio Grande do Sul" = "Sul", "Santa Catarina" = "Sul"
)

df_gapminder <- df |>
  filter(
    local %in% ESTADOS,
    faixa == "todas as idades",
    sexo  == "Ambos"
  ) |>
  mutate(
    taxa   = (n / populacao) * 100000,
    regiao = regiao_map[local],
    local  = reorder(local, taxa)          # ordena por taxa média
  ) |>
  arrange(ano, desc(taxa))

plot_ly(
  df_gapminder,
  x         = ~taxa,
  y         = ~reorder(local, taxa),
  frame     = ~ano,
  type      = "bar",
  orientation = "h",
  color     = ~regiao,
  colors    = c(
    "Norte"       = "#2196F3",
    "Nordeste"    = "#ae8b2d",
    "Centro-Oeste"= "#27ae60",
    "Sudeste"     = "#b91c1c",
    "Sul"         = "#9b59b6"
  ),
  text      = ~paste0(round(taxa, 1), " por 100 mil"),
  textposition = "outside",
  hovertemplate = "<b>%{y}</b><br>Taxa: %{x:.1f} por 100 mil<extra></extra>"
) |>
  layout(
    title = list(
      text = "Taxa de Mortalidade por Câncer de Estômago — Estados",
      font = list(size = 14, color = "#1e293b")
    ),
    xaxis = list(
      title      = "Taxa por 100 mil habitantes",
      showgrid   = TRUE,
      gridcolor  = "#e2e8f0",
      zeroline   = FALSE
    ),
    yaxis = list(
      title    = "",
      tickfont = list(size = 10)
    ),
    legend = list(
      title       = list(text = "Região"),
      orientation = "v"
    ),
    plot_bgcolor  = "white",
    paper_bgcolor = "white",
    margin        = list(l = 180, r = 60, t = 60, b = 50)
  ) |>
  animation_opts(
    frame      = 800,
    transition = 500,
    easing     = "cubic-in-out",
    redraw     = FALSE
  ) |>
  animation_slider(
    currentvalue = list(
      prefix = "Ano: ",
      font   = list(color = "#ae8b2d", size = 14)
    )
  ) |>
  animation_button(
    x = 1, xanchor = "right",
    y = 0, yanchor = "bottom",
    label = "▶ Play"
  )

#  mapa --------------------------------------

estados <- c(
  "Acre",
  "Alagoas",
  "Amapá",
  "Amazonas",
  "Bahia",
  "Ceará",
  "Distrito Federal",
  "Espírito Santo",
  "Goiás",
  "Maranhão",
  "Mato Grosso",
  "Mato Grosso do Sul",
  "Minas Gerais",
  "Pará",
  "Paraíba",
  "Paraná",
  "Pernambuco",
  "Piauí",
  "Rio de Janeiro (Estado)",
  "Rio Grande do Norte",
  "Rio Grande do Sul",
  "Rondônia",
  "Roraima",
  "Santa Catarina",
  "São Paulo",
  "Sergipe",
  "Tocantins"
)

regioes <- c(
  "Norte",
  "Nordeste",
  "Centro-Oeste",
  "Sudeste",
  "Sul"
)
brasil <- "Brasil"

input <- data.frame(territorio = "Regional")
territorio <- switch(
    input$territorio,
    "Estadual" = estados,
    "Regional" = regioes,
    "Brasil" = brasil
  )
ano_selecionado <-  2024

# shp_br <- geobr::read_country(year = 2020) |> 
#   mutate(codigo = 0) |> 
#   select(codigo, geometry)
# shp_estate <- geobr::read_state(year = 2024) |> 
#   rename(codigo = "code_state") |> 
#   select(codigo, geometry)
# shp_region <- geobr::read_region(year = 2020) |> 
#   rename(codigo = "code_region") |> 
#   select(codigo, geometry)
# shapes <- rbind(shp_br, shp_estate, shp_region)
# saveRDS(shapes, "data/shapes.rds")
shapes <- readRDS("data/shapes.rds")

df_mapa_br <- df |>
  filter(
    local %in% territorio,
    ano == ano_selecionado,
    sexo == "Ambos",
    faixa == "todas as idades"
  ) |> 
  mutate(
    taxa = (n/populacao)*100000
  ) |> 
  left_join(shapes)

plot_mapa_br <- ggplot() +
  geom_sf(
    data  = df_mapa_br,
    aes(fill = taxa, geometry = geometry),
    color = "white", linewidth = 0.15
  ) +
  geom_sf(
    data = shapes |> filter(
      codigo %in% switch(input$territorio,
        "Brasil"   = integer(0),
        "Regional" = c(1, 2, 3, 4, 5),
        "Estadual" = unique(shapes$codigo[!shapes$codigo %in% 0:5])
      )
    ),
    fill = NA, color = "#94a3b8", linewidth = 0.3
  ) +
  geom_sf(
    data = shapes |> filter(codigo == 0),
    fill = NA, color = "#1e293b", linewidth = 1
  ) +
  scale_fill_gradientn(
    colors   = c("#fff1f1", "#fca5a5", "#ef4444", "#b91c1c", "#7f1d1d"),
    na.value = "#e2e8f0",
    name     = "Taxa por\n100 mil hab.",
    guide    = guide_colorbar(
      barwidth     = unit(0.35, "cm"),
      barheight    = unit(4, "cm"),
      ticks        = FALSE,
      frame.colour = "#cbd5e1",
      frame.linewidth = 0.4,
      title.hjust  = 0.5,
      title.theme  = element_text(
        size = 6.5, color = "#334155", face = "bold", lineheight = 1.4
      ),
      label.theme  = element_text(size = 6.5, color = "#64748b")
    )
  ) +
  annotation_north_arrow(
    location    = "bl",
    which_north = "true",
    pad_x = unit(0.4, "cm"),
    pad_y = unit(0.4, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col  = "#334155",
      line_col  = "#334155",
      fill      = c("#b91c1c", "#f8fafc"),
      text_size = 6
    ),
    height = unit(1.0, "cm"),
    width  = unit(1.0, "cm")
  ) +
  annotation_scale(
    location   = "br",
    width_hint = 0.2,
    text_cex   = 0.5,
    line_width = 0.5,
    text_col   = "#64748b",
    line_col   = "#94a3b8"
  ) +
  labs(title = glue::glue("{input$territorio} · {input$ano}")) +
  theme_void(base_size = 11) +
  theme(
    plot.title.position = "plot",
    plot.title           = element_text(
      face = "bold", size = 10, color = "#1e293b",
      hjust = 0.5, margin = margin(b = 4)
    ),
    plot.margin          = margin(8, 48, 8, 8),
    legend.position      = "right",
    legend.justification = "center",
    legend.margin        = margin(0, 0, 0, 4)
  )

plot_mapa_br

#  Serie --------------------------------------

estados <- c(
  "Acre",
  "Alagoas",
  "Amapá",
  "Amazonas",
  "Bahia",
  "Ceará",
  "Distrito Federal",
  "Espírito Santo",
  "Goiás",
  "Maranhão",
  "Mato Grosso",
  "Mato Grosso do Sul",
  "Minas Gerais",
  "Pará",
  "Paraíba",
  "Paraná",
  "Pernambuco",
  "Piauí",
  "Rio de Janeiro (Estado)",
  "Rio Grande do Norte",
  "Rio Grande do Sul",
  "Rondônia",
  "Roraima",
  "Santa Catarina",
  "São Paulo",
  "Sergipe",
  "Tocantins"
)

regioes <- c(
  "Norte",
  "Nordeste",
  "Centro-Oeste",
  "Sudeste",
  "Sul"
)

brasil <- "Brasil"


territorio <- estados
territorio <- regioes
territorio <- brasil

df_plot <- df |>
  filter(
    local %in% territorio,
    faixa == "todas as idades",
    sexo  == "Ambos"
  ) |>
  group_by(local, ano) |>
  summarise(
    obitos    = sum(n,         na.rm = TRUE),
    populacao = sum(populacao, na.rm = TRUE),
    .groups = "drop"
  ) |>
  mutate(taxa = (obitos / populacao) * 100000)

n_locais <- n_distinct(df_plot$local)
cores_locais <- setNames(
  colorRampPalette(c("#F72585","#7209B7","#4361EE","#4CC9F0","#06D6A0"))(n_locais),
  unique(df_plot$local)
)

plot <- ggplot(df_plot, aes(x = ano, y = taxa, color = local, group = local)) +
  geom_line(linewidth = 1.1, alpha = 0.9) +
  geom_point(size = 1.8, alpha = 0.75) +
  geom_smooth(method = "loess", se = FALSE,
              linetype = "dashed", linewidth = 0.7, alpha = 0.6) +
  annotate("rect", xmin = 2020, xmax = 2022,
           ymin = -Inf, ymax = Inf,
           alpha = 0.08, fill = "gray70") +
  annotate("text", x = 2021, y = Inf, label = "COVID-19",
           vjust = 1.5, size = 3, color = "gray40", fontface = "italic") +
  scale_x_continuous(breaks = seq(2001, 2024, by = 3), expand = c(0, 0.5)) +
  scale_y_continuous(labels = scales::comma) +
  scale_color_manual(values = cores_locais, name = NULL) +
  labs(
    title = "Taxa de Mortalidade por 100 mil habitantes",
    x = NULL, y = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title         = element_text(face = "bold", size = 15, hjust = 0.5),
    panel.grid.major.y = element_line(color = "gray90", linewidth = 0.4),
    panel.grid.major.x = element_blank(),
    panel.grid.minor   = element_blank(),
    axis.ticks.x       = element_line(color = "gray70"),
    axis.text          = element_text(color = "gray30"),
    plot.margin        = margin(10, 15, 10, 10),
    legend.position    = if (length(territorio) == 1) "none" else "right",
    legend.text        = element_text(size = 9)
  )

plot
ggplotly(plot)

# Piramide --------------------------------------
ano_selecionado  <- 2010
local_selecionado <- "Norte"

niveis_faixa <- c(
  "0-4 anos","5-9 anos","10-14 anos","15-19 anos",
  "20-24 anos","25-29 anos","30-34 anos","35-39 anos",
  "40-44 anos","45-49 anos","50-54 anos","55-59 anos",
  "60-64 anos","65-69 anos","70-74 anos","75-79 anos","80-mais"
)
labels_faixa <- c(
  "0-4","5-9","10-14","15-19",
  "20-24","25-29","30-34","35-39",
  "40-44","45-49","50-54","55-59",
  "60-64","65-69","70-74","75-79","80+"
)

df_plot <- df |>
  filter(
    local  == local_selecionado,
    ano    == ano_selecionado,
    faixa  != "todas as idades",
    sexo   != "Ambos"
  ) |>
  mutate(
    idade_rotulada = factor(faixa, levels = niveis_faixa, labels = labels_faixa)
  )

df_plot$idade_rotulada <- factor(df_plot$idade_rotulada,
                                  levels = rev(labels_faixa))

df_masc   <- filter(df_plot, sexo == "Masculino")
df_fem    <- filter(df_plot, sexo == "Feminino")
df_centro <- distinct(df_plot, idade_rotulada)

lim_x <- max(df_plot$n, na.rm = TRUE) * 1.08
breaks_x <- pretty(c(0, lim_x), n = 5)

## painel masculino ----
p_masc <- ggplot(df_masc, aes(x = n, y = idade_rotulada)) +
  geom_bar(stat = "identity", width = 0.75, fill = "#2196F3", alpha = 0.92) +
  geom_text(aes(label = n), hjust = 1.25,
            size = 2.8, color = "white", fontface = "bold") +
  scale_x_reverse(
    breaks = breaks_x, labels = breaks_x,
    limits = c(lim_x, 0),
    expand = expansion(mult = c(0.02, 0))
  ) +
  labs(title = "Masculino", x = NULL, y = NULL) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title         = element_text(face = "bold", hjust = 0.5,
                                      color = "#2196F3", size = 13),
    axis.text.y        = element_blank(),
    axis.ticks.y       = element_blank(),
    axis.text.x        = element_text(color = "gray50", size = 9),
    panel.grid.major.y = element_blank(),
    panel.grid.minor   = element_blank(),
    plot.margin        = margin(5, 0, 5, 10)
  )

## painel datas -----
p_centro <- ggplot(df_centro, aes(x = 0.5, y = idade_rotulada)) +
  geom_text(aes(label = idade_rotulada),
            color    = "#333333",
            size     = 2.9,
            fontface = "bold") +
  scale_x_continuous(limits = c(0, 1), expand = c(0, 0)) +
  labs(title = "Idade", x = NULL, y = NULL) +
  theme_void() +
  theme(
    plot.title  = element_text(face = "bold", hjust = 0.5,
                               color = "gray40", size = 13),
    plot.margin = margin(5, 2, 5, 2)
  )

## painel feminino ----
p_fem <- ggplot(df_fem, aes(x = n, y = idade_rotulada)) +
  geom_bar(stat = "identity", width = 0.75, fill = "#E91E63", alpha = 0.92) +
  geom_text(aes(label = n), hjust = -0.25,
            size = 2.8, color = "white", fontface = "bold") +
  scale_x_continuous(
    breaks = breaks_x, labels = breaks_x,
    limits = c(0, lim_x),
    expand = expansion(mult = c(0, 0.02))
  ) +
  labs(title = "Feminino", x = NULL, y = NULL) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title         = element_text(face = "bold", hjust = 0.5,
                                      color = "#E91E63", size = 13),
    axis.text.y        = element_blank(),
    axis.ticks.y       = element_blank(),
    axis.text.x        = element_text(color = "gray50", size = 9),
    panel.grid.major.y = element_blank(),
    panel.grid.minor   = element_blank(),
    plot.margin        = margin(5, 10, 5, 0)
  )

## Montar ----
(p_masc | p_centro | p_fem) +
  plot_annotation(
    title    = glue::glue("Pirâmide Etária de Óbitos"),
    subtitle = glue::glue("{local} · {ano} · "),
    theme = theme(
      plot.title    = element_text(face = "bold", hjust = 0.5, size = 17),
      plot.subtitle = element_text(hjust = 0.5, color = "gray50", size = 10)
    )
  ) +
  plot_layout(widths = c(5, 1.3, 5))

# Mapa RJ ----

# shape_rj <- read_municipality(year = 2024, code_muni = "RJ") |> 
#   mutate(code_muni = code_muni %/% 10) |> 
#   rename(codigo = "code_muni") |> 
#   select(codigo, geometry)
# saveRDS(shape_rj, "data/shape_rj.rds")
shape_rj <- readRDS("data/shape_rj.rds")
shapes <- readRDS("data/shapes.rds")

df_mapa_rj <- shape_rj |>
  left_join(
    df |>
      filter(
        ano   == 2024,
        sexo  == "Ambos",
        faixa == "todas as idades"
      ) |>
      mutate(taxa = (n / populacao) * 100000),
    by = "codigo"
  )

plot_mapa_rj <- ggplot() +
  geom_sf_interactive(
    data = df_mapa_rj,
    aes(
      fill    = taxa,
      geometry = geometry,
      tooltip = paste0("<b>", local, "</b><br>Taxa: ", round(taxa, 1), " por 100 mil"),
      data_id = codigo
    ),
    color = "white", linewidth = 0.15
  ) +
  geom_sf(
    data = shape_rj,
    fill = NA, color = "#94a3b8", linewidth = 0.3
  ) +
  geom_sf(
    data = shapes |> filter(codigo == 33),
    fill = NA, color = "#1e293b", linewidth = 0.8
  ) +
  scale_fill_gradientn(
    colors   = c("#fff1f1", "#fca5a5", "#ef4444", "#b91c1c", "#7f1d1d"),
    na.value = "#e2e8f0",
    name     = "Taxa por\n100 mil hab.",
    guide    = guide_colorbar(
      barwidth        = unit(0.35, "cm"),
      barheight       = unit(4, "cm"),
      ticks           = FALSE,
      frame.colour    = "#cbd5e1",
      frame.linewidth = 0.4,
      title.hjust     = 0.5,
      title.theme     = element_text(size = 6.5, color = "#334155",
                                     face = "bold", lineheight = 1.4),
      label.theme     = element_text(size = 6.5, color = "#64748b")
    )
  ) +
  annotation_north_arrow(
    location    = "bl", which_north = "true",
    pad_x = unit(0.4, "cm"), pad_y = unit(0.4, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col  = "#334155", line_col = "#334155",
      fill      = c("#b91c1c", "#f8fafc"), text_size = 6
    ),
    height = unit(1.0, "cm"), width = unit(1.0, "cm")
  ) +
  annotation_scale(
    location = "br", width_hint = 0.2,
    text_cex = 0.5, line_width = 0.5,
    text_col = "#64748b", line_col = "#94a3b8"
  ) +
  labs(title = glue::glue("Rio de Janeiro · input$ano_rj")) +
  theme_void(base_size = 11) +
  theme(
    plot.title.position  = "plot",
    plot.title           = element_text(face = "bold", size = 10,
                                        color = "#1e293b", hjust = 0.5,
                                        margin = margin(b = 4)),
    plot.margin          = margin(8, 48, 8, 8),
    legend.position      = "right",
    legend.justification = "center",
    legend.margin        = margin(0, 0, 0, 4)
  )

girafe(
  ggobj  = plot_mapa_rj,
  width_svg  = 6,
  height_svg = 5,
  options = list(
    opts_hover(css = "fill-opacity:0.75; stroke:#1e293b; stroke-width:1.5px;"),
    opts_hover_inv(css = "fill-opacity:0.2;"),
    opts_tooltip(
      css       = "background:#1e293b; color:white; padding:6px 10px;
                   border-radius:5px; font-size:12px; line-height:1.5;",
      use_fill  = FALSE,
      delay_mouseout = 500
    ),
    opts_sizing(rescale = TRUE)
  )
)

## Saude ----

shape_saude <- geobr::read_health_region(year = 2025, code_state = 33, geometry_level = "micro") %>% 
  filter(!st_is_empty(geometry)) %>% 
  select(name_health_region, geometry) %>%
  mutate(name_health_region = iconv(name_health_region, from = "latin1", to = "UTF-8"))
saveRDS(shape_saude, "data/shape_saude.rds")
shape_saude <- readRDS("data/shape_saude.rds")

shape_rj <- readRDS("data/shape_rj.rds")
shapes <- readRDS("data/shapes.rds")

df_mapa_rj <- shape_rj |>
  left_join(
    df |>
      filter(
        ano   == 2024,
        sexo  == "Ambos",
        faixa == "todas as idades"
      ) |>
      mutate(taxa = (n / populacao) * 100000),
    by = "codigo"
  )

plot_mapa_rj <- ggplot() +
  geom_sf_interactive(
    data = df_mapa_rj,
    aes(
      fill     = taxa,
      geometry = geometry,
      tooltip  = paste0("<b>", local, "</b><br>Taxa: ", round(taxa, 1), " por 100 mil"),
      data_id  = codigo
    ),
    color = "white", linewidth = 0.15
  ) +
  geom_sf(
    data = shape_rj,
    fill = NA, color = "#94a3b8", linewidth = 0.3
  ) +
  # ── Camada regiões de saúde ──────────────────────────────────────────────
  geom_sf(
    data      = shape_saude,
    fill      = NA,
    color     = "#0f172a",
    linewidth = 2,
    linetype  = "dashed"
  ) +
  geom_sf_label(
    data  = shape_saude,
    aes(label = name_health_region),        # ajuste o campo se necessário
    size        = 2.2,
    color       = "#0f172a",
    fill        = alpha("white", 0.65),
    label.size  = 0,
    label.padding = unit(0.15, "lines"),
    fontface    = "bold",
    lineheight  = 0.9,
    check_overlap = TRUE
  ) +
  # ────────────────────────────────────────────────────────────────────────
  geom_sf(
    data = shapes |> filter(codigo == 33),
    fill = NA, color = "#1e293b", linewidth = 0.8
  ) +
  scale_fill_gradientn(
    colors   = c("#fff1f1", "#fca5a5", "#ef4444", "#b91c1c", "#7f1d1d"),
    na.value = "#e2e8f0",
    name     = "Taxa por\n100 mil hab.",
    guide    = guide_colorbar(
      barwidth        = unit(0.35, "cm"),
      barheight       = unit(4, "cm"),
      ticks           = FALSE,
      frame.colour    = "#cbd5e1",
      frame.linewidth = 0.4,
      title.hjust     = 0.5,
      title.theme     = element_text(size = 6.5, color = "#334155",
                                     face = "bold", lineheight = 1.4),
      label.theme     = element_text(size = 6.5, color = "#64748b")
    )
  ) +
  annotation_north_arrow(
    location    = "bl", which_north = "true",
    pad_x = unit(0.4, "cm"), pad_y = unit(0.4, "cm"),
    style = north_arrow_fancy_orienteering(
      text_col  = "#334155", line_col = "#334155",
      fill      = c("#b91c1c", "#f8fafc"), text_size = 6
    ),
    height = unit(1.0, "cm"), width = unit(1.0, "cm")
  ) +
  annotation_scale(
    location = "br", width_hint = 0.2,
    text_cex = 0.5, line_width = 0.5,
    text_col = "#64748b", line_col = "#94a3b8"
  ) +
  labs(title = glue::glue("Rio de Janeiro · input$ano_rj")) +
  theme_void(base_size = 11) +
  theme(
    plot.title.position  = "plot",
    plot.title           = element_text(face = "bold", size = 10,
                                        color = "#1e293b", hjust = 0.5,
                                        margin = margin(b = 4)),
    plot.margin          = margin(8, 48, 8, 8),
    legend.position      = "right",
    legend.justification = "center",
    legend.margin        = margin(0, 0, 0, 4)
  )
plot_mapa_rj

# Serie RJ ----------

muni_selecionado <- "Rio de Janeiro (Município)"

df_serie_rj <- df |>
  filter(
    local %in% muni_selecionado,
    faixa == "todas as idades",
    sexo  == "Ambos"
  ) |>
  mutate(taxa = (n / populacao) * 100000)

ggplot(df_serie_rj, aes(x = ano, y = taxa)) +
  geom_line(aes(color = "Observado"), linewidth = 1.2) +
  geom_point(aes(color = "Observado"), size = 2, alpha = 0.8) +
  geom_smooth(
    method   = "loess", se = FALSE,
    aes(color = "Tendência (LOESS)"),
    linetype = "dashed", linewidth = 0.8
  ) +
  annotate("rect",
    xmin = 2020, xmax = 2022, ymin = -Inf, ymax = Inf,
    alpha = 0.08, fill = "gray70"
  ) +
  annotate("text",
    x = 2021, y = Inf, label = "COVID-19",
    vjust = 1.5, size = 3, color = "gray40", fontface = "italic"
  ) +
  scale_x_continuous(breaks = seq(2001, 2024, by = 3), expand = c(0, 0.5)) +
  scale_y_continuous(labels = scales::number_format(big.mark = "")) +
  scale_color_manual(
    name   = NULL,
    values = c("Observado" = "#2c3e50", "Tendência (LOESS)" = "#f39c12")
  ) +
  labs(
    title    = glue::glue("Taxa de Mortalidade · {muni_selecionado}"),
    subtitle = "Evolução anual (2001–2024) | Todas as idades, ambos os sexos",
    x = NULL, y = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title           = element_text(face = "bold", size = 15, hjust = 0.5),
    plot.subtitle        = element_text(color = "gray40", hjust = 0.5),
    panel.grid.major.y   = element_line(color = "gray90", linewidth = 0.4),
    panel.grid.major.x   = element_blank(),
    panel.grid.minor     = element_blank(),
    axis.ticks.x         = element_line(color = "gray70"),
    axis.text            = element_text(color = "gray30"),
    plot.margin          = margin(10, 15, 10, 10),
    legend.position      = "right",
    legend.justification = "center",
  )
