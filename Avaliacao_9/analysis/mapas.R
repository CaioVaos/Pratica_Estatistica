#########Analise das variaveis numericas########################################

# Setup ------------------------------------------------------------------------
library(readxl)     # Leitura de arquivos excel
library(ggplot2)    # Gráficos
library(geobr)      # Shapefiles do Brasil
library(dplyr)      # Manipulação de bases de dados
# library(ggsflabel)  # Criação de Labels que se repelem
library(ggspatial)  # Rosa dos ventos e escala
library(sf)         # Leitura de shapefiles fora do geobr

# Sys.unsetenv("GITHUB_PAT")
remotes::install_github("ipeaGIT/geobr", subdir = "r-package")

devtools::install_github("yutannihilation/ggsflabel")

# read_health_region -----------------------------------------------------------
args(read_health_region)

Dados <- read_health_region(year = 2025,
                            geometry_level = "micro",
                            simplified = T)
Dados <- Dados[!st_is_empty(Dados),]
ggplot(Dados)+geom_sf()

# read_indigenous_land -----------------------------------------------------------
args(read_indigenous_land)
