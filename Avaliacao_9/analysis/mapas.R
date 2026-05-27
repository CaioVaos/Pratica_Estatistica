#########Analise das variaveis numericas########################################

# Setup ------------------------------------------------------------------------
library(readxl)     # Leitura de arquivos excel
library(ggplot2)    # Gráficos
library(geobr)      # Shapefiles do Brasil
library(dplyr)      # Manipulação de bases de dados
# library(ggsflabel)  # Criação de Labels que se repelem
library(ggspatial)  # Rosa dos ventos e escala
library(sf)         # Leitura de shapefiles fora do geobr
library(ggiraph)    # Mapas interativos

remotes::install_github("ipeaGIT/geobr", subdir = "r-package")

devtools::install_github("yutannihilation/ggsflabel")
