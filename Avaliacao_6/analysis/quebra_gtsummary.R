# Categorica -------------------------------------------------------------------

library(stringr)

# Gera o HTML completo da tabela
# tabela_0 <- tabela_categorica

html_completo <- tabela_categorica %>% as_raw_html()

# Extrai o cabeçalho (thead)
cabecalho <- str_extract(html_completo, "(?s)<thead[^>]*>.*?</thead>")

# Extrai o corpo (tbody)
corpo     <- str_extract(html_completo, "(?s)<tbody[^>]*>.*?</tbody>")

# Extrai o rodapé (tfoot)
rodape    <- str_extract(html_completo, "(?s)<tfoot[^>]*>.*?</tfoot>")

# Salva separadamente
saveRDS(cabecalho, "Avaliacao_6/data/tabela_categorica_cabecalho.rds")
saveRDS(corpo,     "Avaliacao_6/data/tabela_categorica_corpo.rds")
saveRDS(rodape,    "Avaliacao_6/data/tabela_categorica_rodape.rds")

# Numericas --------------------------------------------------------------------

library(stringr)

# Gera o HTML completo da tabela
# tabela_0 <- tabela_numerica

html_completo <- tabela_numerica %>% as_raw_html()

# Extrai o cabeçalho (thead)
cabecalho <- str_extract(html_completo, "(?s)<thead[^>]*>.*?</thead>")

# Extrai o corpo (tbody)
corpo     <- str_extract(html_completo, "(?s)<tbody[^>]*>.*?</tbody>")

# Extrai o rodapé (tfoot)
rodape    <- str_extract(html_completo, "(?s)<tfoot[^>]*>.*?</tfoot>")

# Salva separadamente
saveRDS(cabecalho, "Avaliacao_6/data/tabela_numerica_cabecalho.rds")
saveRDS(corpo,     "Avaliacao_6/data/tabela_numerica_corpo.rds")
saveRDS(rodape,    "Avaliacao_6/data/tabela_numerica_rodape.rds")
