##############Criar Tabela Testes de Independência##############################
# Setup ------------------------------------------------------------------------
require(openxlsx)            # Leitura de base de dados
require(dplyr)               # Manipulação de base de dados
require(gtsummary)           # Tabelas automáticas
require(gt)                  # Tabelas automáticas
require(rstatix)             # Coeficiente de Cramer

Dados = read.xlsx("Avaliacao_2/data/DadosAviao.xlsx")

# linguagem de pontuação em português
# formatação em português (vírgula pra decimais e ponto para milhares)
theme_gtsummary_language("pt", big.mark = ".", decimal.mark = ",")

# Adequando variáveis qualitativas ---------------------------------------------

DadosQuali = Dados %>% select(Conforto_Assento,Genero, Cliente, Idade,
                              Tipo_Viagem, Classe, Distancia,
                              Atraso_Partida, Atraso_Chegada)
DadosQuali <- DadosQuali %>%
  mutate(
    
    Idade = recode(Idade,
                   "<30 anos" = "Menor que 30 anos",
                   "30–50 anos" = "30 a 50 anos",
                   ">50 anos" = "Maior que 50 anos"
    ),
    Distancia= recode(Distancia,
                      
 "500–1500 km"= "500 a 1500 km",
                      ">1500 km"= "Maior que 1500 km",
                      " <500 km"="Menor que 500 km")
  )
# Tabelas de contingencia ------------------------------------------------------
tbl_summary(data = DadosQuali)

tbl_summary(data = DadosQuali,
            by = Conforto_Assento)

tbl_summary(data = DadosQuali,
            by = Conforto_Assento ,
            percent = "row")
# Verificação da Matriz Esperada -----------------------------------------------
chisq.test(Dados$Genero,Dados$Conforto_Assento)$expected

chisq.test(Dados$Cliente,Dados$Conforto_Assento)$expected

chisq.test(Dados$Idade,Dados$Conforto_Assento)$expected

chisq.test(Dados$Tipo_Viagem,Dados$Conforto_Assento)$expected

chisq.test(Dados$Classe,Dados$Conforto_Assento)$expected # menor que 5 -> fisher

chisq.test(Dados$Distancia,Dados$Conforto_Assento)$expected

chisq.test(Dados$Atraso_Partida,Dados$Conforto_Assento)$expected

chisq.test(Dados$Atraso_Chegada,Dados$Conforto_Assento)$expected

# Tabela com o pvalor ----------------------------------------------------------
tbl_summary(data = DadosQuali,
            by = Conforto_Assento,
            percent = "row")%>%
  add_p(
    pvalue_fun = function(x) formatC(x, digits = 4, format = "f")
  )
# Residuos ---------------------------------------------------------------------
# acima de 1,96 ou abaixo de -1,96 na normal padrao sao os caras influentes
chisq.test(Dados$Idade,Dados$Conforto_Assento)$stdres

chisq.test(Dados$Classe,Dados$Conforto_Assento)$stdres

chisq.test(Dados$Tipo_Viagem,Dados$Conforto_Assento)$stdres

chisq.test(Dados$Atraso_Partida,Dados$Conforto_Assento)$stdres

# Coeficiente de Cramer --------------------------------------------------------

cramer_fun <- function(data, variable, by, ...) {
  tab <- table(data[[variable]], data[[by]])
  v <- cramer_v(tab)
  tibble::tibble(`**Cramér**` = round(v, 3))
}

# Criando tabela ---------------------------------------------------------------
tabela <- tbl_summary(data = DadosQuali,
            by = Conforto_Assento,
            percent = "row",
            label = list(
              Genero ~ "Gênero<sup>Q</sup>",
              Cliente ~ "Tipo de Cliente<sup>Q</sup>",
              Idade ~ "Idade<sup>Q</sup>",
              Tipo_Viagem ~"Tipo de Viagem<sup>Q</sup>",
              Classe ~ "Classe<sup>F</sup>",
              Distancia ~ "Distância<sup>Q</sup>",
              Atraso_Partida ~"Atraso na Partida<sup>Q</sup>",
              Atraso_Chegada ~"Atraso na Chegada<sup>Q</sup>"
            )
)%>%
  add_p(pvalue_fun = label_style_pvalue(digits = 3)) %>%
  bold_p(t = 0.05) %>%
  add_stat(fns = everything() ~ cramer_fun)%>%
  modify_spanning_header(all_stat_cols() ~ "**Satisfação com o conforto do assento**") %>%
  modify_header(label ~ "**Variáveis**") %>%
  bold_labels() %>%
  modify_header(all_stat_cols() ~ "**{level}**<br>{n} ({style_percent(p)}%)")%>%
  modify_bold(
    columns = stat_1,  # primeira coluna
    rows = (variable == "Idade" & label == "Menor que 30 anos")| ( label== "Maior que 50 anos") | (variable == "Classe" & label == "Economica")| (label=="Executiva") |(variable== "Atraso_Partida" & label== "Sem atraso")  
  )%>%
  modify_bold( columns = stat_1,  # primeira coluna
   rows= (variable== "Atraso_Partida" & label== "Com atraso")) %>% 
  modify_bold(
    columns = stat_2,  # segunda coluna
    rows = (variable == "Classe" & label == "Executiva")| (variable== "Idade" & label== "Menor que 30 anos") |(variable== "Atraso_Partida" &  label== "Sem atraso"))%>%
    modify_bold( columns = stat_2,  # segunda coluna
                 rows= (variable== "Atraso_Partida" & label== "Com atraso")) %>% 
    modify_bold(
    columns = stat_3,  # terceira coluna
    rows =  (variable == "Idade" & label == "Menor que 30 anos")|( label== "Maior que 50 anos") | (variable == "Classe" & label == "Economica")| (label=="Executiva") |(variable== "Atraso_Partida" & label== "Sem atraso") | (variable== "Tipo_Viagem" & label== "Viagem a Negocios")| (label=="Viagem Pessoal") )%>%
      modify_bold( columns = stat_3,  # terceira coluna
                   rows= (variable== "Atraso_Partida" & label== "Com atraso")) %>% 
  modify_footnote(everything() ~ NA)%>%
  as_gt() %>%                   
  fmt_markdown(columns = label) %>%
  tab_options(
    table.font.size = "20px",    
    heading.title.font.size = "26px",
    column_labels.font.size = "22px"
  ) %>% 
  tab_source_note(
    source_note = md("**Q**: Teste do qui-quadrado de Pearson; <br>
                     **F**: Teste exato de Fisher; <br>
                     Valores em negrito na coluna *Valor-p* indicam significância estatística (p < 0,05); <br>
                     Valores em negrito nas tabelas de contingência indicam células com resíduos padronizados elevados.")
  )

## Colorindo -------------------------------------------------------------------

tabela <- tabela %>%
  tab_style(
    style = list(
      cell_fill(color = "#fef9c3")
    ),
    locations = cells_body(
      rows = variable %in% c("Classe", "Idade", "Tipo_Viagem", "Atraso_Partida") & 
        row_type == "label"
    )
  )

tabela <- tabela %>%
  tab_style(
    style = cell_fill(color = "#dbeafe"),
    locations = cells_body(
      columns = stat_1,
      rows = (variable == "Idade"          & label == "Maior que 50 anos") |
        (variable == "Classe"         & label == "Executiva")         |
        (variable == "Atraso_Partida" & label == "Sem atraso")
    )
  ) %>%
  tab_style(
    style = cell_fill(color = "#dbeafe"),
    locations = cells_body(
      columns = stat_3,
      rows = (variable == "Idade"          & label == "Menor que 30 anos") |
        (variable == "Tipo_Viagem"    & label == "Viagem Pessoal")    |
        (variable == "Classe"         & label == "Economica")         |
        (variable == "Atraso_Partida" & label == "Com atraso")
    )
  ) %>%
  tab_style(
    style = cell_fill(color = "#dcfce7"),
    locations = cells_body(
      columns = stat_1,
      rows = (variable == "Idade"          & label == "Menor que 30 anos") |
        (variable == "Classe"         & label == "Economica")         |
        (variable == "Atraso_Partida" & label == "Com atraso")
    )
  ) %>%
  tab_style(
    style = cell_fill(color = "#dcfce7"),
    locations = cells_body(
      columns = stat_3,
      rows = (variable == "Idade"          & label == "Maior que 50 anos")   |
        (variable == "Tipo_Viagem"    & label == "Viagem a Negocios")   |
        (variable == "Classe"         & label == "Executiva")           |
        (variable == "Atraso_Partida" & label == "Sem atraso")
    )
  )

## Salvar ----------------------------------------------------------------------
gt::gtsave(tabela, "Avaliacao_2/media/tabela.html")
