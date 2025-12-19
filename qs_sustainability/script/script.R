if (!require("readxl")) install.packages("readxl")
if (!require("dplyr")) install.packages("dplyr")
if (!require("stringr")) install.packages("stringr")

library(readxl)
library(dplyr)
library(stringr)

setwd("/home/uff-lagef/Desktop/qs/qs_sustainability")

arquivos <- c(
  "2024 QS Sustainability Rankings.xlsx",
  "2025 QS Sustainability Rankings.xlsx",
  "2026 QS Sustainability Rankings.xlsx"
)

processar_arquivo <- function(arquivo) {
  ano <- gsub("\\D", "", arquivo)
  
  cat(sprintf("\n=== Processando %s ===\n", arquivo))
  
  dados <- read_excel(arquivo)
  
  cat("Colunas encontradas:", paste(names(dados), collapse = ", "), "\n")
  
  names(dados) <- gsub("\\s+|\\.", "_", names(dados))
  
  if ("Overall" %in% names(dados) && !"Overall_Score" %in% names(dados)) {
    dados <- dados %>% rename(Overall_Score = Overall)
  }
  
  dados_brasil <- dados %>%
    filter(
      grepl("brazil|brasil", Location, ignore.case = TRUE) |
        grepl("brazil|brasil", Institution, ignore.case = TRUE)
    )
  
  return(list(ano = ano, dados = dados_brasil))
}

resultados <- list()

for (arquivo in arquivos) {
  resultado <- processar_arquivo(arquivo)
  resultados[[resultado$ano]] <- resultado$dados
  
  if (nrow(resultado$dados) > 0) {
    nome_csv <- paste0("QS_Sustainability_Brazil_", resultado$ano, ".csv")
    write.csv(resultado$dados, nome_csv, row.names = FALSE, na = "")
    cat(sprintf("Arquivo salvo: %s com %d instituições brasileiras\n", 
                nome_csv, nrow(resultado$dados)))
  
    cat(sprintf("\nPreview Brasil %s:\n", resultado$ano))
    
    colunas_disponiveis <- intersect(
      c("Institution", "Location", "Rank_Display", "Overall_Score"),
      names(resultado$dados)
    )
    
    print(resultado$dados %>% select(all_of(colunas_disponiveis)) %>% head(10))
    cat("\n")
  } else {
    cat(sprintf("Nenhuma instituição brasileira encontrada em %s\n", arquivo))
  }
}

cat("\n=== Processamento individual por ano ===\n")

# 2024
dados_2024 <- read_excel("2024 QS Sustainability Rankings.xlsx")
names(dados_2024) <- gsub("\\s+|\\.", "_", names(dados_2024))
brasil_2024 <- dados_2024 %>% 
  filter(grepl("brazil|brasil", Location, ignore.case = TRUE))
write.csv(brasil_2024, "QS_Sustainability_Brazil_2024.csv", row.names = FALSE, na = "")

# 2025 - renomear coluna Overall para Overall_Score
dados_2025 <- read_excel("2025 QS Sustainability Rankings.xlsx")
names(dados_2025) <- gsub("\\s+|\\.", "_", names(dados_2025))
# Renomear se existir a coluna Overall
if ("Overall" %in% names(dados_2025)) {
  dados_2025 <- dados_2025 %>% rename(Overall_Score = Overall)
}
brasil_2025 <- dados_2025 %>% 
  filter(grepl("brazil|brasil", Location, ignore.case = TRUE))
write.csv(brasil_2025, "QS_Sustainability_Brazil_2025.csv", row.names = FALSE, na = "")

# 2026
dados_2026 <- read_excel("2026 QS Sustainability Rankings.xlsx")
names(dados_2026) <- gsub("\\s+|\\.", "_", names(dados_2026))
brasil_2026 <- dados_2026 %>% 
  filter(grepl("brazil|brasil", Location, ignore.case = TRUE))
write.csv(brasil_2026, "QS_Sustainability_Brazil_2026.csv", row.names = FALSE, na = "")