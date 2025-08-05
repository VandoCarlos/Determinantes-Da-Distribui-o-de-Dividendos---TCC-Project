# ========================================================
# Script Completo: Modelos por Regime + Diagnósticos e Exportação Word
# ========================================================

# 0) Instalação e carregamento de pacotes
pacotes <- c("tidyverse", "readxl", "plm", "survival", "flextable", "officer", 
             "modelsummary", "AER", "sandwich", "lmtest")
instalados <- rownames(installed.packages())
if (length(setdiff(pacotes, instalados)) > 0) {
  install.packages(setdiff(pacotes, instalados), repos = "https://cloud.r-project.org/", dependencies = TRUE)
}

library(tidyverse)
library(readxl)
library(plm)
library(survival)
library(flextable)
library(officer)
library(modelsummary)
library(AER)
library(sandwich)
library(lmtest)

# 1) Leitura e pré-processamento dos dados
dados_brutos <- read_excel("Book(Planilha1).xlsx", skip = 3)
colnames(dados_brutos)[1:4] <- c("ticker", "empresa", "pais", "setor")

anos <- 2001:2024
n_anos <- length(anos)

extrair_long <- function(df, start_col, var_name) {
  df %>%
    select(1:4, start_col:(start_col + n_anos - 1)) %>%
    setNames(c("ticker", "empresa", "pais", "setor", as.character(anos))) %>%
    pivot_longer(cols = as.character(anos), names_to = "ano", values_to = var_name) %>%
    mutate(ano = as.integer(ano))
}

dividendos    <- extrair_long(dados_brutos, 5,   "dividendos")
roa           <- extrair_long(dados_brutos, 29,  "roa")
capex         <- extrair_long(dados_brutos, 53,  "capex")
ativos        <- extrair_long(dados_brutos, 77,  "ativos")
alavancagem   <- extrair_long(dados_brutos, 101, "alavancagem")
endividamento <- extrair_long(dados_brutos, 149, "endividamento")
valor_mercado <- extrair_long(dados_brutos, 173, "valor_mercado")

painel_raw <- dividendos %>%
  left_join(roa, by = c("ticker", "empresa", "pais", "setor", "ano")) %>%
  left_join(capex, by = c("ticker", "empresa", "pais", "setor", "ano")) %>%
  left_join(ativos, by = c("ticker", "empresa", "pais", "setor", "ano")) %>%
  left_join(alavancagem, by = c("ticker", "empresa", "pais", "setor", "ano")) %>%
  left_join(endividamento, by = c("ticker", "empresa", "pais", "setor", "ano")) %>%
  left_join(valor_mercado, by = c("ticker", "empresa", "pais", "setor", "ano"))

painel_clean <- painel_raw %>%
  filter(dividendos > 0) %>%
  mutate(
    dividendos = log(dividendos),
    ativos     = log(ativos),
    ano        = as.integer(ano),
    roa = roa * 100,
    capex = capex * 100,
    endividamento = endividamento * 100,
    valor_mercado = valor_mercado * 100,
    alavancagem = alavancagem * 100
  ) %>%
  drop_na(dividendos, roa, capex, ativos, alavancagem, endividamento, valor_mercado) %>%
  filter(across(c(dividendos, roa, capex, ativos, alavancagem, endividamento, valor_mercado, ano), ~ is.finite(.)))

# 2) Windsorização
windsorizar <- function(x, p1 = 0.01, p99 = 0.99) {
  quantis <- quantile(x, probs = c(p1, p99), na.rm = TRUE)
  x[x < quantis[1]] <- quantis[1]
  x[x > quantis[2]] <- quantis[2]
  return(x)
}

vars_para_windsorizar <- c("dividendos", "roa", "capex", "ativos", "alavancagem", "endividamento", "valor_mercado")
painel_clean[vars_para_windsorizar] <- lapply(painel_clean[vars_para_windsorizar], windsorizar)

painel_p <- pdata.frame(painel_clean, index = c("ticker", "ano"), drop.index = TRUE, row.names = TRUE)

anos_quebra <- c(2003, 2006, 2014, 2017, 2020)
limites <- c(2001, anos_quebra, 2024)
labels <- paste0(limites[-length(limites)], "-", limites[-1])
painel_clean$regime <- cut(painel_clean$ano, breaks = limites, labels = labels, include.lowest = TRUE)

# 3) Estimação dos modelos, coleta de resultados e diagnósticos

result <- list(pooled = list(), fixo = list(), aleat = list(), tobit = list())
hausman <- list(); waldF <- list(); bp_tobit <- list()
diag_pooled <- diag_fixo <- diag_aleat <- diag_tobit <- list()

# Funções para diagnóstico
diagnosticos_plm <- function(modelo, nome_modelo, regime) {
  s <- summary(modelo)
  dw <- tryCatch(dwtest(modelo), error = function(e) NA)
  data.frame(
    Regime         = regime,
    Modelo         = nome_modelo,
    R2             = round(s$r.squared["rsq"], 4),
    R2_ajust       = round(s$r.squared["adjrsq"], 4),
    DurbinWatson   = ifelse(is.na(dw), NA, round(dw$statistic, 4)),
    Pval_DW        = ifelse(is.na(dw), NA, round(dw$p.value, 4)),
    N_obs          = nobs(modelo)
  )
}

diagnosticos_tobit <- function(modelo, regime) {
  logLik <- as.numeric(logLik(modelo))
  AICval <- AIC(modelo)
  BICval <- BIC(modelo)
  data.frame(
    Regime         = regime,
    Modelo         = "Tobit (survreg)",
    LogLik         = round(logLik, 2),
    AIC            = round(AICval, 2),
    BIC            = round(BICval, 2)
  )
}

# Teste BP adaptado para Tobit
teste_bp_tobit <- function(modelo_tobit, dados, formula_aux) {
  residuos <- residuals(modelo_tobit, type = "response")
  lm_aux <- lm(residuos^2 ~ ., data = model.matrix(formula_aux, dados)[,-1])
  bptest(lm_aux)
}

extrai_plm <- function(sum_obj, regime) {
  df <- as.data.frame(sum_obj$coefficients)
  names(df)[1:4] <- c("Estimate", "StdError", "Stat", "P_value")
  df %>%
    mutate(
      Term    = rownames(df),
      Regime  = regime,
      P_value = round(P_value, 6),
      Pct = case_when(
        Term == "(Intercept)" ~ (exp(Estimate) - 1) * 100,
        Term == "ativos"     ~ Estimate,
        TRUE                 ~ (exp(Estimate) - 1) * 100
      ),
      Pct = as.numeric(format(Pct, scientific = FALSE, nsmall = 2, decimal.mark = "."))
    ) %>%
    select(Regime, Term, Estimate, Pct, StdError, Stat, P_value)
}

extrai_tobit <- function(sum_obj, regime) {
  coefs <- as.data.frame(sum_obj$table)
  coefs$Term <- rownames(coefs)
  names(coefs)[1:4] <- c("Estimate", "StdError", "Stat", "P_value")
  coefs %>%
    mutate(
      Regime  = regime,
      P_value = round(P_value, 6),
      Pct = case_when(
        Term == "(Intercept)" ~ (exp(Estimate) - 1) * 100,
        Term == "ativos"     ~ Estimate,
        TRUE                 ~ (exp(Estimate) - 1) * 100
      ),
      Pct = as.numeric(format(Pct, scientific = FALSE, nsmall = 2, decimal.mark = "."))
    ) %>%
    select(Regime, Term, Estimate, Pct, StdError, Stat, P_value)
}

for (reg in levels(painel_clean$regime)) {
  df <- filter(painel_clean, regime == reg)
  
  m_pool  <- plm(dividendos ~ roa + capex + ativos + alavancagem + 
                   endividamento + valor_mercado,
                 data = df, model = "pooling", index = c("ticker", "ano"))
  m_fixo  <- plm(dividendos ~ roa + capex + ativos + alavancagem + 
                   endividamento + valor_mercado,
                 data = df, model = "within",  index = c("ticker", "ano"))
  m_aleat <- plm(dividendos ~ roa + capex + ativos + alavancagem + 
                   endividamento + valor_mercado,
                 data = df, model = "random",  index = c("ticker", "ano"))
  m_tobit <- survreg(
    Surv(dividendos, dividendos > 0, type = "left") ~
      roa + capex + ativos + alavancagem + endividamento + valor_mercado,
    data = df, dist = "gaussian"
  )
  
  result$pooled[[reg]] <- summary(m_pool)
  result$fixo[[reg]]   <- summary(m_fixo)
  result$aleat[[reg]]  <- summary(m_aleat)
  result$tobit[[reg]]  <- summary(m_tobit)
  
  hausman[[reg]] <- tryCatch(phtest(m_fixo, m_aleat), error = function(e) NA)
  waldF[[reg]]   <- tryCatch(plmtest(m_pool, effect = "individual", type = "F"), error = function(e) NA)
  bp_tobit[[reg]] <- tryCatch(
    teste_bp_tobit(m_tobit, df, dividendos ~ roa + capex + ativos + alavancagem + endividamento + valor_mercado),
    error = function(e) NA
  )
  
  diag_pooled[[reg]] <- diagnosticos_plm(m_pool, "Pooled OLS", reg)
  diag_fixo[[reg]]   <- diagnosticos_plm(m_fixo, "Efeitos Fixos", reg)
  diag_aleat[[reg]]  <- diagnosticos_plm(m_aleat, "Efeitos Aleatórios", reg)
  diag_tobit[[reg]]  <- diagnosticos_tobit(m_tobit, reg)
}

# 4) Montagem das tabelas
pooled_df <- bind_rows(Map(extrai_plm, result$pooled, names(result$pooled)))
fixo_df <- bind_rows(Map(extrai_plm, result$fixo, names(result$fixo)))
aleat_df <- bind_rows(Map(extrai_plm, result$aleat, names(result$aleat)))
tobit_df <- bind_rows(Map(extrai_tobit, result$tobit, names(result$tobit)))

hausman_tab <- map_df(hausman, ~ if (inherits(.x, "htest")) tibble(Stat = round(.x$statistic, 6), P_value = round(.x$p.value, 6)) else tibble(Stat = NA, P_value = NA), .id = "Periodo")
wald_tab <- map_df(waldF, ~ if (inherits(.x, "htest")) tibble(Stat = round(.x$statistic, 6), P_value = round(.x$p.value, 6)) else tibble(Stat = NA, P_value = NA), .id = "Periodo")
bp_tobit_tab <- map_df(bp_tobit, ~ if (inherits(.x, "htest")) tibble(Stat = round(.x$statistic, 6), P_value = round(.x$p.value, 6)) else tibble(Stat = NA, P_value = NA), .id = "Periodo")

# Tabelas diagnósticas
diag_pooled_df <- bind_rows(diag_pooled)
diag_fixo_df   <- bind_rows(diag_fixo)
diag_aleat_df  <- bind_rows(diag_aleat)
diag_tobit_df  <- bind_rows(diag_tobit)

# 5) Criar flextable e função de exportação
make_ft <- function(df, caption) flextable(df) %>% autofit() %>% set_caption(caption)

ft_pooled <- make_ft(pooled_df, "Pooled OLS por Período (com % contínua)")
ft_fixo <- make_ft(fixo_df, "Efeitos Fixos por Período (com % contínua)")
ft_aleat <- make_ft(aleat_df, "Efeitos Aleatórios por Período (com % contínua)")
ft_tobit <- make_ft(tobit_df, "Tobit por Período (com % contínua)")
ft_hausman <- make_ft(hausman_tab, "Teste de Hausman por Período")
ft_wald <- make_ft(wald_tab, "Teste F para Efeitos Individuais por Período")
ft_bp_tobit <- make_ft(bp_tobit_tab, "Teste de Breusch-Pagan (BP) para Tobit por Período")

ft_diag_pooled <- make_ft(diag_pooled_df, "Diagnóstico Pooled OLS")
ft_diag_fixo   <- make_ft(diag_fixo_df,   "Diagnóstico Efeitos Fixos")
ft_diag_aleat  <- make_ft(diag_aleat_df,  "Diagnóstico Efeitos Aleatórios")
ft_diag_tobit  <- make_ft(diag_tobit_df,  "Diagnóstico Tobit")

# Função para exportar cada tabela em um Word separado
exportar_modelo <- function(ft_obj, nome_arquivo, titulo) {
  doc <- read_docx() %>%
    body_add_par(titulo, style = "heading 1") %>%
    body_add_flextable(ft_obj)
  print(doc, target = nome_arquivo)
}

# 6) Exportação das tabelas dos modelos
exportar_modelo(ft_pooled, "model_pooled.docx", "Resultados Pooled OLS por Período")
exportar_modelo(ft_fixo,   "model_fixo.docx",   "Resultados Efeitos Fixos por Período")
exportar_modelo(ft_aleat,  "model_aleat.docx",  "Resultados Efeitos Aleatórios por Período")
exportar_modelo(ft_tobit,  "model_tobit.docx",  "Resultados Tobit por Período")
exportar_modelo(ft_hausman,"model_hausman.docx","Teste de Hausman por Período")
exportar_modelo(ft_wald,   "model_wald.docx",   "Teste F para Efeitos Individuais por Período")
exportar_modelo(ft_bp_tobit,"model_bp_tobit.docx","Teste BP Tobit por Período")
exportar_modelo(ft_diag_pooled, "diag_pooled.docx", "Diagnóstico Pooled OLS")
exportar_modelo(ft_diag_fixo,   "diag_fixo.docx",   "Diagnóstico Efeitos Fixos")
exportar_modelo(ft_diag_aleat,  "diag_aleat.docx",  "Diagnóstico Efeitos Aleatórios")
exportar_modelo(ft_diag_tobit,  "diag_tobit.docx",  "Diagnóstico Tobit")

cat("Exportação concluída: arquivos Word gerados para cada modelo e diagnóstico.\n")
