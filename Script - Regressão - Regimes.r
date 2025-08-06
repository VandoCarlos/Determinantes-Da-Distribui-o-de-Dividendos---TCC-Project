#Instalação e pacotes - Pode Demorar de 10 a 15 minutos.
pacotes <- c("tidyverse","readxl","censReg","flextable","officer")
instalados <- rownames(installed.packages())
if(length(setdiff(pacotes, instalados))>0){
  install.packages(setdiff(pacotes, instalados),
                   repos="https://cloud.r-project.org/", dependencies=TRUE)
}
options(scipen=999)
library(tidyverse); library(readxl)
library(censReg)
library(flextable); library(officer)

#Leitura e formato longo - Realizando a Base de dados - Colunas Nomeadas Para data frame.
anos   <- 2001:2024; n_anos <- length(anos)
extrair_long <- function(df, start, name) {
  df %>%
    select(1:4, start:(start+n_anos-1)) %>%
    setNames(c("ticker","empresa","pais","setor", as.character(anos))) %>%
    pivot_longer(as.character(anos), names_to="ano", values_to=name) %>%
    mutate(ano = as.integer(ano))
}
dados_brutos <- read_excel("Book(Planilha1).xlsx", skip=3)
colnames(dados_brutos)[1:4] <- c("ticker","empresa","pais","setor")
div_raw <- extrair_long(dados_brutos, 5,   "dividendos_raw")
roa     <- extrair_long(dados_brutos, 29,  "roa")
capex   <- extrair_long(dados_brutos, 53,  "capex")
ativos  <- extrair_long(dados_brutos, 77,  "ativos")
alav    <- extrair_long(dados_brutos, 101, "alavancagem")
liq     <- extrair_long(dados_brutos, 173, "liquidez")
painel  <- div_raw %>% left_join(roa) %>% left_join(capex) %>%
  left_join(ativos) %>% left_join(alav) %>% left_join(liq)

# 2) Variáveis e transformações - Variáveis em Pontos Porcentuais e duas em Logaritimas sendo Dividendos e Ativos.
painel <- painel %>%
  mutate(
    dividendos     = pmax(dividendos_raw,0),
    log_dividendos = log1p(dividendos),
    ativos         = log(ativos),
    roa            = roa*100,
    capex          = capex*100,
    liquidez       = liquidez*100,
    alavancagem    = alavancagem*100
  ) %>%
  drop_na(dividendos,log_dividendos,roa,capex,ativos,alavancagem,liquidez) %>%
  filter(across(c(dividendos,log_dividendos,roa,capex,ativos,alavancagem,liquidez,ano), is.finite))

# 3) Windsorização (5%–95%) - Reduzir pesos das Observações atipicas (extremos)
windsorizar <- function(x,p1=0.05,p99=0.95){
  q <- quantile(x,probs=c(p1,p99),na.rm=TRUE)
  x[x<q[1]] <- q[1]; x[x>q[2]] <- q[2]; x
}
vars_w <- c("dividendos","log_dividendos","roa","capex","ativos","alavancagem","liquidez")
painel[vars_w] <- lapply(painel[vars_w], windsorizar)

# 4) Regimes Escolhidos Por Ciclos Economicos.
quebras <- c(2004,2008,2014,2020)
limites <- c(2001, quebras, 2024)
labels  <- paste0(limites[-length(limites)], "-", limites[-1])
painel  <- painel %>% mutate(regime = cut(ano, breaks=limites, labels=labels, include.lowest=TRUE))
regimes <- levels(painel$regime)

# 5) Função para AME manual do Tobit padrão - Average Marginal Effect
doc_std <- read_docx() %>% body_add_par("Tobit Padrão por Regime", style="heading 1")

# Função AME corrigida
calc_ame_censreg <- function(modelo, data) {
  est <- coef(modelo)
  coef_names <- names(est)
  idx_vars <- !(coef_names %in% c("(Intercept)", "logSigma"))
  coefs_vars <- est[idx_vars]
  
  xb <- as.numeric(model.matrix(modelo) %*% est[coef_names != "logSigma"])
  sigma <- exp(est["logSigma"])
  phi <- pnorm(xb / sigma)
  mean_phi <- mean(phi)
  ame <- coefs_vars * mean_phi
  ame_pct <- paste0(round((exp(ame)-1)*100, 2),"%") # Corrigido: sempre em percentual
  list(ame=round(ame,5), ame_pct=ame_pct)
}

# ...loop do modelo...
for(reg in regimes) {
  df_reg <- painel %>% filter(regime==reg)
  mod_std <- tryCatch(
    censReg(log_dividendos ~ roa + capex + ativos + alavancagem + liquidez,
            data=df_reg, left=0),
    error=function(e) NULL
  )
  if(!is.null(mod_std)){
    est <- coef(mod_std)
    se <- sqrt(diag(vcov(mod_std)))
    ame_res <- calc_ame_censreg(mod_std, df_reg)
    Terms <- names(est)[!(names(est) %in% c("(Intercept)", "logSigma"))]
    tab <- tibble(
      Term      = Terms,
      Estimate  = round(est[Terms],4),
      Pct       = case_when(
        Term == "ativos" ~ as.character(round(Estimate,4)),
        TRUE            ~ paste0(round((exp(Estimate)-1)*100,2),"%")
      ),
      AME      = ame_res$ame[Terms],
      Pct_AME  = case_when(
        Term == "ativos" ~ as.character(round(AME,4)),
        TRUE            ~ paste0(round((exp(AME)-1)*100,2),"%")
      ),
      StdError  = round(se[Terms],4),
      z_value   = round(Estimate/StdError,2),
      p_value   = round(2*pnorm(-abs(Estimate/StdError)),4)
    )
    doc_std <- doc_std %>%
      body_add_par(paste0("Regime ",reg), style="heading 2") %>%
      body_add_flextable(flextable(tab)%>%autofit())
  }
}

# 8) Salvar Word para poder realizar as analises.
print(doc_std, target="Tobit_Padrao_Regimes_AME.docx")

# Definição dos regimes Para um Tobit mais Robusto.
quebras <- c(2004,2008,2014,2020)
limites <- c(2001, quebras, 2024)
labels  <- paste0(limites[-length(limites)], "-", limites[-1])
painel  <- painel %>% mutate(regime = cut(ano, breaks=limites, labels=labels, include.lowest=TRUE))
regimes <- levels(painel$regime)
vars <- c("roa", "capex", "ativos", "alavancagem", "liquidez")

# Função de AME via predição numérica (diferenciação)
calc_ame_pred_crch <- function(modelo, data, vars, h=1e-5){
  ame <- numeric(length(vars))
  names(ame) <- vars
  for(i in seq_along(vars)){
    data_up <- data
    data_up[[vars[i]]] <- data_up[[vars[i]]] + h
    pred_up <- predict(modelo, newdata = data_up, type = "response")
    pred_orig <- predict(modelo, newdata = data, type = "response")
    ame[vars[i]] <- mean((pred_up - pred_orig) / h)
  }
  ame_pct <- sapply(vars, function(v){
    if(v == "ativos") {
      round(ame[v], 4)
    } else {
      paste0(round((exp(ame[v])-1)*100, 2),"%")
    }
  }, USE.NAMES = TRUE)
  list(AME=round(ame,5), Pct_AME=ame_pct)
}

# Documento Word
doc_rob <- read_docx() %>%
  body_add_par("Tobit Robusto (Student-t) - Coeficientes e AME", style="heading 1")

for(reg in regimes){
  df_reg <- painel %>% filter(regime==reg)
  mod_rob <- tryCatch(
    crch(log_dividendos ~ roa + capex + ativos + alavancagem + liquidez | 1,
         data=df_reg, dist="student", df=7, link.scale="log", left=0),
    error=function(e) NULL
  )
  if(!is.null(mod_rob)){
    X <- model.matrix(~ roa + capex + ativos + alavancagem + liquidez, data=df_reg)
    coefs_loc <- coef(mod_rob)[colnames(X)]
    se <- sqrt(diag(vcov(mod_rob)))[colnames(X)]
    ame_res <- calc_ame_pred_crch(mod_rob, df_reg, vars)
    tab <- tibble(
      Term     = vars,
      Estimate = round(coefs_loc[-1], 4),
      Pct      = case_when(
        Term == "ativos" ~ as.character(round(coefs_loc["ativos"],4)),
        TRUE             ~ paste0(round((exp(coefs_loc[Term])-1)*100,2),"%")
      ),
      AME      = ame_res$AME[vars],
      Pct_AME  = ame_res$Pct_AME[vars],
      StdError = round(se[-1], 4),
      z_value  = round(coefs_loc[-1] / se[-1], 2),
      p_value  = round(2 * pnorm(-abs(coefs_loc[-1] / se[-1])), 4)
    )
    doc_rob <- doc_rob %>%
      body_add_par(paste0("Regime ",reg), style="heading 2") %>%
      body_add_flextable(flextable(tab) %>% autofit())
  }
}

print(doc_rob, target="Tobit_Robusto_Regimes_AME.docx")
