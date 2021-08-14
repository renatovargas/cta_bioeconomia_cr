# Naciones Unidas
# Comisión Económica para América Latina
# Banco Central de Costa Rica
# Estudio exploratorio para la elaboración de una 
# Cuenta Satélite de Bioeconomía para Costa Rica
# Supervisor CEPAL: Adrián Rodríguez
# Supervisora BCR: Irene Alvarado
# Consultor: Renato Vargas


# Procesamiento de datos básicos de oferta y utilización.

# Preámbulo

# Ubicación

rm(list = ls())
wd <- "D:/github/cta_bioeconomia_cr/codigo_datos/data"
setwd(wd)

# Librerías

library(readxl)
library(reshape2)

# ==========================
# Ingesta de datos de Oferta
# ==========================

oferta <- as.matrix(read_excel(
    "COU18.xlsx", 
    range = "'COU 2018'!C14:QB199",
    col_names = FALSE,
    col_types = "numeric",
    progress = TRUE
    ))

# Nombramos correlativamente según Excel
nfilas    <- c(sprintf("of%03d", seq(1,dim(oferta)[1]) ))
ncolumnas <- c(sprintf("oc%03d", seq(1,dim(oferta)[2]) ))

rownames(oferta) <- nfilas
colnames(oferta) <- ncolumnas

# Identificamos las columnas vacías y de subtotales y totales
# para omitir

of_omitir_columnas <- c(
    1,4,7,10,13,16,19,22,25,28,31,34,37,40,43,46,49,52,55,58,
    61,64,67,70,73,76,79,82,85,88,91,94,97,100,103,106,107,110,
    113,116,117,120,121,124,127,130,131,134,137,140,144,145,
    146,147,148,151,154,157,158,161,164,167,168,171,174,175,176,
    179,182,185,188,191,194,197,200,203,206,209,212,215,218,221,
    224,227,230,233,236,237,240,243,244,247,250,253,254,257,260,
    263,266,269,272,275,278,281,284,287,288,291,292,295,298,301,
    304,307,310,313,316,319,320,323,326,329,332,335,338,341,344,
    347,350,353,356,359,360,361,362,365,368,371,374,377,380,383,
    384,385,386,389,392,395,398,401,404,407,410,413,416,417,418,
    419,420,421,422,423,426,429,430,432,437,441,442
)

of_omitir_filas <- c(
    185
)

oferta <- oferta[-of_omitir_filas,-of_omitir_columnas]


# ===============================
# Ingesta de datos de Utilización
# ===============================

utilizacion <- as.matrix(read_excel(
    "COU18.xlsx", 
    range = "'COU 2018'!C214:QC399",
    col_names = FALSE,
    col_types = "numeric",
    progress = TRUE
))

# Nombramos correlativamente según Excel
nfilas    <- c(sprintf("uf%03d", seq(1,dim(utilizacion)[1]) ))
ncolumnas <- c(sprintf("uc%03d", seq(1,dim(utilizacion)[2]) ))

rownames(utilizacion) <- nfilas
colnames(utilizacion) <- ncolumnas

# Identificamos las columnas vacías y de subtotales y totales
# para omitir

ut_omitir_columnas <- c(
    1,4,7,10,13,16,19,22,25,28,31,34,37,40,43,46,49,52,55,58,
    61,64,67,70,73,76,79,82,85,88,91,94,97,100,103,106,107,110,
    113,116,117,120,121,124,127,130,131,134,137,140,144,145,
    146,147,148,151,154,157,158,161,164,167,168,171,174,175,176,
    179,182,185,188,191,194,197,200,203,206,209,212,215,218,221,
    224,227,230,233,236,237,240,243,244,247,250,253,254,257,260,
    263,266,269,272,275,278,281,284,287,288,291,292,295,298,301,
    304,307,310,313,316,319,320,323,326,329,332,335,338,341,344,
    347,350,353,356,359,360,361,362,365,368,371,374,377,380,383,
    384,385,386,389,392,395,398,401,404,407,410,413,416,417,418,
    419,420,421,422,423,426,429,430,434,437,438,442,443
)

ut_omitir_filas <- c(
    185
)

utilizacion <- utilizacion[-ut_omitir_filas,-ut_omitir_columnas]

# ==============================================
# Formato de base de datos Flat File para el COU
# ==============================================

db_oferta       <- melt(oferta)
db_utilizacion  <- melt(utilizacion)

db_oferta$"Of/Ut" <- "Oferta"
db_utilizacion$"Of/Ut" <- "Utilización"

db_COU <- rbind(
    db_oferta,
    db_utilizacion
)

colnames(db_COU) <- c("Filas", "Columnas", "Valores", "Of/Ut")
