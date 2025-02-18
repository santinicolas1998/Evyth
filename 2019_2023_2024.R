rm(list = ls())
library(dplyr)
library(readr)
library(stringr)
library(tidyverse)
library(lubridate)
library(foreign)
library(janitor)
library(lubridate)
library(forcats)
library(openxlsx) 

setwd("C:/Users/20416597295/Documents/Santiago/evyth")

evyth_url <- "http://datos.yvera.gob.ar/dataset/b5819e9b-5edf-4aad-bd39-a81158a2b3f3/resource/645e5505-68ee-4cfa-90f9-fcc9a4a34a85/download/evyth_microdatos.csv"

evyth <- "evyth 2019-24.csv"

download.file(url = evyth_url, destfile = evyth)

evyth_2019_24 <- read.csv("evyth 2019-24.csv")

### TIPO DE TRANSPORTE A CABA
anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  visitantes <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(anio == !!anio) %>% 
    group_by(tipo_visitante) %>% 
    mutate(tipo_visitante = as_factor(tipo_visitante)) %>% 
    mutate(tipo_visitante = fct_recode(tipo_visitante,
                                       "Turista" = "1",
                                       "Excursionista" = "2")) %>%
    summarise(casos_ponderados = sum(pondera)) %>% 
    mutate(casos_ponderados = as.integer(casos_ponderados))
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("visitantes_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(visitantes, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}


### AGLOMERADO DE ORIGEN A CABA

anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  aglomerados <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    group_by(aglomerado_origen) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(casos_ponderados = as.integer(casos_ponderados), na.rm = TRUE)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("aglomerados_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(aglomerados, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}


### TIPO DE TRANSPORTE A CABA
anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  transporte_turistas <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    group_by(px09) %>% 
    mutate(px09 = as_factor(px09)) %>%
    mutate(px09 = fct_recode(px09,
                                   "Automovil" = "1",
                                   "Omnibus" = "2",
                                   "Tren" = "3",
                                   "Avión" = "4",
                                   "Embarcación" = "5",
                                   "Taxi o remis" = "6",
                                   "Otro" = "8",
                                   "Ns/ns" = "99")) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados, na.rm = TRUE)) * 100)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("alojamiento_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(transporte_turistas, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}



### TIPO DE ALOJAMENTO EN CABA

anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  alojamiento <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    group_by(px08_agrup) %>% 
    mutate(px08_agrup = as_factor(px08_agrup)) %>%
    mutate(px08_agrup = fct_recode(px08_agrup,
                                   "Segunda vivienda del Hogar"="1",
                                   "Vivienda de familiares y amigos" = "2",
                                   "Vivienda alquilada por temporada" = "3",
                                   "Camping" = "4",
                                   "Hotel h/ 3*" = "5",
                                   "Hotel 4 y 5*" = "6",
                                   "Resto" = "7")) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados, na.rm = TRUE)) * 100)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("alojamiento_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(alojamiento, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}


### TAMAÑO GRUPO A CABA

anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  tamanio_grupo <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    filter(px06_agrup != 99) %>% 
    group_by(px06_agrup) %>% 
    mutate(px06_agrup = as_factor(px06_agrup)) %>%
    mutate(px06_agrup = fct_recode(px06_agrup,
                               "1 o 2 personas" = "1",
                               "3 o 4 personas" = "2",
                               "5 o más personas" = "3")) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados, na.rm = TRUE)) * 100)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("tamaño_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(tamanio_grupo, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}






### MOTIVO DE VIAJE A CABA

anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  motivo <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    filter(px10_1 != 99) %>% 
    group_by(px10_1) %>% 
    mutate(px10_1 = as_factor(px10_1)) %>%
    mutate(px10_1 = fct_recode(px10_1,
                                   "Esparcimiento, ocio" = "1",
                                   "Visitas familiares y amigos" = "2",
                                   "Negocios" = "3",
                                   "Estudios" = "4",
                                   "Salud" = "5",
                                   "Religión" = "6",
                                   "Compras" = "7",
                                   "Otros" = "8",
                                   "Ns/ns" = "99")) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados, na.rm = TRUE)) * 100)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("motivo_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(motivo, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}






# ACTIVIDAD REALIZADA EM CABA
# Años a analizar
anios <- c(2019, 2023, 2024)

# Loop para crear un Excel por cada año
for (anio in anios) {
  
  # Crear una lista para almacenar los resultados
  resultados <- list()

  resultados[["Actividad_0"]] <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    filter(px17_1 != 9, px17_1 != "") %>% 
    mutate(px17_1 = as_factor(px17_1)) %>%
    mutate(px17_1 = fct_recode(px17_1,
                               "Si" = "1",
                               "No" = "2")) %>% 
    group_by(px17_1) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados)) * 100)
  
  # Loop de 1 a 13 para generar las demás hojas
  for (i in 1:13) {
    var_name <- paste0("px17_2_", i)  
    sheet_name <- paste0("Actividad_", i)  
    
    resultados[[sheet_name]] <- evyth_2019_24 %>%
      filter(region_destino == 1) %>%
      filter(aglomerado_origen != 32) %>%
      filter(tipo_visitante == 1) %>%
      filter(anio == !!anio) %>%
      filter(!is.na(.data[[var_name]])) %>%  
      filter(.data[[var_name]] != 9) %>%  
      mutate(!!var_name := as_factor(.data[[var_name]])) %>%
      mutate(!!var_name := fct_recode(.data[[var_name]],
                                      "Si" = "1",
                                      "No" = "2")) %>%
      group_by(.data[[var_name]]) %>%
      summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>%
      mutate(porcentaje = (casos_ponderados / sum(casos_ponderados)) * 100)
  }
  
  # Crear un archivo Excel con todas las hojas
  file_name <- paste0("actividades_turisticas_", anio, ".xlsx")
  wb <- createWorkbook()
  
  # Agregar cada dataset como una hoja en el Excel
  for (sheet in names(resultados)) {
    addWorksheet(wb, sheet)
    writeData(wb, sheet, resultados[[sheet]])
  }
  
  # Guardar el archivo Excel
  saveWorkbook(wb, file_name, overwrite = TRUE)
  print(paste("Archivo guardado como:", file_name))
}



unique(evyth_2019_24$px07_agrup)


anios <- c(2019, 2023, 2024)

# Loop para crear un CSV por cada año
for (anio in anios) {
  
  estadia <- evyth_2019_24 %>% 
    filter(region_destino == 1) %>% 
    filter(aglomerado_origen != 32) %>%
    filter(tipo_visitante == 1) %>% 
    filter(anio == !!anio) %>% 
    filter(px07_agrup != 99) %>% 
    group_by(px07_agrup) %>% 
    mutate(px07_agrup = as_factor(px07_agrup)) %>%
    mutate(px07_agrup = fct_recode(px07_agrup,
                                   "1 a 3 noches" = "1",
                                   "4 a 7 noches" = "2",
                                   "8 a 14 noches" = "3",
                                   "15 a 30 noches" = "4",
                                   "más de 30 noches" = "5",
                                   "Ns/ns" = "")) %>% 
    summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>% 
    mutate(porcentaje = (casos_ponderados / sum(casos_ponderados, na.rm = TRUE)) * 100)
  
  # Crear el nombre del archivo CSV
  file_name <- paste0("estadia_", anio, ".csv")
  
  # Guardar el archivo CSV
  write.csv(estadia, file_name, row.names = FALSE)
  
  # Verificación de archivo guardado
  print(paste("Archivo guardado como:", file_name))
}

write.csv(estadia, "estadia 2024.csv")




# CALIFICACIONES CABA
# Años a analizar
anios <- c(2019, 2023, 2024)

# Loop para crear un Excel por cada año
for (anio in anios) {
  
  # Crear una lista para almacenar los resultados
  resultados <- list()
  
  # Loop de 1 a 7 para generar las demás hojas
  for (i in 1:7) {
    var_name <- paste0("px18_", i)  
    sheet_name <- paste0("Calificacion_", i)  
    
    resultados[[sheet_name]] <- evyth_2019_24 %>%
      filter(region_destino == 1) %>%
      filter(aglomerado_origen != 32) %>%
      filter(tipo_visitante == 1) %>%
      filter(anio == !!anio) %>%
      filter(!is.na(.data[[var_name]])) %>%  
      filter(.data[[var_name]] != 99,.data[[var_name]] != "",.data[[var_name]] != 100) %>%  
      mutate(!!var_name := as_factor(.data[[var_name]])) %>%
      mutate(!!var_name := fct_recode(.data[[var_name]],
                                      "1" = "1",
                                      "2" = "2",
                                      "3" = "3",
                                      "4" = "4",
                                      "5" = "5",
                                      "6" = "6",
                                      "7" = "7",
                                      "8" = "8",
                                      "9" = "9",
                                      "10" = "10",
                                      "Ns./ Nr." = "99",
                                      "No usa" = "100")) %>%
      group_by(.data[[var_name]]) %>%
      summarise(casos_ponderados = sum(pondera, na.rm = TRUE)) %>%
      mutate(porcentaje = (casos_ponderados / sum(casos_ponderados)) * 100)
  }
  
  # Crear un archivo Excel con todas las hojas
  file_name <- paste0("calificaciones_", anio, ".xlsx")
  wb <- createWorkbook()
  
  # Agregar cada dataset como una hoja en el Excel
  for (sheet in names(resultados)) {
    addWorksheet(wb, sheet)
    writeData(wb, sheet, resultados[[sheet]])
  }
  
  # Guardar el archivo Excel
  saveWorkbook(wb, file_name, overwrite = TRUE)
  print(paste("Archivo guardado como:", file_name))
}



# De todos los turistas nacionales, cuáles vienen por motivos religiosos (hacer perfil)




# De todos los turistas nacionales, cuáles vienen por motivos religiosos (hacer perfil)
# correr todo para 2019,2023 y 2024



