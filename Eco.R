# Cargar las librerías necesarias
library(readxl)
library(agricolae)
library(ggplot2)
library(ggsignif)
library(openxlsx)

# Función para convertir los valores con formato en numéricos
convert_to_numeric <- function(x) {
  as.numeric(gsub("±.*", "", x))
}

# Leer los datos desde el archivo Excel
data <- read_excel("dataEco.xlsx")

# Variables a analizar
variables <- c("pH", "Dissolved_O2", "Conductivity", "Temperature_change", "Color", "Phosphate", "Nitrite", "Nitrate", "COD", "BOD5", "Total_solids", "Turbidity", "Pb", "As", "Fecal_coliforms")

# Preprocesamiento de los datos
for (var in variables) {
  data[[var]] <- convert_to_numeric(data[[var]])
}

# Crear una lista para almacenar los resultados de ANOVA y Tukey
resultados_anova_tukey <- list()

# Realizar el análisis de ANOVA y la prueba de Tukey para cada variable
for (var in variables) {
  # Realizar el análisis de ANOVA
  anova_result <- aov(as.formula(paste(var, "~ Station")), data = data)
  
  # Imprimir los resultados del ANOVA
  cat(paste("Variable:", var, "\n"))
  print(summary(anova_result))
  
  # Obtener el valor p del ANOVA
  p_value <- summary(anova_result)[[1]][["Pr(>F)"]][1]
  
  # Realizar la prueba de Tukey si el valor p es significativo (menor que 0.05)
  if (p_value < 0.05) {
    tukey_result <- HSD.test(anova_result, "Station")
    resultados_anova_tukey[[paste0(var, "_tukey")]] <- tukey_result$groups
  } else {
    resultados_anova_tukey[[paste0(var, "_tukey")]] <- "No se encontraron diferencias significativas entre los grupos."
  }
}

# Guardar los resultados de ANOVA y Tukey en un archivo Excel
write.xlsx(resultados_anova_tukey, file = "resultados_anova_tukey.xlsx")

# Generar gráficas de cajas y bigotes para las variables con diferencias significativas
for (var in variables) {
  # Verificar si la variable tiene diferencias significativas (utilizando la prueba de Tukey)
  if (!is.character(resultados_anova_tukey[[paste0(var, "_tukey")]])) {
    # Crear el gráfico de cajas y bigotes con ggplot2
    p <- ggplot(data, aes(x = as.factor(Station), y = data[[var]], fill = as.factor(Station))) +
      geom_boxplot() +
      scale_fill_gradient(low = "#DDDDDD", high = "#444444") +  # Utilizar una paleta de grises degradados
      theme_minimal() +
      labs(x = "Estación", y = var, title = paste("Gráfico de Cajas y Bigotes de", var),
           fill = "Estación") +
      theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1)) +
      geom_signif(comparisons = list(pairwise.t.test(data[[var]], data$Station)$p.adj < 0.05),
                  y_position = max(data[[var]]) + 1,
                  map_signif_level = TRUE,
                  tip_length = 0.01) +
      scale_fill_discrete(name = "Estación", labels = paste("Estación", 1:6))
    
    # Guardar el gráfico en una imagen con resolución de revista
    ggsave(paste0("grafico_", var, ".png"), plot = p, width = 10, height = 6, units = "in", dpi = 300)
    
    # Visualizar el gráfico en la ventana de gráficos
    print(p)
  }
}

# Calcular el ICA y agregarlo al dataframe
data$ICA <- rowSums(data[, c("pH", "Dissolved_O2", "Conductivity", "Temperature_change", "Color", "Phosphate", "Nitrite", "Nitrate", "COD", "BOD5", "Total_solids", "Turbidity", "Pb", "As", "Fecal_coliforms")])

# Generar gráfico de barras para el ICA (con paleta de grises degradados)
p_ica_bar <- ggplot(data, aes(x = as.factor(Station), y = ICA, fill = as.factor(Station))) +
  geom_bar(stat = "identity") +
  scale_fill_gradient(low = "#DDDDDD", high = "#444444") +  # Utilizar una paleta de grises degradados
  theme_minimal() +
  labs(x = "Estación", y = "ICA", title = "Índice de Calidad del Agua (ICA) por Estación",
       fill = "Estación") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1)) +
  scale_fill_discrete(name = "Estación", labels = paste("Estación", 1:6))

# Guardar el gráfico de barras del ICA en una imagen con resolución de revista
ggsave("grafico_ica_bar.png", plot = p_ica_bar, width = 10, height = 6, units = "in", dpi = 300)

# Visualizar el gráfico de barras del ICA en la ventana de gráficos
print(p_ica_bar)

# Crear un nuevo libro de trabajo de Excel para guardar los resultados del ANOVA, Tukey y el ICA
wb <- createWorkbook()

# Guardar los resultados de ANOVA y Tukey en una hoja en el libro de trabajo
addWorksheet(wb, sheetName = "Resultados_ANOVA_Tukey")
writeData(wb, sheet = "Resultados_ANOVA_Tukey", x = resultados_anova_tukey)

# Guardar los valores del ICA en otra hoja en el libro de trabajo
addWorksheet(wb, sheetName = "Resultados_ICA")
writeData(wb, sheet = "Resultados_ICA", x = data.frame(Station = data$Station, ICA = data$ICA))

# Guardar el libro de trabajo en un archivo Excel
saveWorkbook(wb, file = "resultados_analisis_ICA.xlsx")
