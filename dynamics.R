# Скрипт для копирования данных о количестве сообщений по объектам "Ульяновская
# область" и "С.И. Морозов" из еженедельных отчётов ИС "Медиалогия". Входные 
# данные приходят в папку "Медиалогия" в ящике на mail.ru каждый 
# четверг (13 писем). Вложения в формате xlsx необходимо сохранять в
# субдиректорию mlg/ГГГГММДД в домашней папке, не изменяя названия файлов.

# Установка пакетов
#install.packages("xlsx", dependencies = TRUE)
#install.packages("readxl", dependencies = TRUE)
#install.packages("assertive", dependencies = TRUE)
#install.packages("dplyr")

# Загрузка пакетов
library(readxl)
library(xlsx)

# Объявление рабочей директории
setwd("~/Documents/Work/mlg/")

# Функция для проверки "пустых" переменных
IsInteger0 <- function(x)
{
  is.integer(x) && length(x) == 0L
}

# Функция для выгрузки рядов с данными о количестве сообщений по объектам:
CreateRows <- function(x)
{
  period <- as.character(x[10, 3])
  names(x) <- as.character(unlist(x[15, ]))
  x <- x[-c(1:15), ]
  ulo <- as.numeric(x[x$`Название объекта` %in% 
                         "Ульяновская область", ]$`Количество сообщений`)
  ulo <- as.numeric(ifelse(IsInteger0(ulo), "NA", ulo))
  x[x$`Название объекта` %in% "МОРОЗОВ Сергей Иванович (Ульяновская обл.)", ]
  morozov <- as.numeric(x[x$`Название объекта`
                          %in% "МОРОЗОВ Сергей Иванович (Ульяновская обл.)",
                          ]$`Количество сообщений`)
  morozov <- as.numeric(ifelse(IsInteger0(morozov), "NA", morozov))
  y <- cbind(period, ulo, morozov)
  return(y)
}

# Функция для переименования периода в формате ДД.ММ.ГГГГ:
MonthRenamer <- function(x){
  x$Период <- gsub("с ", "", x$Период)
  x$Период <- gsub(" по ", "-", x$Период)
  x$Период <- gsub(" января ", ".01.", x$Период)
  x$Период <- gsub(" февраля ", ".02.", x$Период)
  x$Период <- gsub(" марта ", ".03.", x$Период)
  x$Период <- gsub(" апреля ", ".04.", x$Период)
  x$Период <- gsub(" мая ", ".05.", x$Период)
  x$Период <- gsub(" июня ", ".06.", x$Период)
  x$Период <- gsub(" июля ", ".07.", x$Период)
  x$Период <- gsub(" августа ", ".08.", x$Период)
  x$Период <- gsub(" сентября ", ".09.", x$Период)
  x$Период <- gsub(" октября ", ".10.", x$Период)
  x$Период <- gsub(" ноября ", ".11.", x$Период)
  x$Период <- gsub(" декабря ", ".12.", x$Период)
  return(x)
}

# Записать в переменную список xlsx-файлов по проектам:
files.list <- gsub(".xlsx", "", list.files("20190103/"))
# Записать в переменную данные из файлы по Ульяновской области:
ulo <- list.files(pattern = files.list[13], recursive = TRUE)
ulo.list <- lapply(ulo, read_excel)
# Создать таблицу с данными по Ульяновской области:
data <- lapply(ulo.list, CreateRows)
table <- do.call(rbind.data.frame, data)
table[ , 2] <- as.numeric(levels(table[, 2]))
table[ , 3] <- as.numeric(levels(table[ , 3]))
colnames(table) <- c("Период", "Ульяновская обл.", "С.И. Морозов")
table <- MonthRenamer(table)

# Экспортировать в xlsx-файл:
write.xlsx(table, "динамика.xlsx")
