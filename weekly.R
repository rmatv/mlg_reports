# Скрипт для копирования данных об Ульяновской области из xlsx-выгрузок
# ИС "Медиалогия" в таблицу "недельный_срез.xlsx".
# Входные данные приходят в папку "Медиалогия" в ящике на mail.ru каждый 
# четверг (13 писем). Вложения в формате xlsx необходимо сохранять в
# субдиректорию mlg/ГГГГММДД в домашней папке, не изменяя названия файлов.

# ВАЖНО!
# Необходимо ВРУЧНУЮ указать корректную субдиректорию в строке 19.

# Установка пакетов
#install.packages("xlsx", dependencies = TRUE)
#install.packages("readxl", dependencies = TRUE)

# Загрузка пакетов
library(readxl)
library(xlsx)

# Объявление рабочей директории
setwd("~/Documents/Work/mlg/20190131/")

# Списки регионов и глав субъектов для ранжирования
regions <- read.csv('../regions.csv', header = FALSE, 
                    sep = ';', encoding = 'UTF-8')
persons <- read.csv('../persons.csv', header = FALSE, 
                    sep = ';', encoding = 'UTF-8')
regions <- as.character(levels(regions$V1))[regions$V1]
persons <- as.character(levels(persons$V1))[persons$V1]

## Функция для проверки "пустых" переменных
IsInteger0 <- function(x)
{
  is.integer(x) && length(x) == 0L
}

## Функция для выгрузки рядов данных по Ульяновской области
CreateRows <- function(x)
{
  # Забрать названия столбцов:
  names(x) <- as.character(unlist(x[14, ]))
  # Удалить строки без данных:
  x <- x[-c(1:14), ] 
  # Записать в переменную количество публикаций:
  publ <- as.numeric(x[x$`Название объекта` %in% 
                         "Ульяновская область", ]$`Количество сообщений`)
  publ <- as.numeric(ifelse(IsInteger0(publ), "NA", publ))
  # Записать в переменную охват аудитории:
  coverage <- as.numeric(x[x$`Название объекта` %in%
                             "Ульяновская область", ]
                         $`Охват (из открытых источников)`)
  coverage <- as.numeric(ifelse(IsInteger0(publ), "NA", coverage))
  # Записать в переменную количество негативных упоминаний:
  neg <- as.numeric(x[x$`Название объекта` %in%
                        "Ульяновская область", ]
                    $`Негативный характер упоминаний`)
  neg <- as.numeric(ifelse(IsInteger0(neg), "NA", neg))
  # Записать в переменную количество позитивных упоминаний:
  pos <- as.numeric(x[x$`Название объекта` %in%
                        "Ульяновская область", ]
                    $`Позитивный характер упоминаний`)
  pos <- as.numeric(ifelse(IsInteger0(pos), "NA", pos))
  # Записать в переменную нейтральных упоминаний:
  neu <- publ - (neg + pos)
  
  # Записать в переменную место Ульяновской области среди субъектов РФ:
  filt.reg <- x[x$`Название объекта` %in% regions, ]
  filt.reg <- as.character(filt.reg$`Название объекта`)
  rank.ulo <- which(filt.reg %in% 'Ульяновская область')
  rank.ulo <- as.numeric(ifelse(IsInteger0(rank.ulo), 'NA', rank.ulo))
  
  # Записать в переменную место СИ Морозова среди глав субъектов РФ:
  filt.per <- x[x$`Название объекта` %in% persons, ]
  filt.per <- as.character(filt.per$`Название объекта`)
  rank.mor <- which(filt.per %in% 'МОРОЗОВ Сергей Иванович (Ульяновская обл.)')
  rank.mor <- as.numeric(ifelse(IsInteger0(rank.mor), 'NA', rank.mor))
  
  # Записать все переменные в строку таблицы:
  y <- cbind(rank.ulo, rank.mor, publ, coverage, neg, neu, pos)
  return(y)
}

# Записать в переменную список xlsx-файлов по проектам в субдиректории:
projects <- list.files(pattern = ".xlsx")
projects <- projects[1:10]
projects.names <- gsub(".xlsx", "", list.files())
projects.names <- cbind(projects.names[1:10])

# Записать данные всех xlsx-файлов в список:
projects.list <- lapply(projects, read_excel)

# Вычленить данные по Ульяновской области:
data <- lapply(projects.list, CreateRows)

# Табулировать данные по Ульяновской области:
table <- do.call(rbind.data.frame, data)

# Задать названия рядов
rownames(table) <- c("Малый бизнес", "Культура", "Демография",
                     "Цифровая экономика", "Экология", "Образование",
                     "Экспорт", "Здравоохранение", "Жильё", "Дороги")
# Задать названия столбцов:
colnames(table) <- c("Ульяновская обл.", "С.И.Морозов", "Кол-во публикаций",
                     "Охват аудитории", "Негатив.", "Нейтральн.", "Позитив.")

# Экспортировать в xlsx-файл:
write.xlsx(table, "недельный_срез.xlsx")
