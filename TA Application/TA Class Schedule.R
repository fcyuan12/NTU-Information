######
# If use '###', there are steps or sections.
# If use '##', there is description. 
# If use '#', there is testing instruction or code.
######

######
# Part 0 
#
######

### Delete all things in Global Environment.

rm(list = ls())  

###  delete all things in Console
# 1. Move cursor in consle
# 2. Control + L 


######
# Part 1 
# Install packages
######

# install.packages("xlsx")
library(xlsx)

# install.packages("readxl")
library(readxl)


######
# Part 2
# Import data
######

### Method 1: slower than method 2, but instructions are intuitive
#TA_application <- read.xlsx('TA application.xlsx', sheetIndex = 1, startRow = 9, header = TRUE, encoding ='UTF-8')

### Method 2: faster than method 1 
## skip 8 lines because there are private information, such as ID, name, email, etc.
TA_application <- read_excel('TA志願填寫new.xlsx', skip = 8)
#View(TA_application)

######
# Part 3 
# Delete invalid classes
######

#colnames(TA_application)

### 3.1 Delete '已安排TA。'
TA_application <- TA_application[- which(TA_application$'志願\r\n\r\n實習課TA、教務處TA、GA皆各至少填寫5個志願，請填入1,2,3,4,5等等。' == '已安排TA。'), ]

### 3.2 Delete '已安排GA。'
TA_application <- TA_application[- which(TA_application$'志願\r\n\r\n實習課TA、教務處TA、GA皆各至少填寫5個志願，請填入1,2,3,4,5等等。' == '已安排GA。'), ]

### 3.3 Delete '授課\r\n對象' == '[M]' # This condition is not necessary. 
TA_application <- TA_application[- which(TA_application$'授課\r\n對象' == '[M]'), ]

### 3.4 Delete classes which are same time with required subjects (Microeconomics, Macroeconomics, and Econometrics)

#colnames(TA_application)
#TA_analysis <- data.frame(TA_application[,2], TA_application[,14], TA_application[,16])

# Applied numbers of classes which same time with required subjects:
# 一234910：077, 015, 031, 032
# 二910：None
# 三678910：025, 033, 034, 040, 056
# 四234：047
# 五：None

#class(TA_application$申請序號)
申請序號n <- as.numeric(TA_application$申請序號)
TA_application2 <- cbind(TA_application, 申請序號n)
InvalidApplyNum <- which(TA_application2$申請序號n %in% c(77, 15, 31, 32, 25, 33, 34, 40, 56, 47))

#InvalidApplyNum <- which(TA_application$申請序號n %in% c('077', '015', '031', '032', '025', '033', '034', '040', '056', '047'))

TA_application2 <- TA_application2[- InvalidApplyNum,]


######
# Part 4
# Categorize positions: 實習課TA (TA1), 教務處TA (TA2), GA
######

## Verify categories of TA. 
table(TA_application2$TA類別)

## Define TA1 is '實習課TA'.

#which(TA_application2$TA類別 == '實習課TA') # 1-7
#TA1 <- TA_application2[1:7, ]

TA1 <- TA_application2[TA_application2$TA類別 == '實習課TA', ]

TA1_focus <- TA1[, c('申請序號', '學期', '上課時間', '實習課')]

## Define TA2 is '教務處TA'.

#which(TA_application2$TA類別 == '教務處TA') # 8-25
#TA2 <- TA_application2[8:25, ]
TA2 <- TA_application2[TA_application2$TA類別 == '教務處TA', ]

TA2_focus <- TA2[, c('申請序號', '學期', '上課時間', '實習課')]

## Define GA is 'GA'.

#which(TA_application2$TA類別 == 'GA') # 26-52
#GA <- TA_application2[26:52, ]
GA <- TA_application2[TA_application2$TA類別 == 'GA', ]

GA_focus <- GA[, c('申請序號', '學期', '上課時間', '實習課')]


######
# Part 5
# Export data
######
# Reference:
# http://www.sthda.com/english/wiki/exporting-data-from-r

write.xlsx(TA1_focus, file = 'TA application.xlsx', sheetName='TA1', append=TRUE)
write.xlsx(TA2_focus, file = 'TA application.xlsx', sheetName='TA2', append=TRUE)
write.xlsx(GA_focus, file = 'TA application.xlsx', sheetName='GA', append=TRUE)


